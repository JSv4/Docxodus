using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// The GDI device context that <see cref="EmfImage"/> and <see cref="WmfImage"/> play their records
/// into, and the SVG it draws. Callers pass logical coordinates; every element keeps them and carries
/// the logical-to-device transform in effect when it was drawn, so a pen scales with the mapping as it
/// does in GDI. Every bound throws, so the caller keeps the whole-picture placeholder rather than
/// partial artwork.
/// </summary>
internal sealed class MetafileCanvas
{
    internal static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    internal const int MaxRecords = 200_000;
    /// <summary>Budget for a whole picture: one for each point read from a record, each point of a stroked
    /// line or path bracket, and each text character.</summary>
    internal const int MaxPoints = 250_000;
    /// <summary>Clip regions that may intersect one another before a RESTOREDC or a replacing clip: each
    /// becomes a <c>clipPath</c> that refers to the one before, which a browser resolves recursively.</summary>
    internal const int MaxClipDepth = 64;
    private const int MaxElements = 50_000, MaxSavedStates = 1024, MaxTextLength = 32_768;
    private const long MaxBitmapPixels = 16L * 1024 * 1024;
    /// <summary>Characters of SVG a picture may write, bitmaps included. The converter holds the picture
    /// several times over (elements, serialized SVG, the HTML that inlines it).</summary>
    private const long MaxOutputCharacters = 16L * 1024 * 1024;
    private const double Huge = 1e9;

    internal abstract record GdiObject;
    /// <summary>A pen; <c>Color</c> is "none" for PS_NULL. Cosmetic pens draw one device pixel.</summary>
    internal sealed record Pen(string Color, double Width, bool Cosmetic, string? Dashes, string Cap, string Join) : GdiObject;
    /// <summary>A brush; <c>Color</c> is "none" for BS_NULL. Pattern brushes are kept so that object
    /// slots stay aligned, but filling with one is unsupported.</summary>
    internal sealed record Brush(string Color, int Hatch = -1, bool Pattern = false) : GdiObject;
    internal sealed record Font(string Family, double Height, double Escapement, int Weight, bool Italic, bool Underline, bool Strike, int Charset = 0) : GdiObject;
    /// <summary>A palette, region or colour space: occupies a slot, draws nothing.</summary>
    internal sealed record OtherObject : GdiObject;

    internal static readonly Pen BlackPen = new("#000000", 1, true, null, "round", "round");
    internal static readonly Pen NullPen = new("none", 0, true, null, "round", "round");
    internal static readonly Brush WhiteBrush = new("#ffffff"), NullBrush = new("none");
    internal static readonly Font DefaultFont = new("", 16, 0, 400, false, false, false);

    /// <summary>x' = A·x + C·y + E, y' = B·x + D·y + F (SVG matrix order).</summary>
    internal readonly record struct Matrix(double A, double B, double C, double D, double E, double F)
    {
        internal static readonly Matrix Identity = new(1, 0, 0, 1, 0, 0);
        internal (double X, double Y) Apply(double x, double y) => (A * x + C * y + E, B * x + D * y + F);
        /// <summary>This transform followed by <paramref name="next"/>.</summary>
        internal Matrix Then(Matrix next) => new(A * next.A + B * next.C, A * next.B + B * next.D,
            C * next.A + D * next.C, C * next.B + D * next.D, E * next.A + F * next.C + next.E, E * next.B + F * next.D + next.F);
        internal double Determinant => A * D - B * C;
        internal Matrix Invert()
        {
            double det = Determinant;
            if (det == 0 || !double.IsFinite(det)) throw new NotSupportedException("Metafile transform is not invertible.");
            return new(D / det, -B / det, -C / det, A / det, (C * F - D * E) / det, (B * E - A * F) / det);
        }
    }

    private sealed class State
    {
        internal int MapMode = 1;
        internal double WinOrgX, WinOrgY, WinExtX = 1, WinExtY = 1, VpOrgX, VpOrgY, VpExtX = 1, VpExtY = 1;
        internal Matrix World = Matrix.Identity;
        internal Pen Pen = BlackPen;
        internal Brush Brush = WhiteBrush;
        internal Font Font = DefaultFont;
        internal string TextColor = "#000000", BkColor = "#ffffff";
        internal bool OpaqueBackground = true, EvenOdd = true, Clockwise;
        internal int TextAlign;
        internal string? Clip;
        internal int ClipDepth;
        internal double X, Y;
        internal State Clone() => (State)MemberwiseClone();
    }

    private readonly XElement root;
    private readonly string idPrefix;
    private readonly bool rejectCosmeticPens;
    private readonly double fixedDeviceWidth, fixedDeviceHeight, pixelsPerMmX, pixelsPerMmY;
    private readonly string defaultJoin;
    private State state = new();
    private readonly Stack<State> saved = new();
    private XElement? defs, clipGroup;
    private int elements, points, ids;
    private long bitmapPixels, output;
    private Path? bracket;
    private bool inBracket;
    private Matrix bracketMatrix;
    private Path? stroke; private XElement? strokeElement; private (Pen, Matrix, XElement)? strokeKey;
    private double minX = double.PositiveInfinity, minY = double.PositiveInfinity, maxX = double.NegativeInfinity, maxY = double.NegativeInfinity;

    /// <param name="root">Element the drawing is appended to.</param>
    /// <param name="seed">Input bytes; their checksum keeps generated ids distinct between pictures.</param>
    /// <param name="fixedDevice">WMF: the window always maps onto this device box, whatever the map mode.</param>
    /// <param name="pixelsPerMm">EMF: reference-device pixels per millimeter for the metric map modes.</param>
    /// <param name="rejectCosmeticPens">WMF: a visible device hairline has no resolution-independent width.</param>
    /// <param name="defaultJoin">Line join for pens whose style names none (PS_JOIN_ROUND is zero).</param>
    internal MetafileCanvas(XElement root, byte[] seed, (double Width, double Height)? fixedDevice = null,
        (double X, double Y)? pixelsPerMm = null, bool rejectCosmeticPens = false, string defaultJoin = "round")
    {
        this.root = root;
        idPrefix = "mf" + Crc32.Compute(seed).ToString("x8", CultureInfo.InvariantCulture) + "-";
        (fixedDeviceWidth, fixedDeviceHeight) = fixedDevice ?? (0, 0);
        (pixelsPerMmX, pixelsPerMmY) = pixelsPerMm ?? (1, 1);
        this.rejectCosmeticPens = rejectCosmeticPens;
        this.defaultJoin = defaultJoin;
    }

    /// <summary>Device-space bounds of everything drawn, or null when nothing was.</summary>
    internal (double MinX, double MinY, double MaxX, double MaxY)? Bounds
        => minX <= maxX ? (minX, minY, maxX, maxY) : null;

    internal Pen CurrentPen => state.Pen;
    internal Brush CurrentBrush => state.Brush;
    internal Font CurrentFont => state.Font;
    internal (double X, double Y) Position => (state.X, state.Y);

    // ---- Objects -------------------------------------------------------------------------------

    /// <summary>Builds a pen from a GDI pen style (PS_*), with dash, cap and join bits.</summary>
    internal Pen CreatePen(uint style, double width, string color, bool cosmetic, IReadOnlyList<double>? userDashes = null)
    {
        int kind = (int)(style & 0xf);
        if (kind == 5) return NullPen;
        string cap = ((style >> 8) & 0xf) switch { 0 => "round", 1 => "square", 2 => "butt",
            _ => throw new NotSupportedException("Unsupported metafile pen end cap.") };
        string join = ((style >> 12) & 0xf) switch { 0 => defaultJoin, 1 => "bevel", 2 => "miter",
            _ => throw new NotSupportedException("Unsupported metafile pen join.") };
        if (cosmetic) width = 1;
        if (!double.IsFinite(width) || width < 0 || width > Huge) throw new InvalidDataException("Invalid metafile pen width.");
        double unit = cosmetic ? 1 : Math.Max(width, 1);
        double[]? pattern = kind switch {
            0 or 6 => null, 1 => new double[] { 18, 6 }, 2 => new double[] { 3, 3 }, 3 => new double[] { 9, 6, 3, 6 },
            4 => new double[] { 9, 3, 3, 3, 3, 3 }, 8 => new double[] { 1, 1 },
            7 => userDashes is { Count: > 0 } ? userDashes.Select(d => d / unit).ToArray() : null,
            _ => throw new NotSupportedException("Unsupported metafile pen style.") };
        if (pattern != null && (pattern.Length > 64 || pattern.Any(d => !double.IsFinite(d) || d < 0) || pattern.All(d => d == 0)))
            throw new NotSupportedException("Unsupported metafile pen dash pattern.");
        string? dashes = pattern == null ? null : string.Join(" ", pattern.Select(d => N(d * unit)));
        return new Pen(color, width, cosmetic, dashes, cap, join);
    }

    /// <summary>Builds a brush from a LOGBRUSH style (BS_*) and hatch (HS_*).</summary>
    internal static Brush CreateBrush(uint style, string color, uint hatch) => style switch {
        0 => new Brush(color), 1 => NullBrush,
        2 when hatch <= 5 => new Brush(color, (int)hatch),
        2 => throw new NotSupportedException("Unsupported metafile hatch style."),
        3 or 4 or 5 or 6 => new Brush("none", Pattern: true),
        _ => throw new NotSupportedException("Unsupported metafile brush style.") };

    internal static Font CreateFont(string face, int height, int escapement, int weight, bool italic, bool underline, bool strike)
    {
        var family = new StringBuilder();
        foreach (char c in face)
        {
            if (c == '\0') break;
            if (char.IsLetterOrDigit(c) || c is ' ' or '-' or '_' or '.') family.Append(c);
        }
        return new Font(family.ToString().Trim(), Math.Abs((double)height), escapement / 10.0, weight, italic, underline, strike);
    }

    internal void Select(GdiObject value)
    {
        switch (value)
        {
            case Pen pen: state.Pen = pen; break;
            case Brush brush: state.Brush = brush; break;
            case Font font: state.Font = font; break;
            default: throw new NotSupportedException("Selecting a metafile region or palette object is not supported.");
        }
    }

    // ---- Device context ------------------------------------------------------------------------

    internal void SetMapMode(int mode)
    {
        if (mode is < 1 or > 8) throw new NotSupportedException("Unsupported metafile map mode.");
        state.MapMode = mode;
    }
    internal void SetWindowOrigin(double x, double y) { state.WinOrgX = x; state.WinOrgY = y; }
    internal void SetViewportOrigin(double x, double y) { state.VpOrgX = x; state.VpOrgY = y; }
    // GDI refuses a zero extent and keeps the previous one.
    internal void SetWindowExtent(double x, double y) { if (x != 0 && y != 0) { state.WinExtX = x; state.WinExtY = y; } }
    internal void SetViewportExtent(double x, double y) { if (x != 0 && y != 0) { state.VpExtX = x; state.VpExtY = y; } }
    internal (double X, double Y) WindowOrigin => (state.WinOrgX, state.WinOrgY);
    internal (double X, double Y) WindowExtent => (state.WinExtX, state.WinExtY);
    internal (double X, double Y) ViewportOrigin => (state.VpOrgX, state.VpOrgY);
    internal (double X, double Y) ViewportExtent => (state.VpExtX, state.VpExtY);

    internal void SetWorldTransform(Matrix value) { Finite(value); state.World = value; }
    /// <summary>MODIFYWORLDTRANSFORM: 1 identity, 2 left-multiply, 3 right-multiply, 4 set.</summary>
    internal void ModifyWorldTransform(Matrix value, uint mode)
    {
        Finite(value);
        state.World = mode switch {
            1 => Matrix.Identity, 2 => value.Then(state.World), 3 => state.World.Then(value), 4 => value,
            _ => throw new NotSupportedException("Unsupported world transform mode.") };
        Finite(state.World);
    }

    internal void SetTextColor(string color) => state.TextColor = color;
    internal void SetBackgroundColor(string color) => state.BkColor = color;
    internal void SetBackgroundMode(int mode) => state.OpaqueBackground = mode == 2;
    internal void SetTextAlign(int align) => state.TextAlign = align;
    internal void SetPolyFillMode(int mode) => state.EvenOdd = mode != 2;
    internal void SetArcDirection(int direction) => state.Clockwise = direction == 2;

    internal void Save()
    {
        if (saved.Count >= MaxSavedStates) throw new InvalidDataException("Metafile saves too many device contexts.");
        saved.Push(state.Clone());
    }
    /// <summary>RESTOREDC: a negative value pops that many states; a positive one names a depth.</summary>
    internal void Restore(int which)
    {
        int pops = which < 0 ? -which : saved.Count - which + 1;
        if (pops < 1 || pops > saved.Count) throw new InvalidDataException("Invalid metafile RESTOREDC.");
        for (int i = 1; i < pops; i++) saved.Pop();
        state = saved.Pop();
    }

    /// <summary>The logical-to-device transform now in effect: the world transform, then the page mapping.</summary>
    internal Matrix Transform
    {
        get
        {
            var result = state.World.Then(Page);
            Finite(result);
            return result;
        }
    }

    /// <summary>Window-to-viewport mapping for the current map mode.</summary>
    private Matrix Page
    {
        get
        {
            var s = state;
            double sx, sy;
            if (fixedDeviceWidth > 0) { sx = fixedDeviceWidth / s.WinExtX; sy = fixedDeviceHeight / s.WinExtY; }
            else
            {
                (sx, sy) = s.MapMode switch {
                    2 => (0.1 * pixelsPerMmX, -0.1 * pixelsPerMmY), 3 => (0.01 * pixelsPerMmX, -0.01 * pixelsPerMmY),
                    4 => (0.254 * pixelsPerMmX, -0.254 * pixelsPerMmY), 5 => (0.0254 * pixelsPerMmX, -0.0254 * pixelsPerMmY),
                    6 => (25.4 / 1440 * pixelsPerMmX, -25.4 / 1440 * pixelsPerMmY),
                    7 or 8 => (s.VpExtX / s.WinExtX, s.VpExtY / s.WinExtY), _ => (1.0, 1.0) };
                if (s.MapMode == 7)
                {
                    double scale = Math.Min(Math.Abs(sx), Math.Abs(sy));
                    sx = Math.Sign(sx) * scale; sy = Math.Sign(sy) * scale;
                }
            }
            double vx = fixedDeviceWidth > 0 ? 0 : s.VpOrgX, vy = fixedDeviceWidth > 0 ? 0 : s.VpOrgY;
            return new Matrix(sx, 0, 0, sy, vx - s.WinOrgX * sx, vy - s.WinOrgY * sy);
        }
    }

    // ---- Geometry ------------------------------------------------------------------------------

    internal void MoveTo(double x, double y)
    {
        state.X = x; state.Y = y;
        if (inBracket) bracket!.Open = false;
    }

    internal void LineTo(double x, double y)
    {
        Stroke(p => { p.Continue(state.X, state.Y); p.Op('L', x, y); });
        state.X = x; state.Y = y;
    }

    /// <summary>Kinds: polyline, polygon, polybezier, polylineto, polybezierto.</summary>
    internal enum Poly { Line, Gon, Bezier, LineTo, BezierTo }

    internal void Polygonal(Poly kind, ReadOnlySpan<(double X, double Y)> pts)
    {
        Count(pts.Length);
        if (pts.Length == 0) return;
        if (kind is Poly.Bezier or Poly.BezierTo && (pts.Length - (kind == Poly.Bezier ? 1 : 0)) % 3 != 0)
            throw new InvalidDataException("Metafile Bézier point count is not 1 + 3n.");
        var copy = pts.ToArray();
        void Build(Path p)
        {
            int i = 0;
            if (kind is Poly.LineTo or Poly.BezierTo) p.Continue(state.X, state.Y);
            else { p.Op('M', copy[0].X, copy[0].Y); i = 1; }
            if (kind is Poly.Bezier or Poly.BezierTo)
                for (; i + 2 < copy.Length; i += 3) p.Op('C', copy[i].X, copy[i].Y, copy[i + 1].X, copy[i + 1].Y, copy[i + 2].X, copy[i + 2].Y);
            else for (; i < copy.Length; i++) p.Op('L', copy[i].X, copy[i].Y);
            if (kind == Poly.Gon) p.Close();
        }
        if (kind == Poly.Gon) Closed(Build);
        else Stroke(Build);
        if (kind is Poly.LineTo or Poly.BezierTo) (state.X, state.Y) = copy[^1];
    }

    /// <summary>POLYPOLYLINE / POLYPOLYGON: <paramref name="counts"/> split <paramref name="pts"/>.</summary>
    internal void PolyPolygonal(bool closed, ReadOnlySpan<(double X, double Y)> pts, ReadOnlySpan<int> counts)
    {
        Count(pts.Length);
        var copy = pts.ToArray(); var sizes = counts.ToArray();
        void Build(Path p)
        {
            int at = 0;
            foreach (int n in sizes)
            {
                if (n == 0) continue;
                p.Op('M', copy[at].X, copy[at].Y);
                for (int i = 1; i < n; i++) p.Op('L', copy[at + i].X, copy[at + i].Y);
                if (closed) p.Close();
                at += n;
            }
        }
        if (closed) Closed(Build); else Stroke(Build);
    }

    /// <summary>POLYDRAW: PT_MOVETO 6, PT_LINETO 2, PT_BEZIERTO 4, PT_CLOSEFIGURE flag 1.</summary>
    internal void PolyDraw(ReadOnlySpan<(double X, double Y)> pts, ReadOnlySpan<byte> types)
    {
        Count(pts.Length);
        var copy = pts.ToArray(); var kinds = types.ToArray();
        double x = state.X, y = state.Y;
        void Build(Path p)
        {
            for (int i = 0; i < copy.Length; i++)
            {
                int type = kinds[i] & ~1;
                if (type == 6) p.Op('M', copy[i].X, copy[i].Y);
                else if (type == 2) { p.Continue(state.X, state.Y); p.Op('L', copy[i].X, copy[i].Y); }
                else if (type == 4 && i + 2 < copy.Length && (kinds[i + 1] & ~1) == 4 && (kinds[i + 2] & ~1) == 4)
                {
                    p.Continue(state.X, state.Y);
                    p.Op('C', copy[i].X, copy[i].Y, copy[i + 1].X, copy[i + 1].Y, copy[i + 2].X, copy[i + 2].Y);
                    i += 2;
                }
                else throw new InvalidDataException("Invalid metafile POLYDRAW point type.");
                if ((kinds[i] & 1) != 0) p.Close();
                (x, y) = copy[i];
            }
        }
        Stroke(Build);
        if (copy.Length > 0) (state.X, state.Y) = (x, y);
    }

    internal void Rectangle(double left, double top, double right, double bottom, double cornerWidth = 0, double cornerHeight = 0)
    {
        Order(ref left, ref right); Order(ref top, ref bottom);
        double rx = Math.Min(Math.Abs(cornerWidth) / 2, (right - left) / 2), ry = Math.Min(Math.Abs(cornerHeight) / 2, (bottom - top) / 2);
        if (inBracket)
        {
            if (rx > 0 && ry > 0)
            {
                Into(p => {
                    p.Op('M', left + rx, top); p.Op('L', right - rx, top); p.Arc(rx, ry, false, true, right, top + ry);
                    p.Op('L', right, bottom - ry); p.Arc(rx, ry, false, true, right - rx, bottom);
                    p.Op('L', left + rx, bottom); p.Arc(rx, ry, false, true, left, bottom - ry);
                    p.Op('L', left, top + ry); p.Arc(rx, ry, false, true, left + rx, top); p.Close();
                });
            }
            else Into(p => { p.Op('M', left, top); p.Op('L', right, top); p.Op('L', right, bottom); p.Op('L', left, bottom); p.Close(); });
            return;
        }
        if (right <= left || bottom <= top) return;
        var rect = new XElement(Svg + "rect", new XAttribute("x", N(left)), new XAttribute("y", N(top)),
            new XAttribute("width", N(right - left)), new XAttribute("height", N(bottom - top)));
        if (rx > 0 && ry > 0) rect.Add(new XAttribute("rx", N(rx)), new XAttribute("ry", N(ry)));
        Shape(rect, true, true, Corners(left, top, right, bottom));
    }

    internal void Ellipse(double left, double top, double right, double bottom)
    {
        Order(ref left, ref right); Order(ref top, ref bottom);
        double rx = (right - left) / 2, ry = (bottom - top) / 2, cx = left + rx, cy = top + ry;
        if (inBracket)
        {
            Into(p => { p.Op('M', right, cy); p.Arc(rx, ry, false, true, left, cy); p.Arc(rx, ry, false, true, right, cy); p.Close(); });
            return;
        }
        if (rx <= 0 || ry <= 0) return;
        Shape(new XElement(Svg + "ellipse", new XAttribute("cx", N(cx)), new XAttribute("cy", N(cy)),
            new XAttribute("rx", N(rx)), new XAttribute("ry", N(ry))), true, true,
            Corners(left, top, right, bottom));
    }

    internal enum ArcKind { Arc, ArcTo, Chord, Pie }

    /// <summary>ARC, ARCTO, CHORD and PIE: an elliptical arc in a box, from and to two radial points.</summary>
    internal void Arc(ArcKind kind, double left, double top, double right, double bottom, double xStart, double yStart, double xEnd, double yEnd)
    {
        Order(ref left, ref right); Order(ref top, ref bottom);
        double rx = (right - left) / 2, ry = (bottom - top) / 2, cx = left + rx, cy = top + ry;
        if (rx <= 0 || ry <= 0) return;
        double a1 = Math.Atan2((yStart - cy) / ry, (xStart - cx) / rx), a2 = Math.Atan2((yEnd - cy) / ry, (xEnd - cx) / rx);
        // AD_COUNTERCLOCKWISE is anticlockwise as seen on the device. With y down and an orientation-
        // preserving transform that is a decreasing parametric angle.
        bool positive = state.Clockwise ^ (Transform.Determinant < 0);
        double sweep = positive ? a2 - a1 : a1 - a2;
        sweep %= 2 * Math.PI;
        if (sweep <= 0) sweep += 2 * Math.PI;
        Sweep(kind, cx, cy, rx, ry, a1, positive ? sweep : -sweep);
    }

    /// <summary>ANGLEARC: angles in degrees, anticlockwise on the device.</summary>
    internal void AngleArc(double cx, double cy, double radius, double startDegrees, double sweepDegrees)
    {
        if (!double.IsFinite(startDegrees) || !double.IsFinite(sweepDegrees) || radius < 0)
            throw new InvalidDataException("Invalid metafile ANGLEARC.");
        double flip = Transform.Determinant < 0 ? -1 : 1;
        double sweep = -flip * Math.Clamp(sweepDegrees, -360, 360) * Math.PI / 180;
        Sweep(ArcKind.ArcTo, cx, cy, radius, radius, -flip * startDegrees * Math.PI / 180, sweep);
    }

    private void Sweep(ArcKind kind, double cx, double cy, double rx, double ry, double start, double sweep)
    {
        double sx = cx + rx * Math.Cos(start), sy = cy + ry * Math.Sin(start);
        double end = start + sweep, ex = cx + rx * Math.Cos(end), ey = cy + ry * Math.Sin(end);
        void Build(Path p)
        {
            if (kind == ArcKind.Pie) { p.Op('M', cx, cy); p.Op('L', sx, sy); }
            else if (kind == ArcKind.ArcTo) { p.Continue(state.X, state.Y); p.Op('L', sx, sy); }
            else p.Op('M', sx, sy);
            bool positive = sweep > 0;
            if (Math.Abs(sweep) >= 2 * Math.PI - 1e-9)
            {
                double mid = start + sweep / 2;
                p.Arc(rx, ry, false, positive, cx + rx * Math.Cos(mid), cy + ry * Math.Sin(mid));
                p.Arc(rx, ry, false, positive, ex, ey);
            }
            else if (sweep != 0) p.Arc(rx, ry, Math.Abs(sweep) > Math.PI, positive, ex, ey);
            if (kind is ArcKind.Chord or ArcKind.Pie) p.Close();
        }
        if (kind is ArcKind.Chord or ArcKind.Pie) Closed(Build, Corners(cx - rx, cy - ry, cx + rx, cy + ry));
        else Stroke(Build, Corners(cx - rx, cy - ry, cx + rx, cy + ry));
        if (kind == ArcKind.ArcTo) (state.X, state.Y) = (ex, ey);
    }

    // ---- Path brackets -------------------------------------------------------------------------

    internal void BeginPath()
    {
        bracket = new Path(); inBracket = true; bracketMatrix = Transform;
    }
    internal void EndPath() => inBracket = false;
    internal void AbortPath() { bracket = null; inBracket = false; }
    internal void CloseFigure()
    {
        if (!inBracket) return;
        if (bracket!.Open) { bracket.Close(); (state.X, state.Y) = bracket.Start; }
    }

    /// <summary>FILLPATH, STROKEPATH and STROKEANDFILLPATH.</summary>
    internal void DrawPath(bool fill, bool outline)
    {
        if (bracket is not { } path || inBracket) { AbortPath(); return; }
        bracket = null;
        if (path.Empty) return;
        if (fill) path.CloseOpenFigure();
        var element = new XElement(Svg + "path", new XAttribute("d", path.ToString()));
        Shape(element, fill, outline, path.Points, bracketMatrix, path: true);
    }

    /// <summary>SELECTCLIPPATH: RGN_AND (1) intersects, RGN_COPY (5) replaces.</summary>
    internal void SelectClipPath(int mode)
    {
        if (bracket is not { } path || inBracket) { AbortPath(); return; }
        bracket = null;
        if (mode is not (1 or 5)) throw new NotSupportedException("Unsupported metafile clip-path combination.");
        path.CloseOpenFigure();
        var shape = new XElement(Svg + "path", new XAttribute("d", path.ToString()),
            new XAttribute("clip-rule", state.EvenOdd ? "evenodd" : "nonzero"));
        SetTransformAttribute(shape, bracketMatrix);
        AddClip(shape, mode == 1);
    }

    // ---- Clipping ------------------------------------------------------------------------------

    internal void IntersectClipRect(double left, double top, double right, double bottom)
        => AddClip(new XElement(Svg + "path", new XAttribute("d", DeviceRect(left, top, right, bottom))), true);

    internal void ExcludeClipRect(double left, double top, double right, double bottom)
        => AddClip(new XElement(Svg + "path", new XAttribute("clip-rule", "evenodd"),
            new XAttribute("d", $"M{N(-Huge)} {N(-Huge)}H{N(Huge)}V{N(Huge)}H{N(-Huge)}Z" + DeviceRect(left, top, right, bottom))), true);

    internal void ResetClip() { state.Clip = null; state.ClipDepth = 0; }

    /// <summary>A region of device-space rectangles: RGN_AND (1) intersects, RGN_COPY (5) replaces.</summary>
    internal void ClipToDeviceRects(ReadOnlySpan<(double L, double T, double R, double B)> rects, int mode)
    {
        if (mode is not (1 or 5)) throw new NotSupportedException("Unsupported metafile clip-region combination.");
        Count(rects.Length * 4);
        var d = new StringBuilder();
        foreach (var r in rects) d.Append('M').Append(N(r.L)).Append(' ').Append(N(r.T)).Append('H').Append(N(r.R))
            .Append('V').Append(N(r.B)).Append('H').Append(N(r.L)).Append('Z');
        AddClip(new XElement(Svg + "path", new XAttribute("d", d.Length == 0 ? "M0 0Z" : d.ToString())), mode == 1);
    }

    internal bool HasClip => state.Clip != null;

    private string DeviceRect(double left, double top, double right, double bottom)
    {
        var m = Transform; Order(ref left, ref right); Order(ref top, ref bottom);
        var d = new StringBuilder();
        foreach (var (x, y) in new[] { (left, top), (right, top), (right, bottom), (left, bottom) })
        {
            var (dx, dy) = m.Apply(x, y);
            d.Append(d.Length == 0 ? 'M' : 'L').Append(N(dx)).Append(' ').Append(N(dy));
        }
        return d.Append('Z').ToString();
    }

    private void AddClip(XElement shape, bool intersect)
    {
        string id = NextId();
        var clip = new XElement(Svg + "clipPath", new XAttribute("id", id), new XAttribute("clipPathUnits", "userSpaceOnUse"), shape);
        int depth = 1;
        if (intersect && state.Clip != null)
        {
            if ((depth = state.ClipDepth + 1) > MaxClipDepth) throw new NotSupportedException("Metafile clip regions nest too deeply.");
            clip.Add(new XAttribute("clip-path", $"url(#{state.Clip})"));
        }
        Defs().Add(clip);
        Count(clip);
        state.Clip = id;
        state.ClipDepth = depth;
    }

    // ---- Text and bitmaps ----------------------------------------------------------------------

    /// <summary>ANSI text in the selected font's character set. Western and default-charset fonts read as
    /// Windows-1252; a byte above ASCII in any other set would be mis-decoded.</summary>
    internal string Ansi(ReadOnlySpan<byte> text)
    {
        if (state.Font.Charset is not (0 or 1) && text.IndexOfAnyInRange((byte)0x80, (byte)0xff) >= 0)
            throw new NotSupportedException("Metafile text outside the Western character set is not supported.");
        return EmfImage.Windows1252(text);
    }

    /// <summary>TEXTOUT/EXTTEXTOUT at a logical reference point. <paramref name="advances"/> are the
    /// optional logical character advances. <paramref name="opaque"/> is the ETO_OPAQUE rectangle.</summary>
    internal void Text(double x, double y, string text, IReadOnlyList<double>? advances, (double L, double T, double R, double B)? opaque)
    {
        if (inBracket) throw new NotSupportedException("Text in a metafile path bracket is not supported.");
        if (opaque is { } box)
        {
            var saveBrush = state.Brush; var savePen = state.Pen;
            state.Brush = new Brush(state.BkColor); state.Pen = NullPen;
            try { Rectangle(box.L, box.T, box.R, box.B); }
            finally { state.Brush = saveBrush; state.Pen = savePen; }
        }
        var clean = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) { clean.Append(c).Append(text[++i]); continue; }
            clean.Append(c == '\t' || (c >= 0x20 && XmlConvert.IsXmlChar(c)) ? c : ' ');
        }
        if (clean.Length > MaxTextLength) throw new NotSupportedException("Metafile text run exceeds the supported length.");
        bool updateCp = (state.TextAlign & 1) != 0;
        // Without a spacing array the run advances by the width of its glyphs, which only a font knows.
        if (updateCp && advances == null && clean.Length > 0)
            throw new NotSupportedException("Metafile text that moves the current position by its own width is not supported.");
        if (updateCp) (x, y) = (state.X, state.Y);
        double total = advances?.Sum() ?? 0;
        if (updateCp) state.X = x + total;
        if (clean.ToString().Trim().Length == 0) return;
        // A symbol font draws other glyphs for these characters (Symbol's "a" is an alpha), and private-use
        // characters have no glyph outside the font they were written for.
        if (state.Font.Charset == 2 || text.AsSpan().IndexOfAnyInRange('\uE000', '\uF8FF') >= 0)
            throw new NotSupportedException("Metafile symbol-font text is not supported.");
        Count(clean.Length);

        var font = state.Font;
        double size = font.Height > 0 ? font.Height : 12;
        var s = state;
        var m = Transform;
        var (ax, ay) = m.Apply(x, y);
        // GDI keeps text upright under page-space flips; the world transform still rotates and scales it.
        var page = Page;
        var linear = new Matrix(s.World.A, s.World.B, s.World.C, s.World.D, 0, 0)
            .Then(new Matrix(Math.Abs(page.A), 0, 0, Math.Abs(page.D), 0, 0));
        double angle = font.Escapement * Math.PI / 180;
        var frame = new Matrix(Math.Cos(angle), -Math.Sin(angle), Math.Sin(angle), Math.Cos(angle), 0, 0)
            .Then(linear).Then(new Matrix(1, 0, 0, 1, ax, ay));
        var element = new XElement(Svg + "text", new XAttribute("x", "0"), new XAttribute("y", "0"));
        int horizontal = s.TextAlign & 6, vertical = s.TextAlign & 24;
        if (advances != null && advances.Count > 0)
        {
            double start = horizontal == 2 ? -total : horizontal == 6 ? -total / 2 : 0, at = start;
            var xs = new StringBuilder();
            // Advances are per UTF-16 code unit; SVG positions are per character.
            for (int i = 0; i < clean.Length && i < advances.Count; i++)
            {
                if (!char.IsLowSurrogate(clean[i])) xs.Append(xs.Length == 0 ? "" : " ").Append(N(at));
                at += advances[i];
            }
            element.SetAttributeValue("x", xs.ToString());
        }
        else if (horizontal != 0) element.Add(new XAttribute("text-anchor", horizontal == 2 ? "end" : "middle"));
        if (vertical != 24) element.Add(new XAttribute("dominant-baseline", vertical == 8 ? "text-after-edge" : "text-before-edge"));
        element.Add(new XAttribute("font-size", N(size)),
            new XAttribute("font-family", font.Family.Length == 0 ? "sans-serif" : $"'{font.Family}', sans-serif"),
            new XAttribute("fill", s.TextColor), new XAttribute("style", "white-space:pre"));
        if (font.Weight is > 0 and not 400) element.Add(new XAttribute("font-weight", Math.Clamp(font.Weight / 100 * 100, 100, 900)));
        if (font.Italic) element.Add(new XAttribute("font-style", "italic"));
        if (font.Underline || font.Strike)
            element.Add(new XAttribute("text-decoration", font.Underline && font.Strike ? "underline line-through" : font.Underline ? "underline" : "line-through"));
        element.Add(clean.ToString());
        SetTransformAttribute(element, frame);
        double width = advances != null ? Math.Abs(total) : size * clean.Length;
        foreach (var (px, py) in new[] { (-width, -size), (width, size), (-width, size), (width, -size) })
            Track(frame, px, py, 0);
        Add(element);
    }

    /// <summary>Draws an already-encoded image into a logical destination rectangle (which may be flipped).</summary>
    internal void Image(double x, double y, double width, double height, string href, long pixels)
    {
        if (inBracket) throw new NotSupportedException("Bitmaps in a metafile path bracket are not supported.");
        if ((bitmapPixels += pixels) > MaxBitmapPixels) throw new NotSupportedException("Metafile bitmaps exceed the pixel limit.");
        if (width == 0 || height == 0) return;
        var m = new Matrix(width, 0, 0, height, x, y).Then(Transform);
        var image = new XElement(Svg + "image", new XAttribute("width", "1"), new XAttribute("height", "1"),
            new XAttribute("preserveAspectRatio", "none"), new XAttribute("href", href));
        SetTransformAttribute(image, m, always: true);
        foreach (var (px, py) in new[] { (0.0, 0.0), (1.0, 1.0), (0.0, 1.0), (1.0, 0.0) }) Track(m, px, py, 0);
        Add(image);
    }

    /// <summary>Accounts for bitmap pixels that are decoded but not drawn.</summary>
    internal void ReserveBitmap(long pixels)
    {
        if (bitmapPixels + pixels > MaxBitmapPixels) throw new NotSupportedException("Metafile bitmaps exceed the pixel limit.");
    }

    /// <summary>PATBLT-style fill of a logical rectangle with the current brush or a fixed colour.</summary>
    internal void FillRectangle(double left, double top, double right, double bottom, string? color)
    {
        var saveBrush = state.Brush; var savePen = state.Pen;
        if (color != null) state.Brush = new Brush(color);
        state.Pen = NullPen;
        try { Rectangle(left, top, right, bottom); }
        finally { state.Brush = saveBrush; state.Pen = savePen; }
    }

    /// <summary>Completes pending output; call once after the final record.</summary>
    internal void Finish() => Flush();

    // ---- Emission ------------------------------------------------------------------------------

    private void Into(Action<Path> build)
    {
        if (Transform != bracketMatrix) throw new NotSupportedException("A metafile path bracket changes its transform.");
        int before = bracket!.Points.Count;
        build(bracket);
        Count(bracket.Points.Count - before);
    }

    /// <summary>An open figure: stroked, never filled. Consecutive figures with the same pen share one path.</summary>
    private void Stroke(Action<Path> build, (double, double)[]? extra = null)
    {
        if (inBracket) { Into(build); return; }
        var pen = state.Pen;
        if (pen.Color == "none") { var scratch = new Path(); build(scratch); return; }
        var m = Transform;
        var container = Container();
        var key = (pen, m, container);
        if (stroke == null || strokeKey != key || container.LastNode != strokeElement)
        {
            Flush();
            stroke = new Path();
            strokeElement = new XElement(Svg + "path");
            strokeKey = key;
            Shape(strokeElement, false, true, Array.Empty<(double, double)>(), m, path: true, flush: false);
        }
        int before = stroke.Points.Count;
        // A figure that ends at the current position continues; anything else starts a new one.
        stroke.Open = stroke.Open && before > 0 && stroke.Points[^1] == (state.X, state.Y);
        build(stroke);
        Count(stroke.Points.Count - before);
        double pad = PenPad(m);
        for (int i = before; i < stroke.Points.Count; i++) Track(m, stroke.Points[i].X, stroke.Points[i].Y, pad);
        if (extra != null) foreach (var (x, y) in extra) Track(m, x, y, pad);
    }

    private void Closed(Action<Path> build, (double, double)[]? extra = null)
    {
        if (inBracket) { Into(build); return; }
        var path = new Path();
        build(path);
        if (path.Empty) return;
        var points = extra == null ? path.Points : path.Points.Concat(extra).ToList();
        Shape(new XElement(Svg + "path", new XAttribute("d", path.ToString())), true, true, points, path: true);
    }

    private void Shape(XElement element, bool fill, bool outline, IReadOnlyList<(double X, double Y)> logical,
        Matrix? transform = null, bool path = false, bool flush = true)
    {
        if (flush) Flush();
        var m = transform ?? Transform;
        var s = state;
        element.Add(new XAttribute("fill", fill ? Fill(s.Brush, m) : "none"));
        if (path && fill) element.Add(new XAttribute("fill-rule", s.EvenOdd ? "evenodd" : "nonzero"));
        var pen = s.Pen;
        if (outline && pen.Color != "none")
        {
            if (pen.Cosmetic && rejectCosmeticPens)
                throw new NotSupportedException("Visible device-pixel metafile pens have no resolution-independent width.");
            element.Add(new XAttribute("stroke", pen.Color), new XAttribute("stroke-width", N(pen.Width)),
                new XAttribute("stroke-linejoin", pen.Join));
            if (!pen.Cosmetic && pen.Cap != "butt") element.Add(new XAttribute("stroke-linecap", pen.Cap));
            if (pen.Dashes != null) element.Add(new XAttribute("stroke-dasharray", pen.Dashes));
            if (pen.Cosmetic) element.Add(new XAttribute("vector-effect", "non-scaling-stroke"));
        }
        else element.Add(new XAttribute("stroke", "none"));
        SetTransformAttribute(element, m);
        double pad = outline ? PenPad(m) : 0;
        foreach (var (x, y) in logical) Track(m, x, y, pad);
        Add(element);
    }

    private string Fill(Brush brush, Matrix m)
    {
        if (brush.Pattern) throw new NotSupportedException("Metafile pattern brushes are not supported.");
        if (brush.Hatch < 0 || brush.Color == "none") return brush.Color;
        // Hatches are eight device pixels, anchored at the device origin.
        string id = NextId();
        var pattern = new XElement(Svg + "pattern", new XAttribute("id", id), new XAttribute("patternUnits", "userSpaceOnUse"),
            new XAttribute("width", "8"), new XAttribute("height", "8"));
        SetTransformAttribute(pattern, m.Invert(), "patternTransform");
        if (state.OpaqueBackground)
            pattern.Add(new XElement(Svg + "rect", new XAttribute("width", "8"), new XAttribute("height", "8"), new XAttribute("fill", state.BkColor)));
        string d = brush.Hatch switch {
            0 => "M0 4.5H8", 1 => "M4.5 0V8", 2 => "M0 0L8 8", 3 => "M0 8L8 0", 4 => "M0 4.5H8M4.5 0V8", _ => "M0 0L8 8M0 8L8 0" };
        pattern.Add(new XElement(Svg + "path", new XAttribute("d", d), new XAttribute("stroke", brush.Color), new XAttribute("stroke-width", "1")));
        Defs().Add(pattern);
        Count(pattern);
        return $"url(#{id})";
    }

    private double PenPad(Matrix m)
        => state.Pen.Color == "none" ? 0 : state.Pen.Cosmetic ? 0.5
            : state.Pen.Width / 2 * Math.Max(Math.Sqrt(m.A * m.A + m.B * m.B), Math.Sqrt(m.C * m.C + m.D * m.D));

    private void Flush()
    {
        if (stroke != null && strokeElement != null)
        {
            string d = stroke.Empty ? "M0 0" : stroke.ToString();
            Spend(d.Length);
            strokeElement.SetAttributeValue("d", d);
        }
        stroke = null; strokeElement = null; strokeKey = null;
    }

    private void Add(XElement element)
    {
        if (element != strokeElement) Flush();
        Count(element);
        Container().Add(element);
    }

    private XElement Container()
    {
        if (state.Clip is not { } clip) return root;
        if (clipGroup != null && root.LastNode == clipGroup && (string?)clipGroup.Attribute("clip-path") == $"url(#{clip})") return clipGroup;
        Flush();
        clipGroup = new XElement(Svg + "g", new XAttribute("clip-path", $"url(#{clip})"));
        Count(clipGroup);
        root.Add(clipGroup);
        return clipGroup;
    }

    private XElement Defs()
    {
        if (defs != null) return defs;
        defs = new XElement(Svg + "defs");
        root.AddFirst(defs);
        return defs;
    }

    private string NextId() => idPrefix + (++ids).ToString(CultureInfo.InvariantCulture);

    private void Count(int added)
    {
        if ((points += added) > MaxPoints) throw new NotSupportedException("Metafile exceeds the supported drawing size.");
    }

    /// <summary>Counts an element about to join the drawing, and charges what it will serialize to.</summary>
    private void Count(XElement element)
    {
        if (++elements > MaxElements) throw new NotSupportedException("Metafile exceeds the supported drawing size.");
        foreach (var e in element.DescendantsAndSelf())
        {
            long size = 2 * e.Name.LocalName.Length + 5;
            foreach (var a in e.Attributes()) size += a.Name.LocalName.Length + a.Value.Length + 4;
            foreach (var n in e.Nodes()) if (n is XText t) size += t.Value.Length;
            Spend(size);
        }
    }

    private void Spend(long characters)
    {
        if ((output += characters) > MaxOutputCharacters) throw new NotSupportedException("Metafile exceeds the supported output size.");
    }

    private void Track(Matrix m, double x, double y, double pad)
    {
        var (dx, dy) = m.Apply(x, y);
        if (!double.IsFinite(dx) || !double.IsFinite(dy) || Math.Abs(dx) > Huge || Math.Abs(dy) > Huge)
            throw new InvalidDataException("Metafile coordinates exceed the supported range.");
        minX = Math.Min(minX, dx - pad); minY = Math.Min(minY, dy - pad);
        maxX = Math.Max(maxX, dx + pad); maxY = Math.Max(maxY, dy + pad);
    }

    private static void SetTransformAttribute(XElement element, Matrix m, string name = "transform", bool always = false)
    {
        if (!always && m == Matrix.Identity) return;
        element.SetAttributeValue(name, m.B == 0 && m.C == 0 && m.E == 0 && m.F == 0
            ? $"scale({N(m.A)} {N(m.D)})"
            : $"matrix({N(m.A)} {N(m.B)} {N(m.C)} {N(m.D)} {N(m.E)} {N(m.F)})");
    }

    private static void Finite(Matrix m)
    {
        // NaN fails every comparison, so "not within range" also rejects it.
        static bool Bad(double v) => !(Math.Abs(v) <= Huge);
        if (Bad(m.A) || Bad(m.B) || Bad(m.C) || Bad(m.D) || Bad(m.E) || Bad(m.F))
            throw new InvalidDataException("Invalid metafile transform.");
    }

    private static (double X, double Y)[] Corners(double left, double top, double right, double bottom)
        => new[] { (left, top), (right, bottom), (left, bottom), (right, top) };

    private static void Order(ref double low, ref double high) { if (high < low) (low, high) = (high, low); }

    /// <summary>Ten significant digits: every coordinate the canvas accepts (up to 1e9) keeps its integer
    /// part, a small scale keeps its precision, and binary noise such as 0.30000000000000004 is dropped
    /// (SVG accepts the exponent form). Trigonometric residue such as cos 90° = 6e-17 is zero.</summary>
    internal static string N(double value)
        => Math.Abs(value) < 1e-9 ? "0" : value.ToString("G10", CultureInfo.InvariantCulture);

    /// <summary>SVG path data in logical coordinates, with GDI's notion of an open figure.</summary>
    private sealed class Path
    {
        private readonly StringBuilder data = new();
        internal readonly List<(double X, double Y)> Points = new();
        internal bool Open;
        internal (double X, double Y) Start;
        internal bool Empty => data.Length == 0;

        /// <summary>Starts a figure at the current position unless one is already open.</summary>
        internal void Continue(double x, double y) { if (!Open) Op('M', x, y); }

        internal void Op(char op, params double[] xy)
        {
            data.Append(op);
            for (int i = 0; i < xy.Length; i += 2)
            {
                if (i > 0) data.Append(' ');
                data.Append(N(xy[i])).Append(' ').Append(N(xy[i + 1]));
                Points.Add((xy[i], xy[i + 1]));
            }
            if (op == 'M') { Open = true; Start = (xy[0], xy[1]); }
        }

        internal void Arc(double rx, double ry, bool large, bool positive, double x, double y)
        {
            data.Append('A').Append(N(rx)).Append(' ').Append(N(ry)).Append(" 0 ").Append(large ? '1' : '0')
                .Append(' ').Append(positive ? '1' : '0').Append(' ').Append(N(x)).Append(' ').Append(N(y));
            Points.Add((x, y));
        }

        internal void Close() { if (Open) { data.Append('Z'); Open = false; } }
        internal void CloseOpenFigure() => Close();
        public override string ToString() => data.ToString();
    }
}
