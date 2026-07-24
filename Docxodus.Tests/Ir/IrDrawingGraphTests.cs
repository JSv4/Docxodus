#nullable enable

using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Synthetic coverage for drawing-local relationship-graph identity. These fixtures intentionally construct
/// only the relevant package edges (outer DrawingML → chart/diagram XML → optional nested data), so they prove
/// the IR behavior without depending on corpus or benchmark documents.
/// </summary>
public class IrDrawingGraphTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    [Fact]
    public void Read_ChartXmlChange_FlipsOpaqueDrawingIdentity()
    {
        var left = ChartDocument("before");
        var right = ChartDocument("after");

        var leftParagraph = Paragraph(left);
        var rightParagraph = Paragraph(right);

        Assert.NotEqual(Opaque(leftParagraph).CanonicalHash, Opaque(rightParagraph).CanonicalHash);
        Assert.NotEqual(leftParagraph.ContentHash, rightParagraph.ContentHash);
    }

    [Fact]
    public void Read_ChartGraph_RelationshipIdRenumbering_IsEqual()
    {
        var workbook = new byte[] { 0x50, 0x4B, 0x03, 0x04, 1, 2, 3 };
        var first = ChartDocument("same", "rIdChart", "rIdWorkbook", workbook);
        // Only outer and nested relationship ids differ; the reachable semantic content is identical.
        var second = ChartDocument("same", "rIdRenumbered", "rIdNested", workbook);

        Assert.Equal(Opaque(Paragraph(first)).CanonicalHash, Opaque(Paragraph(second)).CanonicalHash);
        Assert.Equal(Paragraph(first).ContentHash, Paragraph(second).ContentHash);
    }

    [Fact]
    public void Read_ChartEmbeddedWorkbookChange_FlipsDrawingIdentity()
    {
        var left = ChartDocument("same", "rIdChart", "rIdWorkbook", new byte[] { 0x50, 0x4B, 1 });
        var right = ChartDocument("same", "rIdChart", "rIdWorkbook", new byte[] { 0x50, 0x4B, 2 });

        Assert.NotEqual(Opaque(Paragraph(left)).CanonicalHash, Opaque(Paragraph(right)).CanonicalHash);
    }

    [Fact]
    public void Read_SmartArtDataChange_FlipsOpaqueDrawingIdentity()
    {
        var left = SmartArtDocument("before");
        var right = SmartArtDocument("after");

        Assert.NotEqual(Opaque(Paragraph(left)).CanonicalHash, Opaque(Paragraph(right)).CanonicalHash);
        Assert.NotEqual(Paragraph(left).ContentHash, Paragraph(right).ContentHash);
    }

    [Fact]
    public void Read_SmartArtPrebuiltRelationshipRenumbering_IsEqual()
    {
        var first = SmartArtDocument("same", "rIdData", "rIdPrebuilt", "prebuilt");
        var second = SmartArtDocument("same", "rIdDataRenumbered", "rIdPrebuiltRenumbered", "prebuilt");

        Assert.Equal(Opaque(Paragraph(first)).CanonicalHash, Opaque(Paragraph(second)).CanonicalHash);
    }

    [Fact]
    public void Read_HeaderChart_ResolvesItsOwnRelationshipScope()
    {
        // Both owning parts deliberately use rIdChart. A graph hasher rooted at MainDocumentPart would read
        // the unchanged main chart for the header too and miss this header-only chart edit.
        var left = HeaderScopedChartDocument("header-left");
        var right = HeaderScopedChartDocument("header-right");

        var leftIr = IrReader.Read(left);
        var rightIr = IrReader.Read(right);
        var leftHeader = leftIr.Headers.Single().Scope.Blocks.OfType<IrParagraph>().Single();
        var rightHeader = rightIr.Headers.Single().Scope.Blocks.OfType<IrParagraph>().Single();

        Assert.Equal(Opaque(Paragraph(left)).CanonicalHash, Opaque(Paragraph(right)).CanonicalHash);
        Assert.NotEqual(Opaque(leftHeader).CanonicalHash, Opaque(rightHeader).CanonicalHash);
    }

    [Fact]
    public void Compare_ChartGraphChange_AcceptRejectUseLiveRightLeftGraph()
    {
        var left = ChartDocument("left");
        var right = ChartDocument("right");

        var redline = DocxDiff.Compare(left, right);

        Assert.Equal("right", ChartMarker(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("left", ChartMarker(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_SmartArtGraphChange_AcceptRejectUseLiveRightLeftGraph()
    {
        var left = SmartArtDocument("left");
        var right = SmartArtDocument("right");

        var redline = DocxDiff.Compare(left, right);

        Assert.Equal("right", SmartArtMarker(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("left", SmartArtMarker(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_ChartEmbeddedWorkbookChange_AcceptRejectUseLiveRightLeftBytes()
    {
        var leftBytes = new byte[] { 0x50, 0x4B, 1 };
        var rightBytes = new byte[] { 0x50, 0x4B, 2 };
        var redline = DocxDiff.Compare(
            ChartDocument("same", workbookRelationshipId: "rIdWorkbook", workbookBytes: leftBytes),
            ChartDocument("same", workbookRelationshipId: "rIdWorkbook", workbookBytes: rightBytes));

        Assert.Equal(rightBytes, ChartWorkbookBytes(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(leftBytes, ChartWorkbookBytes(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_ChartExternalDataExternalUri_AcceptRejectUseLiveRightLeftUri()
    {
        var redline = DocxDiff.Compare(
            ChartExternalDataDocument("https://left.example/data.xlsx"),
            ChartExternalDataDocument("https://right.example/data.xlsx"));

        Assert.Equal("https://right.example/data.xlsx", ChartExternalDataUri(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("https://left.example/data.xlsx", ChartExternalDataUri(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_SmartArtPrebuiltChange_AcceptRejectUseLiveRightLeftPrebuilt()
    {
        var redline = DocxDiff.Compare(
            SmartArtDocument("same", prebuiltRelationshipId: "rIdPrebuilt", prebuiltMarker: "left"),
            SmartArtDocument("same", prebuiltRelationshipId: "rIdPrebuilt", prebuiltMarker: "right"));

        Assert.Equal("right", SmartArtPrebuiltMarker(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("left", SmartArtPrebuiltMarker(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_RelationshipOnImportedXmlRoot_AcceptRejectRemapIt()
    {
        var leftBytes = new byte[] { 1, 2, 3 };
        var rightBytes = new byte[] { 4, 5, 6 };
        var redline = DocxDiff.Compare(
            ChartDocumentWithRootPayload(leftBytes),
            ChartDocumentWithRootPayload(rightBytes));

        Assert.Equal(rightBytes, ChartRootPayloadBytes(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(leftBytes, ChartRootPayloadBytes(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Read_MalformedXmlGraphBytes_ChangeOpaqueIdentityWithoutThrowing()
    {
        var left = MalformedGraphDocument("<g:root xmlns:g=\"urn:docxdiff:test\">left");
        var right = MalformedGraphDocument("<g:root xmlns:g=\"urn:docxdiff:test\">right");

        var leftParagraph = Paragraph(left);
        var rightParagraph = Paragraph(right);

        Assert.NotEqual(Opaque(leftParagraph).CanonicalHash, Opaque(rightParagraph).CanonicalHash);
        Assert.NotEqual(leftParagraph.ContentHash, rightParagraph.ContentHash);
    }

    [Fact]
    public void Compare_MalformedXmlGraphBytes_AcceptRejectPreserveRawRightLeftBytes()
    {
        const string leftXml = "<g:root xmlns:g=\"urn:docxdiff:test\">left";
        const string rightXml = "<g:root xmlns:g=\"urn:docxdiff:test\">right";
        var redline = DocxDiff.Compare(MalformedGraphDocument(leftXml), MalformedGraphDocument(rightXml));

        Assert.Equal(rightXml, MalformedGraphBytes(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(leftXml, MalformedGraphBytes(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_ChartGraphBehindTextbox_UsesStructuralCarrierForAcceptReject()
    {
        var left = ChartTextboxDocument("left");
        var right = ChartTextboxDocument("right");
        var leftParagraph = Paragraph(left);
        var rightParagraph = Paragraph(right);

        Assert.NotEqual(default, leftParagraph.InlineEnvelopeDigest);
        Assert.NotEqual(leftParagraph.InlineEnvelopeDigest, rightParagraph.InlineEnvelopeDigest);

        var redline = DocxDiff.Compare(left, right);
        Assert.Equal("right", ChartMarker(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("left", ChartMarker(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_AlternateContentChoiceChartGraph_AcceptRejectUseLiveRightLeftGraph()
    {
        var left = AlternateContentChartDocument("left", includeFallbackTextbox: false);
        var right = AlternateContentChartDocument("right", includeFallbackTextbox: false);

        Assert.NotEqual(Opaque(Paragraph(left)).CanonicalHash, Opaque(Paragraph(right)).CanonicalHash);

        var redline = DocxDiff.Compare(left, right);
        Assert.Equal("right", ChartMarker(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("left", ChartMarker(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_AlternateContentChoiceChartWithFallbackTextbox_UsesStructuralCarrier()
    {
        var left = AlternateContentChartDocument("left", includeFallbackTextbox: true);
        var right = AlternateContentChartDocument("right", includeFallbackTextbox: true);
        var leftParagraph = Paragraph(left);
        var rightParagraph = Paragraph(right);

        Assert.NotEqual(default, leftParagraph.InlineEnvelopeDigest);
        Assert.NotEqual(leftParagraph.InlineEnvelopeDigest, rightParagraph.InlineEnvelopeDigest);

        var redline = DocxDiff.Compare(left, right);
        Assert.Equal("right", ChartMarker(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("left", ChartMarker(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Read_OrdinaryImage_DrawingDigestMatchesLegacyResolverCanonicalization()
    {
        var document = ImageDocument();
        var paragraph = Paragraph(document);
        var image = Assert.Single(paragraph.Inlines.OfType<IrInlineImage>());

        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var drawing = main.GetXDocument().Descendants(W + "drawing").Single();

        Assert.Equal(IrHasher.CanonicalHash(drawing, new IrRelResolver(main)), image.DrawingDigest);
    }

    [Fact]
    public void Hash_IncompleteGraphFallback_StreamsTheEntireCutoffXmlPart()
    {
        // The marker is deliberately after the hasher's 81,920-byte I/O chunk. Hold the package fallback
        // constant so this assertion isolates the immediate raw-part component of the incomplete identity.
        var padding = new string('x', 100_000);
        var left = ChartDocument("left", leadingPadding: padding);
        var right = ChartDocument("right", leadingPadding: padding);
        var limits = new IrDrawingGraphHashLimits(
            MaxXmlGraphDepth: 32,
            MaxXmlGraphParts: 256,
            MaxXmlPartBytes: 1024,
            MaxXmlGraphBytes: 1024 * 1024);
        var sharedPackageFingerprint = IrHash.Compute("test-shared-package-fingerprint");

        Assert.NotEqual(
            LimitedDrawingHash(left, limits, sharedPackageFingerprint),
            LimitedDrawingHash(right, limits, sharedPackageFingerprint));
    }

    [Fact]
    public void Hash_IncompleteGraphFallback_DetectsNestedTargetChangesAtEveryLimit()
    {
        // The outer DrawingML and chart XML are byte-identical; only the workbook reachable through
        // c:externalData/@r:id differs. Every synthetic limit must therefore make a conservative package-bound
        // identity rather than collapsing the incomplete graph to a content-type sentinel.
        var left = ChartDocument("same", workbookRelationshipId: "rIdWorkbook", workbookBytes: new byte[] { 1 });
        var right = ChartDocument("same", workbookRelationshipId: "rIdWorkbook", workbookBytes: new byte[] { 2 });
        var limits = new[]
        {
            new IrDrawingGraphHashLimits(0, 256, 1024 * 1024, 1024 * 1024), // depth
            new IrDrawingGraphHashLimits(32, 0, 1024 * 1024, 1024 * 1024), // part count
            new IrDrawingGraphHashLimits(32, 256, 1, 1024 * 1024), // per-part byte budget
            new IrDrawingGraphHashLimits(32, 256, 1024 * 1024, 1), // aggregate byte budget
        };

        foreach (var limit in limits)
        {
            Assert.NotEqual(LimitedDrawingHash(left, limit), LimitedDrawingHash(right, limit));
            Assert.Equal(LimitedDrawingHash(left, limit), LimitedDrawingHash(left, limit));
        }
    }

    [Fact]
    public void MoveRelatedParts_SharedDiagramDataKeepsOwnerScopedPrebuiltEdgesDistinct()
    {
        // SmartArt's dataModelExt stores a relationship id that resolves on the part which links the
        // DiagramDataPart, not on the data part itself. Two XML owners can legally share one data part while
        // using the same local id for two different prebuilt drawings. The importer must clone the data part per
        // owner; otherwise the first copied data part is reused and its extension is wrong for the second owner.
        const string xmlRelationship = "urn:docxdiff:test/xml";
        const string diagramDataRelationship =
            "http://schemas.microsoft.com/office/2007/relationships/diagramData";
        const string diagramDrawingRelationship =
            "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
        XNamespace test = "urn:docxdiff:shared-diagram-data";
        XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        XNamespace dsp = "http://schemas.microsoft.com/office/drawing/2008/diagram";

        using var sourceBytes = new MemoryStream();
        using var destinationBytes = new MemoryStream();
        using var sourcePackage = Package.Open(sourceBytes, FileMode.Create, FileAccess.ReadWrite);
        using var destinationPackage = Package.Open(destinationBytes, FileMode.Create, FileAccess.ReadWrite);

        var sourceRoot = sourcePackage.CreatePart(new Uri("/source/root.xml", UriKind.Relative), "application/xml");
        var sourceOwnerA = sourcePackage.CreatePart(new Uri("/source/owner-a.xml", UriKind.Relative), "application/xml");
        var sourceOwnerB = sourcePackage.CreatePart(new Uri("/source/owner-b.xml", UriKind.Relative), "application/xml");
        var sourceData = sourcePackage.CreatePart(new Uri("/source/data.xml", UriKind.Relative), "application/xml");
        var sourcePrebuiltA = sourcePackage.CreatePart(new Uri("/source/prebuilt-a.xml", UriKind.Relative), "application/xml");
        var sourcePrebuiltB = sourcePackage.CreatePart(new Uri("/source/prebuilt-b.xml", UriKind.Relative), "application/xml");

        WritePackageXml(sourceRoot, $"<t:root xmlns:t=\"{test}\"/>");
        WritePackageXml(sourceOwnerA, $"<dgm:relIds xmlns:dgm=\"{dgm}\" xmlns:r=\"{R}\" r:dm=\"rIdData\"/>");
        WritePackageXml(sourceOwnerB, $"<dgm:relIds xmlns:dgm=\"{dgm}\" xmlns:r=\"{R}\" r:dm=\"rIdData\"/>");
        WritePackageXml(sourceData,
            $"<dgm:dataModel xmlns:dgm=\"{dgm}\" xmlns:dsp=\"{dsp}\"><dsp:dataModelExt relId=\"rIdPrebuilt\"/></dgm:dataModel>");
        WritePackageXml(sourcePrebuiltA, $"<t:prebuilt xmlns:t=\"{test}\">prebuilt-a</t:prebuilt>");
        WritePackageXml(sourcePrebuiltB, $"<t:prebuilt xmlns:t=\"{test}\">prebuilt-b</t:prebuilt>");

        sourceRoot.CreateRelationship(sourceOwnerA.Uri, TargetMode.Internal, xmlRelationship, "rIdA");
        sourceRoot.CreateRelationship(sourceOwnerB.Uri, TargetMode.Internal, xmlRelationship, "rIdB");
        sourceOwnerA.CreateRelationship(sourceData.Uri, TargetMode.Internal, diagramDataRelationship, "rIdData");
        sourceOwnerB.CreateRelationship(sourceData.Uri, TargetMode.Internal, diagramDataRelationship, "rIdData");
        sourceOwnerA.CreateRelationship(sourcePrebuiltA.Uri, TargetMode.Internal, diagramDrawingRelationship, "rIdPrebuilt");
        sourceOwnerB.CreateRelationship(sourcePrebuiltB.Uri, TargetMode.Internal, diagramDrawingRelationship, "rIdPrebuilt");

        var destinationRoot = destinationPackage.CreatePart(
            new Uri("/destination/root.xml", UriKind.Relative), "application/xml");
        var carrier = new XElement(test + "root",
            new XAttribute(XNamespace.Xmlns + "r", R),
            new XElement(test + "a", new XAttribute(R + "id", "rIdA")),
            new XElement(test + "b", new XAttribute(R + "id", "rIdB")));

        WmlComparer.MoveRelatedPartsToDestination(sourceRoot, destinationRoot, carrier);

        var destinationOwnerA = RelatedPart(destinationRoot, (string)carrier.Element(test + "a")!.Attribute(R + "id")!);
        var destinationOwnerB = RelatedPart(destinationRoot, (string)carrier.Element(test + "b")!.Attribute(R + "id")!);
        var destinationDataA = RelatedPart(destinationOwnerA,
            (string)ReadPackageXml(destinationOwnerA).Root!.Attribute(R + "dm")!);
        var destinationDataB = RelatedPart(destinationOwnerB,
            (string)ReadPackageXml(destinationOwnerB).Root!.Attribute(R + "dm")!);

        Assert.NotEqual(destinationDataA.Uri, destinationDataB.Uri);
        Assert.Equal("prebuilt-a", OwnerScopedPrebuiltMarker(destinationOwnerA, destinationDataA, dsp));
        Assert.Equal("prebuilt-b", OwnerScopedPrebuiltMarker(destinationOwnerB, destinationDataB, dsp));
    }

    private static IrParagraph Paragraph(WmlDocument document) =>
        IrReader.Read(document).Body.Blocks.OfType<IrParagraph>().Single();

    private static string OwnerScopedPrebuiltMarker(
        PackagePart owner, PackagePart data, XNamespace dsp)
    {
        var prebuiltRelationshipId = (string)ReadPackageXml(data)
            .Descendants(dsp + "dataModelExt").Single().Attribute("relId")!;
        return ReadPackageXml(RelatedPart(owner, prebuiltRelationshipId)).Root!.Value;
    }

    private static PackagePart RelatedPart(PackagePart owner, string relationshipId)
    {
        var relationship = owner.GetRelationship(relationshipId);
        var targetUri = PackUriHelper.ResolvePartUri(owner.Uri, relationship.TargetUri);
        return owner.Package.GetPart(targetUri);
    }

    private static XDocument ReadPackageXml(PackagePart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream);
    }

    private static void WritePackageXml(PackagePart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }

    private static IrOpaqueInline Opaque(IrParagraph paragraph) =>
        Assert.Single(paragraph.Inlines.OfType<IrOpaqueInline>());

    private static IrHash LimitedDrawingHash(
        WmlDocument document,
        IrDrawingGraphHashLimits limits,
        IrHash? sourceDocumentFingerprint = null)
    {
        var documentBytes = document.DocumentByteArray;
        using var stream = new MemoryStream(documentBytes);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var drawing = main.GetXDocument().Descendants(W + "drawing").Single();
        return new IrDrawingGraphHasher(
            main,
            new Lazy<IrHash>(() => sourceDocumentFingerprint ?? IrHash.Compute(documentBytes)),
            limits).Hash(drawing);
    }

    private static WmlDocument ChartDocument(
        string marker,
        string chartRelationshipId = "rIdChart",
        string? workbookRelationshipId = null,
        byte[]? workbookBytes = null,
        string? leadingPadding = null)
    {
        if ((workbookRelationshipId is null) != (workbookBytes is null))
            throw new ArgumentException("Workbook relationship id and bytes must be supplied together.");

        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var chart = main.AddNewPart<ChartPart>(chartRelationshipId);
            if (workbookRelationshipId is not null)
            {
                var workbook = chart.AddExtendedPart(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ".xlsx",
                    workbookRelationshipId);
                using var workbookStream = workbook.GetStream(FileMode.Create, FileAccess.Write);
                workbookStream.Write(workbookBytes!, 0, workbookBytes!.Length);
            }

            var externalData = workbookRelationshipId is null
                ? string.Empty
                : $"<c:externalData r:id=\"{workbookRelationshipId}\"/>";
            var padding = leadingPadding is null ? string.Empty : $"<c:padding>{leadingPadding}</c:padding>";
            WriteXml(chart,
                $"<c:chartSpace xmlns:c=\"{C}\" xmlns:r=\"{R}\">" +
                $"<c:chart>{padding}<c:marker>{marker}</c:marker></c:chart>{externalData}</c:chartSpace>");

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                $"xmlns:c=\"{C}\"><w:body><w:p><w:r>{ChartDrawing(chartRelationshipId)}" +
                "</w:r></w:p></w:body></w:document>");
        }
        return new WmlDocument("chart-graph.docx", stream.ToArray());
    }

    private static WmlDocument ChartExternalDataDocument(string externalUri)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var chart = main.AddNewPart<ChartPart>("rIdChart");
            chart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
                new Uri(externalUri, UriKind.Absolute), "rIdExternal");
            WriteXml(chart,
                $"<c:chartSpace xmlns:c=\"{C}\" xmlns:r=\"{R}\">" +
                "<c:chart><c:marker>same</c:marker></c:chart><c:externalData r:id=\"rIdExternal\"/>" +
                "</c:chartSpace>");

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                $"xmlns:c=\"{C}\"><w:body><w:p><w:r>{ChartDrawing("rIdChart")}" +
                "</w:r></w:p></w:body></w:document>");
        }
        return new WmlDocument("chart-external-graph.docx", stream.ToArray());
    }

    private static WmlDocument ChartDocumentWithRootPayload(byte[] payload)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var chart = main.AddNewPart<ChartPart>("rIdChart");
            var nested = chart.AddExtendedPart(
                "urn:docxdiff:test-root-payload",
                "application/octet-stream",
                ".bin",
                "rIdRootPayload");
            WriteBytes(nested, payload);
            // c:chartSpace/@r:id is intentionally schema-nonstandard. The generic relationship copier promises
            // to handle every supported relationship attribute, including one located on an imported XML ROOT.
            WriteXml(chart,
                $"<c:chartSpace xmlns:c=\"{C}\" xmlns:r=\"{R}\" r:id=\"rIdRootPayload\">" +
                "<c:chart><c:marker>same</c:marker></c:chart></c:chartSpace>");

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                $"xmlns:c=\"{C}\"><w:body><w:p><w:r>{ChartDrawing("rIdChart")}" +
                "</w:r></w:p></w:body></w:document>");
        }
        return new WmlDocument("chart-root-payload.docx", stream.ToArray());
    }

    private static WmlDocument MalformedGraphDocument(string rawXml)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var graph = main.AddExtendedPart(
                "urn:docxdiff:test-graph",
                "application/vnd.docxdiff.test+xml",
                ".xml",
                "rIdGraph");
            WriteBytes(graph, System.Text.Encoding.UTF8.GetBytes(rawXml));

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">" +
                "<w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"1\" cy=\"1\"/>" +
                "<a:graphic><a:graphicData uri=\"urn:docxdiff:test-graph\">" +
                "<g:node xmlns:g=\"urn:docxdiff:test\" r:id=\"rIdGraph\"/>" +
                "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>");
        }
        return new WmlDocument("malformed-graph.docx", stream.ToArray());
    }

    private static WmlDocument ChartTextboxDocument(string marker)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var chart = main.AddNewPart<ChartPart>("rIdChart");
            WriteXml(chart,
                $"<c:chartSpace xmlns:c=\"{C}\"><c:chart><c:marker>{marker}</c:marker></c:chart></c:chartSpace>");

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
                $"xmlns:c=\"{C}\"><w:body><w:p><w:r>{ChartTextboxDrawing("rIdChart")}" +
                "</w:r></w:p></w:body></w:document>");
        }
        return new WmlDocument("chart-textbox-graph.docx", stream.ToArray());
    }

    private static WmlDocument AlternateContentChartDocument(string marker, bool includeFallbackTextbox)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var chart = main.AddNewPart<ChartPart>("rIdChart");
            WriteXml(chart,
                $"<c:chartSpace xmlns:c=\"{C}\"><c:chart><c:marker>{marker}</c:marker></c:chart></c:chartSpace>");
            var fallback = includeFallbackTextbox
                ? "<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>" +
                  "<w:p><w:r><w:t>Fallback textbox</w:t></w:r></w:p>" +
                  "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>"
                : string.Empty;

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
                "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                $"xmlns:c=\"{C}\"><w:body><w:p><w:r><mc:AlternateContent>" +
                $"<mc:Choice Requires=\"wps\">{ChartDrawing("rIdChart")}</mc:Choice>{fallback}" +
                "</mc:AlternateContent></w:r></w:p></w:body></w:document>");
        }
        return new WmlDocument("alternate-content-chart.docx", stream.ToArray());
    }

    private static WmlDocument ImageDocument()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            var image = main.AddImagePart(ImagePartType.Png, "rIdImage");
            WriteBytes(image, new byte[] { 0x89, 0x50, 0x4E, 0x47 });

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">" +
                "<w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"200\"/>" +
                "<a:graphic><a:graphicData><a:blip r:embed=\"rIdImage\"/>" +
                "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>");
        }
        return new WmlDocument("ordinary-image.docx", stream.ToArray());
    }

    private static WmlDocument SmartArtDocument(
        string marker,
        string dataRelationshipId = "rIdData",
        string? prebuiltRelationshipId = null,
        string? prebuiltMarker = null)
    {
        if ((prebuiltRelationshipId is null) != (prebuiltMarker is null))
            throw new ArgumentException("Prebuilt relationship id and marker must be supplied together.");

        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var data = main.AddNewPart<DiagramDataPart>(dataRelationshipId);
            if (prebuiltRelationshipId is not null)
            {
                var prebuilt = main.AddExtendedPart(
                    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
                    "application/vnd.ms-office.drawingml.diagramDrawing+xml",
                    ".xml",
                    prebuiltRelationshipId);
                WriteXml(prebuilt,
                    "<dsp:drawing xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\">" +
                    $"<dsp:marker>{prebuiltMarker}</dsp:marker></dsp:drawing>");
            }

            var prebuiltEdge = prebuiltRelationshipId is null
                ? string.Empty
                : "<dsp:dataModelExt xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" " +
                  $"relId=\"{prebuiltRelationshipId}\"/>";
            WriteXml(data,
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">" +
                $"<dgm:pt modelId=\"{marker}\"/>{prebuiltEdge}</dgm:dataModel>");

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                "xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">" +
                $"<w:body><w:p><w:r>{SmartArtDrawing(dataRelationshipId)}</w:r></w:p></w:body></w:document>");
        }
        return new WmlDocument("smartart-graph.docx", stream.ToArray());
    }

    private static WmlDocument HeaderScopedChartDocument(string headerMarker)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var mainChart = main.AddNewPart<ChartPart>("rIdChart");
            WriteXml(mainChart,
                $"<c:chartSpace xmlns:c=\"{C}\"><c:chart><c:marker>main</c:marker></c:chart></c:chartSpace>");

            var header = main.AddNewPart<HeaderPart>("rIdHeader");
            var headerChart = header.AddNewPart<ChartPart>("rIdChart");
            WriteXml(headerChart,
                $"<c:chartSpace xmlns:c=\"{C}\"><c:chart><c:marker>{headerMarker}</c:marker></c:chart></c:chartSpace>");
            WriteXml(header,
                $"<w:hdr xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                $"xmlns:c=\"{C}\"><w:p><w:r>{ChartDrawing("rIdChart")}</w:r></w:p></w:hdr>");

            WriteXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                $"xmlns:c=\"{C}\"><w:body><w:p><w:r>{ChartDrawing("rIdChart")}</w:r></w:p>" +
                "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("header-chart-graph.docx", stream.ToArray());
    }

    private static string ChartDrawing(string relationshipId) =>
        "<w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"200\"/><wp:docPr id=\"1\" name=\"Chart\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
        $"<c:chart r:id=\"{relationshipId}\"/></a:graphicData></a:graphic></wp:inline></w:drawing>";

    private static string ChartTextboxDrawing(string relationshipId) =>
        "<w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"200\"/><wp:docPr id=\"1\" name=\"Chart\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
        $"<c:chart r:id=\"{relationshipId}\"/>" +
        "<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>Inside</w:t></w:r></w:p>" +
        "</w:txbxContent></wps:txbx></wps:wsp>" +
        "</a:graphicData></a:graphic></wp:inline></w:drawing>";

    private static string SmartArtDrawing(string relationshipId) =>
        "<w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"200\"/><wp:docPr id=\"1\" name=\"Diagram\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">" +
        $"<dgm:relIds r:dm=\"{relationshipId}\"/></a:graphicData></a:graphic></wp:inline></w:drawing>";

    private static string ChartMarker(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var relationshipId = (string?)main.GetXDocument().Descendants(C + "chart").Single().Attribute(R + "id");
        var chart = Assert.IsType<ChartPart>(main.GetPartById(relationshipId!));
        return chart.GetXDocument().Descendants(C + "marker").Single().Value;
    }

    private static byte[] ChartWorkbookBytes(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var chartId = (string?)main.GetXDocument().Descendants(C + "chart").Single().Attribute(R + "id");
        var chart = Assert.IsType<ChartPart>(main.GetPartById(chartId!));
        var workbookId = (string?)chart.GetXDocument().Descendants(C + "externalData").Single().Attribute(R + "id");
        return PartBytes(chart.GetPartById(workbookId!));
    }

    private static string ChartExternalDataUri(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var chartId = (string?)main.GetXDocument().Descendants(C + "chart").Single().Attribute(R + "id");
        var chart = Assert.IsType<ChartPart>(main.GetPartById(chartId!));
        var externalId = (string?)chart.GetXDocument().Descendants(C + "externalData").Single().Attribute(R + "id");
        return chart.ExternalRelationships.Single(relationship => relationship.Id == externalId).Uri.ToString();
    }

    private static byte[] ChartRootPayloadBytes(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var chartId = (string?)main.GetXDocument().Descendants(C + "chart").Single().Attribute(R + "id");
        var chart = Assert.IsType<ChartPart>(main.GetPartById(chartId!));
        var payloadId = (string?)chart.GetXDocument().Root!.Attribute(R + "id");
        return PartBytes(chart.GetPartById(payloadId!));
    }

    private static string SmartArtMarker(WmlDocument document)
    {
        XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var relationshipId = (string?)main.GetXDocument().Descendants(dgm + "relIds")
            .Single().Attribute(R + "dm");
        var data = Assert.IsType<DiagramDataPart>(main.GetPartById(relationshipId!));
        return (string?)data.GetXDocument().Descendants(dgm + "pt").Single().Attribute("modelId") ?? string.Empty;
    }

    private static string SmartArtPrebuiltMarker(WmlDocument document)
    {
        XNamespace dsp = "http://schemas.microsoft.com/office/drawing/2008/diagram";
        XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var dataId = (string?)main.GetXDocument().Descendants(dgm + "relIds").Single().Attribute(R + "dm");
        var data = Assert.IsType<DiagramDataPart>(main.GetPartById(dataId!));
        var prebuiltId = (string?)data.GetXDocument().Descendants(dsp + "dataModelExt")
            .Single().Attribute("relId");
        var prebuilt = main.GetPartById(prebuiltId!);
        return ReadXml(prebuilt).Descendants(dsp + "marker").Single().Value;
    }

    private static string MalformedGraphBytes(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart!;
        var graphId = (string?)main.GetXDocument().Descendants()
            .Single(element => element.Name.NamespaceName == "urn:docxdiff:test" && element.Name.LocalName == "node")
            .Attribute(R + "id");
        return System.Text.Encoding.UTF8.GetString(PartBytes(main.GetPartById(graphId!)));
    }

    private static XDocument ReadXml(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream);
    }

    private static byte[] PartBytes(OpenXmlPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var destination = new MemoryStream();
        source.CopyTo(destination);
        return destination.ToArray();
    }

    private static void WriteBytes(OpenXmlPart part, byte[] bytes)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }
}
