#nullable enable

using System.Globalization;
using Docxodus;
using DocxodusWeb.Data;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

// Configure SQLite — store DB in a persistent data directory
var dataDir = Environment.GetEnvironmentVariable("DATA_DIR") ?? Path.Combine(Directory.GetCurrentDirectory(), "appdata");
Directory.CreateDirectory(dataDir);
var uploadsDir = Path.Combine(dataDir, "uploads");
Directory.CreateDirectory(uploadsDir);

builder.Services.AddDbContext<TicketDbContext>(opt =>
    opt.UseSqlite($"Data Source={Path.Combine(dataDir, "tickets.db")}"));

// Allow large file uploads (100 MB)
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = 100 * 1024 * 1024);

var app = builder.Build();

// Auto-migrate on startup
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<TicketDbContext>();
    db.Database.EnsureCreated();
}

app.UseStaticFiles();

// ──────────────────────────────────────────────
// Redline API
// ──────────────────────────────────────────────

app.MapPost("/api/compare", async (HttpRequest request) =>
{
    var form = await request.ReadFormAsync();
    var originalFile = form.Files.GetFile("original");
    var modifiedFile = form.Files.GetFile("modified");

    if (originalFile is null || modifiedFile is null)
        return Results.BadRequest(new { error = "Both 'original' and 'modified' .docx files are required." });

    var author = form["author"].FirstOrDefault() ?? "Docxodus";
    var detailThreshold = 0.0;
    if (form.ContainsKey("detailThreshold") &&
        double.TryParse(form["detailThreshold"], NumberStyles.Float, CultureInfo.InvariantCulture, out var dt))
        detailThreshold = dt;

    var settings = new WmlComparerSettings
    {
        AuthorForRevisions = author,
        DetailThreshold = detailThreshold,
        DetectMoves = true,
        DetectFormatChanges = true,
    };

    try
    {
        var originalBytes = await ReadFormFile(originalFile);
        var modifiedBytes = await ReadFormFile(modifiedFile);

        var originalDoc = new WmlDocument("original.docx", originalBytes);
        var modifiedDoc = new WmlDocument("modified.docx", modifiedBytes);

        var result = WmlComparer.Compare(originalDoc, modifiedDoc, settings);

        return Results.File(result.DocumentByteArray,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "redline.docx");
    }
    catch (Exception ex)
    {
        return Results.Problem(detail: ex.Message, statusCode: 500);
    }
}).DisableAntiforgery();

app.MapPost("/api/compare/html", async (HttpRequest request) =>
{
    var form = await request.ReadFormAsync();
    var originalFile = form.Files.GetFile("original");
    var modifiedFile = form.Files.GetFile("modified");

    if (originalFile is null || modifiedFile is null)
        return Results.BadRequest(new { error = "Both 'original' and 'modified' .docx files are required." });

    var author = form["author"].FirstOrDefault() ?? "Docxodus";
    var detailThreshold = 0.0;
    if (form.ContainsKey("detailThreshold") &&
        double.TryParse(form["detailThreshold"], NumberStyles.Float, CultureInfo.InvariantCulture, out var dt))
        detailThreshold = dt;

    var settings = new WmlComparerSettings
    {
        AuthorForRevisions = author,
        DetailThreshold = detailThreshold,
        DetectMoves = true,
        DetectFormatChanges = true,
    };

    try
    {
        var originalBytes = await ReadFormFile(originalFile);
        var modifiedBytes = await ReadFormFile(modifiedFile);

        var originalDoc = new WmlDocument("original.docx", originalBytes);
        var modifiedDoc = new WmlDocument("modified.docx", modifiedBytes);

        var result = WmlComparer.Compare(originalDoc, modifiedDoc, settings);
        var htmlSettings = new WmlToHtmlConverterSettings
        {
            RenderTrackedChanges = true,
        };
        var html = WmlToHtmlConverter.ConvertToHtml(result, htmlSettings);

        return Results.Content(html.ToString(), "text/html");
    }
    catch (Exception ex)
    {
        return Results.Problem(detail: ex.Message, statusCode: 500);
    }
}).DisableAntiforgery();

// ──────────────────────────────────────────────
// Ticket API
// ──────────────────────────────────────────────

app.MapGet("/api/tickets", async (TicketDbContext db, string? status, int page = 1, int pageSize = 25) =>
{
    var query = db.Tickets.AsQueryable();

    if (!string.IsNullOrEmpty(status) && Enum.TryParse<TicketStatus>(status, true, out var s))
        query = query.Where(t => t.Status == s);

    var total = await query.CountAsync();
    var tickets = await query
        .OrderByDescending(t => t.CreatedAt)
        .Skip((page - 1) * pageSize)
        .Take(pageSize)
        .Select(t => new
        {
            t.Id,
            t.Title,
            t.Description,
            t.SubmitterEmail,
            Status = t.Status.ToString(),
            t.CreatedAt,
            t.UpdatedAt,
            t.OriginalFileName,
            t.ModifiedFileName,
            t.RevisionCount,
            t.ComparisonLog,
        })
        .ToListAsync();

    return Results.Ok(new { total, page, pageSize, tickets });
});

app.MapGet("/api/tickets/{id:int}", async (int id, TicketDbContext db) =>
{
    var t = await db.Tickets.FindAsync(id);
    if (t is null) return Results.NotFound();

    return Results.Ok(new
    {
        t.Id,
        t.Title,
        t.Description,
        t.SubmitterEmail,
        Status = t.Status.ToString(),
        t.CreatedAt,
        t.UpdatedAt,
        t.OriginalFileName,
        t.ModifiedFileName,
        t.RevisionCount,
        t.ComparisonLog,
    });
});

app.MapPost("/api/tickets", async (HttpRequest request, TicketDbContext db) =>
{
    var form = await request.ReadFormAsync();
    var originalFile = form.Files.GetFile("original");
    var modifiedFile = form.Files.GetFile("modified");

    if (originalFile is null || modifiedFile is null)
        return Results.BadRequest(new { error = "Both 'original' and 'modified' .docx files are required." });

    var title = form["title"].FirstOrDefault();
    var description = form["description"].FirstOrDefault();
    if (string.IsNullOrWhiteSpace(title))
        return Results.BadRequest(new { error = "'title' is required." });

    var ticketDir = Path.Combine(uploadsDir, Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(ticketDir);

    var originalPath = Path.Combine(ticketDir, "original.docx");
    var modifiedPath = Path.Combine(ticketDir, "modified.docx");
    await SaveFormFile(originalFile, originalPath);
    await SaveFormFile(modifiedFile, modifiedPath);

    // Run redline and store result
    string? redlinePath = null;
    string? comparisonLog = null;
    int? revisionCount = null;

    try
    {
        var originalBytes = await File.ReadAllBytesAsync(originalPath);
        var modifiedBytes = await File.ReadAllBytesAsync(modifiedPath);
        var originalDoc = new WmlDocument("original.docx", originalBytes);
        var modifiedDoc = new WmlDocument("modified.docx", modifiedBytes);

        var log = new ComparisonLog();
        var settings = new WmlComparerSettings
        {
            AuthorForRevisions = "Docxodus",
            DetailThreshold = 0,
            DetectMoves = true,
            DetectFormatChanges = true,
            Log = log,
        };

        var result = WmlComparer.Compare(originalDoc, modifiedDoc, settings);
        var revisions = WmlComparer.GetRevisions(result, settings);

        redlinePath = Path.Combine(ticketDir, "redline.docx");
        await File.WriteAllBytesAsync(redlinePath, result.DocumentByteArray);
        revisionCount = revisions.Count;

        if (log.HasWarnings || log.HasErrors)
        {
            var parts = new List<string>();
            foreach (var w in log.Warnings) parts.Add($"WARN: [{w.Code}] {w.Message}");
            foreach (var e in log.Errors) parts.Add($"ERROR: [{e.Code}] {e.Message}");
            comparisonLog = string.Join("\n", parts);
        }
    }
    catch (Exception ex)
    {
        comparisonLog = $"Comparison failed: {ex.Message}";
    }

    var ticket = new Ticket
    {
        Title = title,
        Description = description ?? string.Empty,
        SubmitterEmail = form["email"].FirstOrDefault(),
        OriginalFilePath = originalPath,
        OriginalFileName = originalFile.FileName,
        ModifiedFilePath = modifiedPath,
        ModifiedFileName = modifiedFile.FileName,
        RedlineFilePath = redlinePath,
        ComparisonLog = comparisonLog,
        RevisionCount = revisionCount,
    };

    db.Tickets.Add(ticket);
    await db.SaveChangesAsync();

    return Results.Created($"/api/tickets/{ticket.Id}", new
    {
        ticket.Id,
        ticket.Title,
        Status = ticket.Status.ToString(),
        ticket.RevisionCount,
        ticket.ComparisonLog,
    });
}).DisableAntiforgery();

app.MapPatch("/api/tickets/{id:int}", async (int id, HttpRequest request, TicketDbContext db) =>
{
    var ticket = await db.Tickets.FindAsync(id);
    if (ticket is null) return Results.NotFound();

    var body = await request.ReadFromJsonAsync<TicketUpdateDto>();
    if (body is null) return Results.BadRequest();

    if (body.Status is not null && Enum.TryParse<TicketStatus>(body.Status, true, out var s))
        ticket.Status = s;
    if (body.Title is not null)
        ticket.Title = body.Title;
    if (body.Description is not null)
        ticket.Description = body.Description;

    ticket.UpdatedAt = DateTime.UtcNow;
    await db.SaveChangesAsync();

    return Results.Ok(new { ticket.Id, Status = ticket.Status.ToString(), ticket.UpdatedAt });
});

app.MapGet("/api/tickets/{id:int}/files/{which}", async (int id, string which, TicketDbContext db) =>
{
    var ticket = await db.Tickets.FindAsync(id);
    if (ticket is null) return Results.NotFound();

    string? filePath = which.ToLowerInvariant() switch
    {
        "original" => ticket.OriginalFilePath,
        "modified" => ticket.ModifiedFilePath,
        "redline" => ticket.RedlineFilePath,
        _ => null,
    };

    if (filePath is null || !File.Exists(filePath))
        return Results.NotFound();

    var fileName = which.ToLowerInvariant() switch
    {
        "original" => ticket.OriginalFileName,
        "modified" => ticket.ModifiedFileName,
        "redline" => "redline.docx",
        _ => "file.docx",
    };

    return Results.File(filePath,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        fileName);
});

// SPA fallback — serve index.html for non-API, non-file routes
app.MapFallbackToFile("index.html");

app.Run();

// ──────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────

static async Task<byte[]> ReadFormFile(IFormFile file)
{
    using var ms = new MemoryStream();
    await file.CopyToAsync(ms);
    return ms.ToArray();
}

static async Task SaveFormFile(IFormFile file, string path)
{
    using var stream = File.Create(path);
    await file.CopyToAsync(stream);
}

record TicketUpdateDto(string? Status, string? Title, string? Description);
