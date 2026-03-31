#nullable enable

using Microsoft.EntityFrameworkCore;

namespace DocxodusWeb.Data;

public class TicketDbContext : DbContext
{
    public TicketDbContext(DbContextOptions<TicketDbContext> options) : base(options) { }

    public DbSet<Ticket> Tickets => Set<Ticket>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Ticket>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.Id).ValueGeneratedOnAdd();
            entity.Property(e => e.Status).HasConversion<string>();
        });
    }
}

public class Ticket
{
    public int Id { get; set; }
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string? SubmitterEmail { get; set; }
    public TicketStatus Status { get; set; } = TicketStatus.Open;
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;

    /// <summary>Path to the original .docx file on disk.</summary>
    public string OriginalFilePath { get; set; } = string.Empty;
    public string OriginalFileName { get; set; } = string.Empty;

    /// <summary>Path to the modified .docx file on disk.</summary>
    public string ModifiedFilePath { get; set; } = string.Empty;
    public string ModifiedFileName { get; set; } = string.Empty;

    /// <summary>Path to the redline .docx produced by Docxodus (generated on submission).</summary>
    public string? RedlineFilePath { get; set; }

    /// <summary>Comparison log warnings/errors, if any.</summary>
    public string? ComparisonLog { get; set; }

    /// <summary>Number of revisions detected in the redline.</summary>
    public int? RevisionCount { get; set; }
}

public enum TicketStatus
{
    Open,
    InProgress,
    Resolved,
    WontFix,
    Duplicate
}
