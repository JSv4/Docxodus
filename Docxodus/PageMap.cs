#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace Docxodus;

/// <summary>The layout substrate that produced a <see cref="PageMap"/>.</summary>
public enum PageMapMode
{
    /// <summary>Fixed page boxes exist and geometry is authoritative.</summary>
    Paginated,

    /// <summary>No fixed pages exist. Citations are deliberately unavailable.</summary>
    Continuous,
}

/// <summary>Whether a materialized map can answer page-citation requests.</summary>
public enum PageMapAvailability
{
    Unavailable,
    Available,
}

/// <summary>The OOXML story that owns a rendered anchor fragment.</summary>
public enum PageMapStory
{
    Body,
    Header,
    Footer,
    Footnote,
    Endnote,
    Comment,
}

/// <summary>A rectangle in page-relative points, with the page's top-left as (0, 0).</summary>
public sealed record PageMapRect(double X, double Y, double Width, double Height);

/// <summary>Physical page-box geometry. Page numbers are 1-based and document-global.</summary>
public sealed record PageMapPage
{
    required public int PageNumber { get; init; }
    required public int PageInSection { get; init; }
    required public double Width { get; init; }
    required public double Height { get; init; }
    public int? SectionIndex { get; init; }
    /// <summary>The renderer's stable page-style identity (for example a named CSS @page rule).</summary>
    required public string PageName { get; init; }
}

/// <summary>
/// One visible piece of one source anchor. A paragraph, table, or note split across pages has
/// multiple fragments with the same <see cref="AnchorId"/> and distinct page-qualified
/// <see cref="FragmentId"/> values.
/// </summary>
public sealed record PageMapFragment
{
    required public string FragmentId { get; init; }
    required public string AnchorId { get; init; }
    required public int FragmentIndex { get; init; }
    required public int PageNumber { get; init; }
    required public PageMapRect Geometry { get; init; }
    required public PageMapStory Story { get; init; }

    /// <summary>True when the fragment belongs to an addressable table cell.</summary>
    public bool InTableCell { get; init; }
}

/// <summary>
/// Portable output of a paginated renderer. Core Docxodus validates and consumes this contract;
/// it never invents page numbers. Version 1 uses page-relative point geometry.
/// </summary>
public sealed record PageMap
{
    public const int CurrentSchemaVersion = 1;

    public int SchemaVersion { get; init; } = CurrentSchemaVersion;
    required public PageMapMode Mode { get; init; }
    required public PageMapAvailability Availability { get; init; }
    required public long DocumentVersion { get; init; }
    required public string RendererFingerprint { get; init; }
    public IReadOnlyList<PageMapPage> Pages { get; init; } = Array.Empty<PageMapPage>();
    public IReadOnlyList<PageMapFragment> Fragments { get; init; } = Array.Empty<PageMapFragment>();
}

/// <summary>Why a PageMap registration was rejected.</summary>
public enum PageMapRegistrationError
{
    UnsupportedSchemaVersion,
    StaleDocumentVersion,
    RendererFingerprintMismatch,
    InvalidMap,
}

/// <summary>Typed result of <see cref="DocxSession.RegisterPageMap"/>.</summary>
public sealed record PageMapRegistrationResult
{
    required public bool Success { get; init; }
    public PageMapRegistrationError? Error { get; init; }
    public string? Message { get; init; }
}

/// <summary>
/// Identifies the exact rendered layout a read wants citations from. Both fields are mandatory:
/// accepting "the latest" map would make concurrent render/edit workflows ambiguous.
/// </summary>
public sealed record PageCitationRequest(long DocumentVersion, string RendererFingerprint);

/// <summary>Why an optional page citation could not be supplied.</summary>
public enum PageCitationUnavailableReason
{
    NoPageMap,
    ContinuousMode,
    StaleDocumentVersion,
    RendererFingerprintMismatch,
    AnchorNotMapped,
}

/// <summary>Explicit citation result for one source anchor.</summary>
public sealed record PageCitation
{
    required public string AnchorId { get; init; }
    required public PageMapAvailability Availability { get; init; }
    public PageCitationUnavailableReason? UnavailableReason { get; init; }
    required public long DocumentVersion { get; init; }
    required public string RendererFingerprint { get; init; }
    /// <summary>Physical descriptors for the pages referenced by <see cref="Fragments"/>.</summary>
    public IReadOnlyList<PageMapPage> Pages { get; init; } = Array.Empty<PageMapPage>();
    public IReadOnlyList<PageMapFragment> Fragments { get; init; } = Array.Empty<PageMapFragment>();
}

/// <summary>Session-level state of the registered layout map.</summary>
public sealed record PageMapStatus
{
    required public PageMapAvailability Availability { get; init; }
    public PageCitationUnavailableReason? UnavailableReason { get; init; }
    required public long DocumentVersion { get; init; }
    public string? RendererFingerprint { get; init; }
    public PageMapMode? Mode { get; init; }
}
