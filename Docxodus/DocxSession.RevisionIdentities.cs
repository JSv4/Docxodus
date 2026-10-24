using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

public sealed partial class DocxSession
{
    private static readonly XNamespace PeopleNamespace = "http://schemas.microsoft.com/office/word/2012/wordml";

    // Configuration is immutable after capture. Ordinary Undo/Redo changes the package, not
    // this baseline; transactions and preview shadows retain it along with the account list.
    internal sealed record RevisionIdentityContext(
        IReadOnlyList<RevisionIdentity> Identities, IReadOnlySet<string> UsedAuthors);

    private RevisionIdentityContext? _revisionIdentityContext;

    /// <summary>The account identities configured for subsequent editing.</summary>
    public IReadOnlyList<RevisionIdentity> RevisionIdentities =>
        _revisionIdentityContext?.Identities ?? Array.Empty<RevisionIdentity>();

    /// <summary>
    /// Configure account-based revision ownership and remember which reviewer names are already
    /// in use. Null or empty restores display-name ownership. This changes session configuration,
    /// takes no undo step, and does not write the document's reviewer list until an edit needs it.
    /// </summary>
    public void SetRevisionIdentities(IReadOnlyList<RevisionIdentity>? identities)
    {
        if (_disposed) return;
        if (identities is null || identities.Count == 0)
        {
            _revisionIdentityContext = null;
            return;
        }
        var copy = identities.ToArray();
        if (copy.Any(identity => identity is null || string.IsNullOrWhiteSpace(identity.ProviderId)
            || string.IsNullOrWhiteSpace(identity.UserId) || !Internal.XmlText.IsValid(identity.ProviderId)
            || !Internal.XmlText.IsValid(identity.UserId)))
            throw new ArgumentException("revision identities need nonblank XML-compatible provider and user identifiers", nameof(identities));
        _revisionIdentityContext = new RevisionIdentityContext(Array.AsReadOnly(copy), UsedRevisionAuthorNames());
    }

    /// <summary>
    /// Query a proposed display name, or the current revision author when omitted. With accounts
    /// configured, a name already used without a matching account is taken. The host can choose
    /// a distinguishing name before recording a revision or comment.
    /// </summary>
    public RevisionAuthorNameStatus GetRevisionAuthorNameStatus(string? author = null)
    {
        ThrowIfDisposed();
        var name = author ?? _revisionAuthor ?? "docxodus";
        if (RecordedPeople(name) is { Count: > 0 } people && _revisionIdentityContext is not null)
            return PeopleAreOwn(people) ? RevisionAuthorNameStatus.Own : RevisionAuthorNameStatus.Taken;
        var used = UsedRevisionAuthorNames().Contains(name);
        if (_revisionIdentityContext is null)
            return !used ? RevisionAuthorNameStatus.Free
                : name == (_revisionAuthor ?? "docxodus") ? RevisionAuthorNameStatus.Own : RevisionAuthorNameStatus.Taken;
        if (_revisionIdentityContext.UsedAuthors.Contains(name)) return RevisionAuthorNameStatus.Taken;
        return used ? RevisionAuthorNameStatus.Own : RevisionAuthorNameStatus.Free;
    }

    private bool IsOwnRevisionAuthor(string? recorded, string author)
    {
        if (_revisionIdentityContext is null)
            return string.Equals(recorded, author, StringComparison.Ordinal);
        if (recorded is null) return false;
        var people = RecordedPeople(recorded);
        if (people.Count > 0) return PeopleAreOwn(people);
        return string.Equals(recorded, author, StringComparison.Ordinal)
            && string.Equals(author, _revisionAuthor ?? "docxodus", StringComparison.Ordinal)
            && !_revisionIdentityContext.UsedAuthors.Contains(recorded);
    }

    private List<XElement> RecordedPeople(string author) =>
        _doc!.MainDocumentPart?.WordprocessingPeoplePart?.GetXDocument().Root?
            .Elements(PeopleNamespace + "person")
            .Where(person => (string?)person.Attribute(PeopleNamespace + "author") == author).ToList() ?? new();

    private bool PeopleAreOwn(IReadOnlyList<XElement> people) => people.All(person =>
    {
        var presence = person.Elements(PeopleNamespace + "presenceInfo").ToList();
        if (presence.Count != 1) return false;
        var provider = (string?)presence[0].Attribute(PeopleNamespace + "providerId");
        var user = (string?)presence[0].Attribute(PeopleNamespace + "userId");
        return _revisionIdentityContext!.Identities.Any(identity => IdentityMatches(identity, provider, user));
    });

    private static bool IdentityMatches(RevisionIdentity identity, string? provider, string? user)
    {
        if (identity.ProviderId != provider || string.IsNullOrWhiteSpace(user)) return false;
        if (identity.UserId == user) return true;
        if (provider != "AD") return false;
        var left = identity.UserId.Split(new[] { "::" }, StringSplitOptions.None);
        var right = user.Split(new[] { "::" }, StringSplitOptions.None);
        if (left.Length != 3 || right.Length != 3 || left[0] != "S" || right[0] != "S") return false;
        return (!string.IsNullOrWhiteSpace(left[1]) && !string.IsNullOrWhiteSpace(right[1])
                && string.Equals(left[1], right[1], StringComparison.OrdinalIgnoreCase))
            || (!string.IsNullOrWhiteSpace(left[2]) && !string.IsNullOrWhiteSpace(right[2])
                && string.Equals(left[2], right[2], StringComparison.OrdinalIgnoreCase));
    }

    private HashSet<string> UsedRevisionAuthorNames()
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var part in _doc!.GetAllParts().Where(part => part is VmlDrawingPart
            || part.ContentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)
            || part.ContentType is "application/xml" or "text/xml"))
        {
            if (part.GetXDocument().Root is not { } root) continue;
            foreach (var attribute in root.DescendantsAndSelf().Attributes()
                .Where(attribute => attribute.Name == W.author || attribute.Name == PeopleNamespace + "author"))
                names.Add(attribute.Value);
        }
        return names;
    }

    private void RecordRevisionIdentity(string author)
    {
        if (_revisionIdentityContext is null || author != (_revisionAuthor ?? "docxodus")
            || _revisionIdentityContext.UsedAuthors.Contains(author) || RecordedPeople(author).Count > 0) return;
        var main = _doc!.MainDocumentPart!;
        var part = main.WordprocessingPeoplePart ?? main.AddNewPart<WordprocessingPeoplePart>();
        var document = part.GetXDocument();
        if (document.Root is null)
            document.Add(new XElement(PeopleNamespace + "people", new XAttribute(XNamespace.Xmlns + "w15", PeopleNamespace)));
        var identity = _revisionIdentityContext.Identities[0];
        document.Root!.Add(new XElement(PeopleNamespace + "person",
            new XAttribute(PeopleNamespace + "author", author), new XElement(PeopleNamespace + "presenceInfo",
                new XAttribute(PeopleNamespace + "providerId", identity.ProviderId),
                new XAttribute(PeopleNamespace + "userId", identity.UserId))));
        part.PutXDocument();
    }

    private static bool SameRevisionIdentityContext(RevisionIdentityContext? left, RevisionIdentityContext? right) =>
        ReferenceEquals(left, right) || (left is not null && right is not null
            && left.Identities.SequenceEqual(right.Identities) && left.UsedAuthors.SetEquals(right.UsedAuthors));
}
