using System;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Immutable snapshot of thread-local COM rich error details captured from <c>IErrorInfo</c>
/// (<c>oleaut32!GetErrorInfo</c>).
/// </summary>
public sealed class ComErrorInfo
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ComErrorInfo"/> class.
    /// </summary>
    /// <param name="interfaceId">The interface identifier (<c>IErrorInfo::GetGUID</c>) that defined the error.</param>
    /// <param name="source">The ProgID or component source (<c>IErrorInfo::GetSource</c>).</param>
    /// <param name="description">The human-readable error description (<c>IErrorInfo::GetDescription</c>).</param>
    /// <param name="helpFile">Optional help file path (<c>IErrorInfo::GetHelpFile</c>).</param>
    /// <param name="helpContext">Optional help context identifier (<c>IErrorInfo::GetHelpContext</c>).</param>
    public ComErrorInfo(
        Guid interfaceId,
        string? source,
        string? description,
        string? helpFile,
        uint helpContext)
    {
        InterfaceId = interfaceId;
        Source = source;
        Description = description;
        HelpFile = helpFile;
        HelpContext = helpContext;
    }

    /// <summary>
    /// Gets the COM interface identifier (<c>IID</c>) associated with the error, or <see cref="Guid.Empty"/>.
    /// </summary>
    public Guid InterfaceId { get; }

    /// <summary>
    /// Gets the error source string (typically the server's ProgID or module name), if provided.
    /// </summary>
    public string? Source { get; }

    /// <summary>
    /// Gets the human-readable error description reported by the COM server, if provided.
    /// </summary>
    public string? Description { get; }

    /// <summary>
    /// Gets the help file path associated with the error, if provided.
    /// </summary>
    public string? HelpFile { get; }

    /// <summary>
    /// Gets the help topic context ID associated with <see cref="HelpFile"/>.
    /// </summary>
    public uint HelpContext { get; }

    /// <summary>
    /// Formats a concise diagnostic summary combining <see cref="Source"/> and <see cref="Description"/>,
    /// omitting empty or whitespace segments.
    /// </summary>
    /// <returns>The formatted summary string, or <see cref="string.Empty"/> when neither source nor description is set.</returns>
    public string ToSummaryString()
    {
        var hasSource = Source is not null && !string.IsNullOrWhiteSpace(Source);
        var hasDescription = Description is not null && !string.IsNullOrWhiteSpace(Description);
        if (hasSource && hasDescription)
        {
            return Source!.Trim() + ": " + Description!.Trim();
        }

        if (hasDescription)
        {
            return Description!.Trim();
        }

        return hasSource ? Source!.Trim() : string.Empty;
    }

    /// <inheritdoc />
    public override string ToString()
        => ToSummaryString();
}
