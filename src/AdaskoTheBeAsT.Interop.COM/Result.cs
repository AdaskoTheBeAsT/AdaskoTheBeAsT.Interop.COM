using System;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Represents the result of a COM execution operation.
/// </summary>
public sealed class Result
{
    /// <summary>
    /// Gets or sets the exception that occurred during COM execution, if any.
    /// </summary>
    public Exception? Exception { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the COM execution was successful.
    /// </summary>
    public bool Success { get; set; }
}
