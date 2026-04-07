using System;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Represents the outcome of an operation performed by <see cref="Executor"/>.
/// </summary>
public class Result
{
    /// <summary>
    /// Gets or sets the captured exception when the operation fails.
    /// This value is typically <see langword="null"/> when <see cref="Success"/> is <see langword="true"/>.
    /// </summary>
    public Exception? Exception { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the operation completed successfully.
    /// When this value is <see langword="false"/>, inspect <see cref="Exception"/> for the failure cause.
    /// </summary>
    public bool Success { get; set; }
}
