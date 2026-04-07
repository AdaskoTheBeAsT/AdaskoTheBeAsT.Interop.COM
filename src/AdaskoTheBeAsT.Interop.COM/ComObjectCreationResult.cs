namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Represents the outcome of creating a COM object together with the handle required to release it later.
/// </summary>
/// <typeparam name="T">The COM object type.</typeparam>
public sealed class ComObjectCreationResult<T> : Result
    where T : class
{
    /// <summary>
    /// Gets or sets the created COM object handle.
    /// This value is populated only when <see cref="Result.Success"/> is <see langword="true"/>; otherwise it is <see langword="null"/>.
    /// </summary>
    public ComObjectHandle<T>? Value { get; set; }
}
