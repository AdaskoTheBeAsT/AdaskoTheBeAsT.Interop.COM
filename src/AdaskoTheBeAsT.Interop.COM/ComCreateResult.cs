namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Represents the outcome of creating a lifetime-scoped COM resource such as <see cref="IComActivationScope"/>.
/// </summary>
/// <typeparam name="T">The created resource type.</typeparam>
public sealed class ComCreateResult<T> : Result
    where T : class
{
    /// <summary>
    /// Gets or sets the created resource instance when <see cref="Result.Success"/> is <see langword="true"/>;
    /// otherwise <see langword="null"/>.
    /// </summary>
    public T? Value { get; set; }
}
