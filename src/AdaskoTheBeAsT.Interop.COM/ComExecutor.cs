using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Default, recommended <see cref="IComExecutor"/> implementation. Delegates every call to the static
/// <see cref="Executor"/> with an immutable message-pumping policy, so it is safe to register as a singleton.
/// </summary>
/// <remarks>
/// <para>
/// This is the preferred entry point in v3.0+. Depend on <see cref="IComExecutor"/> in your classes and
/// register <see cref="ComExecutor"/> once in the composition root:
/// <c>services.AddSingleton&lt;IComExecutor, ComExecutor&gt;();</c>.
/// </para>
/// <para>
/// The static <see cref="Executor"/> class remains available for backward compatibility with v2.x callers.
/// </para>
/// </remarks>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public sealed class ComExecutor : IComExecutor
{
    private readonly bool _pumpPendingMessages;

    /// <summary>
    /// Initializes a new instance of the <see cref="ComExecutor"/> class with automatic message pumping enabled.
    /// </summary>
    public ComExecutor()
        : this(pumpPendingMessages: true)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ComExecutor"/> class with a fixed message-pumping policy.
    /// </summary>
    /// <param name="pumpPendingMessages">
    /// Whether to pump pending messages after successful callbacks and during handle release.
    /// Set to <see langword="false"/> when the host owns message processing.
    /// </param>
    /// <remarks>
    /// The built-in pump does not run host-specific thread-message handlers, accelerators, or dialog preprocessing.
    /// Created handles retain this policy, including when disposed or freed through another executor.
    /// </remarks>
    public ComExecutor(bool pumpPendingMessages)
    {
        _pumpPendingMessages = pumpPendingMessages;
    }

    /// <inheritdoc />
    public Result Execute(string comAssemblyPath, string manifestPath, Action action)
        => Executor.Execute(comAssemblyPath, manifestPath, action, _pumpPendingMessages);

    /// <inheritdoc />
    public Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action)
        => Executor.Execute(comPathDescriptors, action, _pumpPendingMessages);

    /// <inheritdoc />
    public ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class
        => Executor.Create(comAssemblyPath, manifestPath, factory, _pumpPendingMessages);

    /// <inheritdoc />
    public ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)
        where T : class
        => Executor.Create(comPathDescriptors, factory, _pumpPendingMessages);

    /// <inheritdoc />
    public Result Free<T>(ComObjectHandle<T> comObjectHandle)
        where T : class
        => Executor.Free(comObjectHandle);
}
