using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Default, recommended <see cref="IComExecutor"/> implementation. Delegates every call to the static
/// <see cref="Executor"/> and carries no state, so it is safe to register as a singleton.
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
    /// <inheritdoc />
    public Result Execute(string comAssemblyPath, string manifestPath, Action action)
        => Executor.Execute(comAssemblyPath, manifestPath, action);

    /// <inheritdoc />
    public Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action)
        => Executor.Execute(comPathDescriptors, action);

    /// <inheritdoc />
    public ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class
        => Executor.Create(comAssemblyPath, manifestPath, factory);

    /// <inheritdoc />
    public ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)
        where T : class
        => Executor.Create(comPathDescriptors, factory);

    /// <inheritdoc />
    public Result Free<T>(ComObjectHandle<T> comObjectHandle)
        where T : class
        => Executor.Free(comObjectHandle);
}
