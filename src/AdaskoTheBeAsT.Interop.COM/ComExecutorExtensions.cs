using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Extension methods that expose <see cref="IComActivationScope"/> and <see cref="ComActivationOptions"/>
/// capabilities on <see cref="IComExecutor"/> without adding breaking abstract members to <see cref="IComExecutor"/>.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class ComExecutorExtensions
{
    /// <summary>
    /// Activates a registration-free COM scope on the calling thread using <paramref name="executor"/>.
    /// </summary>
    /// <param name="executor">The executor instance.</param>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the scope activation.</returns>
    public static ComCreateResult<IComActivationScope> Activate(
        this IComExecutor executor,
        string comAssemblyPath,
        string manifestPath,
        ComActivationOptions? options = null)
    {
        ThrowHelper.ThrowIfNull(executor, nameof(executor));
        return executor is IComActivationExecutor activationExecutor
            ? activationExecutor.Activate(comAssemblyPath, manifestPath, options)
            : Executor.Activate(comAssemblyPath, manifestPath, options);
    }

    /// <summary>
    /// Activates multiple registration-free COM contexts as a single scope on the calling thread using <paramref name="executor"/>.
    /// </summary>
    /// <param name="executor">The executor instance.</param>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the scope activation.</returns>
    public static ComCreateResult<IComActivationScope> Activate(
        this IComExecutor executor,
        ICollection<ComPathDescriptor> comPathDescriptors,
        ComActivationOptions? options = null)
    {
        ThrowHelper.ThrowIfNull(executor, nameof(executor));
        return executor is IComActivationExecutor activationExecutor
            ? activationExecutor.Activate(comPathDescriptors, options)
            : Executor.Activate(comPathDescriptors, options);
    }

    /// <summary>
    /// Deactivates an <see cref="IComActivationScope"/> using <paramref name="executor"/>.
    /// </summary>
    /// <param name="executor">The executor instance.</param>
    /// <param name="activationScope">The scope to deactivate.</param>
    /// <returns>The outcome of the deactivation.</returns>
    public static Result Deactivate(this IComExecutor executor, IComActivationScope activationScope)
    {
        ThrowHelper.ThrowIfNull(executor, nameof(executor));
        return executor is IComActivationExecutor activationExecutor
            ? activationExecutor.Deactivate(activationScope)
            : Executor.Deactivate(activationScope);
    }

    /// <summary>
    /// Executes <paramref name="action"/> inside a single registration-free COM context with explicit <paramref name="options"/>.
    /// </summary>
    /// <param name="executor">The executor instance.</param>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file.</param>
    /// <param name="action">Callback to invoke inside the active context.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the execution.</returns>
    public static Result Execute(
        this IComExecutor executor,
        string comAssemblyPath,
        string manifestPath,
        Action action,
        ComActivationOptions? options)
    {
        ThrowHelper.ThrowIfNull(executor, nameof(executor));
        return executor is IComActivationExecutor activationExecutor
            ? activationExecutor.Execute(comAssemblyPath, manifestPath, action, options)
            : Executor.Execute(comAssemblyPath, manifestPath, action, options);
    }

    /// <summary>
    /// Executes <paramref name="action"/> inside multiple registration-free COM contexts with explicit <paramref name="options"/>.
    /// </summary>
    /// <param name="executor">The executor instance.</param>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="action">Callback to invoke inside the active contexts.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the execution.</returns>
    public static Result Execute(
        this IComExecutor executor,
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action,
        ComActivationOptions? options)
    {
        ThrowHelper.ThrowIfNull(executor, nameof(executor));
        return executor is IComActivationExecutor activationExecutor
            ? activationExecutor.Execute(comPathDescriptors, action, options)
            : Executor.Execute(comPathDescriptors, action, options);
    }

    /// <summary>
    /// Creates a COM object inside a single registration-free COM context with explicit <paramref name="options"/>.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="executor">The executor instance.</param>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file.</param>
    /// <param name="factory">Factory creating the COM object.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the creation.</returns>
    public static ComObjectCreationResult<T> Create<T>(
        this IComExecutor executor,
        string comAssemblyPath,
        string manifestPath,
        Func<T> factory,
        ComActivationOptions? options)
        where T : class
    {
        ThrowHelper.ThrowIfNull(executor, nameof(executor));
        return executor is IComActivationExecutor activationExecutor
            ? activationExecutor.Create(comAssemblyPath, manifestPath, factory, options)
            : Executor.Create(comAssemblyPath, manifestPath, factory, options);
    }

    /// <summary>
    /// Creates a COM object inside multiple registration-free COM contexts with explicit <paramref name="options"/>.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="executor">The executor instance.</param>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="factory">Factory creating the COM object.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the creation.</returns>
    public static ComObjectCreationResult<T> Create<T>(
        this IComExecutor executor,
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory,
        ComActivationOptions? options)
        where T : class
    {
        ThrowHelper.ThrowIfNull(executor, nameof(executor));
        return executor is IComActivationExecutor activationExecutor
            ? activationExecutor.Create(comPathDescriptors, factory, options)
            : Executor.Create(comPathDescriptors, factory, options);
    }
}
