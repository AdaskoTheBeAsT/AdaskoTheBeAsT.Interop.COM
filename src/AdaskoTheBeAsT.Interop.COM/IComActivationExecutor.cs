using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Extends <see cref="IComExecutor"/> with explicit <see cref="IComActivationScope"/> management and
/// <see cref="ComActivationOptions"/>-aware execution overloads while keeping <see cref="IComExecutor"/>
/// source- and binary-compatible for existing implementations and decorators.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public interface IComActivationExecutor : IComExecutor
{
    /// <summary>
    /// Activates a single registration-free COM context with explicit <paramref name="options"/> and executes <paramref name="action"/>.
    /// </summary>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="action">Action to execute within the activation context.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the operation.</returns>
    Result Execute(string comAssemblyPath, string manifestPath, Action action, ComActivationOptions? options);

    /// <summary>
    /// Activates multiple registration-free COM contexts with explicit <paramref name="options"/> and executes <paramref name="action"/>.
    /// </summary>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="action">Action to execute within all activation contexts.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the operation.</returns>
    Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action, ComActivationOptions? options);

    /// <summary>
    /// Creates a COM object with explicit <paramref name="options"/> that keeps its activation context alive until released.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="factory">Factory that creates the COM object while the activation context is active.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the creation.</returns>
    ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory, ComActivationOptions? options)
        where T : class;

    /// <summary>
    /// Creates a COM object inside multiple activation contexts with explicit <paramref name="options"/>.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="factory">Factory that creates the COM object while the activation contexts are active.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the creation.</returns>
    ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory, ComActivationOptions? options)
        where T : class;

    /// <summary>
    /// Activates a registration-free COM scope on the calling thread for multi-coclass sessions.
    /// </summary>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the scope activation.</returns>
    ComCreateResult<IComActivationScope> Activate(
        string comAssemblyPath,
        string manifestPath,
        ComActivationOptions? options = null);

    /// <summary>
    /// Activates multiple registration-free COM contexts as a single LIFO scope on the calling thread.
    /// </summary>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the scope activation.</returns>
    ComCreateResult<IComActivationScope> Activate(
        ICollection<ComPathDescriptor> comPathDescriptors,
        ComActivationOptions? options = null);

    /// <summary>
    /// Deactivates an <see cref="IComActivationScope"/> previously returned by <c>Activate</c>,
    /// returning a structured <see cref="Result"/> without throwing on wrong-thread or non-LIFO order.
    /// </summary>
    /// <param name="activationScope">The scope to deactivate.</param>
    /// <returns>The outcome of the deactivation.</returns>
    Result Deactivate(IComActivationScope activationScope);
}
