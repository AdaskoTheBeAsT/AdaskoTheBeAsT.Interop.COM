using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Primary entry point of the library. Abstraction over the static <see cref="Executor"/> helpers that allows
/// dependency injection, unit-test substitution, and composition with other services.
/// </summary>
/// <remarks>
/// <para>
/// New code should depend on this interface rather than on the static <see cref="Executor"/> class.
/// The default implementation <see cref="ComExecutor"/> is stateless and should be registered as a singleton.
/// </para>
/// </remarks>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public interface IComExecutor
{
    /// <summary>
    /// Activates a single registration-free COM context and executes <paramref name="action"/> inside it.
    /// </summary>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="action">Action to execute within the activation context.</param>
    /// <returns>The outcome of the operation.</returns>
    Result Execute(string comAssemblyPath, string manifestPath, Action action);

    /// <summary>
    /// Activates multiple registration-free COM contexts and executes <paramref name="action"/> inside them.
    /// </summary>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="action">Action to execute within all activation contexts.</param>
    /// <returns>The outcome of the operation.</returns>
    Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action);

    /// <summary>
    /// Creates a COM object that keeps its activation context alive until it is released.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="factory">Factory that creates the COM object while the activation context is active.</param>
    /// <returns>The outcome of the creation.</returns>
    ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class;

    /// <summary>
    /// Creates a COM object inside multiple activation contexts that stay alive until the object is released.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="factory">Factory that creates the COM object while the activation contexts are active.</param>
    /// <returns>The outcome of the creation.</returns>
    ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)
        where T : class;

    /// <summary>
    /// Releases a COM object handle previously created by one of the <c>Create</c> overloads.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comObjectHandle">The handle to release.</param>
    /// <returns>The outcome of the release.</returns>
    Result Free<T>(ComObjectHandle<T> comObjectHandle)
        where T : class;
}
