using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Default, recommended <see cref="IComExecutor"/> and <see cref="IComActivationExecutor"/> implementation.
/// Delegates every call to the static <see cref="Executor"/> with an immutable default policy, so it is safe
/// to register as a singleton.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public sealed class ComExecutor : IComActivationExecutor
{
    private readonly ComActivationOptions _defaultOptions;

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
    /// Whether to pump pending messages after successful callbacks and during handle or scope release.
    /// Set to <see langword="false"/> when the host owns message processing.
    /// </param>
    public ComExecutor(bool pumpPendingMessages)
    {
        _defaultOptions = new ComActivationOptions
        {
            PumpPendingMessages = pumpPendingMessages,
        };
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ComExecutor"/> class with default <paramref name="defaultOptions"/>
    /// applied whenever per-call options are omitted.
    /// </summary>
    /// <param name="defaultOptions">Default activation options cloned into this executor instance.</param>
    public ComExecutor(ComActivationOptions defaultOptions)
    {
        ThrowHelper.ThrowIfNull(defaultOptions, nameof(defaultOptions));
        _defaultOptions = defaultOptions.Clone();
    }

    /// <inheritdoc />
    public Result Execute(string comAssemblyPath, string manifestPath, Action action)
        => Executor.Execute(comAssemblyPath, manifestPath, action, _defaultOptions);

    /// <inheritdoc />
    public Result Execute(string comAssemblyPath, string manifestPath, Action action, ComActivationOptions? options)
        => Executor.Execute(comAssemblyPath, manifestPath, action, options ?? _defaultOptions);

    /// <inheritdoc />
    public Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action)
        => Executor.Execute(comPathDescriptors, action, _defaultOptions);

    /// <inheritdoc />
    public Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action, ComActivationOptions? options)
        => Executor.Execute(comPathDescriptors, action, options ?? _defaultOptions);

    /// <inheritdoc />
    public ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class
        => Executor.Create(comAssemblyPath, manifestPath, factory, _defaultOptions);

    /// <inheritdoc />
    public ComObjectCreationResult<T> Create<T>(
        string comAssemblyPath,
        string manifestPath,
        Func<T> factory,
        ComActivationOptions? options)
        where T : class
        => Executor.Create(comAssemblyPath, manifestPath, factory, options ?? _defaultOptions);

    /// <inheritdoc />
    public ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)
        where T : class
        => Executor.Create(comPathDescriptors, factory, _defaultOptions);

    /// <inheritdoc />
    public ComObjectCreationResult<T> Create<T>(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory,
        ComActivationOptions? options)
        where T : class
        => Executor.Create(comPathDescriptors, factory, options ?? _defaultOptions);

    /// <inheritdoc />
    public ComCreateResult<IComActivationScope> Activate(
        string comAssemblyPath,
        string manifestPath,
        ComActivationOptions? options = null)
        => Executor.Activate(comAssemblyPath, manifestPath, options ?? _defaultOptions);

    /// <inheritdoc />
    public ComCreateResult<IComActivationScope> Activate(
        ICollection<ComPathDescriptor> comPathDescriptors,
        ComActivationOptions? options = null)
        => Executor.Activate(comPathDescriptors, options ?? _defaultOptions);

    /// <inheritdoc />
    public Result Deactivate(IComActivationScope activationScope)
        => Executor.Deactivate(activationScope);

    /// <inheritdoc />
    public Result Free<T>(ComObjectHandle<T> comObjectHandle)
        where T : class
        => Executor.Free(comObjectHandle);
}
