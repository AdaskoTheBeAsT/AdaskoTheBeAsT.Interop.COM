using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Executes method of COM dll in Single Threaded Apartment.
/// </summary>
public static class Executor
{
    /// <summary>
    /// Executes an action within a single COM activation context.
    /// Creates activation context from manifest, executes the action in STA, then cleans up.
    /// </summary>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="action">Action to execute within the activation context.</param>
    /// <returns>Result indicating success or failure with optional exception details.</returns>
    /// <exception cref="ArgumentNullException">Thrown when comAssemblyPath or manifestPath is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comAssemblyPath or manifestPath is empty or whitespace.</exception>
    public static Result Execute(
        string comAssemblyPath,
        string manifestPath,
        Action action)
    {
        if (comAssemblyPath is null)
        {
            throw new ArgumentNullException(nameof(comAssemblyPath));
        }

        if (manifestPath is null)
        {
            throw new ArgumentNullException(nameof(manifestPath));
        }

        if (string.IsNullOrWhiteSpace(comAssemblyPath))
        {
            throw new ArgumentException("COM assembly path cannot be empty or whitespace.", nameof(comAssemblyPath));
        }

        if (string.IsNullOrWhiteSpace(manifestPath))
        {
            throw new ArgumentException("Manifest path cannot be empty or whitespace.", nameof(manifestPath));
        }

        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        var ctx = PrepareContext(descriptor);
        var hActCtx = CreateContext(ctx);

        var result = new Result { Success = false };
        try
        {
            var cookie = ActivateContext(hActCtx);

            try
            {
                action?.Invoke();

                // Pump COM messages in STA apartment
                NativeMethods.PumpPendingMessages();

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Exception = ex;
            }
            finally
            {
                NativeMethods.DeactivateActCtx(0, cookie);
            }
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            NativeMethods.ReleaseActCtx(hActCtx);
        }

        return result;
    }

    /// <summary>
    /// Executes an action within multiple COM activation contexts.
    /// Creates activation contexts from all descriptors, executes the action in STA,
    /// then cleans up all contexts in reverse order (LIFO).
    /// </summary>
    /// <param name="comPathDescriptors">Collection of COM path descriptors containing DLL and manifest paths.</param>
    /// <param name="action">Action to execute within all activation contexts.</param>
    /// <returns>Result indicating success or failure with optional exception details.</returns>
    /// <exception cref="ArgumentNullException">Thrown when comPathDescriptors is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comPathDescriptors is empty.</exception>
    public static Result Execute(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action)
    {
        ValidateComPathDescriptors(comPathDescriptors);
        var hActCtxs = CreateActivationContexts(comPathDescriptors);

        var result = new Result { Success = false };
        try
        {
            var cookies = new List<IntPtr>(hActCtxs.Count);
            foreach (var hActCtx in hActCtxs)
            {
                var cookie = ActivateContext(hActCtx);
                cookies.Add(cookie);
            }

            try
            {
                action?.Invoke();

                // Pump COM messages after callback
                NativeMethods.PumpPendingMessages();

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Exception = ex;
            }
            finally
            {
                for (int i = cookies.Count - 1; i >= 0; i--)
                {
                    NativeMethods.DeactivateActCtx(0, cookies[i]);
                }
            }
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            for (int i = hActCtxs.Count - 1; i >= 0; i--)
            {
                NativeMethods.ReleaseActCtx(hActCtxs[i]);
            }
        }

        return result;
    }

    private static void ValidateComPathDescriptors(ICollection<ComPathDescriptor> comPathDescriptors)
    {
        if (comPathDescriptors is null)
        {
            throw new ArgumentNullException(nameof(comPathDescriptors));
        }

        if (comPathDescriptors.Count == 0)
        {
            throw new ArgumentException("COM path descriptors collection cannot be empty.", nameof(comPathDescriptors));
        }
    }

    private static List<IntPtr> CreateActivationContexts(ICollection<ComPathDescriptor> comPathDescriptors)
    {
        var hActCtxs = new List<IntPtr>(comPathDescriptors.Count);

        foreach (var comPathDescriptor in comPathDescriptors)
        {
            var ac = PrepareContext(comPathDescriptor);
            var hActCtx = CreateContext(ac);
            hActCtxs.Add(hActCtx);
        }

        return hActCtxs;
    }

    private static ActCtx PrepareContext(ComPathDescriptor comPathDescriptor)
    {
        var ac = default(ActCtx);
        ac.cbSize = Marshal.SizeOf(typeof(ActCtx));
        var expected = IntPtr.Size == 4 ? 0x20 : 0x38;
        if (ac.cbSize != expected)
        {
            throw new ActCtxWrongSizeException("ActCtx.cbSize is wrong");
        }

        ac.lpAssemblyDirectory = comPathDescriptor.ComAssemblyPath;
        ac.lpSource = comPathDescriptor.ComManifestPath;
        ac.dwFlags = NativeMethods.ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID;
        return ac;
    }

    private static IntPtr CreateContext(ActCtx actCtx)
    {
        var hActCtx = NativeMethods.CreateActCtx(ref actCtx);
        if (hActCtx == (IntPtr)(-1))
        {
            throw new Win32Exception();
        }

        return hActCtx;
    }

    private static IntPtr ActivateContext(IntPtr hActCtx)
    {
        if (!NativeMethods.ActivateActCtx(hActCtx, out var cookie))
        {
            throw new Win32Exception();
        }

        return cookie;
    }
}
