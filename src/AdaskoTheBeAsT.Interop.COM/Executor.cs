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
    public static Result Execute(
        string comAssemblyPath,
        string manifestPath,
        Action action)
    {
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

    public static Result Execute(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action)
    {
        var hActCtxs = new List<IntPtr>(comPathDescriptors.Count);

        foreach (var comPathDescriptor in comPathDescriptors)
        {
            var ac = PrepareContext(comPathDescriptor);
            var hActCtx = CreateContext(ac);
            hActCtxs.Add(hActCtx);
        }

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
