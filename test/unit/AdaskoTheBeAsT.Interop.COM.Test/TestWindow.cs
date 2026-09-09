using System.ComponentModel;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace AdaskoTheBeAsT.Interop.COM.Test;

// Delegate-based window-class registration is intentionally shared by all target frameworks.
#pragma warning disable SYSLIB1054
#pragma warning disable CA2216 // Windows must be destroyed on their creating thread, not the finalizer thread.
internal sealed class TestWindow : IDisposable
{
    internal const uint TestMessage = 0x8001;
    internal const uint CharacterMessage = 0x0102;
    private readonly string _className = "ComTestWindow" + Guid.NewGuid().ToString("N");
    private readonly WindowProcedure _procedure;
    private readonly IntPtr _instance = NativeMethods.GetModuleHandle(null);
    private IntPtr _window;

    internal TestWindow()
    {
        _procedure = ProcessMessage;
        var windowClass = new WindowClass
        {
            Procedure = _procedure,
            Instance = _instance,
            ClassName = _className,
        };
        if (NativeMethods.RegisterClass(ref windowClass) == 0)
        {
            throw new Win32Exception();
        }

        _window = NativeMethods.CreateWindowEx(0, _className, string.Empty, 0, 0, 0, 0, 0, new IntPtr(-3), IntPtr.Zero, _instance, IntPtr.Zero);
        if (_window == IntPtr.Zero)
        {
            var error = new Win32Exception();
            _ = NativeMethods.UnregisterClass(_className, _instance);
            throw error;
        }
    }

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate IntPtr WindowProcedure(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);

    internal int DeliveryCount { get; private set; }

    internal UIntPtr LastWParam { get; private set; }

    internal IntPtr LastLParam { get; private set; }

    internal Action? OnDelivery { get; set; }

    public void Dispose()
    {
        if (_window != IntPtr.Zero)
        {
            _ = NativeMethods.DestroyWindow(_window);
            _window = IntPtr.Zero;
            _ = NativeMethods.UnregisterClass(_className, _instance);
        }

        GC.KeepAlive(_procedure);
    }

    internal static void OnStaThread(Action action)
    {
        ExceptionDispatchInfo? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                action?.Invoke();
            }
            catch (Exception ex)
            {
                failure = ExceptionDispatchInfo.Capture(ex);
            }
        })
        {
            IsBackground = true,
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        if (!thread.Join(TimeSpan.FromSeconds(30)))
        {
            throw new TimeoutException("The test STA thread did not finish.");
        }

        failure?.Throw();
    }

    internal void Post(uint message = TestMessage, uint wParam = 0, int lParam = 0)
    {
        if (!NativeMethods.PostMessage(_window, message, new UIntPtr(wParam), new IntPtr(lParam)))
        {
            throw new Win32Exception();
        }
    }

    private IntPtr ProcessMessage(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam)
    {
        if (message is TestMessage or CharacterMessage)
        {
            DeliveryCount++;
            LastWParam = wParam;
            LastLParam = lParam;
            OnDelivery?.Invoke();
            return IntPtr.Zero;
        }

        return NativeMethods.DefWindowProc(window, message, wParam, lParam);
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WindowClass
    {
        public uint Style;
        public WindowProcedure Procedure;
        public int ClassExtra;
        public int WindowExtra;
        public IntPtr Instance;
        public IntPtr Icon;
        public IntPtr Cursor;
        public IntPtr Background;
        public string? MenuName;
        public string ClassName;
    }

    private static class NativeMethods
    {
        [DllImport("kernel32.dll", EntryPoint = "GetModuleHandleW", CharSet = CharSet.Unicode, ExactSpelling = true)]
        internal static extern IntPtr GetModuleHandle(string? moduleName);

        [DllImport("user32.dll", EntryPoint = "RegisterClassW", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        internal static extern ushort RegisterClass(ref WindowClass windowClass);

        [DllImport("user32.dll", EntryPoint = "CreateWindowExW", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        internal static extern IntPtr CreateWindowEx(uint extendedStyle, string className, string windowName, uint style, int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr parameter);

        [DllImport("user32.dll", EntryPoint = "DefWindowProcW", ExactSpelling = true)]
        internal static extern IntPtr DefWindowProc(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "PostMessageW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool PostMessage(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DestroyWindow(IntPtr window);

        [DllImport("user32.dll", EntryPoint = "UnregisterClassW", CharSet = CharSet.Unicode, ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool UnregisterClass(string className, IntPtr instance);
    }
}
#pragma warning restore CA2216
#pragma warning restore SYSLIB1054
