using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable MA0158 // Use System.Threading.Lock
#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable S101
[SuppressUnmanagedCodeSecurity]
internal static class NanoServerWindowManager
{
    private const uint PmRemove = 0x0001;
    private const uint WmQuit = 0x0012;
    private const int GwlpWndProc = -4;
    private const int GwlpHInstance = -6;
    private const int GwlpId = -12;
    private const int GwlStyle = -16;
    private const int GwlExStyle = -20;
    private const int GwlpUserData = -21;
    private const int UoiName = 2;
    private const int ErrorInvalidWindowHandle = 1400;
    private const int ErrorClassDoesNotExist = 1411;

    private static readonly IntPtr FakeWinSta = new(0x1001);
    private static readonly IntPtr FakeDesktop = new(0x1002);

    private static readonly object ClassLock = new();
    private static readonly Dictionary<ushort, WindowClassEntry> ClassesByAtom = new();
    private static readonly Dictionary<string, ushort> AtomsByClassName = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<IntPtr, WindowEntry> Windows = new();
    private static readonly ConcurrentDictionary<uint, ThreadMessageQueue> ThreadQueues = new();
    private static readonly ConcurrentDictionary<IntPtr, uint> Timers = new();
    private static readonly ConcurrentDictionary<string, uint> ClipboardFormats = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<uint, string> ClipboardFormatNames = new();

    // Static fields rooting all unmanaged function delegates for the lifetime of the process.
    private static readonly RegisterClassWDelegate RegisterClassWInstance = RegisterClassWImpl;
    private static readonly UnregisterClassWDelegate UnregisterClassWInstance = UnregisterClassWImpl;
    private static readonly CreateWindowExWDelegate CreateWindowExWInstance = CreateWindowExWImpl;
    private static readonly DestroyWindowDelegate DestroyWindowInstance = DestroyWindowImpl;
    private static readonly DefWindowProcWDelegate DefWindowProcWInstance = DefWindowProcWImpl;
    private static readonly IsWindowDelegate IsWindowInstance = IsWindowImpl;
    private static readonly IsChildDelegate IsChildInstance = IsChildImpl;
    private static readonly SetWindowLongPtrWDelegate SetWindowLongPtrWInstance = SetWindowLongPtrWImpl;
    private static readonly SetWindowLongWDelegate SetWindowLongWInstance = SetWindowLongWImpl;
    private static readonly GetWindowLongPtrWDelegate GetWindowLongPtrWInstance = GetWindowLongPtrWImpl;
    private static readonly GetWindowLongWDelegate GetWindowLongWInstance = GetWindowLongWImpl;
    private static readonly GetWindowThreadProcessIdDelegate GetWindowThreadProcessIdInstance = GetWindowThreadProcessIdImpl;
    private static readonly PeekMessageWDelegate PeekMessageWInstance = PeekMessageWImpl;
    private static readonly GetMessageWDelegate GetMessageWInstance = GetMessageWImpl;
    private static readonly TranslateMessageDelegate TranslateMessageInstance = TranslateMessageImpl;
    private static readonly DispatchMessageWDelegate DispatchMessageWInstance = DispatchMessageWImpl;
    private static readonly PostMessageWDelegate PostMessageWInstance = PostMessageWImpl;
    private static readonly PostThreadMessageWDelegate PostThreadMessageWInstance = PostThreadMessageWImpl;
    private static readonly PostQuitMessageDelegate PostQuitMessageInstance = PostQuitMessageImpl;
    private static readonly SendMessageWDelegate SendMessageWInstance = SendMessageWImpl;
    private static readonly GetQueueStatusDelegate GetQueueStatusInstance = GetQueueStatusImpl;
    private static readonly InSendMessageExDelegate InSendMessageExInstance = InSendMessageExImpl;
    private static readonly IsGuiThreadDelegate IsGuiThreadInstance = IsGuiThreadImpl;
    private static readonly AllowSetForegroundWindowDelegate AllowSetForegroundWindowInstance = AllowSetForegroundWindowImpl;
    private static readonly SetCoalescableTimerDelegate SetCoalescableTimerInstance = SetCoalescableTimerImpl;
    private static readonly SetTimerDelegate SetTimerInstance = SetTimerImpl;
    private static readonly KillTimerDelegate KillTimerInstance = KillTimerImpl;
    private static readonly MsgWaitForMultipleObjectsExDelegate MsgWaitForMultipleObjectsExInstance = MsgWaitForMultipleObjectsExImpl;
    private static readonly IsThreadMessageQueueAttachedDelegate IsThreadMessageQueueAttachedInstance = IsThreadMessageQueueAttachedImpl;
    private static readonly SetThreadInputBlockedDelegate SetThreadInputBlockedInstance = SetThreadInputBlockedImpl;
    private static readonly GetProcessWindowStationDelegate GetProcessWindowStationInstance = GetProcessWindowStationImpl;
    private static readonly GetThreadDesktopDelegate GetThreadDesktopInstance = GetThreadDesktopImpl;
    private static readonly GetUserObjectInformationWDelegate GetUserObjectInformationWInstance = GetUserObjectInformationWImpl;
    private static readonly GetClipboardFormatNameWDelegate GetClipboardFormatNameWInstance = GetClipboardFormatNameWImpl;
    private static readonly MessageBeepDelegate MessageBeepInstance = MessageBeepImpl;
    private static readonly RegisterClipboardFormatWDelegate RegisterClipboardFormatWInstance = RegisterClipboardFormatWImpl;
    private static readonly CoInitializeDelegate CoInitializeInstance = CoInitializeImpl;
    private static readonly Dictionary<string, IntPtr> ExportTable = CreateExportTable();

    private static int _nextAtom = 0xC000;
    private static long _nextHwnd = 0x00010000;
    private static long _nextTimerId = 0x2000;
    private static int _nextClipboardFormat = 0xC100;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate ushort RegisterClassWDelegate(IntPtr lpWndClass);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int UnregisterClassWDelegate(IntPtr lpClassName, IntPtr hInstance);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr CreateWindowExWDelegate(
        uint dwExStyle,
        IntPtr lpClassName,
        IntPtr lpWindowName,
        uint dwStyle,
        int x,
        int y,
        int nWidth,
        int nHeight,
        IntPtr hWndParent,
        IntPtr hMenu,
        IntPtr hInstance,
        IntPtr lpParam);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int DestroyWindowDelegate(IntPtr hWnd);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr DefWindowProcWDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int IsWindowDelegate(IntPtr hWnd);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int IsChildDelegate(IntPtr hWndParent, IntPtr hWnd);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr SetWindowLongPtrWDelegate(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int SetWindowLongWDelegate(IntPtr hWnd, int nIndex, int dwNewLong);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr GetWindowLongPtrWDelegate(IntPtr hWnd, int nIndex);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int GetWindowLongWDelegate(IntPtr hWnd, int nIndex);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate uint GetWindowThreadProcessIdDelegate(IntPtr hWnd, IntPtr lpdwProcessId);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int PeekMessageWDelegate(
        IntPtr lpMsg,
        IntPtr hWnd,
        uint wMsgFilterMin,
        uint wMsgFilterMax,
        uint wRemoveMsg);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int GetMessageWDelegate(
        IntPtr lpMsg,
        IntPtr hWnd,
        uint wMsgFilterMin,
        uint wMsgFilterMax);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TranslateMessageDelegate(IntPtr lpMsg);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr DispatchMessageWDelegate(IntPtr lpMsg);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int PostMessageWDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int PostThreadMessageWDelegate(uint idThread, uint msg, IntPtr wParam, IntPtr lParam);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate void PostQuitMessageDelegate(int nExitCode);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr SendMessageWDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate uint GetQueueStatusDelegate(uint flags);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate uint InSendMessageExDelegate(IntPtr lpReserved);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int IsGuiThreadDelegate(int bConvert);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int AllowSetForegroundWindowDelegate(uint dwProcessId);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr SetCoalescableTimerDelegate(
        IntPtr hWnd,
        IntPtr nIdEvent,
        uint uElapse,
        IntPtr lpTimerFunc,
        uint uToleranceDelay);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr SetTimerDelegate(
        IntPtr hWnd,
        IntPtr nIdEvent,
        uint uElapse,
        IntPtr lpTimerFunc);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int KillTimerDelegate(IntPtr hWnd, IntPtr uIdEvent);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate uint MsgWaitForMultipleObjectsExDelegate(
        uint nCount,
        IntPtr pHandles,
        uint dwMilliseconds,
        uint dwWakeMask,
        uint dwFlags);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int IsThreadMessageQueueAttachedDelegate();

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int SetThreadInputBlockedDelegate(int blocked);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr GetProcessWindowStationDelegate();

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr GetThreadDesktopDelegate(uint dwThreadId);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int GetUserObjectInformationWDelegate(
        IntPtr hObj,
        int nIndex,
        IntPtr pvInfo,
        uint nLength,
        IntPtr lpnLengthNeeded);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int GetClipboardFormatNameWDelegate(uint format, IntPtr lpszFormatName, int cchMaxCount);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int MessageBeepDelegate(uint uType);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate uint RegisterClipboardFormatWDelegate(IntPtr lpszFormat);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int CoInitializeDelegate(IntPtr pvReserved);

    internal static IntPtr ResolveFunction(string functionName)
        => ExportTable.TryGetValue(functionName, out var ptr) ? ptr : IntPtr.Zero;

    internal static void RunHostMessageLoop(ManualResetEventSlim stopEvent)
    {
        var queue = GetOrCreateQueue(NativeMethods.GetCurrentThreadId());
        while (!stopEvent.IsSet)
        {
            queue.ProcessSentMessages();
            while (queue.TryDequeue(IntPtr.Zero, 0, 0, remove: true, out var msg))
            {
                if (msg.Message == WmQuit)
                {
                    return;
                }

                _ = InvokeWindowProc(msg.HWnd, msg.Message, msg.WParam, msg.LParam);
            }

            if (stopEvent.IsSet)
            {
                break;
            }

            _ = NativeMethods.WaitForSingleObject(queue.WakeEvent, 50);
        }
    }

    internal static void WakeThreadQueue(uint threadId)
    {
        if (ThreadQueues.TryGetValue(threadId, out var queue))
        {
            _ = NativeMethods.SetEvent(queue.WakeEvent);
        }
    }

    private static Dictionary<string, IntPtr> CreateExportTable()
        => new(StringComparer.Ordinal)
        {
            ["RegisterClassW"] = Marshal.GetFunctionPointerForDelegate(RegisterClassWInstance),
            ["UnregisterClassW"] = Marshal.GetFunctionPointerForDelegate(UnregisterClassWInstance),
            ["CreateWindowExW"] = Marshal.GetFunctionPointerForDelegate(CreateWindowExWInstance),
            ["DestroyWindow"] = Marshal.GetFunctionPointerForDelegate(DestroyWindowInstance),
            ["DefWindowProcW"] = Marshal.GetFunctionPointerForDelegate(DefWindowProcWInstance),
            ["IsWindow"] = Marshal.GetFunctionPointerForDelegate(IsWindowInstance),
            ["IsChild"] = Marshal.GetFunctionPointerForDelegate(IsChildInstance),
            ["SetWindowLongPtrW"] = Marshal.GetFunctionPointerForDelegate(SetWindowLongPtrWInstance),
            ["SetWindowLongW"] = Marshal.GetFunctionPointerForDelegate(SetWindowLongWInstance),
            ["GetWindowLongPtrW"] = Marshal.GetFunctionPointerForDelegate(GetWindowLongPtrWInstance),
            ["GetWindowLongW"] = Marshal.GetFunctionPointerForDelegate(GetWindowLongWInstance),
            ["GetWindowThreadProcessId"] = Marshal.GetFunctionPointerForDelegate(GetWindowThreadProcessIdInstance),
            ["PeekMessageW"] = Marshal.GetFunctionPointerForDelegate(PeekMessageWInstance),
            ["GetMessageW"] = Marshal.GetFunctionPointerForDelegate(GetMessageWInstance),
            ["TranslateMessage"] = Marshal.GetFunctionPointerForDelegate(TranslateMessageInstance),
            ["DispatchMessageW"] = Marshal.GetFunctionPointerForDelegate(DispatchMessageWInstance),
            ["PostMessageW"] = Marshal.GetFunctionPointerForDelegate(PostMessageWInstance),
            ["PostThreadMessageW"] = Marshal.GetFunctionPointerForDelegate(PostThreadMessageWInstance),
            ["PostQuitMessage"] = Marshal.GetFunctionPointerForDelegate(PostQuitMessageInstance),
            ["SendMessageW"] = Marshal.GetFunctionPointerForDelegate(SendMessageWInstance),
            ["GetQueueStatus"] = Marshal.GetFunctionPointerForDelegate(GetQueueStatusInstance),
            ["InSendMessageEx"] = Marshal.GetFunctionPointerForDelegate(InSendMessageExInstance),
            ["IsGUIThread"] = Marshal.GetFunctionPointerForDelegate(IsGuiThreadInstance),
            ["AllowSetForegroundWindow"] = Marshal.GetFunctionPointerForDelegate(AllowSetForegroundWindowInstance),
            ["SetCoalescableTimer"] = Marshal.GetFunctionPointerForDelegate(SetCoalescableTimerInstance),
            ["SetTimer"] = Marshal.GetFunctionPointerForDelegate(SetTimerInstance),
            ["KillTimer"] = Marshal.GetFunctionPointerForDelegate(KillTimerInstance),
            ["MsgWaitForMultipleObjectsEx"] = Marshal.GetFunctionPointerForDelegate(MsgWaitForMultipleObjectsExInstance),
            ["IsThreadMessageQueueAttached"] = Marshal.GetFunctionPointerForDelegate(IsThreadMessageQueueAttachedInstance),
            ["SetThreadInputBlocked"] = Marshal.GetFunctionPointerForDelegate(SetThreadInputBlockedInstance),
            ["GetProcessWindowStation"] = Marshal.GetFunctionPointerForDelegate(GetProcessWindowStationInstance),
            ["GetThreadDesktop"] = Marshal.GetFunctionPointerForDelegate(GetThreadDesktopInstance),
            ["GetUserObjectInformationW"] = Marshal.GetFunctionPointerForDelegate(GetUserObjectInformationWInstance),
            ["GetClipboardFormatNameW"] = Marshal.GetFunctionPointerForDelegate(GetClipboardFormatNameWInstance),
            ["MessageBeep"] = Marshal.GetFunctionPointerForDelegate(MessageBeepInstance),
            ["RegisterClipboardFormatW"] = Marshal.GetFunctionPointerForDelegate(RegisterClipboardFormatWInstance),
            ["CoInitialize"] = Marshal.GetFunctionPointerForDelegate(CoInitializeInstance),
        };

    private static ThreadMessageQueue GetOrCreateQueue(uint threadId)
        => ThreadQueues.GetOrAdd(threadId, static _ => new ThreadMessageQueue());

    private static ushort RegisterClassWImpl(IntPtr lpWndClass)
    {
        if (lpWndClass == IntPtr.Zero)
        {
            return 0;
        }

        var style = (uint)Marshal.ReadInt32(lpWndClass, 0);
        var wndProc = Marshal.ReadIntPtr(lpWndClass, IntPtr.Size);
        var clsExtra = Marshal.ReadInt32(lpWndClass, IntPtr.Size * 2);
        var wndExtra = Marshal.ReadInt32(lpWndClass, (IntPtr.Size * 2) + 4);
        var hInstanceOffset = (IntPtr.Size * 2) + 8;
        var classNameOffset = hInstanceOffset + (IntPtr.Size * 5);
        var hInstance = Marshal.ReadIntPtr(lpWndClass, hInstanceOffset);
        var classNamePtr = Marshal.ReadIntPtr(lpWndClass, classNameOffset);
        var className = ExtractClassName(classNamePtr);
        if (string.IsNullOrEmpty(className))
        {
            return 0;
        }

        lock (ClassLock)
        {
            if (!AtomsByClassName.TryGetValue(className, out var atom))
            {
                var next = Interlocked.Increment(ref _nextAtom);
                atom = unchecked((ushort)(0xC000 | (next & 0x3FFF)));
                AtomsByClassName[className] = atom;
            }

            ClassesByAtom[atom] = new WindowClassEntry(atom, className, style, wndProc, clsExtra, wndExtra, hInstance);
            return atom;
        }
    }

    private static int UnregisterClassWImpl(IntPtr lpClassName, IntPtr hInstance)
    {
        _ = hInstance;
        if (lpClassName == IntPtr.Zero)
        {
            return 0;
        }

        lock (ClassLock)
        {
            var raw = unchecked((ulong)lpClassName.ToInt64());
            if ((raw & ~0xFFFFUL) == 0)
            {
                var atom = unchecked((ushort)raw);
                if (ClassesByAtom.TryGetValue(atom, out var existing))
                {
                    _ = AtomsByClassName.Remove(existing.ClassName);
                    _ = ClassesByAtom.Remove(atom);
                }

                return 1;
            }

            var className = Marshal.PtrToStringUni(lpClassName);
            if (className is { Length: > 0 } && AtomsByClassName.TryGetValue(className, out var mappedAtom))
            {
                _ = AtomsByClassName.Remove(className);
                _ = ClassesByAtom.Remove(mappedAtom);
            }

            return 1;
        }
    }

    private static IntPtr CreateWindowExWImpl(
        uint dwExStyle,
        IntPtr lpClassName,
        IntPtr lpWindowName,
        uint dwStyle,
        int x,
        int y,
        int nWidth,
        int nHeight,
        IntPtr hWndParent,
        IntPtr hMenu,
        IntPtr hInstance,
        IntPtr lpParam)
    {
        _ = lpWindowName;
        _ = x;
        _ = y;
        _ = nWidth;
        _ = nHeight;
        _ = lpParam;

        var classEntry = FindWindowClass(lpClassName);
        if (classEntry is null)
        {
            NativeMethods.SetLastError(ErrorClassDoesNotExist);
            return IntPtr.Zero;
        }

        var threadId = NativeMethods.GetCurrentThreadId();
        _ = GetOrCreateQueue(threadId);
        var hwnd = new IntPtr(Interlocked.Increment(ref _nextHwnd));
        var entry = new WindowEntry(
            hwnd,
            threadId,
            classEntry.WndProc,
            hInstance != IntPtr.Zero ? hInstance : classEntry.HInstance,
            hWndParent,
            hMenu,
            dwStyle,
            dwExStyle,
            classEntry.WndExtra);
        Windows[hwnd] = entry;
        return hwnd;
    }

    private static WindowClassEntry? FindWindowClass(IntPtr lpClassName)
    {
        if (lpClassName == IntPtr.Zero)
        {
            return null;
        }

        lock (ClassLock)
        {
            var raw = unchecked((ulong)lpClassName.ToInt64());
            if ((raw & ~0xFFFFUL) == 0)
            {
                var atom = unchecked((ushort)raw);
                return ClassesByAtom.TryGetValue(atom, out var byAtom) ? byAtom : null;
            }

            var name = Marshal.PtrToStringUni(lpClassName);
            if (name is { Length: > 0 }
                && AtomsByClassName.TryGetValue(name, out var mappedAtom)
                && ClassesByAtom.TryGetValue(mappedAtom, out var byName))
            {
                return byName;
            }

            return null;
        }
    }

    private static string ExtractClassName(IntPtr classNamePtr)
    {
        if (classNamePtr == IntPtr.Zero)
        {
            return string.Empty;
        }

        var raw = unchecked((ulong)classNamePtr.ToInt64());
        return (raw & ~0xFFFFUL) == 0
            ? $"#Atom{raw:X4}"
            : Marshal.PtrToStringUni(classNamePtr) ?? string.Empty;
    }

    private static int DestroyWindowImpl(IntPtr hWnd)
    {
        if (!Windows.TryRemove(hWnd, out _))
        {
            NativeMethods.SetLastError(ErrorInvalidWindowHandle);
            return 0;
        }

        return 1;
    }

    private static IntPtr DefWindowProcWImpl(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        _ = hWnd;
        _ = msg;
        _ = wParam;
        _ = lParam;
        return IntPtr.Zero;
    }

    private static int IsWindowImpl(IntPtr hWnd) => Windows.ContainsKey(hWnd) ? 1 : 0;

    private static int IsChildImpl(IntPtr hWndParent, IntPtr hWnd)
        => Windows.TryGetValue(hWnd, out var entry) && entry.Parent == hWndParent && hWndParent != IntPtr.Zero ? 1 : 0;

    private static IntPtr SetWindowLongPtrWImpl(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
    {
        if (!Windows.TryGetValue(hWnd, out var entry))
        {
            NativeMethods.SetLastError(ErrorInvalidWindowHandle);
            return IntPtr.Zero;
        }

        return entry.SetLongPtr(nIndex, dwNewLong);
    }

    private static int SetWindowLongWImpl(IntPtr hWnd, int nIndex, int dwNewLong)
        => unchecked((int)SetWindowLongPtrWImpl(hWnd, nIndex, new IntPtr(dwNewLong)).ToInt64());

    private static IntPtr GetWindowLongPtrWImpl(IntPtr hWnd, int nIndex)
    {
        if (!Windows.TryGetValue(hWnd, out var entry))
        {
            NativeMethods.SetLastError(ErrorInvalidWindowHandle);
            return IntPtr.Zero;
        }

        return entry.GetLongPtr(nIndex);
    }

    private static int GetWindowLongWImpl(IntPtr hWnd, int nIndex)
        => unchecked((int)GetWindowLongPtrWImpl(hWnd, nIndex).ToInt64());

    private static uint GetWindowThreadProcessIdImpl(IntPtr hWnd, IntPtr lpdwProcessId)
    {
        if (lpdwProcessId != IntPtr.Zero)
        {
            Marshal.WriteInt32(lpdwProcessId, unchecked((int)NativeMethods.GetCurrentProcessId()));
        }

        return Windows.TryGetValue(hWnd, out var entry) ? entry.ThreadId : NativeMethods.GetCurrentThreadId();
    }

    private static int PeekMessageWImpl(
        IntPtr lpMsg,
        IntPtr hWnd,
        uint wMsgFilterMin,
        uint wMsgFilterMax,
        uint wRemoveMsg)
    {
        var queue = GetOrCreateQueue(NativeMethods.GetCurrentThreadId());
        queue.ProcessSentMessages();
        if (!queue.TryDequeue(hWnd, wMsgFilterMin, wMsgFilterMax, (wRemoveMsg & PmRemove) != 0, out var msg))
        {
            ZeroMsgStruct(lpMsg);
            return 0;
        }

        WriteMsgStruct(lpMsg, msg);
        return 1;
    }

    private static int GetMessageWImpl(
        IntPtr lpMsg,
        IntPtr hWnd,
        uint wMsgFilterMin,
        uint wMsgFilterMax)
    {
        var queue = GetOrCreateQueue(NativeMethods.GetCurrentThreadId());
        while (true)
        {
            queue.ProcessSentMessages();
            if (queue.TryDequeue(hWnd, wMsgFilterMin, wMsgFilterMax, remove: true, out var msg))
            {
                WriteMsgStruct(lpMsg, msg);
                return msg.Message == WmQuit ? 0 : 1;
            }

            _ = NativeMethods.WaitForSingleObject(queue.WakeEvent, 50);
        }
    }

    private static int TranslateMessageImpl(IntPtr lpMsg)
    {
        _ = lpMsg;
        return 0;
    }

    private static IntPtr DispatchMessageWImpl(IntPtr lpMsg)
    {
        if (lpMsg == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        var hWnd = Marshal.ReadIntPtr(lpMsg, 0);
        var msg = unchecked((uint)Marshal.ReadInt32(lpMsg, IntPtr.Size));
        var wParam = Marshal.ReadIntPtr(lpMsg, IntPtr.Size * 2);
        var lParam = Marshal.ReadIntPtr(lpMsg, IntPtr.Size * 3);
        return InvokeWindowProc(hWnd, msg, wParam, lParam);
    }

    private static int PostMessageWImpl(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        uint targetThreadId;
        if (hWnd == IntPtr.Zero)
        {
            targetThreadId = NativeMethods.GetCurrentThreadId();
        }
        else if (Windows.TryGetValue(hWnd, out var entry))
        {
            targetThreadId = entry.ThreadId;
        }
        else
        {
            NativeMethods.SetLastError(ErrorInvalidWindowHandle);
            return 0;
        }

        var queue = GetOrCreateQueue(targetThreadId);
        queue.Enqueue(new QueuedMessage(hWnd, msg, wParam, lParam, unchecked((uint)Environment.TickCount)));
        return 1;
    }

    private static int PostThreadMessageWImpl(uint idThread, uint msg, IntPtr wParam, IntPtr lParam)
    {
        var queue = GetOrCreateQueue(idThread);
        queue.Enqueue(new QueuedMessage(IntPtr.Zero, msg, wParam, lParam, unchecked((uint)Environment.TickCount)));
        return 1;
    }

    private static void PostQuitMessageImpl(int nExitCode)
    {
        var queue = GetOrCreateQueue(NativeMethods.GetCurrentThreadId());
        queue.Enqueue(new QueuedMessage(IntPtr.Zero, WmQuit, new IntPtr(nExitCode), IntPtr.Zero, unchecked((uint)Environment.TickCount)));
    }

    private static IntPtr SendMessageWImpl(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        if (!Windows.TryGetValue(hWnd, out var entry))
        {
            NativeMethods.SetLastError(ErrorInvalidWindowHandle);
            return IntPtr.Zero;
        }

        var currentThreadId = NativeMethods.GetCurrentThreadId();
        if (entry.ThreadId == currentThreadId)
        {
            return InvokeWindowProc(hWnd, msg, wParam, lParam);
        }

        var targetQueue = GetOrCreateQueue(entry.ThreadId);
        return targetQueue.SendCrossThread(hWnd, msg, wParam, lParam);
    }

#pragma warning disable S6640 // Unmanaged function pointer avoids Marshal.GetDelegateForFunctionPointer reverse-P/Invoke unwrapping on .NET 8+.
    private static unsafe IntPtr InvokeWindowProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        if (hWnd == IntPtr.Zero || !Windows.TryGetValue(hWnd, out var entry) || entry.WndProc == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        var callback = (delegate* unmanaged[Stdcall]<IntPtr, uint, IntPtr, IntPtr, IntPtr>)entry.WndProc;
        return callback(hWnd, msg, wParam, lParam);
    }
#pragma warning restore S6640

    private static uint GetQueueStatusImpl(uint flags)
    {
        _ = flags;
        var queue = GetOrCreateQueue(NativeMethods.GetCurrentThreadId());
        return queue.HasMessages ? 0x00FF00FFu : 0u;
    }

    private static uint InSendMessageExImpl(IntPtr lpReserved)
    {
        _ = lpReserved;
        var queue = GetOrCreateQueue(NativeMethods.GetCurrentThreadId());
        return queue.InSendMessageDepth > 0 ? 1u : 0u;
    }

    private static int IsGuiThreadImpl(int bConvert)
    {
        _ = bConvert;
        _ = GetOrCreateQueue(NativeMethods.GetCurrentThreadId());
        return 1;
    }

    private static int AllowSetForegroundWindowImpl(uint dwProcessId)
    {
        _ = dwProcessId;
        return 1;
    }

    private static IntPtr SetCoalescableTimerImpl(
        IntPtr hWnd,
        IntPtr nIdEvent,
        uint uElapse,
        IntPtr lpTimerFunc,
        uint uToleranceDelay)
    {
        _ = uToleranceDelay;
        return SetTimerImpl(hWnd, nIdEvent, uElapse, lpTimerFunc);
    }

    private static IntPtr SetTimerImpl(
        IntPtr hWnd,
        IntPtr nIdEvent,
        uint uElapse,
        IntPtr lpTimerFunc)
    {
        _ = hWnd;
        _ = lpTimerFunc;
        var timerId = nIdEvent != IntPtr.Zero
            ? nIdEvent
            : new IntPtr(Interlocked.Increment(ref _nextTimerId));
        Timers[timerId] = uElapse;
        return timerId;
    }

    private static int KillTimerImpl(IntPtr hWnd, IntPtr uIdEvent)
    {
        _ = hWnd;
        _ = Timers.TryRemove(uIdEvent, out _);
        return 1;
    }

    private static uint MsgWaitForMultipleObjectsExImpl(
        uint nCount,
        IntPtr pHandles,
        uint dwMilliseconds,
        uint dwWakeMask,
        uint dwFlags)
    {
        _ = dwWakeMask;
        var queue = GetOrCreateQueue(NativeMethods.GetCurrentThreadId());
        queue.ProcessSentMessages();
        if (queue.HasMessages)
        {
            return nCount;
        }

        var handles = new IntPtr[nCount + 1];
        for (uint i = 0; i < nCount; i++)
        {
            handles[i] = Marshal.ReadIntPtr(pHandles, checked((int)(i * (uint)IntPtr.Size)));
        }

        handles[nCount] = queue.WakeEvent;
        var waitAll = (dwFlags & 0x0001u) != 0;
        var alertable = (dwFlags & 0x0002u) != 0;
        var result = NativeMethods.WaitForMultipleObjectsEx(nCount + 1, handles, waitAll, dwMilliseconds, alertable);
        queue.ProcessSentMessages();
        return result;
    }

#pragma warning disable S3400
    private static int IsThreadMessageQueueAttachedImpl() => 1;
#pragma warning restore S3400

    private static int SetThreadInputBlockedImpl(int blocked)
    {
        _ = blocked;
        return 1;
    }

    private static IntPtr GetProcessWindowStationImpl() => FakeWinSta;

    private static IntPtr GetThreadDesktopImpl(uint dwThreadId)
    {
        _ = dwThreadId;
        return FakeDesktop;
    }

    private static int GetUserObjectInformationWImpl(
        IntPtr hObj,
        int nIndex,
        IntPtr pvInfo,
        uint nLength,
        IntPtr lpnLengthNeeded)
    {
        if (nIndex != UoiName)
        {
            if (lpnLengthNeeded != IntPtr.Zero)
            {
                Marshal.WriteInt32(lpnLengthNeeded, 0);
            }

            return 1;
        }

        var name = hObj == FakeWinSta ? "WinSta0" : "Default";
        var requiredBytes = (name.Length + 1) * 2;
        if (lpnLengthNeeded != IntPtr.Zero)
        {
            Marshal.WriteInt32(lpnLengthNeeded, requiredBytes);
        }

        if (pvInfo == IntPtr.Zero || nLength < (uint)requiredBytes)
        {
            return 0;
        }

        for (var i = 0; i < name.Length; i++)
        {
            Marshal.WriteInt16(pvInfo, i * 2, unchecked((short)name[i]));
        }

        Marshal.WriteInt16(pvInfo, name.Length * 2, 0);
        return 1;
    }

    private static int GetClipboardFormatNameWImpl(uint format, IntPtr lpszFormatName, int cchMaxCount)
    {
        if (lpszFormatName == IntPtr.Zero || cchMaxCount <= 0)
        {
            return 0;
        }

        var name = ClipboardFormatNames.TryGetValue(format, out var registered)
            ? registered
            : $"Format_{format:X}";
        var copyLength = Math.Min(name.Length, cchMaxCount - 1);
        for (var i = 0; i < copyLength; i++)
        {
            Marshal.WriteInt16(lpszFormatName, i * 2, unchecked((short)name[i]));
        }

        Marshal.WriteInt16(lpszFormatName, copyLength * 2, 0);
        return copyLength;
    }

    private static int MessageBeepImpl(uint uType)
    {
        _ = uType;
        return 1;
    }

    private static uint RegisterClipboardFormatWImpl(IntPtr lpszFormat)
    {
        if (lpszFormat == IntPtr.Zero)
        {
            return 0;
        }

        var formatName = Marshal.PtrToStringUni(lpszFormat);
        if (formatName is not { Length: > 0 })
        {
            return 0;
        }

        return ClipboardFormats.GetOrAdd(
            formatName,
            static key =>
            {
                var id = unchecked((uint)Interlocked.Increment(ref _nextClipboardFormat));
                ClipboardFormatNames[id] = key;
                return id;
            });
    }

    private static int CoInitializeImpl(IntPtr pvReserved)
        => NativeMethods.CoInitializeEx(pvReserved, 0x2u /* COINIT_APARTMENTTHREADED */);

    private static void WriteMsgStruct(IntPtr lpMsg, QueuedMessage msg)
    {
        if (lpMsg == IntPtr.Zero)
        {
            return;
        }

        Marshal.WriteIntPtr(lpMsg, 0, msg.HWnd);
        Marshal.WriteInt32(lpMsg, IntPtr.Size, unchecked((int)msg.Message));
        Marshal.WriteIntPtr(lpMsg, IntPtr.Size * 2, msg.WParam);
        Marshal.WriteIntPtr(lpMsg, IntPtr.Size * 3, msg.LParam);
        Marshal.WriteInt32(lpMsg, IntPtr.Size * 4, unchecked((int)msg.Time));
        Marshal.WriteInt32(lpMsg, (IntPtr.Size * 4) + 4, 0);
        Marshal.WriteInt32(lpMsg, (IntPtr.Size * 4) + 8, 0);
    }

    private static void ZeroMsgStruct(IntPtr lpMsg)
    {
        if (lpMsg == IntPtr.Zero)
        {
            return;
        }

        WriteMsgStruct(lpMsg, default);
    }

    [StructLayout(LayoutKind.Auto)]
    private readonly struct QueuedMessage
    {
        internal QueuedMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam, uint time)
        {
            HWnd = hWnd;
            Message = message;
            WParam = wParam;
            LParam = lParam;
            Time = time;
        }

        internal IntPtr HWnd { get; }

        internal uint Message { get; }

        internal IntPtr WParam { get; }

        internal IntPtr LParam { get; }

        internal uint Time { get; }
    }

    private sealed class WindowClassEntry
    {
        internal WindowClassEntry(
            ushort atom,
            string className,
            uint style,
            IntPtr wndProc,
            int clsExtra,
            int wndExtra,
            IntPtr hInstance)
        {
            Atom = atom;
            ClassName = className;
            Style = style;
            WndProc = wndProc;
            ClsExtra = clsExtra;
            WndExtra = wndExtra;
            HInstance = hInstance;
        }

        internal ushort Atom { get; }

        internal string ClassName { get; }

        internal uint Style { get; }

        internal IntPtr WndProc { get; }

        internal int ClsExtra { get; }

        internal int WndExtra { get; }

        internal IntPtr HInstance { get; }
    }

    private sealed class WindowEntry
    {
        private readonly object _lock = new();
        private readonly byte[] _extraBytes;
        private IntPtr _wndProc;
        private IntPtr _hInstance;
        private IntPtr _id;
        private IntPtr _userData;
        private uint _style;
        private uint _exStyle;

        internal WindowEntry(
            IntPtr hWnd,
            uint threadId,
            IntPtr wndProc,
            IntPtr hInstance,
            IntPtr parent,
            IntPtr id,
            uint style,
            uint exStyle,
            int wndExtra)
        {
            HWnd = hWnd;
            ThreadId = threadId;
            _wndProc = wndProc;
            _hInstance = hInstance;
            Parent = parent;
            _id = id;
            _style = style;
            _exStyle = exStyle;
            _extraBytes = new byte[Math.Max(64, Math.Max(0, wndExtra) + IntPtr.Size)];
        }

        internal IntPtr HWnd { get; }

        internal uint ThreadId { get; }

        internal IntPtr Parent { get; }

        internal IntPtr WndProc
        {
            get
            {
                lock (_lock)
                {
                    return _wndProc;
                }
            }
        }

        internal IntPtr GetLongPtr(int nIndex)
        {
            lock (_lock)
            {
                if (nIndex >= 0 && nIndex + IntPtr.Size <= _extraBytes.Length)
                {
                    return IntPtr.Size == 8
                        ? new IntPtr(BitConverter.ToInt64(_extraBytes, nIndex))
                        : new IntPtr(BitConverter.ToInt32(_extraBytes, nIndex));
                }

                return nIndex switch
                {
                    GwlpWndProc => _wndProc,
                    GwlpHInstance => _hInstance,
                    GwlpId => _id,
                    GwlStyle => new IntPtr(unchecked((int)_style)),
                    GwlExStyle => new IntPtr(unchecked((int)_exStyle)),
                    GwlpUserData => _userData,
                    _ => IntPtr.Zero,
                };
            }
        }

        internal IntPtr SetLongPtr(int nIndex, IntPtr newValue)
        {
            lock (_lock)
            {
                var previous = GetLongPtr(nIndex);
                if (nIndex >= 0 && nIndex + IntPtr.Size <= _extraBytes.Length)
                {
                    var bytes = IntPtr.Size == 8
                        ? BitConverter.GetBytes(newValue.ToInt64())
                        : BitConverter.GetBytes(newValue.ToInt32());
                    Buffer.BlockCopy(bytes, 0, _extraBytes, nIndex, bytes.Length);
                    return previous;
                }

                switch (nIndex)
                {
                    case GwlpWndProc:
                        _wndProc = newValue;
                        break;
                    case GwlpHInstance:
                        _hInstance = newValue;
                        break;
                    case GwlpId:
                        _id = newValue;
                        break;
                    case GwlStyle:
                        _style = unchecked((uint)newValue.ToInt64());
                        break;
                    case GwlExStyle:
                        _exStyle = unchecked((uint)newValue.ToInt64());
                        break;
                    case GwlpUserData:
                        _userData = newValue;
                        break;
                    default:
                        break;
                }

                return previous;
            }
        }
    }

    private sealed class SentMessageItem : IDisposable
    {
        private readonly ManualResetEventSlim _completed = new(initialState: false);
        private long _result;

        internal SentMessageItem(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam)
        {
            HWnd = hWnd;
            Message = message;
            WParam = wParam;
            LParam = lParam;
        }

        internal IntPtr HWnd { get; }

        internal uint Message { get; }

        internal IntPtr WParam { get; }

        internal IntPtr LParam { get; }

        internal IntPtr Result => new(Interlocked.Read(ref _result));

        public void Dispose() => _completed.Dispose();

        internal void Complete(IntPtr result)
        {
            _ = Interlocked.Exchange(ref _result, result.ToInt64());
            _completed.Set();
        }

        internal bool Wait(int millisecondsTimeout) => _completed.Wait(millisecondsTimeout);
    }

    private sealed class ThreadMessageQueue
    {
        private readonly object _lock = new();
        private readonly List<QueuedMessage> _posted = [];
        private readonly Queue<SentMessageItem> _sent = new();

        internal ThreadMessageQueue()
        {
            WakeEvent = NativeMethods.CreateEventW(IntPtr.Zero, bManualReset: true, bInitialState: false, lpName: IntPtr.Zero);
        }

        internal IntPtr WakeEvent { get; }

        internal int InSendMessageDepth { get; private set; }

        internal bool HasMessages
        {
            get
            {
                lock (_lock)
                {
                    return _posted.Count > 0 || _sent.Count > 0;
                }
            }
        }

        internal void Enqueue(QueuedMessage message)
        {
            lock (_lock)
            {
                _posted.Add(message);
                _ = NativeMethods.SetEvent(WakeEvent);
            }
        }

        internal IntPtr SendCrossThread(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam)
        {
            using var item = new SentMessageItem(hWnd, message, wParam, lParam);
            lock (_lock)
            {
                _sent.Enqueue(item);
                _ = NativeMethods.SetEvent(WakeEvent);
            }

            if (item.Wait(2000))
            {
                return item.Result;
            }

            return InvokeWindowProc(hWnd, message, wParam, lParam);
        }

        internal void ProcessSentMessages()
        {
            while (true)
            {
                SentMessageItem? item;
                lock (_lock)
                {
                    if (_sent.Count == 0)
                    {
                        if (_posted.Count == 0)
                        {
                            _ = NativeMethods.ResetEvent(WakeEvent);
                        }

                        return;
                    }

                    item = _sent.Dequeue();
                }

                InSendMessageDepth++;
                try
                {
                    var result = InvokeWindowProc(item.HWnd, item.Message, item.WParam, item.LParam);
                    item.Complete(result);
                }
                finally
                {
                    InSendMessageDepth--;
                }
            }
        }

        internal bool TryDequeue(
            IntPtr hWndFilter,
            uint minFilter,
            uint maxFilter,
            bool remove,
            out QueuedMessage message)
        {
            lock (_lock)
            {
                for (var i = 0; i < _posted.Count; i++)
                {
                    var candidate = _posted[i];
                    if (!MatchesWindowFilter(candidate.HWnd, hWndFilter)
                        || !MatchesMessageFilter(candidate.Message, minFilter, maxFilter))
                    {
                        continue;
                    }

                    if (remove)
                    {
                        _posted.RemoveAt(i);
                        if (_posted.Count == 0 && _sent.Count == 0)
                        {
                            _ = NativeMethods.ResetEvent(WakeEvent);
                        }
                    }

                    message = candidate;
                    return true;
                }

                message = default;
                return false;
            }
        }

        private static bool MatchesWindowFilter(IntPtr messageHwnd, IntPtr filterHwnd)
        {
            if (filterHwnd == IntPtr.Zero)
            {
                return true;
            }

            if (filterHwnd == new IntPtr(-1))
            {
                return messageHwnd == IntPtr.Zero;
            }

            return messageHwnd == filterHwnd;
        }

        private static bool MatchesMessageFilter(uint message, uint minFilter, uint maxFilter)
            => (minFilter == 0 && maxFilter == 0) || (message >= minFilter && message <= maxFilter);
    }
}
#pragma warning restore S101
#pragma warning restore SA1121
#pragma warning restore MA0158
