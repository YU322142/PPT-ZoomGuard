using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Extensibility;

[assembly: ComVisible(false)]
[assembly: AssemblyTitle("Lesson Zoom Guard")]
[assembly: AssemblyVersion("0.2.1.0")]

namespace LessonZoomGuard
{
    [ComVisible(true), Guid("3215887B-2452-46EB-9B59-01F4169CFBD1"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGuardDiagnostics
    {
        [DispId(1)] string GetStatus();
        [DispId(2)] void SetTracing(bool enabled);
        [DispId(3)] string GetLogPath();
    }

    [ComVisible(true), Guid("F1C4E46D-337F-4504-9F69-81A47FE21183"), ProgId("LessonZoomGuard.Connect"),
        ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_IDTExtensibility2))]
    public sealed class Connect : _IDTExtensibility2, IGuardDiagnostics
    {
        public Connect() { Log.Write("component constructed"); }
        private readonly Dictionary<IntPtr, ShowGuard> shows = new Dictionary<IntPtr, ShowGuard>();
        private Timer timer;
        private object application;
        private Native.WinEventCallback eventCallback;
        private IntPtr eventHook;
        private IntPtr foregroundHook;
        private bool stopping;
        private bool checkingNavigation;
        private long navigationClosures;
        private string navigationError = "";

        public void OnConnection(object app, ext_ConnectMode mode, object addIn, ref Array custom)
        {
            try
            {
                application = app;
                Log.Write("connected version=0.2.1-preview arch=" + (IntPtr.Size * 8));
                Start();
                addIn.GetType().InvokeMember("Object", BindingFlags.SetProperty, null, addIn,
                    new object[] { new GuardDiagnostics(this) });
            }
            catch (Exception ex) { Log.Write("connect error " + ex.GetType().Name + " " + ex.Message); }
        }

        private void Start()
        {
            if (timer != null) return;
            stopping = false;
            eventCallback = WindowEvent;
            eventHook = Native.SetWinEventHook(0x8000, 0x8003, IntPtr.Zero, eventCallback,
                (uint)Process.GetCurrentProcess().Id, 0, 0);
            foregroundHook = Native.SetWinEventHook(3, 3, IntPtr.Zero, eventCallback, 0, 0, 0);
            timer = new Timer();
            timer.Interval = 60;
            timer.Tick += Tick;
            timer.Start();
            Tick(null, EventArgs.Empty);
        }

        private void WindowEvent(IntPtr hook, uint ev, IntPtr h, int obj, int child, uint thread, uint time)
        {
            if (stopping || h == IntPtr.Zero) return;
            if (ev == 0x8002) CloseNavigation();
            if (obj != 0 || child != 0) return;
            try
            {
                if (Native.ClassName(h) == "screenClass" && ev != 0x8001) AddShow(h);
                foreach (ShowGuard guard in new List<ShowGuard>(shows.Values)) guard.Refresh();
            }
            catch (Exception ex) { Log.Write("window event error " + ex.GetType().Name); }
        }

        private void AddShow(IntPtr h)
        {
            if (shows.ContainsKey(h)) return;
            uint pid;
            uint thread = Native.GetWindowThreadProcessId(h, out pid);
            if (pid != (uint)Process.GetCurrentProcess().Id || thread != Native.GetCurrentThreadId()) return;
            shows.Add(h, new ShowGuard(h, delegate { return ReadViewInt("CurrentShowPosition"); }));
        }

        private int ReadViewInt(string property)
        {
            object windows = null, show = null, view = null;
            try
            {
                windows = ComGet(application, "SlideShowWindows");
                show = ComGet(windows, "Item", 1);
                view = ComGet(show, "View");
                return Convert.ToInt32(ComGet(view, property));
            }
            catch { return 0; }
            finally
            {
                foreach (object item in new object[] { view, show, windows })
                    if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
            }
        }
        private static object ComGet(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(name, args.Length == 0 ? BindingFlags.GetProperty : BindingFlags.InvokeMethod, null, target, args);
        }

        private void Tick(object sender, EventArgs args)
        {
            if (stopping) return;
            try
            {
                CloseNavigation();
                foreach (IntPtr h in Native.ShowWindows((uint)Process.GetCurrentProcess().Id)) AddShow(h);
                foreach (var pair in new List<KeyValuePair<IntPtr, ShowGuard>>(shows))
                {
                    if (!Native.IsWindow(pair.Key)) { pair.Value.Dispose(); shows.Remove(pair.Key); }
                    else pair.Value.Refresh();
                }
            }
            catch (Exception ex) { Log.Write("tick error " + ex.GetType().Name + " " + ex.Message); }
        }

        // Input filtering is best-effort across touch drivers. Enforce the
        // user-visible invariant separately through PowerPoint's own navigation
        // object, including paths that do not deliver WM_POINTER to screenClass.
        private void CloseNavigation()
        {
            if (stopping || application == null || checkingNavigation) return;
            checkingNavigation = true;
            object windows = null;
            try
            {
                windows = ComGet(application, "SlideShowWindows");
                int count = Convert.ToInt32(ComGet(windows, "Count"));
                for (int index = 1; index <= count; index++)
                {
                    object show = null, navigation = null;
                    try
                    {
                        show = ComGet(windows, "Item", index);
                        navigation = ComGet(show, "SlideNavigation");
                        if (!Convert.ToBoolean(ComGet(navigation, "Visible"))) continue;
                        navigation.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null,
                            navigation, new object[] { false });
                        navigationClosures++;
                        Log.Write("navigation closed count=" + navigationClosures);
                        foreach (ShowGuard guard in new List<ShowGuard>(shows.Values)) guard.NavigationClosed();
                    }
                    finally
                    {
                        foreach (object item in new object[] { navigation, show })
                            if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
                    }
                }
                navigationError = "";
            }
            catch (Exception ex)
            {
                string error = ex.GetType().Name + ":" + ex.HResult.ToString("X");
                if (navigationError != error) Log.Write("navigation check error " + error);
                navigationError = error;
            }
            finally
            {
                if (windows != null && Marshal.IsComObject(windows)) Marshal.ReleaseComObject(windows);
                checkingNavigation = false;
            }
        }

        private void Stop()
        {
            if (stopping) return;
            stopping = true;
            if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; }
            if (eventHook != IntPtr.Zero) { Native.UnhookWinEvent(eventHook); eventHook = IntPtr.Zero; }
            if (foregroundHook != IntPtr.Zero) { Native.UnhookWinEvent(foregroundHook); foregroundHook = IntPtr.Zero; }
            foreach (ShowGuard guard in shows.Values) guard.Dispose();
            shows.Clear();
            application = null;
            Log.Write("disconnected");
        }

        public string GetStatus()
        {
            Tick(null, EventArgs.Empty);
            var result = new StringBuilder("LessonZoomGuard 0.2.1-preview");
            result.Append("\nnavigationClosed=").Append(navigationClosures).Append(" navigationError=").Append(navigationError);
            foreach (ShowGuard guard in shows.Values) result.Append("\n").Append(guard.Status);
            if (shows.Count == 0) result.Append("\nNo slide-show window.");
            return result.ToString();
        }
        public void SetTracing(bool enabled) { Log.Tracing = enabled; Log.Write("tracing=" + enabled); }
        public string GetLogPath() { return Log.Path; }
        public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { Stop(); }
        public void OnBeginShutdown(ref Array custom) { Stop(); }
        public void OnStartupComplete(ref Array custom) { Start(); }
        public void OnAddInsUpdate(ref Array custom) { }
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IGuardDiagnostics))]
    public sealed class GuardDiagnostics : IGuardDiagnostics
    {
        private readonly Connect owner;
        internal GuardDiagnostics(Connect owner) { this.owner = owner; }
        public string GetStatus() { return owner.GetStatus(); }
        public void SetTracing(bool enabled) { owner.SetTracing(enabled); }
        public string GetLogPath() { return owner.GetLogPath(); }
    }

    internal sealed class ShowGuard : NativeWindow, IDisposable
    {
        private readonly IntPtr window;
        private Native.GestureConfig original;
        private bool haveOriginal;
        private bool enabled;
        private bool disposed;
        private bool selectorSeen;
        private bool applying;
        private string lastConfiguration = "";
        private int traceBudget = 400;
        private long blocked;
        private readonly HashSet<uint> touching = new HashSet<uint>();
        private readonly Dictionary<uint, IntPtr> delivered = new Dictionary<uint, IntPtr>();
        private bool suppressTouchSession;
        private readonly Func<int> readPosition;
        private int position;

        public ShowGuard(IntPtr h, Func<int> getPosition)
        {
            readPosition = getPosition;
            window = h;
            original = Native.ReadZoom(h, out haveOriginal);
            AssignHandle(h);
            Apply();
            Log.Write("show attached h=" + h + " original=" + original.Want + "/" + original.Block + " read=" + haveOriginal);
        }

        public string Status
        {
            get { return "hwnd=" + window + " mode=" + (enabled ? "zoom-enabled" : "zoom-locked") +
                " config=" + lastConfiguration + " blocked=" + blocked + " native-input=true suppressed=" + suppressTouchSession; }
        }

        public void Refresh()
        {
            if (disposed) return;
            bool selection = Native.HasVisibleSelector(window);
            int current = readPosition();
            if (current > 0 && position > 0 && current != position) SetMode(false, "slide changed");
            if (current > 0) position = current;
            if (selection && !selectorSeen) SetMode(true, "native magnifier selector");
            selectorSeen = selection;
        }

        public void NavigationClosed()
        {
            SetMode(false, "navigation safety net");
            // Do not immediately interpret a lingering selector as a fresh,
            // intentional request to enable zoom after closing the overview.
            selectorSeen = Native.HasVisibleSelector(window);
        }

        public bool FilterPointer(IntPtr recipient, uint message, IntPtr wParam)
        {
            uint id = (uint)(wParam.ToInt64() & 0xffff);
            Native.PointerInfo pointer;
            if (!Native.GetPointerInfo(id, out pointer) || pointer.Type != 2) return false;
            if (message == 0x246) touching.Add(id);
            if (enabled) return false;
            uint count = Native.TouchContactsInFrame(id);
            if (!enabled && !suppressTouchSession && (count > 1 || touching.Count > 1))
            {
                suppressTouchSession = true;
                blocked++;
                Log.Write("multi-touch cancelled frame=" + pointer.Frame + " count=" + count);
                foreach (var prior in new List<KeyValuePair<uint, IntPtr>>(delivered))
                {
                    // Bypass our filter so PowerPoint receives the cancellation.
                    Message cancellation = Message.Create(prior.Value, 0x24c, new IntPtr(prior.Key), IntPtr.Zero);
                    base.WndProc(ref cancellation);
                }
                delivered.Clear();
            }
            bool reject = suppressTouchSession && !enabled;
            if (!reject && message == 0x246) delivered[id] = recipient;
            if (message == 0x247 || message == 0x24c)
            {
                touching.Remove(id); delivered.Remove(id);
                if (touching.Count == 0 && count == 0) suppressTouchSession = false;
            }
            return reject;
        }

        private void SetMode(bool allow, string reason)
        {
            if (enabled == allow) return;
            enabled = allow;
            touching.Clear(); delivered.Clear(); suppressTouchSession = false;
            Apply();
            Log.Write("mode h=" + window + " allow=" + allow + " reason=" + reason);
        }

        private void Apply()
        {
            if (applying || disposed || !Native.IsWindow(window)) return;
            applying = true;
            try
            {
                var setting = enabled ? original : new Native.GestureConfig { Id = 3, Want = 0, Block = 1 };
                if (!haveOriginal && enabled) setting = new Native.GestureConfig { Id = 3, Want = 1, Block = 0 };
                bool ok = Native.SetGestureConfig(window, 0, 1, ref setting, 12);
                string value = ok ? (enabled ? "restored" : "zoom-blocked") : "error-" + Marshal.GetLastWin32Error();
                if (lastConfiguration != value) { lastConfiguration = value; Log.Write("config h=" + window + " " + value); }
            }
            finally { applying = false; }
        }

        protected override void WndProc(ref Message m)
        {
            try
            {
                // Preserve original single-finger/pen/mouse input. Only touch
                // sequences with multiple contacts are consumed, at dispatch time.
                // GetPointerFrameInfo is valid here, unlike queued-message hooks.
                if (m.Msg >= 0x245 && m.Msg <= 0x24c && FilterPointer(window, (uint)m.Msg, m.WParam))
                { m.Result = IntPtr.Zero; return; }
                if (Log.Tracing && traceBudget > 0 && (m.Msg >= 0x400 || (m.Msg >= 0x240 && m.Msg <= 0x24F) || m.Msg == 0x111 || m.Msg == 0x100 || m.Msg == 0x201 || m.Msg == 0x202 || m.Msg == 0x210))
                {
                    traceBudget--;
                    Log.Write("msg h=" + window + " m=" + m.Msg.ToString("X") + " w=" + m.WParam.ToInt64().ToString("X") + "");
                }
                if (m.Msg == 0x11A) Apply(); // WM_GESTURENOTIFY: do not discard pan or pen input.
                if (m.Msg == 0x0119 && !enabled)
                {
                    Native.GestureInfo info = new Native.GestureInfo();
                    info.Size = (uint)Marshal.SizeOf(typeof(Native.GestureInfo));
                    if (Native.GetGestureInfo(m.LParam, ref info) && info.Id == 3)
                    {
                        Native.CloseGestureInfoHandle(m.LParam);
                        blocked++;
                        m.Result = IntPtr.Zero;
                        return;
                    }
                }
                bool escape = ((m.Msg == 0x100 || m.Msg == 0x104) && m.WParam.ToInt64() == 27)
                    || (m.Msg == 0x111 && ((m.WParam.ToInt64() & 0xffff) == 0x19b || (m.WParam.ToInt64() & 0xffff) == 0x834));
                bool nativeZoomExit = enabled && (m.Msg == 0x205 || m.Msg == 0x7b);
                base.WndProc(ref m);
                if (escape) { selectorSeen = false; SetMode(false, "Escape"); }
                if (nativeZoomExit) { selectorSeen = false; SetMode(false, "native right-click zoom exit"); }
                if (m.Msg == 0x210) Refresh();
                if (m.Msg == 0x11A) Apply(); // WM_GESTURENOTIFY follows the native window's setup.
            }
            catch (Exception ex)
            {
                Log.Write("message error " + ex.GetType().Name);
                // Managed exceptions must never escape into the Office message loop.
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            if (Native.IsWindow(window) && haveOriginal)
            {
                var setting = original;
                Native.SetGestureConfig(window, 0, 1, ref setting, 12);
            }
            if (Handle != IntPtr.Zero) ReleaseHandle();
            Log.Write("show detached h=" + window);
        }
    }

    internal static class Native
    {
        internal delegate bool EnumCallback(IntPtr h, IntPtr l);
        internal delegate void WinEventCallback(IntPtr hook, uint ev, IntPtr h, int obj, int child, uint thread, uint time);
        [StructLayout(LayoutKind.Sequential)] internal struct Point { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] internal struct Rect { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] internal struct PointerInfo
        {
            public uint Type, Id, Frame, Flags; public IntPtr Device, Target;
            public Point Pixel, Himetric, PixelRaw, HimetricRaw;
            public uint Time, History; public int InputData; public uint Keys;
            public ulong Performance; public uint Change;
        }
        [StructLayout(LayoutKind.Sequential)] internal struct GestureConfig { public uint Id, Want, Block; }
        [StructLayout(LayoutKind.Sequential)] internal struct PointS { public short X, Y; }
        [StructLayout(LayoutKind.Sequential)] internal struct GestureInfo
        {
            public uint Size, Flags, Id; public IntPtr Target; public PointS Location;
            public uint Instance, Sequence; public ulong Arguments; public uint Extra;
        }
        [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr h);
        [DllImport("user32.dll")] internal static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] internal static extern bool GetPointerInfo(uint id, out PointerInfo info);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
        [DllImport("user32.dll", SetLastError = true)] internal static extern bool GetPointerFrameInfo(uint id, ref uint count, IntPtr pointers);
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("kernel32.dll")] internal static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumCallback callback, IntPtr l);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr h, EnumCallback callback, IntPtr l);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder b, int count);
        [DllImport("user32.dll", SetLastError = true)] internal static extern bool SetGestureConfig(IntPtr h, uint reserved, uint count, ref GestureConfig config, uint size);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool GetGestureConfig(IntPtr h, uint reserved, uint flags, ref uint count, [In, Out] GestureConfig[] config, uint size);
        [DllImport("user32.dll")] internal static extern bool GetGestureInfo(IntPtr h, ref GestureInfo info);
        [DllImport("user32.dll")] internal static extern bool CloseGestureInfoHandle(IntPtr h);
        [DllImport("user32.dll")] internal static extern IntPtr SetWinEventHook(uint first, uint last, IntPtr module, WinEventCallback callback, uint pid, uint thread, uint flags);
        [DllImport("user32.dll")] internal static extern bool UnhookWinEvent(IntPtr hook);

        internal static uint TouchContactsInFrame(uint id)
        {
            uint count = 0;
            GetPointerFrameInfo(id, ref count, IntPtr.Zero);
            if (count == 0 || count > 256) return 0;
            int size = Marshal.SizeOf(typeof(PointerInfo));
            IntPtr buffer = Marshal.AllocHGlobal(checked(size * (int)count));
            try
            {
                if (!GetPointerFrameInfo(id, ref count, buffer)) return 0;
                uint contacts = 0;
                for (int i = 0; i < count; i++)
                {
                    var p = (PointerInfo)Marshal.PtrToStructure(IntPtr.Add(buffer, i * size), typeof(PointerInfo));
                    if (p.Type == 2 && (p.Flags & 4) != 0) contacts++;
                }
                return contacts;
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }

        internal static string ClassName(IntPtr h) { var b = new StringBuilder(256); GetClassName(h, b, b.Capacity); return b.ToString(); }
        internal static List<IntPtr> ShowWindows(uint process)
        {
            var result = new List<IntPtr>();
            EnumWindows(delegate(IntPtr h, IntPtr l)
            {
                uint pid; GetWindowThreadProcessId(h, out pid);
                if (pid == process && ClassName(h) == "screenClass") result.Add(h);
                return true;
            }, IntPtr.Zero);
            return result;
        }
        internal static bool HasVisibleSelector(IntPtr window)
        {
            bool found = false;
            EnumChildWindows(window, delegate(IntPtr h, IntPtr l)
            {
                if (ClassName(h) == "PPTOpticalZoom" && IsWindowVisible(h)) { found = true; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }
        internal static GestureConfig ReadZoom(IntPtr h, out bool ok)
        {
            var config = new GestureConfig[] { new GestureConfig { Id = 3 } };
            uint count = 1;
            ok = GetGestureConfig(h, 0, 0, ref count, config, 12);
            return config[0];
        }
    }

    internal static class Log
    {
        internal static bool Tracing;
        internal static readonly string Path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LessonZoomGuard", "guard.log");
        internal static void Write(string value)
        {
            try
            {
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Path));
                if (File.Exists(Path) && new FileInfo(Path).Length > 1048576) File.Delete(Path);
                File.AppendAllText(Path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + value + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
        }
    }
}
