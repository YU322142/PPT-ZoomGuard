using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Extensibility;
using System.Drawing;

[assembly: ComVisible(false)]
[assembly: AssemblyTitle("Lesson Zoom Guard")]
[assembly: AssemblyVersion("0.1.0.0")]

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

        public void OnConnection(object app, ext_ConnectMode mode, object addIn, ref Array custom)
        {
            try
            {
                application = app;
                Log.Write("connected version=0.1.0 arch=" + (IntPtr.Size * 8));
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
            if (stopping || obj != 0 || child != 0 || h == IntPtr.Zero) return;
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
            shows.Add(h, new ShowGuard(h, IsInkActive, delegate { return ReadViewInt("CurrentShowPosition"); }));
        }

        private bool IsInkActive()
        {
            int type = ReadViewInt("PointerType");
            return type == 2 || type == 5 || type == 6;
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
                foreach (IntPtr h in Native.ShowWindows((uint)Process.GetCurrentProcess().Id)) AddShow(h);
                foreach (var pair in new List<KeyValuePair<IntPtr, ShowGuard>>(shows))
                {
                    if (!Native.IsWindow(pair.Key)) { pair.Value.Dispose(); shows.Remove(pair.Key); }
                    else pair.Value.Refresh();
                }
            }
            catch (Exception ex) { Log.Write("tick error " + ex.GetType().Name + " " + ex.Message); }
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
            var result = new StringBuilder("LessonZoomGuard 0.1.0");
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
        private readonly InputShield shield;
        private readonly Func<int> readPosition;
        private int position;

        public ShowGuard(IntPtr h, Func<bool> isInk, Func<int> getPosition)
        {
            readPosition = getPosition;
            window = h;
            original = Native.ReadZoom(h, out haveOriginal);
            AssignHandle(h);
            shield = new InputShield(h, isInk);
            Apply();
            Log.Write("show attached h=" + h + " original=" + original.Want + "/" + original.Block + " read=" + haveOriginal);
        }

        public string Status
        {
            get { return "hwnd=" + window + " mode=" + (enabled ? "zoom-enabled" : "zoom-locked") +
                " config=" + lastConfiguration + " blocked=" + blocked + " shield=" + shield.Visible + " rejected=" + shield.Rejected; }
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
            shield.RefreshGate(!enabled);
        }

        

        

        private void SetMode(bool allow, string reason)
        {
            if (enabled == allow) return;
            enabled = allow;
            if (shield != null) shield.RefreshGate(!enabled);
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
            shield.Dispose();
            if (Native.IsWindow(window) && haveOriginal)
            {
                var setting = original;
                Native.SetGestureConfig(window, 0, 1, ref setting, 12);
            }
            if (Handle != IntPtr.Zero) ReleaseHandle();
            Log.Write("show detached h=" + window);
        }
    }

    internal sealed class InputShield : Form
    {
        private readonly IntPtr target;
        private readonly Func<bool> isInk;
        private readonly HashSet<uint> contacts = new HashSet<uint>();
        private bool multi, drawing;
        private uint primary;
        private Point start, last;
        private long started;
        internal long Rejected;

        internal InputShield(IntPtr window, Func<bool> ink)
        {
            target = window; isInk = ink;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.Black;
            Opacity = 1.0 / 255.0;
            TopMost = true;
        }
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams
        {
            get { var cp = base.CreateParams; cp.ExStyle |= 0x08000080; cp.Parent = target; return cp; }
        }
        protected override void OnVisibleChanged(EventArgs e)
        {
            if (!Visible) { contacts.Clear(); multi = false; drawing = false; }
            base.OnVisibleChanged(e);
        }
        internal void RefreshGate(bool locked)
        {
            IntPtr foreground = Native.GetForegroundWindow();
            bool active = foreground == target || foreground == Handle;
            if (!locked || !active || !Native.IsWindowVisible(target))
            {
                if (Visible) Hide();
                return;
            }
            Native.Rect r;
            if (!Native.GetWindowRect(target, out r)) return;
            var wanted = Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom);
            if (Bounds != wanted) Bounds = wanted;
            if (!Visible) Show();
        }

        private Point ClientPoint(Native.Point point)
        {
            Native.ScreenToClient(target, ref point);
            return new Point(point.X, point.Y);
        }
        private IntPtr Pack(Point p) { return new IntPtr(unchecked((int)((uint)(ushort)p.X | ((uint)(ushort)p.Y << 16)))); }
        private void Mouse(uint message, int keys, Point point)
        {
            Native.PostMessage(target, message, new IntPtr(keys), Pack(point));
        }
        
        private void FinishTap(Point point)
        {
            int dx = point.X - start.X, dy = point.Y - start.Y;
            if (Math.Abs(dx) > 100 && Math.Abs(dx) > Math.Abs(dy) * 2)
            {
                int key = dx < 0 ? 0x27 : 0x25;
                Native.PostMessage(target, 0x100, new IntPtr(key), IntPtr.Zero);
                Native.PostMessage(target, 0x101, new IntPtr(key), IntPtr.Zero);
            }
            else if (Math.Abs(dx) < 60 && Math.Abs(dy) < 60 && Environment.TickCount - started >= 650)
            {
                Mouse(0x204, 2, point); Mouse(0x205, 0, point);
            }
            else if (Math.Abs(dx) < 60 && Math.Abs(dy) < 60)
            {
                Mouse(0x200, 0, point); Mouse(0x201, 1, point); Mouse(0x202, 0, point);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x21 || m.Msg == 0x24b) { m.Result = new IntPtr(3); return; }
            if (m.Msg >= 0x245 && m.Msg <= 0x24f)
            {
                uint id = (uint)(m.WParam.ToInt64() & 0xffff);
                Native.PointerInfo info;
                if (Native.GetPointerInfo(id, out info) && (info.Type == 2 || info.Type == 3))
                {
                    Point p = ClientPoint(info.Pixel);
                    if (m.Msg == 0x246)
                    {
                        if (contacts.Count == 0) { primary = id; multi = false; start = last = p; started = Environment.TickCount; drawing = isInk() || info.Type == 3; }
                        contacts.Add(id);
                        uint count = 0; Native.GetPointerFrameInfo(id, ref count, IntPtr.Zero);
                        if (contacts.Count > 1 || count > 1)
                        {
                            if (!multi) Rejected++;
                            multi = true;
                            if (drawing) Native.SendMessage(target, 0x24c, new IntPtr(primary), IntPtr.Zero);
                            drawing = false;
                        }
                        else if (drawing && id == primary)
                        {
                            Native.SendMessage(target, 0x249, m.WParam, m.LParam);
                            Native.SendMessage(target, (uint)m.Msg, m.WParam, m.LParam);
                        }
                    }
                    else if (m.Msg == 0x245 && id == primary && !multi)
                    {
                        last = p;
                        if (drawing) Native.SendMessage(target, (uint)m.Msg, m.WParam, m.LParam);
                    }
                    else if (m.Msg == 0x247)
                    {
                        if (id == primary && !multi)
                        {
                            if (drawing) Native.SendMessage(target, (uint)m.Msg, m.WParam, m.LParam);
                            else if (!drawing) FinishTap(p);
                        }
                        contacts.Remove(id);
                        if (contacts.Count == 0) { multi = false; drawing = false; }
                    }
                    else if (m.Msg == 0x24c)
                    {
                        contacts.Remove(id);
                        if (contacts.Count == 0) { multi = false; drawing = false; }
                    }
                    m.Result = IntPtr.Zero; return;
                }
            }
            if ((m.Msg >= 0x200 && m.Msg <= 0x20e) || m.Msg == 0x2a3)
            {
                // Windows can also promote touch to mouse; the pointer path above
                // already handled those contacts, so only forward real mouse input.
                long extra = Native.GetMessageExtraInfo().ToInt64();
                if ((extra & 0xffffff00L) != 0xff515700L)
                    Native.PostMessage(target, (uint)m.Msg, m.WParam, m.LParam);
                m.Result = IntPtr.Zero; return;
            }
            if (m.Msg == 0x20)
            {
                m.Result = Native.SendMessage(target, 0x20, target, m.LParam); return;
            }
            base.WndProc(ref m);
        }
    }

    internal static class Native
    {
        internal delegate bool EnumCallback(IntPtr h, IntPtr l);
        internal delegate void WinEventCallback(IntPtr hook, uint ev, IntPtr h, int obj, int child, uint thread, uint time);
        internal delegate IntPtr MessageHook(int code, IntPtr w, IntPtr l);
        [StructLayout(LayoutKind.Sequential)] internal struct Point { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] internal struct Rect { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] internal struct MouseInput { public int X, Y; public uint Data, Flags, Time; public IntPtr Extra; }
        [StructLayout(LayoutKind.Sequential)] internal struct Input { public uint Type; public MouseInput Mouse; }
        [StructLayout(LayoutKind.Sequential)] internal struct PointerInfo
        {
            public uint Type, Id, Frame, Flags; public IntPtr Device, Target;
            public Point Pixel, Himetric, PixelRaw, HimetricRaw;
            public uint Time, History; public int InputData; public uint Keys;
            public ulong Performance; public uint Change;
        }
        [StructLayout(LayoutKind.Sequential)] internal struct QueuedMessage
        {
            public IntPtr Window; public uint Message; public IntPtr WParam, LParam;
            public uint Time; public Point Position; public uint Private;
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
        [DllImport("user32.dll")] internal static extern bool IsTouchWindow(IntPtr h, out uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern IntPtr DefWindowProc(IntPtr h, uint message, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] internal static extern bool GetPointerType(uint id, out uint type);
        [DllImport("user32.dll")] internal static extern bool GetPointerInfo(uint id, out PointerInfo info);
        [DllImport("user32.dll")] internal static extern IntPtr GetMessageExtraInfo();
        [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] internal static extern bool GetWindowRect(IntPtr h, out Rect r);
        [DllImport("user32.dll")] internal static extern bool ScreenToClient(IntPtr h, ref Point p);
        [DllImport("user32.dll")] internal static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
        [DllImport("user32.dll", SetLastError = true)] internal static extern bool GetPointerFrameInfo(uint id, ref uint count, IntPtr pointers);
        [DllImport("user32.dll")] internal static extern IntPtr GetAncestor(IntPtr h, uint flags);
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
