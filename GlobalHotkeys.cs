using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace JarvisPowerPoint
{
    internal enum PresentationHotkey { Next, Previous, Resume, NextResult, BlackScreen, RestoreSlides }

    internal sealed class HotkeyEventArgs : EventArgs
    {
        public HotkeyEventArgs(PresentationHotkey command) { Command = command; }
        public PresentationHotkey Command { get; private set; }
    }

    internal sealed class HotkeyRegistrationException : InvalidOperationException
    {
        internal HotkeyRegistrationException(PresentationHotkey command, string shortcut, int nativeError,
            bool cleanupSucceeded)
            : base("Cannot register " + shortcut + " (Windows error "
                + nativeError.ToString(CultureInfo.InvariantCulture) + "). "
                + (cleanupSucceeded ? "No keyboard shortcuts are enabled."
                    : "Keyboard commands are disabled; cleanup failed. Disable shortcuts again or close the application."))
        {
            Command = command;
            Shortcut = shortcut;
            NativeErrorCode = nativeError;
        }

        public PresentationHotkey Command { get; private set; }
        public string Shortcut { get; private set; }
        public int NativeErrorCode { get; private set; }
    }

    internal interface IHotkeyNative
    {
        bool Register(IntPtr window, int id, uint modifiers, uint key, out int error);
        bool Unregister(IntPtr window, int id, out int error);
    }

    internal sealed class GlobalHotkeys : NativeWindow, IDisposable
    {
        internal const int HotkeyMessage = 0x0312;
        internal const uint Modifiers = 0x0001 | 0x0002 | 0x0004;
        internal const uint NoRepeat = 0x4000;
        private const long MinimumIntervalMilliseconds = 180;
        private static readonly HotkeySpec[] Specifications = {
            new HotkeySpec(PresentationHotkey.Next, Keys.Right, "Right"),
            new HotkeySpec(PresentationHotkey.Previous, Keys.Left, "Left"),
            new HotkeySpec(PresentationHotkey.Resume, Keys.Enter, "Enter"),
            new HotkeySpec(PresentationHotkey.NextResult, Keys.N, "N"),
            new HotkeySpec(PresentationHotkey.BlackScreen, Keys.B, "B"),
            new HotkeySpec(PresentationHotkey.RestoreSlides, Keys.S, "S")
        };
        private readonly IHotkeyNative native;
        private readonly Func<long> milliseconds;
        private readonly int ownerThread;
        private readonly Dictionary<int, HotkeySpec> registered = new Dictionary<int, HotkeySpec>();
        private readonly Dictionary<PresentationHotkey, long> lastDispatch =
            new Dictionary<PresentationHotkey, long>();
        private bool enabled;
        private bool disposed;
        private int nextId = 0x4000;

        public GlobalHotkeys() : this(new Win32HotkeyNative(), CreateClock()) { }

        internal GlobalHotkeys(IHotkeyNative native, Func<long> milliseconds)
        {
            if (native == null) throw new ArgumentNullException("native");
            if (milliseconds == null) throw new ArgumentNullException("milliseconds");
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Keyboard shortcuts require the UI STA thread.");
            this.native = native;
            this.milliseconds = milliseconds;
            ownerThread = Thread.CurrentThread.ManagedThreadId;
            CreateHandle(new CreateParams { Caption = "Jarvis shortcut receiver", Parent = new IntPtr(-3) });
        }

        public bool Enabled { get { return enabled; } }
        public event EventHandler<HotkeyEventArgs> CommandPressed;

        public static IDictionary<PresentationHotkey, string> ShortcutDescriptions
        {
            get
            {
                var descriptions = new Dictionary<PresentationHotkey, string>();
                foreach (HotkeySpec spec in Specifications) descriptions.Add(spec.Command, spec.Description);
                return descriptions;
            }
        }

        public void Enable()
        {
            CheckThread();
            if (disposed) throw new ObjectDisposedException("GlobalHotkeys");
            if (enabled) return;
            if (registered.Count != 0) Disable();
            // Fresh IDs keep already-queued notifications from an earlier enable cycle inert.
            if (nextId > 0xBFFF - Specifications.Length + 1)
                throw new InvalidOperationException("Restart the application before enabling shortcuts again.");
            try
            {
                foreach (HotkeySpec spec in Specifications)
                {
                    int id = nextId++;
                    int error;
                    if (!native.Register(Handle, id, Modifiers | NoRepeat, (uint)spec.Key, out error))
                    {
                        int cleanupError = ReleaseRegistrations();
                        throw new HotkeyRegistrationException(spec.Command, spec.Description, error, cleanupError == 0);
                    }
                    registered.Add(id, spec);
                }
                lastDispatch.Clear();
                enabled = true;
            }
            catch
            {
                enabled = false;
                ReleaseRegistrations();
                throw;
            }
        }

        public void Disable()
        {
            CheckThread();
            if (disposed) return;
            enabled = false;
            lastDispatch.Clear();
            int error = ReleaseRegistrations();
            if (error != 0)
                throw new InvalidOperationException("Keyboard shortcut cleanup failed (Windows error "
                    + error.ToString(CultureInfo.InvariantCulture) + "). Commands are disabled; retry cleanup.");
        }

        public void Dispose()
        {
            CheckThread();
            if (disposed) return;
            try { Disable(); }
            finally
            {
                disposed = true;
                enabled = false;
                CommandPressed = null;
                DestroyHandle();
                registered.Clear();
            }
            GC.SuppressFinalize(this);
        }

        protected override void WndProc(ref Message message)
        {
            if (message.Msg == HotkeyMessage)
            {
                if (message.HWnd == Handle)
                {
                    long id = message.WParam.ToInt64();
                    if (id >= 0 && id <= 0xBFFF) ProcessHotkeyMessage((int)id, message.LParam);
                }
                return;
            }
            base.WndProc(ref message);
        }

        internal bool ProcessHotkeyMessage(int id, IntPtr data)
        {
            CheckThread();
            HotkeySpec spec;
            if (disposed || !enabled || !registered.TryGetValue(id, out spec)) return false;
            long packet = data.ToInt64();
            if ((packet & 0xFFFF) != Modifiers || ((packet >> 16) & 0xFFFF) != (uint)spec.Key) return false;
            long now = milliseconds();
            long previous;
            if (lastDispatch.TryGetValue(spec.Command, out previous)
                && (now < previous || now - previous < MinimumIntervalMilliseconds)) return false;
            lastDispatch[spec.Command] = now;
            EventHandler<HotkeyEventArgs> handler = CommandPressed;
            if (handler != null) handler(this, new HotkeyEventArgs(spec.Command));
            return true;
        }

        private int ReleaseRegistrations()
        {
            int firstError = 0;
            foreach (int id in new List<int>(registered.Keys))
            {
                int error;
                if (native.Unregister(Handle, id, out error)) registered.Remove(id);
                else if (firstError == 0) firstError = error == 0 ? 1 : error;
            }
            return firstError;
        }

        private void CheckThread()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("Keyboard shortcuts must be managed on their owning UI thread.");
        }

        private static Func<long> CreateClock()
        {
            Stopwatch clock = Stopwatch.StartNew();
            return delegate { return clock.ElapsedMilliseconds; };
        }

        private sealed class HotkeySpec
        {
            public readonly PresentationHotkey Command;
            public readonly Keys Key;
            public readonly string Description;

            public HotkeySpec(PresentationHotkey command, Keys key, string label)
            {
                Command = command;
                Key = key;
                Description = "Ctrl+Alt+Shift+" + label;
            }
        }

        private sealed class Win32HotkeyNative : IHotkeyNative
        {
            public bool Register(IntPtr window, int id, uint modifiers, uint key, out int error)
            {
                bool result = RegisterHotKey(window, id, modifiers, key);
                error = result ? 0 : Marshal.GetLastWin32Error();
                return result;
            }

            public bool Unregister(IntPtr window, int id, out int error)
            {
                bool result = UnregisterHotKey(window, id);
                error = result ? 0 : Marshal.GetLastWin32Error();
                return result;
            }

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint virtualKey);

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool UnregisterHotKey(IntPtr window, int id);
        }
    }
}
