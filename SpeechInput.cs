using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Speech.AudioFormat;
using System.Speech.Recognition;
using System.Text;
using System.Threading;

namespace JarvisPowerPoint
{
    internal sealed class AudioInputDevice
    {
        private readonly string id;
        private readonly string name;

        public AudioInputDevice(string id, string name)
        {
            this.id = id;
            this.name = name;
        }

        public string Id { get { return id; } }
        public string Name { get { return name; } }
        public override string ToString() { return name; }
    }

    internal static class SpeechInput
    {
        private static Exception captureFailure;
        internal static bool CaptureBlocked { get { return Interlocked.CompareExchange(ref captureFailure, null, null) != null; } }

        internal static void EnsureCaptureAvailable()
        {
            Exception error = Interlocked.CompareExchange(ref captureFailure, null, null);
            if (error != null)
                throw new IOException(RestartRequiredMessage(true), error);
        }

        internal static string RestartRequiredMessage(bool english)
        {
            return english
                ? "Microphone cleanup failed. Cancel setup, quit Jarvis completely, check the device, then restart Jarvis manually. "
                    + "No new capture or settings save is allowed until restart."
                : "Échec de libération du microphone. Annulez la configuration, quittez complètement Jarvis, vérifiez le périphérique, "
                    + "puis redémarrez Jarvis manuellement. Toute nouvelle écoute ou sauvegarde des réglages est bloquée jusqu’au redémarrage.";
        }

        internal static SpeechCleanupException BlockCaptureFailure(Exception failure)
        {
            var error = new SpeechCleanupException(RestartRequiredMessage(true), failure);
            Interlocked.CompareExchange(ref captureFailure, error, null);
            return error;
        }

        private static SpeechCleanupException BlockCapture(string message)
        {
            var error = new SpeechCleanupException(message);
            Interlocked.CompareExchange(ref captureFailure, error, null);
            return error;
        }

        public static IList<AudioInputDevice> GetDevices()
        {
            var devices = new List<AudioInputDevice>();
            devices.Add(new AudioInputDevice("default", "Default (Windows)"));
            uint count = WaveInNative.waveInGetNumDevs();
            for (uint index = 0; index < count; index++)
            {
                WaveInNative.Capabilities caps;
                WaveInNative.Check(WaveInNative.waveInGetDevCaps(new UIntPtr(index), out caps,
                    (uint)Marshal.SizeOf(typeof(WaveInNative.Capabilities))), "enumerate microphones");
                devices.Add(new AudioInputDevice(CreateDeviceId(index, caps.Name,
                    caps.Manufacturer, caps.Product), caps.Name));
            }
            return devices.AsReadOnly();
        }

        public static IDisposable Attach(SpeechRecognitionEngine engine, string deviceId)
        {
            if (engine == null) throw new ArgumentNullException("engine");
            EnsureCaptureAvailable();
            if (deviceId == "default")
            {
                engine.SetInputToDefaultAudioDevice();
                return null;
            }

            uint index = ResolveDeviceIndex(deviceId, GetDevices());
            var input = new WaveInInput(index, deviceId);
            try
            {
                engine.SetInputToAudioStream(input.Stream,
                    new SpeechAudioFormatInfo(16000, AudioBitsPerSample.Sixteen, AudioChannel.Mono));
                return input;
            }
            catch (Exception failure)
            {
                SpeechInputErrors.CleanupPreservingFailure(input.Dispose, failure);
                throw;
            }
        }

        // Polling is useful because SAPI does not always forward a stream read failure promptly.
        public static Exception GetError(IDisposable input)
        {
            var capture = input as WaveInInput;
            return capture == null ? null : capture.Stream.Error;
        }

        public static int GetLevel(IDisposable input)
        {
            var capture = input as WaveInInput;
            return capture == null ? 0 : capture.Level;
        }

        internal static int MeasureLevel(byte[] pcm, int count)
        {
            if (pcm == null) throw new ArgumentNullException("pcm");
            if (count < 0 || count > pcm.Length || (count & 1) != 0)
                throw new ArgumentOutOfRangeException("count");
            if (count == 0) return 0;
            double sum = 0;
            for (int i = 0; i < count; i += 2)
            {
                double sample = (short)(pcm[i] | (pcm[i + 1] << 8));
                sum += sample * sample;
            }
            if (sum == 0) return 0;
            double rms = Math.Sqrt(sum / (count / 2)) / 32768.0;
            // Map -60..0 dBFS onto the UI's 0..100 range, without retaining another audio copy.
            return (int)Math.Round(Math.Max(0, Math.Min(100, (20 * Math.Log10(rms) + 60) * 100 / 60)));
        }

        internal static string CreateDeviceId(uint index, string name, ushort manufacturer, ushort product)
        {
            return "wavein:" + index.ToString(CultureInfo.InvariantCulture) + ":"
                + manufacturer.ToString(CultureInfo.InvariantCulture) + ":"
                + product.ToString(CultureInfo.InvariantCulture) + ":"
                + Convert.ToBase64String(Encoding.UTF8.GetBytes(name));
        }

        internal static uint ResolveDeviceIndex(string id, IList<AudioInputDevice> devices)
        {
            if (!string.IsNullOrEmpty(id) && id.StartsWith("wavein:", StringComparison.Ordinal))
            {
                string[] parts = id.Split(':');
                uint index;
                if (parts.Length == 5 && uint.TryParse(parts[1], NumberStyles.None,
                    CultureInfo.InvariantCulture, out index))
                {
                    foreach (AudioInputDevice device in devices)
                        if (device.Id == id) return index;
                }
            }
            throw new IOException("The selected microphone is missing or its device index/name has changed. "
                + "Reconnect it, open microphone setup, refresh the list and select it again. "
                + "No other microphone was selected automatically.");
        }

        private sealed class WaveInInput : IDisposable
        {
            private const int BufferBytes = 3200;
            private const int BufferCount = 4;
            private readonly ManualResetEvent stop = new ManualResetEvent(false);
            private readonly object startupGate = new object();
            private readonly PcmAudioStream stream = new PcmAudioStream(128000);
            private readonly Thread worker;
            private readonly uint deviceIndex;
            private readonly string deviceId;
            private readonly NativeBuffer[] buffers = new NativeBuffer[BufferCount];
            private IntPtr handle;
            private Exception startupError;
            private Exception cleanupError;
            private bool startupFinished;
            private int disposed;
            private int level;

            public WaveInInput(uint index, string id)
            {
                deviceIndex = index;
                deviceId = id;
                worker = new Thread(Capture);
                worker.IsBackground = true;
                worker.Name = "Jarvis local microphone input";
                try
                {
                    worker.Start();
                }
                catch
                {
                    stream.Dispose();
                    stop.Dispose();
                    throw;
                }
                bool initialized;
                lock (startupGate)
                {
                    if (!startupFinished) Monitor.Wait(startupGate, 5000);
                    initialized = startupFinished;
                }
                if (!initialized || startupError != null)
                {
                    Exception error = startupError ??
                        new TimeoutException("The microphone driver did not initialize within five seconds.");
                    Exception cleanupError = SpeechInputErrors.CaptureExpectedFailure(Dispose);
                    if (cleanupError != null) error = new AggregateException(error, cleanupError);
                    throw new IOException("Cannot start the selected microphone. Check Windows microphone "
                        + "privacy permissions, reconnect the device or choose another microphone.",
                        error);
                }
            }

            public PcmAudioStream Stream { get { return stream; } }
            public int Level
            {
                get
                {
                    return Interlocked.CompareExchange(ref disposed, 0, 0) != 0 ? 0
                        : Interlocked.CompareExchange(ref level, 0, 0);
                }
            }

            private void Capture()
            {
                bool started = false;
                try
                {
                    EnsureCaptureAvailable();
                    var format = new WaveInNative.Format();
                    format.FormatTag = 1;
                    format.Channels = 1;
                    format.SamplesPerSecond = 16000;
                    format.AverageBytesPerSecond = 32000;
                    format.BlockAlign = 2;
                    format.BitsPerSample = 16;
                    // CALLBACK_NULL: one worker owns every native call; no native-to-managed callback
                    // or callback delegate can race disposal. Polling adds at most 20 ms of latency.
                    WaveInNative.Check(WaveInNative.waveInOpen(out handle, deviceIndex, ref format,
                        IntPtr.Zero, IntPtr.Zero, 0), "open microphone (16 kHz, 16-bit mono)");
                    if (stop.WaitOne(0)) return;
                    WaveInNative.Capabilities openedCaps;
                    UIntPtr handleValue = IntPtr.Size == 8
                        ? new UIntPtr(unchecked((ulong)handle.ToInt64()))
                        : new UIntPtr(unchecked((uint)handle.ToInt32()));
                    WaveInNative.Check(WaveInNative.waveInGetDevCaps(handleValue, out openedCaps,
                        (uint)Marshal.SizeOf(typeof(WaveInNative.Capabilities))), "verify opened microphone");
                    if (CreateDeviceId(deviceIndex, openedCaps.Name, openedCaps.Manufacturer,
                        openedCaps.Product) != deviceId)
                        throw new IOException("The microphone changed while opening it. Refresh setup and select it again.");

                    for (int i = 0; i < buffers.Length; i++)
                    {
                        if (stop.WaitOne(0)) return;
                        buffers[i] = new NativeBuffer(BufferBytes);
                        WaveInNative.Check(WaveInNative.waveInPrepareHeader(handle, buffers[i].Header,
                            WaveInNative.HeaderSize), "prepare microphone buffer");
                        buffers[i].Prepared = true;
                        WaveInNative.Check(WaveInNative.waveInAddBuffer(handle, buffers[i].Header,
                            WaveInNative.HeaderSize), "queue microphone buffer");
                    }
                    if (stop.WaitOne(0)) return;
                    WaveInNative.Check(WaveInNative.waveInStart(handle), "start microphone");
                    started = true;
                    SignalStartup();
                    byte[] data = new byte[BufferBytes];
                    var lastData = Stopwatch.StartNew();
                    int nextBuffer = 0;
                    while (!stop.WaitOne(20))
                    {
                        bool exhausted = true;
                        foreach (NativeBuffer pending in buffers)
                        {
                            var header = (WaveInNative.Header)Marshal.PtrToStructure(
                                pending.Header, typeof(WaveInNative.Header));
                            if ((header.Flags & 1) == 0) { exhausted = false; break; }
                        }
                        if (exhausted)
                            throw new IOException("All microphone capture buffers filled before they could "
                                + "be read. Audio may have been lost; stop listening and restart it.");
                        for (int completed = 0; completed < buffers.Length; completed++)
                        {
                            if (stop.WaitOne(0)) break;
                            NativeBuffer buffer = buffers[nextBuffer];
                            WaveInNative.Header header = (WaveInNative.Header)Marshal.PtrToStructure(
                                buffer.Header, typeof(WaveInNative.Header));
                            if ((header.Flags & 1) == 0) break; // WHDR_DONE, in queue order
                            if (header.BytesRecorded > BufferBytes || (header.BytesRecorded & 1) != 0)
                                throw new IOException("The microphone driver returned invalid PCM audio.");
                            if (header.BytesRecorded > 0)
                            {
                                int bytes = (int)header.BytesRecorded;
                                Marshal.Copy(buffer.Data, data, 0, bytes);
                                Interlocked.Exchange(ref level, MeasureLevel(data, bytes));
                                if (!stream.Append(data, bytes)) return;
                                lastData.Restart();
                            }
                            if (!stop.WaitOne(0))
                                WaveInNative.Check(WaveInNative.waveInAddBuffer(handle, buffer.Header,
                                    WaveInNative.HeaderSize), "read microphone; it may have been disconnected");
                            nextBuffer = (nextBuffer + 1) % buffers.Length;
                        }
                        if (lastData.ElapsedMilliseconds > 5000)
                            throw new IOException("The microphone stopped supplying audio. It may have been "
                                + "unplugged or disabled. Reconnect it and select it again in setup.");
                    }
                }
                catch (IOException error) { RecordFailure(error, started); }
                catch (InvalidOperationException error) { RecordFailure(error, started); }
                catch (ArgumentException error) { RecordFailure(error, started); }
                catch (SecurityException error) { RecordFailure(error, started); }
                catch (UnauthorizedAccessException error) { RecordFailure(error, started); }
                catch (DllNotFoundException error) { RecordFailure(error, started); }
                catch (EntryPointNotFoundException error) { RecordFailure(error, started); }
                catch (BadImageFormatException error) { RecordFailure(error, started); }
                finally
                {
                    // Publish startup failures before cleanup so the constructor can begin disposal.
                    if (!started) SignalStartup();
                    try
                    {
                        Exception failure = SpeechInputErrors.CaptureExpectedFailure(ReleaseNativeResources);
                        if (failure != null)
                        {
                            cleanupError = BlockCaptureFailure(failure);
                            stream.Fail(cleanupError);
                            Trace.TraceError("{0}", cleanupError);
                        }
                    }
                    finally { stop.Dispose(); }
                }
            }

            private void RecordFailure(Exception error, bool started)
            {
                if (!started) startupError = error;
                stream.Fail(error);
            }

            private void SignalStartup()
            {
                lock (startupGate)
                {
                    startupFinished = true;
                    Monitor.PulseAll(startupGate);
                }
            }

            private void ReleaseNativeResources()
            {
                if (handle != IntPtr.Zero)
                {
                    ReportCleanupError(WaveInNative.waveInReset(handle), "stop microphone");
                    foreach (NativeBuffer buffer in buffers)
                    {
                        if (buffer != null && buffer.Prepared)
                            ReportCleanupError(WaveInNative.waveInUnprepareHeader(handle, buffer.Header,
                                WaveInNative.HeaderSize), "release microphone buffer");
                    }
                    uint result = WaveInNative.waveInClose(handle);
                    // A failed close must not free memory still owned by a driver. Retry finitely,
                    // then retain that native allocation until process exit and block new captures.
                    // INVALHANDLE means the handle is already gone.
                    for (int retry = 0; result != 0 && result != 5 && retry < 20; retry++)
                    {
                        ReportCleanupError(result, "close microphone");
                        WaveInNative.waveInReset(handle);
                        foreach (NativeBuffer buffer in buffers)
                            if (buffer != null && buffer.Prepared)
                                WaveInNative.waveInUnprepareHeader(handle, buffer.Header, WaveInNative.HeaderSize);
                        Thread.Sleep(100);
                        result = WaveInNative.waveInClose(handle);
                    }
                    if (result != 0 && result != 5)
                    {
                        cleanupError = BlockCapture("The microphone driver did not release its handle "
                            + "after bounded cleanup. Audio delivery is stopped. Restart Jarvis manually after "
                            + "checking the device; another capture must not be opened.");
                        stream.Fail(cleanupError);
                        Trace.TraceError("{0}", cleanupError);
                        return;
                    }
                    ReportCleanupError(result, "close microphone");
                    handle = IntPtr.Zero;
                }
                foreach (NativeBuffer buffer in buffers)
                    if (buffer != null) buffer.Dispose();
            }

            private void ReportCleanupError(uint result, string operation)
            {
                if (result != 0) stream.Fail(WaveInNative.CreateError(result, operation));
            }

            public void Dispose()
            {
                if (Interlocked.Exchange(ref disposed, 1) != 0) return;
                // SAPI may be blocked in Read: wake it before waiting for the native worker.
                stream.Dispose();
                try { stop.Set(); }
                catch (ObjectDisposedException) { } // The worker may already have failed and cleaned up.
                bool released = worker.Join(3000);
                if (!released)
                {
                    var error = BlockCapture("The microphone driver did not stop within three seconds. "
                        + "Audio delivery has stopped; no new capture is safe. Restart Jarvis manually after "
                        + "checking the device or restart Windows if it remains busy.");
                    stream.Fail(error);
                    throw error;
                }
                if (cleanupError != null) throw cleanupError;
            }

            private sealed class NativeBuffer : IDisposable
            {
                public IntPtr Data;
                public IntPtr Header;
                public bool Prepared;

                public NativeBuffer(int size)
                {
                    try
                    {
                        Data = Marshal.AllocHGlobal(size);
                        Header = Marshal.AllocHGlobal((int)WaveInNative.HeaderSize);
                        var header = new WaveInNative.Header();
                        header.Data = Data;
                        header.BufferLength = (uint)size;
                        Marshal.StructureToPtr(header, Header, false);
                    }
                    catch { Dispose(); throw; }
                }

                public void Dispose()
                {
                    if (Header != IntPtr.Zero) Marshal.FreeHGlobal(Header);
                    if (Data != IntPtr.Zero) Marshal.FreeHGlobal(Data);
                    Header = Data = IntPtr.Zero;
                }
            }
        }
    }

    internal sealed class SpeechCleanupException : IOException
    {
        public SpeechCleanupException(string message) : base(message) { }
        public SpeechCleanupException(string message, Exception innerException) : base(message, innerException) { }
    }

    internal static class SpeechInputErrors
    {
        public static bool HasUnsafeCleanup(Exception error)
        {
            if (error == null) return false;
            if (error is SpeechCleanupException) return true;
            var aggregate = error as AggregateException;
            if (aggregate != null)
                foreach (Exception nested in aggregate.InnerExceptions)
                    if (HasUnsafeCleanup(nested)) return true;
            return HasUnsafeCleanup(error.InnerException);
        }

        public static Exception CaptureExpectedFailure(Action operation)
        {
            try { operation(); return null; }
            catch (IOException error) { return error; }
            catch (InvalidOperationException error) { return error; }
            catch (ArgumentException error) { return error; }
            catch (COMException error) { return error; }
            catch (SecurityException error) { return error; }
            catch (UnauthorizedAccessException error) { return error; }
            catch (DllNotFoundException error) { return error; }
            catch (EntryPointNotFoundException error) { return error; }
            catch (BadImageFormatException error) { return error; }
        }

        public static void CleanupPreservingFailure(Action cleanup, Exception primaryFailure = null)
        {
            // Called only from a cleanup-and-rethrow catch. A teardown error must not
            // replace the startup/attachment error; unsafe native ownership must also reach recovery.
            Exception error = CaptureExpectedFailure(cleanup);
            if (HasUnsafeCleanup(error))
            {
                throw new IOException("Speech attachment failed and the microphone driver could not be released.",
                    primaryFailure == null ? error : new AggregateException(primaryFailure, error));
            }
            if (error != null) Trace.TraceError("Local speech cleanup also failed: {0}", error);
        }

        public static void DisposeCapture(Action releaseInput, Action cancelRecognition, Action releaseRecognizer)
        {
            var failures = new List<Exception>();
            try
            {
                Exception failure = CaptureExpectedFailure(releaseInput);
                if (failure != null) failures.Add(failure);
            }
            finally
            {
                try
                {
                    Exception failure = CaptureExpectedFailure(cancelRecognition);
                    if (failure != null) failures.Add(failure);
                }
                finally
                {
                    Exception failure = CaptureExpectedFailure(releaseRecognizer);
                    if (failure != null) failures.Add(failure);
                    // Failed teardown cannot establish that SAPI (including its default input)
                    // or a native driver has relinquished capture ownership.
                    if (failures.Count != 0)
                        throw SpeechInput.BlockCaptureFailure(new AggregateException(failures));
                }
            }
        }
    }

    internal sealed class PcmAudioStream : Stream
    {
        private readonly object gate = new object();
        private readonly byte[] buffer;
        private int head;
        private int length;
        private long position;
        private bool closed;
        private Exception error;

        public PcmAudioStream(int capacity)
        {
            if (capacity <= 0 || (capacity & 1) != 0)
                throw new ArgumentOutOfRangeException("capacity");
            buffer = new byte[capacity];
        }

        public Exception Error { get { lock (gate) return error; } }
        internal int BufferedBytes { get { lock (gate) return length; } }
        internal int Capacity { get { return buffer.Length; } }

        public bool Append(byte[] data, int count)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (count < 0 || count > data.Length || (count & 1) != 0)
                throw new ArgumentOutOfRangeException("count");
            lock (gate)
            {
                if (closed || error != null) return false;
                if (count > buffer.Length - length)
                {
                    Fail(new IOException("Microphone audio overflowed its bounded buffer; recognition "
                        + "could not keep up. Stop listening and restart it. No audio was silently dropped."));
                    return false;
                }
                int tail = (head + length) % buffer.Length;
                int first = Math.Min(count, buffer.Length - tail);
                Buffer.BlockCopy(data, 0, buffer, tail, first);
                Buffer.BlockCopy(data, first, buffer, 0, count - first);
                length += count;
                Monitor.PulseAll(gate);
                return true;
            }
        }

        public void Fail(Exception failure)
        {
            if (failure == null) throw new ArgumentNullException("failure");
            lock (gate)
            {
                if (error == null) error = failure;
                Monitor.PulseAll(gate);
            }
        }

        public override int Read(byte[] data, int offset, int count)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (offset < 0 || count < 0 || offset > data.Length - count)
                throw new ArgumentOutOfRangeException("offset");
            if (count == 0) return 0;
            lock (gate)
            {
                int total = 0;
                // SAPI's IStream reader treats a short read as end-of-stream. Fill its request
                // across capture packets, returning fewer bytes only when disposal signals EOF.
                while (total < count)
                {
                    while (!closed && error == null && length == 0) Monitor.Wait(gate);
                    if (closed) return total;
                    if (error != null) throw new IOException("Local microphone capture failed.", error);
                    int copied = Math.Min(count - total, length);
                    int first = Math.Min(copied, buffer.Length - head);
                    Buffer.BlockCopy(buffer, head, data, offset + total, first);
                    Buffer.BlockCopy(buffer, 0, data, offset + total + first, copied - first);
                    head = (head + copied) % buffer.Length;
                    length -= copied;
                    position += copied;
                    total += copied;
                }
                return total;
            }
        }

        protected override void Dispose(bool disposing)
        {
            lock (gate)
            {
                closed = true;
                length = 0;
                Array.Clear(buffer, 0, buffer.Length);
                Monitor.PulseAll(gate);
            }
            base.Dispose(disposing);
        }

        public override bool CanRead { get { lock (gate) return !closed; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        // System.Speech's IStream adapter queries Length and Position even when CanSeek
        // is false. A live input has no finite length; EOF is signalled only on disposal.
        public override long Length { get { return long.MaxValue; } }
        public override long Position
        {
            get { lock (gate) return position; }
            set { Seek(value, SeekOrigin.Begin); }
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin)
        {
            lock (gate)
            {
                // Allow SAPI's position queries, but never rewind, skip or fabricate samples.
                if ((origin == SeekOrigin.Current && offset == 0)
                    || (origin == SeekOrigin.Begin && offset == position)) return position;
                throw new NotSupportedException("Live microphone audio cannot be repositioned.");
            }
        }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] data, int offset, int count) { throw new NotSupportedException(); }
    }

    internal static class WaveInNative
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct Capabilities
        {
            public ushort Manufacturer;
            public ushort Product;
            public uint DriverVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string Name;
            public uint Formats;
            public ushort Channels;
            public ushort Reserved;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 2)]
        internal struct Format
        {
            public ushort FormatTag;
            public ushort Channels;
            public uint SamplesPerSecond;
            public uint AverageBytesPerSecond;
            public ushort BlockAlign;
            public ushort BitsPerSample;
            public ushort ExtraSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct Header
        {
            public IntPtr Data;
            public uint BufferLength;
            public uint BytesRecorded;
            public UIntPtr User;
            public uint Flags;
            public uint Loops;
            public IntPtr Next;
            public UIntPtr Reserved;
        }

        internal static readonly uint HeaderSize = (uint)Marshal.SizeOf(typeof(Header));

        internal static void Check(uint result, string operation)
        {
            if (result != 0) throw CreateError(result, operation);
        }

        internal static IOException CreateError(uint result, string operation)
        {
            var message = new StringBuilder(256);
            waveInGetErrorText(result, message, (uint)message.Capacity);
            return new IOException("Cannot " + operation + " (WinMM " + result.ToString(
                CultureInfo.InvariantCulture) + "): " + message.ToString()
                + " Check microphone permissions, connection and setup selection.");
        }

        [DllImport("winmm.dll")]
        internal static extern uint waveInGetNumDevs();
        [DllImport("winmm.dll", EntryPoint = "waveInGetDevCapsW", CharSet = CharSet.Unicode)]
        internal static extern uint waveInGetDevCaps(UIntPtr device, out Capabilities caps, uint size);
        [DllImport("winmm.dll")]
        internal static extern uint waveInOpen(out IntPtr handle, uint device, ref Format format,
            IntPtr callback, IntPtr instance, uint flags);
        [DllImport("winmm.dll")]
        internal static extern uint waveInPrepareHeader(IntPtr handle, IntPtr header, uint size);
        [DllImport("winmm.dll")]
        internal static extern uint waveInUnprepareHeader(IntPtr handle, IntPtr header, uint size);
        [DllImport("winmm.dll")]
        internal static extern uint waveInAddBuffer(IntPtr handle, IntPtr header, uint size);
        [DllImport("winmm.dll")]
        internal static extern uint waveInStart(IntPtr handle);
        [DllImport("winmm.dll")]
        internal static extern uint waveInReset(IntPtr handle);
        [DllImport("winmm.dll")]
        internal static extern uint waveInClose(IntPtr handle);
        [DllImport("winmm.dll", EntryPoint = "waveInGetErrorTextW", CharSet = CharSet.Unicode)]
        private static extern uint waveInGetErrorText(uint error, StringBuilder text, uint size);
    }
}
