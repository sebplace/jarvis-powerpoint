[CmdletBinding()]
param(
    [ValidateSet("anycpu", "x86")]
    [string]$Platform = "anycpu",
    [switch]$ProbeMicrophone
)

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $PSScriptRoot
$compiler = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $compiler) {
    throw "The Windows .NET Framework C# compiler and System.Speech are required."
}
$speech = Join-Path (Split-Path -Parent $compiler) "WPF\System.Speech.dll"
if (-not (Test-Path -LiteralPath $speech)) {
    throw "The Windows .NET Framework System.Speech assembly is required."
}

# All scratch outputs stay in this project's tests directory, never in a global temp directory.
$output = Join-Path $PSScriptRoot ("speech-input-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $output | Out-Null
$source = Join-Path $output "SpeechInputTestProgram.cs"
$executable = Join-Path $output "SpeechInputTestProgram.exe"
$testCode = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Speech.Recognition;
using System.Threading;
using System.Windows.Forms;
using JarvisPowerPoint;

internal static class SpeechInputTestProgram
{
    private static int assertions;

    private static void Assert(bool condition, string description)
    {
        if (!condition) throw new Exception(description);
        assertions++;
    }

    private static void Throws<T>(Action action, string description) where T : Exception
    {
        try { action(); }
        catch (T) { assertions++; return; }
        throw new Exception(description);
    }

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length == 1 && args[0] == "--probe-microphone")
            {
                HardwareProbe();
                return 0;
            }
            if (args.Length == 2 && args[0] == "--cleanup-failure")
            {
                CleanupFailureStage(int.Parse(args[1]));
                return 0;
            }
            DeviceIdentity();
            ExpectedFailures();
            InputLevels();
            StreamBoundsAndOrdering();
            SapiStreamContract();
            BlockedReaders();
            ConcurrentOrdering();
            NativeLayouts();
            DialogWithoutCapture();
            EnumerateWithoutCapture();
            ProcessCaptureSafety();
            Console.WriteLine("Passed " + assertions + " assertions; no microphone was opened.");
            Console.WriteLine("No-capture suite excludes live speech, unplug and driver lifecycle; "
                + "-ProbeMicrophone additionally checks bounded capture/disposal/reopen.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void ProcessCaptureSafety()
    {
        Assert(!SpeechInput.CaptureBlocked, "Capture was blocked without a native ownership failure.");
        SpeechInput.EnsureCaptureAvailable();
        string cleanupOrder = "";
        SpeechInputErrors.DisposeCapture(delegate { cleanupOrder += "input "; },
            delegate { cleanupOrder += "cancel "; }, delegate { cleanupOrder += "recognizer"; });
        Assert(cleanupOrder == "input cancel recognizer" && !SpeechInput.CaptureBlocked,
            "Successful teardown reordered cleanup or blocked future capture.");
        var flags = BindingFlags.NonPublic | BindingFlags.Instance;
        var inputFailure = new FailingInput();
        Type sessionType = typeof(SetupDialog).GetNestedType("TestSession", BindingFlags.NonPublic);
        object testSession = System.Runtime.Serialization.FormatterServices.GetUninitializedObject(sessionType);
        sessionType.GetField("input", flags).SetValue(testSession, inputFailure);
        using (var dialog = new SetupDialog("en-US", "default", false))
        {
            MakeSelectionAvailable(dialog);
            Field<CheckBox>(dialog, "skip").Checked = true;
            Assert(Field<Button>(dialog, "test").Enabled && Field<Button>(dialog, "finish").Enabled,
                "Synthetic available selection did not enable test and explicit skip/save.");
            typeof(SetupDialog).GetField("tested", flags).SetValue(dialog, true);
            typeof(SetupDialog).GetField("session", flags).SetValue(dialog, testSession);
            typeof(SetupDialog).GetMethod("Finish", flags).Invoke(dialog, new object[] { null, EventArgs.Empty });
            Assert(inputFailure.DisposeCount == 1 && Field<object>(dialog, "session") == null
                && sessionType.GetField("input", flags).GetValue(testSession) == null,
                "Failed cleanup retained or skipped a capture without a recognizer.");
            Assert(SpeechInput.CaptureBlocked && !Field<Button>(dialog, "finish").Enabled
                && !Field<Button>(dialog, "test").Enabled && dialog.DialogResult != DialogResult.OK,
                "Cleanup during Finish saved settings or did not process-latch test/save.");
            Assert(!Field<bool>(dialog, "tested") && Field<string>(dialog, "failure") != null,
                "Failed teardown preserved a successful test or hid the failure.");
            ((IDisposable)testSession).Dispose();
            Assert(inputFailure.DisposeCount == 1, "Repeated test teardown disposed the retained input twice.");
        }

        var original = new Exception[] { new IOException("input"), new COMException("cancel"),
            new InvalidOperationException("recognizer") };
        cleanupOrder = "";
        try
        {
            SpeechInputErrors.DisposeCapture(delegate { cleanupOrder += "input "; throw original[0]; },
                delegate { cleanupOrder += "cancel "; throw original[1]; },
                delegate { cleanupOrder += "recognizer"; throw original[2]; });
            throw new Exception("Uncertain cleanup was not reported.");
        }
        catch (SpeechCleanupException error)
        {
            var failures = error.InnerException as AggregateException;
            Assert(cleanupOrder == "input cancel recognizer" && failures != null
                && failures.InnerExceptions.Count == 3, "Cleanup failure skipped teardown or lost an error.");
            for (int i = 0; i < original.Length; i++)
                Assert(object.ReferenceEquals(original[i], failures.InnerExceptions[i]),
                    "Cleanup replaced an original exception.");
        }
        Exception firstFailure = (Exception)typeof(SpeechInput).GetField("captureFailure",
            BindingFlags.NonPublic | BindingFlags.Static).GetValue(null);
        typeof(SpeechInput).GetMethod("BlockCapture", BindingFlags.NonPublic | BindingFlags.Static)
            .Invoke(null, new object[] { "Simulated unsafe ownership; no native calls." });
        SpeechInputErrors.DisposeCapture(delegate { }, delegate { }, delegate { });
        Assert(SpeechInput.CaptureBlocked, "Unsafe ownership did not latch across capture sessions.");
        try
        {
            SpeechInput.EnsureCaptureAvailable();
            throw new Exception("Unsafe ownership allowed another capture session.");
        }
        catch (IOException failure)
        {
            Assert(SpeechInputErrors.HasUnsafeCleanup(failure), "Process latch lost the unsafe ownership reason.");
            Assert(object.ReferenceEquals(firstFailure, failure.InnerException),
                "Later cleanup replaced the first unsafe ownership failure.");
        }
        foreach (string culture in new[] { "fr-FR", "en-US" })
        {
            using (var dialog = new SetupDialog(culture, "default", false))
            {
                MakeSelectionAvailable(dialog);
                Field<CheckBox>(dialog, "skip").Checked = true;
                typeof(SetupDialog).GetField("tested", flags).SetValue(dialog, true);
                typeof(SetupDialog).GetMethod("PollTest", flags).Invoke(dialog, new object[] { null, EventArgs.Empty });
                var test = (Control)typeof(SetupDialog).GetField("test", flags).GetValue(dialog);
                var finish = (Control)typeof(SetupDialog).GetField("finish", flags).GetValue(dialog);
                var notice = (Control)typeof(SetupDialog).GetField("availability", flags).GetValue(dialog);
                Assert(!test.Enabled && !finish.Enabled, "Unsafe cleanup left setup actions enabled.");
                Assert(notice.Text.Contains(culture == "fr-FR" ? "red\u00e9marrez" : "restart"),
                    "Unsafe setup cleanup did not explain the manual restart requirement.");
                Assert(Field<Label>(dialog, "status").Text == SpeechInput.RestartRequiredMessage(culture == "en-US"),
                    "Test/skip success text hid the restart requirement.");
                Assert(!Field<System.Windows.Forms.Timer>(dialog, "timer").Enabled,
                    "Restart-required setup kept polling after capture became permanently blocked.");
                typeof(SetupDialog).GetMethod("StartTest", flags).Invoke(dialog, null);
                typeof(SetupDialog).GetMethod("Finish", flags).Invoke(dialog, new object[] { null, EventArgs.Empty });
                Assert(Field<object>(dialog, "session") == null && dialog.DialogResult != DialogResult.OK,
                    "Direct setup test/finish calls bypassed unsafe ownership.");
                Field<ComboBox>(dialog, "languages").SelectedIndex = culture == "fr-FR" ? 1 : 0;
                Field<CheckBox>(dialog, "skip").Checked = true;
                Assert(!test.Enabled && !finish.Enabled && SpeechInput.CaptureBlocked,
                    "Changing language or skipping the test cleared unsafe ownership.");
                typeof(SetupDialog).GetMethod("LoadChoices", flags).Invoke(dialog, new object[] { "default" });
                Assert(!test.Enabled && !finish.Enabled && SpeechInput.CaptureBlocked,
                    "Refreshing setup cleared unsafe ownership.");
            }
        }
        try
        {
            Activator.CreateInstance(sessionType, BindingFlags.Instance | BindingFlags.Public,
                null, new object[] { null, "default" }, null);
            throw new Exception("Test-session construction bypassed the process latch.");
        }
        catch (TargetInvocationException error)
        {
            Assert(SpeechInputErrors.HasUnsafeCleanup(error.InnerException),
                "Test-session construction reached recognizer setup before checking ownership.");
        }
    }

    private sealed class FailingInput : IDisposable
    {
        public int DisposeCount;
        public void Dispose() { DisposeCount++; throw new IOException("Synthetic native ownership failure."); }
    }

    private static void CleanupFailureStage(int stage)
    {
        Assert(!SpeechInput.CaptureBlocked, "Isolated cleanup started blocked.");
        int calls = 0;
        var original = new IOException("Synthetic cleanup stage " + stage);
        Action operation = delegate { if (calls++ == stage) throw original; };
        try
        {
            SpeechInputErrors.DisposeCapture(operation, operation, operation);
            throw new Exception("Isolated cleanup failure was hidden.");
        }
        catch (SpeechCleanupException error)
        {
            var errors = error.InnerException as AggregateException;
            Assert(SpeechInput.CaptureBlocked && calls == 3, "Cleanup stage did not block capture or skipped later teardown.");
            Assert(errors != null && errors.InnerExceptions.Count == 1
                && object.ReferenceEquals(errors.InnerExceptions[0], original), "Cleanup lost its original failure.");
        }
        SpeechInputErrors.DisposeCapture(delegate { }, delegate { }, delegate { });
        Throws<IOException>(SpeechInput.EnsureCaptureAvailable, "Successful later teardown cleared the safety latch.");
        Console.WriteLine("Passed " + assertions + " isolated cleanup-stage " + stage + " assertions; no capture.");
    }

    private static void MakeSelectionAvailable(SetupDialog dialog)
    {
        Field<Dictionary<string, RecognizerInfo>>(dialog, "recognizers").Add("fr-FR", null);
        Field<Dictionary<string, RecognizerInfo>>(dialog, "recognizers").Add("en-US", null);
        typeof(SetupDialog).GetField("devices", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(dialog,
            new List<AudioInputDevice> { new AudioInputDevice("default", "Default"),
                new AudioInputDevice(SpeechInput.CreateDeviceId(0, "Synthetic", 0, 0), "Synthetic") });
    }

    private static void InputLevels()
    {
        Assert(SpeechInput.GetLevel(null) == 0, "Default capture lifetime should not supply a PCM level.");
        Assert(SpeechInput.MeasureLevel(new byte[8], 8) == 0, "Silence did not read zero.");
        Assert(SpeechInput.MeasureLevel(new byte[0], 0) == 0, "Empty packet did not read zero.");
        Assert(SpeechInput.MeasureLevel(new byte[] { 0, 128, 255, 127 }, 4) == 100,
            "Full-scale signed PCM did not read 100.");
        int quiet = SpeechInput.MeasureLevel(new byte[] { 0, 1 }, 2);
        int loud = SpeechInput.MeasureLevel(new byte[] { 0, 16 }, 2);
        Assert(quiet > 0 && quiet < loud && loud < 100, "PCM level does not track amplitude.");
        Throws<ArgumentOutOfRangeException>(delegate { SpeechInput.MeasureLevel(new byte[4], 3); },
            "Meter accepted an incomplete PCM sample.");
    }

    private static void ExpectedFailures()
    {
        Exception[] expected = {
            new IOException(), new InvalidOperationException(), new ArgumentException(),
            new COMException(), new SecurityException(), new UnauthorizedAccessException(),
            new DllNotFoundException(), new EntryPointNotFoundException(), new BadImageFormatException()
        };
        foreach (Exception error in expected)
        {
            Exception original = error;
            Assert(object.ReferenceEquals(original,
                SpeechInputErrors.CaptureExpectedFailure(delegate { throw original; })),
                "Expected failure was swallowed or replaced.");
        }
        Assert(SpeechInputErrors.CaptureExpectedFailure(delegate { }) == null,
            "A successful operation reported a failure.");
        Throws<NullReferenceException>(delegate {
            SpeechInputErrors.CaptureExpectedFailure(delegate { throw new NullReferenceException(); });
        }, "Programming error was swallowed by an overly broad production catch.");
        Assert(SpeechInputErrors.HasUnsafeCleanup(new IOException("attach",
            new AggregateException(new IOException(), new SpeechCleanupException("busy driver")))),
            "Nested native cleanup failure did not block automatic capture restart.");
        Assert(!SpeechInputErrors.HasUnsafeCleanup(new IOException("unplug")),
            "Ordinary disconnect was treated as unsafe native cleanup.");
        Assert(!SpeechInputErrors.HasUnsafeCleanup(null), "Null cleanup failure was unsafe.");
        try
        {
            SpeechInputErrors.CleanupPreservingFailure(
                delegate { throw new SpeechCleanupException("driver retained native ownership"); },
                new ArgumentException("SAPI rejected stream"));
            throw new Exception("Unsafe cleanup after attach was hidden, allowing another capture.");
        }
        catch (IOException failure)
        {
            Assert(SpeechInputErrors.HasUnsafeCleanup(failure) && failure.InnerException is AggregateException,
                "Unsafe attach cleanup lost either the original or cleanup failure.");
        }
        var startup = new IOException("Original attach failure.");
        try
        {
            try { throw startup; }
            catch
            {
                SpeechInputErrors.CleanupPreservingFailure(delegate { throw new IOException("Cleanup failure."); });
                throw;
            }
        }
        catch (IOException error)
        {
            Assert(object.ReferenceEquals(startup, error), "Cleanup masked the original failure.");
        }
    }

    private static void DeviceIdentity()
    {
        string id = SpeechInput.CreateDeviceId(2, "Microphone \u00e9 : USB", 10, 20);
        var device = new AudioInputDevice(id, "Microphone \u00e9 : USB");
        var devices = new List<AudioInputDevice>();
        devices.Add(new AudioInputDevice("default", "Default (Windows)"));
        devices.Add(device);
        Assert(device.Name == device.ToString(), "Device display name mismatch.");
        Assert(SpeechInput.ResolveDeviceIndex(id, devices) == 2, "Selected index was not resolved.");
        Assert(id == SpeechInput.CreateDeviceId(2, device.Name, 10, 20), "IDs must be deterministic.");
        Throws<IOException>(delegate { SpeechInput.ResolveDeviceIndex("default", devices); },
            "Default must not be mapped to waveIn index zero.");
        Throws<IOException>(delegate { SpeechInput.ResolveDeviceIndex(null, devices); }, "Null ID accepted.");
        Throws<IOException>(delegate { SpeechInput.ResolveDeviceIndex("", devices); }, "Empty ID accepted.");
        Throws<IOException>(delegate { SpeechInput.ResolveDeviceIndex("wavein:xyz", devices); }, "Malformed ID accepted.");
        Throws<IOException>(delegate {
            SpeechInput.ResolveDeviceIndex(SpeechInput.CreateDeviceId(1, device.Name, 10, 20), devices);
        }, "Reordered device silently substituted.");
        Throws<IOException>(delegate {
            SpeechInput.ResolveDeviceIndex(SpeechInput.CreateDeviceId(2, "Other microphone", 10, 20), devices);
        }, "Changed name silently substituted.");
        Throws<IOException>(delegate {
            SpeechInput.ResolveDeviceIndex(SpeechInput.CreateDeviceId(2, device.Name, 10, 21), devices);
        }, "Changed product silently substituted.");
        devices.Remove(device);
        Throws<IOException>(delegate { SpeechInput.ResolveDeviceIndex(id, devices); }, "Missing device silently substituted.");
        Throws<ArgumentNullException>(delegate { SpeechInput.Attach(null, "default"); }, "Null engine accepted.");
        Assert(SpeechInput.GetError(null) == null, "Default lifetime error must be null.");
    }

    private static void StreamBoundsAndOrdering()
    {
        Throws<ArgumentOutOfRangeException>(delegate { new PcmAudioStream(0); }, "Zero capacity accepted.");
        Throws<ArgumentOutOfRangeException>(delegate { new PcmAudioStream(3); }, "Non-PCM-aligned capacity accepted.");
        using (var stream = new PcmAudioStream(8))
        {
            Assert(stream.Capacity == 8 && stream.BufferedBytes == 0, "Incorrect empty capacity.");
            Assert(stream.Length == long.MaxValue && stream.Position == 0, "Live stream metadata is incorrect.");
            Assert(stream.CanRead && !stream.CanWrite && !stream.CanSeek, "Stream capabilities are incorrect.");
            Assert(stream.Read(new byte[0], 0, 0) == 0, "Zero-count Read blocked.");
            Throws<ArgumentNullException>(delegate { stream.Read(null, 0, 1); }, "Null read accepted.");
            Throws<ArgumentOutOfRangeException>(delegate { stream.Read(new byte[2], 1, 2); }, "Invalid read bounds accepted.");
            Throws<ArgumentOutOfRangeException>(delegate { stream.Append(new byte[2], 1); }, "Odd PCM append accepted.");
            Assert(stream.Seek(0, SeekOrigin.Current) == 0, "SAPI position query failed.");
            Throws<NotSupportedException>(delegate { stream.Seek(1, SeekOrigin.Begin); }, "Repositioning allowed.");
            Throws<NotSupportedException>(delegate { stream.Write(new byte[2], 0, 2); }, "Write allowed.");
            Assert(stream.Append(new byte[] { 0, 1, 2, 3, 4, 5 }, 6), "Initial append failed.");
            byte[] read = new byte[12];
            Assert(stream.Read(read, 2, 3) == 3, "A short, unaligned consumer read was not supported.");
            Assert(read[2] == 0 && read[3] == 1 && read[4] == 2, "Initial read changed data.");
            Assert(stream.Append(new byte[] { 6, 7, 8, 9 }, 4), "Wrapped append failed.");
            Assert(stream.Read(read, 1, 7) == 7, "Read failed to return requested bytes.");
            for (int i = 0; i < 7; i++) Assert(read[i + 1] == i + 3, "Wrapped FIFO order broken.");
            Assert(stream.BufferedBytes == 0, "Drain left queued bytes.");
            Assert(stream.Position == 10 && stream.Seek(0, SeekOrigin.Current) == 10,
                "Read position did not advance monotonically.");
            Throws<NotSupportedException>(delegate { stream.Position = 0; }, "Live audio rewind allowed.");
            Assert(stream.Append(new byte[8], 8), "Exact capacity append rejected.");
            Assert(!stream.Append(new byte[2], 2), "Overflow silently accepted.");
            Assert(stream.Error is IOException, "Overflow was not surfaced.");
            Assert(stream.BufferedBytes <= stream.Capacity, "Bounded buffer exceeded capacity.");
            Throws<IOException>(delegate { stream.Read(read, 0, 1); }, "Overflow read did not fail.");
            Assert(!stream.Append(new byte[0], 0), "Failed stream resumed silently.");
        }
        var disposed = new PcmAudioStream(8);
        disposed.Append(new byte[8], 8);
        disposed.Dispose();
        disposed.Dispose();
        Assert(!disposed.CanRead && disposed.BufferedBytes == 0, "Disposed audio not cleared.");
        Assert(disposed.Read(new byte[2], 0, 2) == 0, "Disposed Read did not return EOF.");
        Assert(!disposed.Append(new byte[2], 2), "Disposed append accepted.");
    }

    private static void SapiStreamContract()
    {
        using (var stream = new PcmAudioStream(8))
        {
            Type wrapperType = typeof(SpeechRecognitionEngine).Assembly.GetType(
                "System.Speech.Internal.SapiInterop.SpStreamWrapper", true);
            object wrapper = Activator.CreateInstance(wrapperType,
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                null, new object[] { stream }, null);
            Assert(wrapper != null, "System.Speech could not adapt the live stream.");
            stream.Append(new byte[] { 1, 2, 3, 4 }, 4);
            byte[] read = new byte[4];
            wrapperType.GetMethod("Read").Invoke(wrapper, new object[] { read, read.Length, IntPtr.Zero });
            Assert(read[0] == 1 && read[3] == 4 && stream.Position == 4,
                "System.Speech adapter failed to read PCM data.");
            wrapperType.GetMethod("Seek").Invoke(wrapper, new object[] { 0L, (int)SeekOrigin.Current, IntPtr.Zero });
            object[] stat = { null, 0 };
            wrapperType.GetMethod("Stat").Invoke(wrapper, stat);
            Assert(((System.Runtime.InteropServices.ComTypes.STATSTG)stat[0]).cbSize == long.MaxValue,
                "System.Speech saw a finite stream length.");
        }
    }

    private static void BlockedReaders()
    {
        for (int scenario = 0; scenario < 5; scenario++)
        {
            var stream = new PcmAudioStream(8);
            var entered = new ManualResetEvent(false);
            var done = new ManualResetEvent(false);
            int result = -1;
            Exception failure = null;
            var thread = new Thread(delegate()
            {
                entered.Set();
                try { result = stream.Read(new byte[8], 0, 8); }
                catch (Exception error) { failure = error; }
                finally { done.Set(); }
            });
            thread.IsBackground = true;
            thread.Start();
            try
            {
                Assert(entered.WaitOne(2000), "Reader did not start.");
                Assert(!done.WaitOne(100), "Empty Read did not block.");
                if (scenario == 0)
                {
                    stream.Append(new byte[4], 4);
                    Assert(!done.WaitOne(100), "A partial native packet was mistaken for end-of-stream.");
                    stream.Append(new byte[4], 4);
                }
                if (scenario == 1) stream.Dispose();
                if (scenario == 2) stream.Fail(new IOException("Simulated unplug."));
                if (scenario == 3 || scenario == 4)
                {
                    stream.Append(new byte[4], 4);
                    Assert(!done.WaitOne(100), "Partial SAPI read unexpectedly completed.");
                    if (scenario == 3) stream.Dispose();
                    else stream.Fail(new IOException("Simulated unplug after partial read."));
                }
                Assert(done.WaitOne(2000), "Append/dispose/failure did not unblock Read.");
                Assert(thread.Join(2000), "Reader did not terminate.");
                if (scenario == 0) Assert(result == 8 && failure == null, "Append wake-up failed.");
                if (scenario == 1) Assert(result == 0 && failure == null, "Dispose wake-up failed.");
                if (scenario == 2)
                    Assert(failure is IOException && failure.InnerException.Message == "Simulated unplug.",
                        "Native error did not reach the blocked reader.");
                if (scenario == 3) Assert(result == 4 && failure == null, "Dispose did not release a partial read.");
                if (scenario == 4)
                    Assert(failure is IOException && failure.InnerException.Message == "Simulated unplug after partial read.",
                        "A partial read concealed a capture failure.");
            }
            finally
            {
                stream.Dispose();
                thread.Join(2000);
                entered.Dispose();
                done.Dispose();
            }
        }
    }

    private static void ConcurrentOrdering()
    {
        using (var stream = new PcmAudioStream(4096))
        {
            const int bytes = 20000;
            Exception failure = null;
            int consumed = 0;
            var reader = new Thread(delegate()
            {
                try
                {
                    byte[] chunk = new byte[73];
                    while (consumed < bytes)
                    {
                        int read = stream.Read(chunk, 0, Math.Min(chunk.Length, bytes - consumed));
                        if (read == 0) throw new Exception("Unexpected EOF.");
                        for (int i = 0; i < read; i++)
                            if (chunk[i] != (byte)((consumed + i) % 251)) throw new Exception("Concurrent FIFO corruption.");
                        consumed += read;
                    }
                }
                catch (Exception error) { failure = error; }
            });
            reader.IsBackground = true;
            reader.Start();
            try
            {
                var timeout = System.Diagnostics.Stopwatch.StartNew();
                for (int offset = 0; offset < bytes; offset += 100)
                {
                    while (stream.BufferedBytes > 3900)
                    {
                        if (timeout.ElapsedMilliseconds > 5000) throw new Exception("Concurrent reader stalled.");
                        Thread.Sleep(1);
                    }
                    byte[] chunk = new byte[100];
                    for (int i = 0; i < chunk.Length; i++) chunk[i] = (byte)((offset + i) % 251);
                    if (!stream.Append(chunk, chunk.Length)) throw new Exception("Unexpected concurrent overflow.");
                }
                Assert(reader.Join(5000), "Concurrent reader did not finish.");
                Assert(failure == null && consumed == bytes, "Concurrent audio ordering failed.");
            }
            finally { stream.Dispose(); reader.Join(2000); }
        }
    }

    private static void NativeLayouts()
    {
        Assert(Marshal.SizeOf(typeof(WaveInNative.Format)) == 18, "WAVEFORMATEX layout mismatch.");
        Assert(Marshal.SizeOf(typeof(WaveInNative.Capabilities)) == 80, "WAVEINCAPSW layout mismatch.");
        Assert(WaveInNative.HeaderSize == (IntPtr.Size == 8 ? 48 : 32), "WAVEHDR layout mismatch.");
        Assert(Marshal.OffsetOf(typeof(WaveInNative.Header), "Flags").ToInt32()
            == (IntPtr.Size == 8 ? 24 : 16), "WAVEHDR flags offset mismatch.");
    }

    private static T Field<T>(SetupDialog dialog, string name)
    {
        return (T)typeof(SetupDialog).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(dialog);
    }

    private static void DialogWithoutCapture()
    {
        using (var dialog = new SetupDialog("en-US", "default"))
        {
            Assert(dialog.SelectedCulture == "en-US", "Initial English culture ignored.");
            Assert(dialog.SelectedMicrophoneId == "default", "Initial default microphone ignored.");
            Assert(Field<object>(dialog, "session") == null, "Opening setup activated capture.");
            Assert(!Field<Button>(dialog, "finish").Enabled, "Untested save did not require explicit skip.");
            Assert(Field<Label>(dialog, "privacy").Text.Contains("Audio stays local"), "English privacy notice missing.");
            Field<ComboBox>(dialog, "languages").SelectedIndex = 0;
            Assert(dialog.SelectedCulture == "fr-FR", "Language change failed.");
            Assert(dialog.Text.Contains("Configuration"), "French UI did not update.");
            Assert(Field<Label>(dialog, "privacy").Text.Contains("audio reste local"), "French privacy notice missing.");
            Assert(Field<ComboBox>(dialog, "microphones").Items[0].ToString() == "Par d\u00e9faut (Windows)",
                "Default device label was not localized.");
            Field<CheckBox>(dialog, "skip").Checked = true;
            Assert(Field<object>(dialog, "session") == null, "Skipping test activated capture.");
            Field<ComboBox>(dialog, "languages").SelectedIndex = 1;
            Assert(!Field<CheckBox>(dialog, "skip").Checked, "Selection change did not reset skip.");
            dialog.DialogResult = DialogResult.Cancel;
            Assert(Field<object>(dialog, "session") == null, "Cancel activated capture.");
        }
        using (var dialog = new SetupDialog("fr-FR", "wavein:99:0:0:bm90LXByZXNlbnQ="))
        {
            Assert(dialog.SelectedMicrophoneId == "wavein:99:0:0:bm90LXByZXNlbnQ=",
                "Missing saved device silently fell back to default.");
            Assert(!Field<Button>(dialog, "test").Enabled, "Missing device allowed test.");
            Field<CheckBox>(dialog, "skip").Checked = true;
            Assert(!Field<Button>(dialog, "finish").Enabled, "Missing device allowed save even after skip.");
            Assert(Field<object>(dialog, "session") == null, "Missing-device setup activated capture.");
        }
        using (var dialog = new SetupDialog("fr-FR", "default"))
        {
            Field<Dictionary<string, System.Speech.Recognition.RecognizerInfo>>(dialog, "recognizers").Clear();
            Field<ComboBox>(dialog, "languages").SelectedIndex = 1;
            Field<CheckBox>(dialog, "skip").Checked = true;
            Assert(!Field<Button>(dialog, "test").Enabled, "Missing recognizer allowed test.");
            Assert(!Field<Button>(dialog, "finish").Enabled, "Missing recognizer allowed save.");
            Assert(Field<Label>(dialog, "availability").Text.Contains("not installed"),
                "Missing recognizer installation guidance was not shown.");
            Assert(Field<object>(dialog, "session") == null, "Missing-recognizer setup activated capture.");
            dialog.Dispose();
            Assert(!Field<System.Windows.Forms.Timer>(dialog, "timer").Enabled, "Dialog disposal left timer running.");
        }
        using (var dialog = new SetupDialog("fr-FR", "default"))
        {
            typeof(SetupDialog).GetField("enumerationFailure", BindingFlags.NonPublic | BindingFlags.Instance)
                .SetValue(dialog, "Simulated enumeration failure.");
            typeof(SetupDialog).GetField("devices", BindingFlags.NonPublic | BindingFlags.Instance)
                .SetValue(dialog, new List<AudioInputDevice>());
            Field<ComboBox>(dialog, "languages").SelectedIndex = 1;
            Field<CheckBox>(dialog, "skip").Checked = true;
            Assert(Field<ComboBox>(dialog, "microphones").Items[0].ToString().Contains("enumeration failed"),
                "Enumeration failure looked like a usable default microphone.");
            Assert(!Field<Button>(dialog, "test").Enabled && !Field<Button>(dialog, "finish").Enabled,
                "Enumeration failure allowed test/save.");
            Assert(Field<Label>(dialog, "availability").Text.Contains("Simulated enumeration failure."),
                "Enumeration failure was hidden.");
        }
    }

    private static void EnumerateWithoutCapture()
    {
        IList<AudioInputDevice> devices = SpeechInput.GetDevices();
        Assert(devices.Count >= 1 && devices[0].Id == "default", "Enumeration omitted Default.");
        var ids = new HashSet<string>();
        foreach (AudioInputDevice device in devices)
        {
            Assert(ids.Add(device.Id), "Duplicate enumerated ID.");
            Assert(!string.IsNullOrEmpty(device.Name), "Device has no display name.");
            if (device.Id != "default")
                SpeechInput.ResolveDeviceIndex(device.Id, devices);
        }
        Console.WriteLine("Read-only enumeration: " + (devices.Count - 1) + " waveIn device(s), plus Default.");
    }

    private static void HardwareProbe()
    {
        AudioInputDevice selected = null;
        foreach (AudioInputDevice device in SpeechInput.GetDevices())
            if (device.Id != "default") { selected = device; break; }
        if (selected == null) throw new IOException("No enumerated microphone is available for the opt-in probe.");
        RecognizerInfo recognizer = null;
        foreach (RecognizerInfo candidate in SpeechRecognitionEngine.InstalledRecognizers())
            if (candidate.Culture.Name == "fr-FR" || candidate.Culture.Name == "en-US")
            { recognizer = candidate; break; }
        if (recognizer == null) throw new InvalidOperationException("No installed FR/EN recognizer is available.");

        Console.WriteLine("Opt-in probe: two 2-second selected-device sessions, memory-only audio; no phrase output.");
        for (int attempt = 1; attempt <= 2; attempt++)
        {
            var errors = new List<Exception>();
            var engine = new SpeechRecognitionEngine(recognizer);
            IDisposable input = null;
            Exception completedError = null;
            int levelEvents = 0;
            int peakLevel = 0;
            EventHandler<AudioLevelUpdatedEventArgs> onLevel = delegate { Interlocked.Increment(ref levelEvents); };
            EventHandler<RecognizeCompletedEventArgs> onCompleted = delegate(object sender, RecognizeCompletedEventArgs e)
            {
                if (e.Error != null) Interlocked.CompareExchange(ref completedError, e.Error, null);
            };
            double millisecondsRead = 0;
            long pcmBytesRead = 0;
            long disposalMilliseconds = 0;
            try
            {
                var grammar = new GrammarBuilder();
                grammar.Culture = recognizer.Culture;
                grammar.Append("Jarvis test");
                engine.LoadGrammar(new Grammar(grammar));
                engine.AudioLevelUpdated += onLevel;
                engine.RecognizeCompleted += onCompleted;
                input = SpeechInput.Attach(engine, selected.Id);
                if (input == null) throw new IOException("Explicit selection unexpectedly used Default.");
                engine.RecognizeAsync(RecognizeMode.Multiple);
                var duration = System.Diagnostics.Stopwatch.StartNew();
                while (duration.ElapsedMilliseconds < 2000)
                {
                    Application.DoEvents();
                    Exception error = SpeechInput.GetError(input)
                        ?? Interlocked.CompareExchange(ref completedError, null, null);
                    if (error != null) throw new IOException("Hardware probe capture failed.", error);
                    peakLevel = Math.Max(peakLevel, SpeechInput.GetLevel(input));
                    Thread.Sleep(20);
                }
                millisecondsRead = engine.AudioPosition.TotalMilliseconds;
                var pcm = (PcmAudioStream)input.GetType().GetProperty("Stream").GetValue(input, null);
                pcmBytesRead = pcm.Position;
                if (pcmBytesRead < 32000 || millisecondsRead <= 0)
                    throw new IOException("The recognizer did not continuously consume selected-device PCM audio; bytes "
                        + pcmBytesRead + ", audio position " + millisecondsRead + " ms, audio state "
                        + engine.AudioState + ", level events " + levelEvents + ".");
            }
            catch (Exception error) { errors.Add(error); }
            finally
            {
                engine.AudioLevelUpdated -= onLevel;
                engine.RecognizeCompleted -= onCompleted;
                var disposal = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    if (input != null) input.Dispose();
                }
                catch (Exception error) { errors.Add(error); }
                finally
                {
                    try { engine.Dispose(); }
                    catch (Exception error) { errors.Add(error); }
                }
                disposalMilliseconds = disposal.ElapsedMilliseconds;
                Exception inputError = SpeechInput.GetError(input);
                if (inputError != null) errors.Add(inputError);
            }
            if (errors.Count != 0)
                throw new IOException("Selected-device hardware probe " + attempt + " failed.",
                    new AggregateException(errors));
            Console.WriteLine("Probe " + attempt + ": audio advanced " + millisecondsRead.ToString("F0")
                + " ms; PCM bytes read " + pcmBytesRead + "; level events " + levelEvents
                + "; PCM peak level " + peakLevel + "; GetError=null; disposal " + disposalMilliseconds + " ms.");
        }
        Console.WriteLine("Selected-device capture/disposal/reopen verified; no audio file or recognized text was produced.");
    }
}
'@

try {
    Set-Content -LiteralPath $source -Value $testCode -Encoding UTF8
    & $compiler /nologo /langversion:5 /target:exe "/platform:$Platform" /warnaserror+ `
        "/out:$executable" /reference:System.dll /reference:System.Core.dll `
        /reference:System.Drawing.dll /reference:System.Windows.Forms.dll "/reference:$speech" `
        (Join-Path $projectDirectory "SpeechInput.cs") `
        (Join-Path $projectDirectory "UiAccessibility.cs") `
        (Join-Path $projectDirectory "SetupDialog.cs") $source
    if ($LASTEXITCODE -ne 0) { throw "Speech input test compilation failed: $LASTEXITCODE" }
    & $executable
    if ($LASTEXITCODE -ne 0) { throw "Speech input tests failed: $LASTEXITCODE" }
    foreach ($stage in 0..2) {
        & $executable --cleanup-failure $stage
        if ($LASTEXITCODE -ne 0) { throw "Speech cleanup stage $stage failed: $LASTEXITCODE" }
    }
    if ($ProbeMicrophone) {
        $probe = New-Object Diagnostics.Process
        $probe.StartInfo.FileName = $executable
        $probe.StartInfo.Arguments = "--probe-microphone"
        $probe.StartInfo.UseShellExecute = $false
        $probe.StartInfo.CreateNoWindow = $true
        try {
            if (-not $probe.Start()) { throw "Could not start the microphone probe." }
            if (-not $probe.WaitForExit(30000)) {
                Stop-Process -Id $probe.Id -Force
                [void]$probe.WaitForExit(5000)
                throw "The bounded microphone probe exceeded 30 seconds; only its new test process was stopped."
            }
            if ($probe.ExitCode -ne 0) { throw "Selected-device hardware probe failed: $($probe.ExitCode)" }
        } finally {
            $probe.Dispose()
        }
    }
} finally {
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Recurse -Force }
}
