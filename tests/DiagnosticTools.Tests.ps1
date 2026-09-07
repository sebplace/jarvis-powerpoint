[CmdletBinding()]
param(
    [ValidateSet("anycpu", "x86")]
    [string]$Platform = "anycpu",
    [switch]$ProbeHotkeys,
    [string]$Executable
)

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $PSScriptRoot
$compiler = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $compiler) { throw "The Windows .NET Framework C# compiler is required." }
$speech = Join-Path (Split-Path -Parent $compiler) "WPF\System.Speech.dll"
if (-not (Test-Path -LiteralPath $speech)) { throw "System.Speech is required." }
$output = Join-Path $PSScriptRoot ("diagnostics-" + [Guid]::NewGuid().ToString("N"))
$source = Join-Path $output "DiagnosticTestProgram.cs"
$testExecutable = Join-Path $output "DiagnosticTestProgram.exe"
$testCode = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using JarvisPowerPoint;

// Standalone tests isolate the navigation failure contract from the Office dependency graph.
// -Executable additionally verifies this contract against a built application without starting it.
namespace JarvisPowerPoint
{
    internal enum PresentationUnavailableReason { NoSlideShow, MultipleSlideShows }
    internal sealed class PresentationUnavailableException : InvalidOperationException
    {
        public PresentationUnavailableException(string message)
            : this(message, PresentationUnavailableReason.NoSlideShow) { }
        public PresentationUnavailableException(string message, PresentationUnavailableReason reason)
            : base(message) { Reason = reason; }
        public PresentationUnavailableReason Reason { get; private set; }
    }
}

internal static class DiagnosticTestProgram
{
    private static int assertions;
    private const string CanaryTitle = "CONFIDENTIAL_DECK_TITLE_78126";
    private const string CanaryPhrase = "recognized_secret_phrase_78126";
    private const string CanaryPassword = "PasswordLike!78126$dontkeep";
    private const string CanaryPath = @"C:\Users\PrivateUser78126\PrivateDeck78126.pptx";
    private const string CanaryDevice = "Private_microphone_name_78126";

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            bool probeHotkeys = false;
            string applicationPath = null;
            for (int i = 1; i < args.Length; i++)
            {
                if (args[i] == "--probe-hotkeys") probeHotkeys = true;
                else if (args[i] == "--app" && i + 1 < args.Length) applicationPath = args[++i];
                else throw new ArgumentException("Unexpected test argument.");
            }
            DiagnosticPrivacy(args[0]);
            DiagnosticBoundsAndConcurrency();
            PreflightEvaluation();
            PreflightReadOnlyAndErrors();
            PreflightDialogRefresh();
            PreflightEnumerationWithoutCapture();
            HotkeyBindingsAndLifetime();
            HotkeyRollback();
            HotkeyThreadAffinity();
            VerifyTypedContract(Assembly.GetExecutingAssembly());
            if (applicationPath != null) VerifyBuiltApplication(applicationPath);
            if (probeHotkeys) ProbeHotkeys();
            Console.WriteLine("Passed " + assertions + " assertions.");
            Console.WriteLine("No microphone capture, real settings writes, keystrokes, or presentation actions.");
            if (!probeHotkeys) Console.WriteLine("Native shortcut registration excluded; use -ProbeHotkeys to opt in.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void Assert(bool condition, string description)
    {
        if (!condition) throw new Exception(description);
        assertions++;
    }

    private static T Throws<T>(Action action, string description) where T : Exception
    {
        try { action(); }
        catch (T error) { assertions++; return error; }
        throw new Exception(description);
    }

    private static DiagnosticSnapshot DiagnosticState()
    {
        return new DiagnosticSnapshot {
            CultureName = "fr-FR", UsesDefaultMicrophone = true, State = ListeningState.Listening,
            SlideshowAvailable = true, ScreenCount = 2, HotkeysEnabled = false
        };
    }

    private static void PrivateTextAbsent(string text)
    {
        foreach (string canary in new[] { CanaryTitle, CanaryPhrase, CanaryPassword,
            CanaryPath, "PrivateUser78126", "PrivateDeck78126", CanaryDevice,
            "PrivateAlias78126", "PrivateRoute78126", "PrivateMachine78126" })
            Assert(text.IndexOf(canary, StringComparison.Ordinal) < 0, "Private canary leaked into output.");
    }

    private static void DiagnosticPrivacy(string output)
    {
        var recorder = new DiagnosticRecorder();
        var state = DiagnosticState();
        Assert(new DiagnosticSnapshot().State == ListeningState.Unavailable, "Default diagnostic state did not fail closed.");
        state.CultureName = CanaryTitle + "\r\n" + CanaryPhrase + "\r\n" + CanaryPath + "\r\n" + CanaryPassword;
        var error = new IOException(CanaryPassword,
            new FileNotFoundException(CanaryTitle, CanaryPath));
        error.Data[CanaryDevice] = CanaryPhrase;
        error.Source = "PrivateMachine78126";
        error.HelpLink = CanaryPath;
        recorder.Record(DiagnosticCategory.Speech, false, error);
        recorder.Record(DiagnosticCategory.Navigation, false, new COMException(CanaryTitle, unchecked((int)0x80070005)));
        recorder.Record(DiagnosticCategory.Profiles, false, new PrivateAlias78126Exception());
        recorder.Record(DiagnosticCategory.Settings, false, new PoisonMessageException());
        recorder.Record(DiagnosticCategory.Hotkeys, true);
        string report = recorder.BuildReport(state);
        PrivateTextAbsent(report);
        Assert(report.Contains("Culture: unknown"), "Culture input was not allow-listed.");
        Assert(report.Contains("System.IO.IOException"), "Safe error type missing.");
        Assert(report.Contains("0x80070005"), "HRESULT missing or not invariant hexadecimal.");
        Assert(!report.Contains("PoisonMessageException"), "Custom error type was not allow-listed.");
        Assert(report.Contains("RecentEvents: 5"), "Event count incorrect.");
        Assert(report.Contains(" | Success | none | 0x00000000"), "Success format incorrect.");
        state.CultureName = "EN-us";
        state.State = (ListeningState)int.MaxValue;
        state.ScreenCount = -1;
        report = recorder.BuildReport(state);
        Assert(report.Contains("Culture: en-US"), "Allowed culture was not normalized.");
        Assert(report.Contains("ListeningState: unknown"), "Invalid enum was emitted.");
        Assert(report.Contains("ScreenCount: 0"), "Negative screen count not sanitized.");
        state.ScreenCount = int.MaxValue;
        Assert(recorder.BuildReport(state).Contains("ScreenCount: 64"), "Screen count was not bounded.");
        Throws<ArgumentOutOfRangeException>(delegate {
            recorder.Record((DiagnosticCategory)int.MaxValue, true);
        }, "Invalid category was accepted.");
        Throws<ArgumentNullException>(delegate { recorder.BuildReport(null); }, "Null snapshot accepted.");
        string path = Path.Combine(output, "private-export.txt");
        recorder.Export(path, state);
        Assert(File.ReadAllText(path, Encoding.UTF8) == recorder.BuildReport(state), "Export differs from safe report.");
        byte[] bytes = File.ReadAllBytes(path);
        Assert(bytes.Length > 3 && !(bytes[0] == 239 && bytes[1] == 187 && bytes[2] == 191), "Unexpected UTF-8 BOM.");
        PrivateTextAbsent(File.ReadAllText(path, Encoding.UTF8));
        Throws<ArgumentException>(delegate { recorder.Export(@"\\not-a-server\share\private.txt", state); },
            "Network path was not rejected before file IO.");
        Throws<ArgumentException>(delegate { recorder.Export(@"\\?\C:\private.txt", state); },
            "Device path was not rejected.");
        Exception pathError = Throws<UnauthorizedAccessException>(delegate { recorder.Export(output, state); },
            "Writing over the output directory should fail.");
        recorder.Record(DiagnosticCategory.Settings, false, pathError);
        Assert(!recorder.BuildReport(state).Contains(output), "Export error leaked a disk path.");
        File.Delete(path);
    }

    private static void DiagnosticBoundsAndConcurrency()
    {
        var recorder = new DiagnosticRecorder();
        recorder.Record(DiagnosticCategory.Updates, false, new IOException(CanaryPath));
        for (int i = 0; i < 120; i++) recorder.Record(DiagnosticCategory.Navigation, true);
        string report = recorder.BuildReport(DiagnosticState());
        Assert(report.Contains("RecentEvents: 100"), "Event history was not bounded.");
        Assert(!report.Contains(" | Updates | "), "Oldest event was not removed.");
        int count = 0;
        foreach (string line in report.Split('\n')) if (line.Contains(" | Navigation | ")) count++;
        Assert(count == 100, "Bounded report row count incorrect.");
        var workers = new List<Thread>();
        Exception failure = null;
        for (int worker = 0; worker < 4; worker++)
        {
            var thread = new Thread(delegate() {
                try
                {
                    for (int i = 0; i < 500; i++)
                    {
                        recorder.Record(DiagnosticCategory.Speech, true);
                        if ((i % 20) == 0) recorder.BuildReport(DiagnosticState());
                    }
                }
                catch (Exception error) { Interlocked.CompareExchange(ref failure, error, null); }
            });
            workers.Add(thread);
            thread.Start();
        }
        foreach (Thread thread in workers) Assert(thread.Join(10000), "Recorder thread did not finish.");
        Assert(failure == null, "Recorder concurrency failed.");
        Assert(recorder.BuildReport(DiagnosticState()).Contains("RecentEvents: 100"), "Concurrent history exceeded bound.");
    }

    private static PreflightContext Context()
    {
        return new PreflightContext {
            CultureName = "fr-FR", MicrophoneId = "device-id", IsListening = true,
            HasRecentAudio = true, IndicatorEnabled = true, IndicatorDisplay = @"\\.\DISPLAY1"
        };
    }

    private static PreflightEvidence Evidence()
    {
        return new PreflightEvidence {
            RecognizersRead = true, InstalledCultures = new[] { "fr-FR", "en-US" },
            MicrophonesRead = true, MicrophoneIds = new[] { "default", "device-id" },
            PresentationState = PreflightPresentationState.Available,
            Displays = new PreflightDisplays {
                ScreenIds = new[] { @"\\.\DISPLAY1", @"\\.\DISPLAY2" },
                PrimaryScreenId = @"\\.\DISPLAY1", PresentationScreenId = @"\\.\DISPLAY2"
            }
        };
    }

    private static PreflightItem Find(List<PreflightItem> items, string key)
    {
        foreach (PreflightItem item in items) if (item.CheckKey == key) return item;
        throw new Exception("Missing preflight row: " + key);
    }

    private static PreflightItem Evaluate(string key, PreflightContext context, PreflightEvidence evidence)
    {
        return Find(PreflightService.Evaluate(context, evidence, true), key);
    }

    private static void PreflightEvaluation()
    {
        var context = Context();
        var evidence = Evidence();
        Assert(Evaluate("Slideshow", context, new PreflightEvidence()).Severity == PreflightSeverity.Fail,
            "Missing presentation evidence did not fail closed.");
        List<PreflightItem> ready = PreflightService.Evaluate(context, evidence, true);
        Assert(ready.Count == 6, "Unexpected number of preflight checks.");
        foreach (PreflightItem item in ready) Assert(item.Severity == PreflightSeverity.Pass, "Ready check did not pass.");
        context.CultureName = "en-US";
        Assert(Evaluate("Language", context, evidence).Severity == PreflightSeverity.Pass, "EN-US rejected.");
        context.CultureName = CanaryPassword;
        Assert(Evaluate("Language", context, evidence).Severity == PreflightSeverity.Fail, "Unapproved culture accepted.");
        PrivateTextAbsent(Evaluate("Language", context, evidence).Detail);
        context.CultureName = "fr-FR";
        evidence.InstalledCultures = new[] { "en-US" };
        Assert(Evaluate("Language", context, evidence).Severity == PreflightSeverity.Fail, "Missing recognizer accepted.");
        evidence = Evidence();
        context.MicrophoneId = "default";
        Assert(Evaluate("Microphone", context, evidence).Severity == PreflightSeverity.Pass, "Default input with hardware rejected.");
        evidence.MicrophoneIds = new[] { "default" };
        Assert(Evaluate("Microphone", context, evidence).Severity == PreflightSeverity.Fail, "Default placeholder treated as hardware.");
        context.MicrophoneId = "DEVICE-ID";
        evidence = Evidence();
        Assert(Evaluate("Microphone", context, evidence).Severity == PreflightSeverity.Fail, "Microphone IDs were not exact.");
        context.MicrophoneId = CanaryDevice;
        PrivateTextAbsent(Evaluate("Microphone", context, evidence).Detail);
        context = Context();
        context.HasRecentAudio = false;
        PreflightItem silent = Evaluate("Audio", context, evidence);
        Assert(silent.Severity == PreflightSeverity.Warning && silent.Detail.Contains("Silence does not mean"),
            "Silence was diagnosed as broken microphone.");
        context.IsListening = false;
        context.HasRecentAudio = true;
        Assert(Evaluate("Listening", context, evidence).Severity == PreflightSeverity.Warning, "Paused state not warned.");
        Assert(Evaluate("Audio", context, evidence).Severity == PreflightSeverity.Warning, "Stale audio passed while paused.");
        context = Context();
        foreach (PreflightPresentationState state in new[] { PreflightPresentationState.NoShow,
            PreflightPresentationState.MultipleShows, PreflightPresentationState.Unavailable })
        {
            evidence.PresentationState = state;
            Assert(Evaluate("Slideshow", context, evidence).Severity == PreflightSeverity.Fail, "Unavailable show passed.");
        }
        evidence = Evidence();
        Assert(Evaluate("Indicator", context, evidence).Detail.Contains("may still expose"),
            "Different-screen pass lacks capture caveat.");
        context.IndicatorDisplay = null;
        Assert(Evaluate("Indicator", context, evidence).Severity == PreflightSeverity.Warning, "Null display assumed primary.");
        context.IndicatorDisplay = "";
        Assert(Evaluate("Indicator", context, evidence).Severity == PreflightSeverity.Warning, "Empty display assumed primary.");
        context.IndicatorDisplay = "primary";
        Assert(Evaluate("Indicator", context, evidence).Severity == PreflightSeverity.Fail, "Display alias accepted instead of exact device ID.");
        context.IndicatorDisplay = @"\\.\DISPLAY2";
        Assert(Evaluate("Indicator", context, evidence).Severity == PreflightSeverity.Warning, "Same-screen risk not warned.");
        context.IndicatorDisplay = CanaryPath;
        Assert(Evaluate("Indicator", context, evidence).Severity == PreflightSeverity.Fail, "Missing selected display passed.");
        PrivateTextAbsent(Evaluate("Indicator", context, evidence).Detail);
        context.IndicatorDisplay = @"\\.\DISPLAY1";
        evidence.Displays.PresentationScreenId = null;
        Assert(Evaluate("Indicator", context, evidence).Severity == PreflightSeverity.Warning, "Unknown slide-show screen passed.");
        evidence.Displays.ScreenIds = new[] { @"\\.\DISPLAY1" };
        Assert(Evaluate("Indicator", context, evidence).Severity == PreflightSeverity.Warning, "Single-display risk not warned.");
        evidence.Displays = null;
        Assert(Evaluate("Indicator", context, evidence).Severity == PreflightSeverity.Warning, "Failed display enumeration passed.");
        context.IndicatorEnabled = false;
        Assert(Evaluate("Indicator", context, evidence).Detail.Contains("disabled"), "Disabled indicator not reported.");
        string french = Find(PreflightService.Evaluate(Context(), Evidence(), false), "Listening").Detail;
        Assert(french.Contains("reconnaissance vocale"), "French localization missing.");
    }

    private static PresentationSnapshot PrivateSnapshot()
    {
        return new PresentationSnapshot {
            PresentationName = CanaryTitle, PresentationKey = CanaryPassword, SessionKey = CanaryPhrase,
            Title = CanaryTitle, SavedPath = CanaryPath, WindowHandle = new IntPtr(78126)
        };
    }

    private static void PreflightReadOnlyAndErrors()
    {
        var environment = new FakeEnvironment();
        int snapshots = 0;
        var context = Context();
        var items = PreflightService.Run(context, delegate { snapshots++; return PrivateSnapshot(); }, true, environment);
        Assert(snapshots == 1 && environment.Reads == 3, "Preflight did not use bounded read-only probes.");
        Assert(environment.LastWindow == new IntPtr(78126), "Snapshot window was not used for display check.");
        Assert(context.MicrophoneId == "device-id" && context.IndicatorDisplay == @"\\.\DISPLAY1", "Context was mutated.");
        foreach (PreflightItem item in items) PrivateTextAbsent(item.Detail);
        items = PreflightService.Run(context, delegate { return null; }, true, environment);
        string noShow = Find(items, "Slideshow").Detail;
        Assert(noShow.Contains("No active PowerPoint"), "No-show result not distinguished.");
        Assert(environment.LastWindow == IntPtr.Zero, "No-show lookup used stale window.");
        items = PreflightService.Run(context, delegate {
            throw new PresentationUnavailableException(CanaryPassword);
        }, true, environment);
        Assert(Find(items, "Slideshow").Detail == noShow, "Typed no-show was not distinguished.");
        items = PreflightService.Run(context, delegate {
            throw new PresentationUnavailableException(CanaryPhrase, PresentationUnavailableReason.MultipleSlideShows);
        }, true, environment);
        Assert(Find(items, "Slideshow").Detail.Contains("More than one"), "Ambiguous shows were not distinguished.");
        Assert(Find(items, "Slideshow").Severity == PreflightSeverity.Fail, "Ambiguous shows passed.");
        foreach (PreflightItem item in items) PrivateTextAbsent(item.Detail);
        items = PreflightService.Run(context, delegate {
            throw new PresentationUnavailableException(CanaryPath, (PresentationUnavailableReason)int.MaxValue);
        }, true, environment);
        Assert(Find(items, "Slideshow").Detail != noShow, "Unknown navigation reason was treated as no show.");
        items = PreflightService.Run(context, delegate { throw new COMException(CanaryPath); }, true, environment);
        Assert(Find(items, "Slideshow").Detail != noShow, "COM failure was misreported as no show.");
        foreach (PreflightItem item in items) PrivateTextAbsent(item.Detail);
        environment.Fail = true;
        items = PreflightService.Run(context, delegate { throw new InvalidOperationException(CanaryPhrase); }, true, environment);
        Assert(Find(items, "Language").Severity == PreflightSeverity.Fail, "Recognizer failure did not become a row.");
        Assert(Find(items, "Microphone").Severity == PreflightSeverity.Fail, "Device failure did not become a row.");
        Assert(Find(items, "Indicator").Severity == PreflightSeverity.Warning, "Screen failure did not become a warning.");
        foreach (PreflightItem item in items) PrivateTextAbsent(item.Detail);
        Throws<NullReferenceException>(delegate {
            PreflightService.Run(context, delegate { throw new NullReferenceException(); }, true, new FakeEnvironment());
        }, "Programming error was swallowed.");
        Assert(PreflightService.Run(null, delegate { return null; }, true, environment).Count == 1,
            "Missing current context was not reported.");
        Assert(!PreflightService.IsExpected(new OutOfMemoryException()), "Fatal exception classified as expected.");
        Assert(!PreflightService.IsExpected(new NullReferenceException()), "Programming exception classified as expected.");
    }

    private static void PreflightDialogRefresh()
    {
        var context = Context();
        int contexts = 0;
        int setupCalls = 0;
        var environment = new FakeEnvironment();
        using (var dialog = new PreflightDialog(true,
            delegate { contexts++; return context; }, delegate { return PrivateSnapshot(); },
            delegate { setupCalls++; context = Context(); context.IsListening = false; }, environment))
        {
            Assert(environment.Reads == 0, "Dialog construction probed hardware.");
            dialog.RefreshResults();
            Assert(contexts == 1 && environment.Reads == 3, "Dialog refresh did not read current context.");
            dialog.OpenSetupAndRefresh();
            Assert(contexts == 2 && setupCalls == 1, "Setup did not refresh current context afterward.");
            DataGridView grid = FindGrid(dialog);
            Assert(grid.ReadOnly && !grid.AllowUserToAddRows && !grid.AllowUserToDeleteRows, "Diagnostic rows are editable.");
            Assert((string)grid.Rows[2].Cells[0].Value == "Warning", "Dialog reused stale listening state after setup.");
        }
        using (var dialog = new PreflightDialog(false, delegate { return Context(); }, delegate { return null; },
            delegate { throw new IOException(CanaryPassword); }, new FakeEnvironment()))
        {
            dialog.OpenSetupAndRefresh();
            DataGridView grid = FindGrid(dialog);
            Assert(grid.Rows.Count == 7, "Expected setup failure did not become an additional row.");
            foreach (DataGridViewRow row in grid.Rows) PrivateTextAbsent((string)row.Cells[2].Value);
        }
    }

    private static DataGridView FindGrid(Control control)
    {
        DataGridView grid = control as DataGridView;
        if (grid != null) return grid;
        foreach (Control child in control.Controls)
        {
            grid = FindGrid(child);
            if (grid != null) return grid;
        }
        return null;
    }

    private static void PreflightEnumerationWithoutCapture()
    {
        List<PreflightItem> items = PreflightService.Run(Context(), delegate { return null; }, true);
        Assert(items.Count == 6, "Read-only Windows enumeration did not produce all checks.");
        Assert(Find(items, "Slideshow").Severity == PreflightSeverity.Fail, "No-show production probe passed.");
        foreach (PreflightItem item in items) PrivateTextAbsent(item.Detail);
    }

    private static IntPtr Packet(uint key, uint modifiers)
    {
        return new IntPtr((long)(key << 16) | modifiers);
    }

    private static void HotkeyBindingsAndLifetime()
    {
        var native = new FakeHotkeyNative();
        long now = 1000;
        var hotkeys = new GlobalHotkeys(native, delegate { return now; });
        IntPtr handle = hotkeys.Handle;
        Assert(handle != IntPtr.Zero && IsWindow(handle) && !IsWindowVisible(handle),
            "Hotkey receiver is not a hidden live window.");
        Assert(!hotkeys.Enabled && native.Attempts.Count == 0, "Shortcuts enabled without opt-in.");
        var received = new List<PresentationHotkey>();
        hotkeys.CommandPressed += delegate(object sender, HotkeyEventArgs args) { received.Add(args.Command); };
        try
        {
            hotkeys.Enable();
            hotkeys.Enable();
            Assert(hotkeys.Enabled && native.Attempts.Count == 6 && native.Active.Count == 6, "Enable not complete/idempotent.");
            uint[] expectedKeys = { (uint)Keys.Right, (uint)Keys.Left, (uint)Keys.Enter,
                (uint)Keys.N, (uint)Keys.B, (uint)Keys.S };
            PresentationHotkey[] expectedCommands = { PresentationHotkey.Next, PresentationHotkey.Previous,
                PresentationHotkey.Resume, PresentationHotkey.NextResult,
                PresentationHotkey.BlackScreen, PresentationHotkey.RestoreSlides };
            for (int i = 0; i < 6; i++)
            {
                Registration registration = native.Attempts[i];
                Assert(registration.Window == handle && registration.Key == expectedKeys[i]
                    && registration.Modifiers == (GlobalHotkeys.Modifiers | GlobalHotkeys.NoRepeat), "Shortcut registration differs.");
                Assert(hotkeys.ProcessHotkeyMessage(registration.Id, Packet(registration.Key, GlobalHotkeys.Modifiers)),
                    "Registered shortcut did not dispatch.");
                Assert(received[i] == expectedCommands[i], "Shortcut dispatched wrong command.");
                Assert(!hotkeys.ProcessHotkeyMessage(registration.Id, Packet(registration.Key, GlobalHotkeys.Modifiers)),
                    "Repeat noise dispatched.");
            }
            Registration first = native.Attempts[0];
            Assert(!hotkeys.ProcessHotkeyMessage(-1, Packet(first.Key, GlobalHotkeys.Modifiers)), "Unknown ID dispatched.");
            now += 200;
            Assert(!hotkeys.ProcessHotkeyMessage(first.Id, Packet((uint)Keys.F1, GlobalHotkeys.Modifiers)),
                "Wrong key packet dispatched.");
            Assert(!hotkeys.ProcessHotkeyMessage(first.Id, Packet(first.Key, 0)), "Wrong modifiers dispatched.");
            Assert(hotkeys.ProcessHotkeyMessage(first.Id, Packet(first.Key, GlobalHotkeys.Modifiers)), "Fresh shortcut suppressed.");
            int accepted = received.Count;
            hotkeys.Disable();
            int unregisters = native.UnregisterCalls;
            hotkeys.Disable();
            Assert(!hotkeys.Enabled && native.Active.Count == 0 && native.UnregisterCalls == unregisters,
                "Disable did not clean all keys or was not idempotent.");
            Assert(!hotkeys.ProcessHotkeyMessage(first.Id, Packet(first.Key, GlobalHotkeys.Modifiers)),
                "Disabled shortcut dispatched.");
            hotkeys.Enable();
            Assert(!hotkeys.ProcessHotkeyMessage(first.Id, Packet(first.Key, GlobalHotkeys.Modifiers)),
                "Queued old-generation shortcut dispatched after re-enable.");
            Assert(received.Count == accepted, "Unexpected command event.");
            IDictionary<PresentationHotkey, string> descriptions = GlobalHotkeys.ShortcutDescriptions;
            Assert(descriptions.Count == 6 && descriptions[PresentationHotkey.Resume] == "Ctrl+Alt+Shift+Enter",
                "Shortcut descriptions missing or incorrect.");
            descriptions.Clear();
            Assert(GlobalHotkeys.ShortcutDescriptions.Count == 6, "Caller mutated static bindings.");
        }
        finally { hotkeys.Dispose(); }
        hotkeys.Dispose();
        hotkeys.Disable();
        Assert(!hotkeys.Enabled && native.Active.Count == 0 && !IsWindow(handle), "Dispose leaked registrations/window.");
        Throws<ObjectDisposedException>(delegate { hotkeys.Enable(); }, "Disposed receiver enabled.");
    }

    private static void HotkeyRollback()
    {
        for (int failedAt = 0; failedAt < 6; failedAt++)
        {
            var native = new FakeHotkeyNative { FailRegistration = failedAt };
            using (var hotkeys = new GlobalHotkeys(native, delegate { return 0; }))
            {
                HotkeyRegistrationException error = Throws<HotkeyRegistrationException>(hotkeys.Enable,
                    "Registration conflict did not throw.");
                Assert(!hotkeys.Enabled && native.Active.Count == 0, "Conflict left a partially enabled set.");
                Assert(native.Attempts.Count == failedAt + 1 && native.UnregisterCalls == failedAt,
                    "Rollback did not release exactly the successfully registered prefix.");
                Assert(error.NativeErrorCode == 1409 && error.Message.Contains(error.Shortcut)
                    && error.Message.Contains("1409"), "Conflict lacks exact shortcut/Windows error.");
                native.FailRegistration = -1;
                hotkeys.Enable();
                Assert(hotkeys.Enabled && native.Active.Count == 6, "Could not retry after conflict.");
            }
            Assert(native.Active.Count == 0, "Retry registrations were not disposed.");
        }
        var failingCleanup = new FakeHotkeyNative();
        using (var hotkeys = new GlobalHotkeys(failingCleanup, delegate { return 0; }))
        {
            hotkeys.Enable();
            failingCleanup.FailUnregisterId = failingCleanup.Attempts[2].Id;
            Throws<InvalidOperationException>(hotkeys.Disable, "Unregistration failure was hidden.");
            Assert(!hotkeys.Enabled && failingCleanup.Active.Count == 1 && failingCleanup.UnregisterCalls == 6,
                "Cleanup did not attempt every registration.");
            Assert(!hotkeys.ProcessHotkeyMessage(failingCleanup.Attempts[2].Id,
                Packet(failingCleanup.Attempts[2].Key, GlobalHotkeys.Modifiers)), "Cleanup failure left commands active.");
            failingCleanup.FailUnregisterId = -1;
            hotkeys.Disable();
            Assert(failingCleanup.Active.Count == 0, "Cleanup failure was not retryable.");
        }
    }

    private static void HotkeyThreadAffinity()
    {
        Exception failure = null;
        var worker = new Thread(delegate() {
            try { using (var hotkeys = new GlobalHotkeys(new FakeHotkeyNative(), delegate { return 0; })) { } }
            catch (Exception error) { failure = error; }
        });
        worker.SetApartmentState(ApartmentState.MTA);
        worker.Start();
        Assert(worker.Join(5000) && failure is InvalidOperationException, "MTA construction was not rejected.");
        using (var hotkeys = new GlobalHotkeys(new FakeHotkeyNative(), delegate { return 0; }))
        {
            failure = null;
            worker = new Thread(delegate() {
                try { hotkeys.Enable(); }
                catch (Exception error) { failure = error; }
            });
            worker.Start();
            Assert(worker.Join(5000) && failure is InvalidOperationException, "Cross-thread enable accepted.");
            Assert(!hotkeys.Enabled, "Cross-thread call changed opt-in state.");
        }
    }

    private static void ProbeHotkeys()
    {
        using (var hotkeys = new GlobalHotkeys())
        {
            try
            {
                hotkeys.Enable();
                hotkeys.Disable();
                hotkeys.Enable();
                Assert(hotkeys.Enabled, "Native hotkey probe did not enable.");
                Console.WriteLine("Native probe registered, released and re-registered six shortcuts; disposing now.");
            }
            catch (HotkeyRegistrationException error)
            {
                Console.WriteLine("Native probe skipped due to an existing shortcut conflict: " + error.Message);
            }
        }
    }

    private static void VerifyBuiltApplication(string path)
    {
        VerifyTypedContract(Assembly.LoadFrom(path));
        Console.WriteLine("Built application typed navigation/preflight contract verified without starting the application.");
    }

    private static void VerifyTypedContract(Assembly app)
    {
        Type contextType = app.GetType("JarvisPowerPoint.PreflightContext", true);
        Type snapshotType = app.GetType("JarvisPowerPoint.PresentationSnapshot", true);
        Type errorType = app.GetType("JarvisPowerPoint.PresentationUnavailableException", true);
        Type reasonType = app.GetType("JarvisPowerPoint.PresentationUnavailableReason", true);
        Type serviceType = app.GetType("JarvisPowerPoint.PreflightService", true);
        Type snapshotDelegateType = typeof(Func<>).MakeGenericType(snapshotType);
        MethodInfo run = serviceType.GetMethod("Run", BindingFlags.Public | BindingFlags.Static, null,
            new[] { contextType, snapshotDelegateType, typeof(bool) }, null);
        Assert(run != null, "Built application is missing the preflight API.");
        object context = Activator.CreateInstance(contextType, true);
        contextType.GetProperty("CultureName").SetValue(context, "fr-FR", null);
        contextType.GetProperty("MicrophoneId").SetValue(context, "default", null);
        object defaultError = Activator.CreateInstance(errorType, new object[] { CanaryPassword });
        Assert(errorType.GetProperty("Reason").GetValue(defaultError, null).ToString() == "NoSlideShow",
            "Built navigation one-argument exception changed meaning.");
        string[] reasons = { "NoSlideShow", "MultipleSlideShows" };
        string[] expected = { "No active PowerPoint", "More than one" };
        for (int i = 0; i < reasons.Length; i++)
        {
            object reason = Enum.Parse(reasonType, reasons[i]);
            Exception error = (Exception)Activator.CreateInstance(errorType, new object[] { CanaryPassword, reason });
            Delegate snapshot = Expression.Lambda(snapshotDelegateType,
                Expression.Throw(Expression.Constant(error), snapshotType)).Compile();
            var rows = (System.Collections.IEnumerable)run.Invoke(null, new object[] { context, snapshot, true });
            bool found = false;
            foreach (object row in rows)
            {
                Type itemType = row.GetType();
                string detail = (string)itemType.GetProperty("Detail").GetValue(row, null);
                PrivateTextAbsent(detail);
                if ((string)itemType.GetProperty("CheckKey").GetValue(row, null) != "Slideshow") continue;
                found = true;
                Assert(detail.Contains(expected[i]), "Built preflight did not classify the typed navigation failure.");
                Assert(itemType.GetProperty("Severity").GetValue(row, null).ToString() == "Fail",
                    "Built preflight passed an unavailable show.");
            }
            Assert(found, "Built preflight omitted the slideshow check.");
        }
    }

    private sealed class PrivateAlias78126Exception : Exception { }
    private sealed class PoisonMessageException : Exception
    {
        public override string Message { get { throw new Exception("The private Message getter must never run."); } }
        public override string StackTrace { get { throw new Exception("The private StackTrace getter must never run."); } }
        public override string ToString() { throw new Exception("Private exception text must never be formatted."); }
    }

    private sealed class FakeEnvironment : IPreflightEnvironment
    {
        public int Reads;
        public bool Fail;
        public IntPtr LastWindow;
        public IList<string> GetRecognizerCultures()
        {
            Reads++;
            if (Fail) throw new COMException(CanaryPassword);
            return Evidence().InstalledCultures;
        }
        public IList<string> GetMicrophoneIds()
        {
            Reads++;
            if (Fail) throw new IOException(CanaryDevice);
            return Evidence().MicrophoneIds;
        }
        public PreflightDisplays GetDisplays(IntPtr window)
        {
            Reads++;
            LastWindow = window;
            if (Fail) throw new InvalidOperationException(CanaryPath);
            PreflightDisplays result = Evidence().Displays;
            if (window == IntPtr.Zero) result.PresentationScreenId = null;
            return result;
        }
    }

    private sealed class Registration
    {
        public IntPtr Window;
        public int Id;
        public uint Modifiers;
        public uint Key;
    }

    private sealed class FakeHotkeyNative : IHotkeyNative
    {
        public readonly List<Registration> Attempts = new List<Registration>();
        public readonly HashSet<int> Active = new HashSet<int>();
        public int FailRegistration = -1;
        public int FailUnregisterId = -1;
        public int UnregisterCalls;
        public bool Register(IntPtr window, int id, uint modifiers, uint key, out int error)
        {
            Attempts.Add(new Registration { Window = window, Id = id, Modifiers = modifiers, Key = key });
            if (Attempts.Count - 1 == FailRegistration) { error = 1409; return false; }
            if (!Active.Add(id)) throw new Exception("Registration ID reused.");
            error = 0;
            return true;
        }
        public bool Unregister(IntPtr window, int id, out int error)
        {
            UnregisterCalls++;
            if (id == FailUnregisterId) { error = 5; return false; }
            if (!Active.Remove(id)) throw new Exception("Unknown ID unregistered.");
            error = 0;
            return true;
        }
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr window);
}
'@

try {
    New-Item -ItemType Directory -Path $output | Out-Null
    [System.IO.File]::WriteAllText($source, $testCode, [System.Text.Encoding]::ASCII)
    $arguments = @(
        "/nologo", "/target:exe", "/langversion:5", "/warnaserror+", "/platform:$Platform",
        "/out:$testExecutable", "/reference:System.dll", "/reference:System.Core.dll",
        "/reference:System.Drawing.dll", "/reference:System.Windows.Forms.dll", "/reference:$speech",
        (Join-Path $projectDirectory "PresentationContracts.cs"),
        (Join-Path $projectDirectory "SpeechInput.cs"),
        (Join-Path $projectDirectory "DiagnosticTools.cs"),
        (Join-Path $projectDirectory "PreflightDialog.cs"),
        (Join-Path $projectDirectory "UiAccessibility.cs"),
        (Join-Path $projectDirectory "GlobalHotkeys.cs"), $source
    )
    & $compiler @arguments
    if ($LASTEXITCODE -ne 0) { throw "Diagnostic test compilation failed." }
    $runArguments = @($output)
    if ($ProbeHotkeys) { $runArguments += "--probe-hotkeys" }
    if ($Executable) {
        $candidate = (Resolve-Path -LiteralPath $Executable).Path
        $runArguments += @("--app", $candidate)
    }
    & $testExecutable @runArguments
    if ($LASTEXITCODE -ne 0) { throw "Diagnostic tests failed." }
}
finally {
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Recurse -Force }
}
