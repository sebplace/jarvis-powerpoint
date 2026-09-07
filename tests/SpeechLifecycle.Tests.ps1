[CmdletBinding()]
param([string]$Executable = (Join-Path (Split-Path -Parent $PSScriptRoot) "bin\JarvisPowerPoint.exe"))
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$output = Join-Path $PSScriptRoot ("speech-lifecycle-" + [Guid]::NewGuid().ToString("N"))
$compiler = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
$code = @'
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Speech.Recognition;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
internal static class SpeechLifecycleTests
{
    private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
    private static int assertions;
    private static object Field(object obj, string name) { return obj.GetType().GetField(name, Private).GetValue(obj); }
    private static void Set(object obj, string name, object value) { obj.GetType().GetField(name, Private).SetValue(obj, value); }
    private static object Call(object obj, string name, params object[] args)
    {
        try { return obj.GetType().GetMethod(name, Private | BindingFlags.Public).Invoke(obj, args); }
        catch (TargetInvocationException e) { throw e.InnerException; }
    }
    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
        assertions++;
    }
    private sealed class ControlProbe : NativeWindow, IDisposable
    {
        private readonly Control control;
        public int SetterMessages;
        public int Invalidations;
        public int TextChanges;
        public ControlProbe(Control value)
        {
            control = value;
            AssignHandle(control.Handle);
            control.Invalidated += Invalidated;
            control.TextChanged += TextChanged;
        }
        private void Invalidated(object sender, InvalidateEventArgs args) { Invalidations++; }
        private void TextChanged(object sender, EventArgs args) { TextChanges++; }
        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0x000c || message.Msg == 0x0402) SetterMessages++;
            base.WndProc(ref message);
        }
        public void Dispose()
        {
            control.Invalidated -= Invalidated;
            control.TextChanged -= TextChanged;
            ReleaseHandle();
        }
    }
    private static void CheckUnchangedStatus(object context)
    {
        object panel = Field(context, "panel");
        object indicator = Field(panel, "indicator");
        var probes = new List<ControlProbe>();
        try
        {
            foreach (string name in new[] { "status", "heard", "outcome", "audioLevel" })
                probes.Add(new ControlProbe((Control)Field(panel, name)));
            probes.Add(new ControlProbe((Control)Field(indicator, "text")));
            probes.Add(new ControlProbe((Control)Field(indicator, "level")));
            Call(context, "UpdateStatus");
            foreach (ControlProbe probe in probes)
                probe.SetterMessages = probe.Invalidations = probe.TextChanges = 0;
            for (int i = 0; i < 1000; i++)
            {
                Set(context, "audioLevel", i % 101);
                Call(context, "UpdateStatus");
            }
            int setters = 0, invalidations = 0, textChanges = 0;
            foreach (ControlProbe probe in probes)
            {
                setters += probe.SetterMessages;
                invalidations += probe.Invalidations;
                textChanges += probe.TextChanges;
            }
            Assert(setters == 0 && invalidations == 0 && textChanges == 0,
                "Unchanged hidden status generated redundant control updates.");
            Set(context, "lastOutcome", "Synthetic changed outcome");
            Call(context, "UpdateStatus");
            int changed = 0;
            foreach (ControlProbe probe in probes) changed += probe.TextChanges;
            Assert(changed > 0, "Status benchmark probes missed a real changed outcome.");
            Console.WriteLine("1000 unchanged tray-status updates (varying hidden audio): "
                + setters + " WM_SETTEXT/PBM_SETPOS messages, " + invalidations
                + " managed invalidations, " + textChanges + " text changes. Not a CPU or memory measurement.");
        }
        finally
        {
            foreach (ControlProbe probe in probes) probe.Dispose();
            Set(context, "audioLevel", 0);
        }
    }
    private static void CheckSpeechGuards(object context, Assembly assembly)
    {
        object recovery = Field(context, "speechRecovery");
        Set(context, "speechFailureNotified", true);
        Set(context, "listeningRequested", true);
        Call(recovery, "RequestStart");
        Call(recovery, "Fail", TimeSpan.Zero, true);
        using (var setup = (Form)Activator.CreateInstance(assembly.GetType("JarvisPowerPoint.SetupDialog", true),
            BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { "fr-FR", "default", false }, null))
        {
            Set(context, "activeSetup", setup);
            try
            {
                int generation = (int)recovery.GetType().GetProperty("Generation").GetValue(recovery, null);
                Call(context, "InitializeSpeechRecognition");
                Call(context, "PollSpeechHealth");
                Call(context, "OnRecognitionCompleted", null, null);
                Assert((int)recovery.GetType().GetProperty("Generation").GetValue(recovery, null) == generation,
                    "Active setup allowed initialization, retry, or a completion failure.");
                Assert((int)recovery.GetType().GetProperty("Attempts").GetValue(recovery, null) == 0,
                    "Active setup consumed a retry.");
                Call(recovery, "Succeeded");
                Set(context, "acceptsCommands", true);
                Call(context, "OnSpeechRecognized", null, null);
                Assert(Field(context, "recognizer") == null, "Active setup allowed a speech command.");
                Call(recovery, "Setup", true);
                Call(context, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Suspend));
                Call(context, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Resume));
                Call(context, "SpeechFailed", new System.IO.IOException("Synthetic setup failure"));
                Assert(!(bool)recovery.GetType().GetProperty("CanStart").GetValue(recovery, null)
                    && !((System.Windows.Forms.Timer)Field(context, "speechHealthTimer")).Enabled,
                    "Resume or failure restarted listening during setup.");
            }
            finally
            {
                Set(context, "activeSetup", null);
                Call(recovery, "Setup", false);
            }
        }
        Assert((bool)recovery.GetType().GetProperty("RestartRequired").GetValue(recovery, null),
            "Unsafe cleanup did not persist its process-restart requirement.");
        Set(context, "listeningRequested", false);
        Call(recovery, "Pause");
        Call(context, "SpeechFailed", new System.IO.IOException("Synthetic paused failure"));
        Call(context, "OnRecognitionCompleted", null, null);
        Call(context, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Suspend));
        Call(context, "InitializeSpeechRecognition");
        Call(context, "PollSpeechHealth");
        Call(context, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Resume));
        Assert(!(bool)recovery.GetType().GetProperty("Requested").GetValue(recovery, null)
            && !((System.Windows.Forms.Timer)Field(context, "speechHealthTimer")).Enabled,
            "Failure or suspend/resume reversed the user's pause.");
        Assert(Field(context, "recognizer") == null && Field(context, "microphoneInput") == null,
            "Guard validation created a capture.");
    }
    private static void CheckModalSpeechGuard(Type type)
    {
        object context = Activator.CreateInstance(type);
        Application.Idle -= (EventHandler)Delegate.CreateDelegate(typeof(EventHandler), context,
            type.GetMethod("OnFirstIdle", Private));
        try
        {
            object recovery = Field(context, "speechRecovery");
            Call(recovery, "RequestStart");
            Call(recovery, "Succeeded");
            Set(context, "acceptsCommands", true);
            Set(context, "modalDepth", 1);
            Set(context, "lastOutcome", "");
            Call(context, "OnSpeechRecognized", null, null);
            Assert(((string)Field(context, "lastOutcome")).Length > 0,
                "Modal dialog did not report blocked voice navigation.");
            Assert(Field(context, "recognizer") == null && Field(context, "microphoneInput") == null,
                "Synthetic modal callback opened capture.");
        }
        finally
        {
            Set(context, "modalDepth", 0);
            Call(context, "ExitApplication", null, EventArgs.Empty);
        }
    }
    private static void CheckProcessCaptureBlock(Type type, Assembly assembly)
    {
        object context = Activator.CreateInstance(type);
        Application.Idle -= (EventHandler)Delegate.CreateDelegate(typeof(EventHandler), context,
            type.GetMethod("OnFirstIdle", Private));
        try
        {
            object recovery = Field(context, "speechRecovery");
            assembly.GetType("JarvisPowerPoint.SpeechInput", true)
                .GetMethod("BlockCaptureFailure", BindingFlags.NonPublic | BindingFlags.Static)
                .Invoke(null, new object[] { new System.IO.IOException("Synthetic retained ownership; no capture.") });
            Call(recovery, "Succeeded");
            Set(context, "acceptsCommands", true);
            Set(context, "lastHeard", "");
            Call(context, "OnSpeechRecognized", null, null);
            Assert(!(bool)Field(context, "acceptsCommands") && (string)Field(context, "lastHeard") == "",
                "A queued recognition callback crossed the process-wide latch before health polling.");
            Set(context, "speechFailureNotified", true);
            foreach (string culture in new[] { "fr-FR", "en-US" })
            {
                Set(context, "currentCultureName", culture);
                Set(context, "cleanupBlocked", false);
                Set(context, "listeningRequested", true);
                Call(context, "PrepareSpeechStart");
                Call(context, "InitializeSpeechRecognition");
                Call(context, "PollSpeechHealth");
                Call(context, "OpenSetup", false);
                Assert(Field(context, "activeSetup") == null && Field(context, "recognizer") == null
                    && Field(context, "microphoneInput") == null, "Global ownership latch allowed setup or restart.");
                Assert((bool)recovery.GetType().GetProperty("RestartRequired").GetValue(recovery, null)
                    && !(bool)recovery.GetType().GetProperty("CanStart").GetValue(recovery, null)
                    && (int)recovery.GetType().GetProperty("Attempts").GetValue(recovery, null) == 0,
                    "Global ownership latch consumed retries or allowed capture.");
                Set(context, "listeningRequested", false);
                Call(recovery, "Pause");
                Call(context, "SetRecoveryStatus");
                string restart = culture == "fr-FR" ? "red\u00e9marrez" : "restart";
                Assert(((string)Field(context, "listeningStatus")).Contains(restart),
                    "Paused unsafe cleanup lost its localized restart instruction.");
                Call(context, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Suspend));
                Assert(((string)Field(context, "listeningStatus")).Contains(restart),
                    "Suspended unsafe cleanup lost its localized restart instruction.");
                Call(context, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Resume));
                Assert(((string)Field(context, "listeningStatus")).Contains(restart)
                    && !((System.Windows.Forms.Timer)Field(context, "speechHealthTimer")).Enabled,
                    "Resume hid unsafe cleanup or scheduled capture.");
            }
        }
        finally { Call(context, "ExitApplication", null, EventArgs.Empty); }
    }
    private sealed class FailingInput : IDisposable
    {
        public int DisposeCount;
        public void Dispose()
        {
            DisposeCount++;
            throw new System.IO.IOException("Synthetic retained microphone ownership; no native capture.");
        }
    }
    private static void AssertRestartRequired(object context, Assembly assembly, string culture)
    {
        object recovery = Field(context, "speechRecovery");
        string restart = culture == "fr-FR" ? "red\u00e9marrez" : "restart";
        Assert((bool)assembly.GetType("JarvisPowerPoint.SpeechInput", true)
            .GetProperty("CaptureBlocked", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null, null),
            "Main cleanup did not latch unsafe ownership process-wide.");
        Assert((bool)recovery.GetType().GetProperty("RestartRequired").GetValue(recovery, null)
            && !(bool)recovery.GetType().GetProperty("CanStart").GetValue(recovery, null),
            "Main cleanup did not require an application restart.");
        Assert(((string)Field(context, "listeningStatus")).Contains(restart)
            && ((Control)Field(Field(context, "panel"), "status")).Text.Contains(restart),
            "Paused main UI hid the localized restart instruction.");
        Assert(!((ToolStripMenuItem)Field(context, "toggleItem")).Enabled
            && !((ToolStripMenuItem)Field(context, "setupItem")).Enabled,
            "Main UI still offered listening restart or setup while ownership was unsafe.");
        Assert(!((Control)Field(Field(context, "panel"), "listenButton")).Enabled
            && !((Control)Field(Field(context, "panel"), "microphoneSetupButton")).Enabled
            && !((ToolStripMenuItem)Field(context, "frenchLanguageItem")).Enabled
            && !((ToolStripMenuItem)Field(context, "englishLanguageItem")).Enabled,
            "Panel or language menus still offered capture settings after unsafe cleanup.");
        Assert(!((System.Windows.Forms.Timer)Field(context, "speechHealthTimer")).Enabled
            && Field(context, "recognizer") == null && Field(context, "microphoneInput") == null,
            "Main cleanup retained capture or scheduled another attempt.");
    }
    private static void CheckPauseCleanupGuard(Type type, Assembly assembly, string culture)
    {
        object context = Activator.CreateInstance(type);
        Application.Idle -= (EventHandler)Delegate.CreateDelegate(typeof(EventHandler), context,
            type.GetMethod("OnFirstIdle", Private));
        try
        {
            var input = new FailingInput();
            Set(context, "currentCultureName", culture);
            Set(context, "microphoneInput", input);
            Set(context, "listeningRequested", true);
            Call(context, "ToggleListening", null, EventArgs.Empty);
            Assert(input.DisposeCount == 1 && Field(context, "microphoneInput") == null,
                "Pausing skipped an input whose recognizer was already absent.");
            Assert(!(bool)Field(context, "listeningRequested"), "Pause cleanup failure reversed the user's pause.");
            AssertRestartRequired(context, assembly, culture);
            Call(context, "ToggleListening", null, EventArgs.Empty);
            Call(context, "InitializeSpeechRecognition");
            AssertRestartRequired(context, assembly, culture);
            string previousCulture = (string)Field(context, "currentCultureName");
            Call(context, "ChangeLanguage", Field(context, culture == "fr-FR" ? "englishLanguageItem" : "frenchLanguageItem"),
                EventArgs.Empty);
            Assert((string)Field(context, "currentCultureName") == previousCulture,
                "Blocked language action changed capture preferences.");
        }
        finally { Call(context, "ExitApplication", null, EventArgs.Empty); }
    }
    private static void CheckSetupCloseCleanupGuard(Type type, Assembly assembly, string culture)
    {
        object context = Activator.CreateInstance(type);
        Application.Idle -= (EventHandler)Delegate.CreateDelegate(typeof(EventHandler), context,
            type.GetMethod("OnFirstIdle", Private));
        using (var closeSetup = new System.Windows.Forms.Timer())
        {
            try
            {
                Set(context, "currentCultureName", culture);
                Set(context, "listeningRequested", false);
                Call(Field(context, "speechRecovery"), "Pause");
                var input = new FailingInput();
                bool opened = false;
                Exception injectionFailure = null;
                closeSetup.Interval = 50;
                closeSetup.Tick += delegate
                {
                    Form dialog = Field(context, "activeSetup") as Form;
                    if (dialog == null) return;
                    closeSetup.Stop();
                    opened = true;
                    try
                    {
                        Type testType = dialog.GetType().GetNestedType("TestSession", BindingFlags.NonPublic);
                        object testSession = System.Runtime.Serialization.FormatterServices.GetUninitializedObject(testType);
                        testType.GetField("input", Private).SetValue(testSession, input);
                        Set(dialog, "session", testSession);
                    }
                    catch (Exception error) { injectionFailure = error; }
                    finally
                    {
                        dialog.DialogResult = DialogResult.Cancel;
                        dialog.Close();
                    }
                };
                closeSetup.Start();
                Call(context, "OpenSetup", false);
                Assert(opened && injectionFailure == null && input.DisposeCount == 1,
                    "Paused setup-close fixture did not dispose its synthetic input.");
                Assert(!(bool)Field(context, "listeningRequested") && Field(context, "activeSetup") == null
                    && !(bool)Field(context, "speechRecovery").GetType().GetProperty("InSetup")
                        .GetValue(Field(context, "speechRecovery"), null),
                    "Cancelled setup reversed pause or retained its modal guard.");
                AssertRestartRequired(context, assembly, culture);
            }
            finally
            {
                closeSetup.Stop();
                Call(context, "ExitApplication", null, EventArgs.Empty);
            }
        }
    }
    private static void CheckSessionExitFailure(Type type)
    {
        object context = Activator.CreateInstance(type);
        Application.Idle -= (EventHandler)Delegate.CreateDelegate(typeof(EventHandler), context,
            type.GetMethod("OnFirstIdle", Private));
        Set(context, "cleanupBlocked", true);
        Set(context, "listeningRequested", false);
        Call(Field(context, "speechRecovery"), "Pause");
        object session = Field(context, "session");
        int threadId = (int)Field(session, "threadId");
        bool timerDisposed = false;
        ((System.Windows.Forms.Timer)Field(context, "speechHealthTimer")).Disposed += delegate { timerDisposed = true; };
        try
        {
            // Force a deterministic session teardown failure without opening a presentation.
            Set(session, "threadId", -1);
            bool failed = false;
            try { Call(context, "ExitApplication", null, EventArgs.Empty); }
            catch (InvalidOperationException) { failed = true; }
            Assert(failed, "Synthetic presentation-session teardown failure was not raised.");
            Assert(timerDisposed && !(bool)Field(context, "systemEventsSubscribed")
                && ((Control)Field(context, "panel")).IsDisposed && ((Control)Field(context, "dispatcher")).IsDisposed,
                "Presentation-session teardown failure skipped speech/UI cleanup.");
        }
        finally
        {
            Set(session, "threadId", threadId);
            Call(session, "Dispose");
            if (!(bool)Field(context, "exiting")) Call(context, "ExitApplication", null, EventArgs.Empty);
        }
    }
    [STAThread]
    private static int Main(string[] args)
    {
        object context = null;
        try
        {
            Assembly assembly = Assembly.LoadFrom(args[0]);
            Type type = assembly.GetType("JarvisPowerPoint.JarvisApplicationContext", true);
            if (args.Length == 3 && (args[1] == "--pause-cleanup" || args[1] == "--setup-close-cleanup"))
            {
                if (args[1] == "--pause-cleanup") CheckPauseCleanupGuard(type, assembly, args[2]);
                else CheckSetupCloseCleanupGuard(type, assembly, args[2]);
                Console.WriteLine("Passed " + assertions + " " + args[1] + " " + args[2]
                    + " main-UI integration assertions; no native capture or PowerPoint.");
                return 0;
            }
            CheckModalSpeechGuard(type);
            context = Activator.CreateInstance(type);
            EventHandler firstIdle = (EventHandler)Delegate.CreateDelegate(typeof(EventHandler), context,
                type.GetMethod("OnFirstIdle", Private));
            Application.Idle -= firstIdle;
            // Triple guard: never initialize speech, deny opening capture, keep requests paused.
            Set(context, "cleanupBlocked", true);
            Set(context, "listeningRequested", false);
            object recovery = Field(context, "speechRecovery");
            Call(recovery, "Pause");
            Call(context, "UpdatePollingSchedule");
            Assert(!((System.Windows.Forms.Timer)Field(context, "presentationTimer")).Enabled, "Paused tray presentation timer enabled.");
            Assert(!((System.Windows.Forms.Timer)Field(context, "speechHealthTimer")).Enabled, "Paused tray speech timer enabled.");
            Assert((bool)Field(context, "systemEventsSubscribed"), "Power recovery event not subscribed.");

            object captured = context;
            var thread = new Thread(delegate() { Call(captured, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Suspend)); });
            thread.Start(); thread.Join();
            Assert((string)Field(context, "listeningStatus") == "", "SystemEvents callback mutated UI state off the STA.");
            Application.DoEvents();
            Assert((bool)recovery.GetType().GetProperty("IsSuspended").GetValue(recovery, null), "Queued suspend not marshaled to UI.");
            Assert(!((System.Diagnostics.Stopwatch)Field(context, "clock")).IsRunning, "Rehearsal clock includes suspended time.");
            Assert(Field(context, "recognizer") == null && Field(context, "microphoneInput") == null, "Suspend unexpectedly opened capture.");
            Call(context, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Resume));
            Assert(((System.Diagnostics.Stopwatch)Field(context, "clock")).IsRunning, "Resume did not restart monotonic clock.");
            Assert(!(bool)recovery.GetType().GetProperty("CanStart").GetValue(recovery, null), "Resume bypassed pause.");
            Assert(!((System.Windows.Forms.Timer)Field(context, "speechHealthTimer")).Enabled, "Paused resume scheduled speech.");
            CheckUnchangedStatus(context);
            CheckSpeechGuards(context, assembly);
            Call(context, "OnSpeechRecognized", new object(), null);
            Call(context, "OnRecognitionCompleted", new object(), null);
            Assert(Field(context, "recognizer") == null, "Stale callbacks created capture.");
            var meterArgs = (AudioLevelUpdatedEventArgs)Activator.CreateInstance(typeof(AudioLevelUpdatedEventArgs),
                BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public, null, new object[] { 80 }, null);
            for (int i = 0; i < 1000; i++) Call(context, "OnAudioLevelUpdated", new object(), meterArgs);
            Assert((int)Field(context, "audioLevel") == 0 && (bool)Field(context, "audioEventPending"),
                "High-frequency audio callbacks updated the UI instead of coalescing.");
            Call(context, "DisposeRecognizer");
            Assert(!(bool)Field(context, "audioEventPending"), "Dispose retained pending audio callbacks.");
            Call(context, "OpenSetup", false);
            Assert(Field(context, "activeSetup") == null && Field(context, "recognizer") == null
                && Field(context, "microphoneInput") == null, "Unsafe cleanup allowed a setup test capture.");
            Assert(!(bool)recovery.GetType().GetProperty("InSetup").GetValue(recovery, null),
                "Blocked setup retained its setup guard.");

            int queued = 0;
            bool presentationTimerDisposed = false, speechTimerDisposed = false, iconDisposed = false;
            ((System.Windows.Forms.Timer)Field(context, "presentationTimer")).Disposed += delegate { presentationTimerDisposed = true; };
            ((System.Windows.Forms.Timer)Field(context, "speechHealthTimer")).Disposed += delegate { speechTimerDisposed = true; };
            ((NotifyIcon)Field(context, "notifyIcon")).Disposed += delegate { iconDisposed = true; };
            thread = new Thread(delegate() { Call(captured, "OnUi", (Action)delegate { queued++; }); });
            thread.Start(); thread.Join();
            Call(context, "ExitApplication", null, EventArgs.Empty);
            Assert(!(bool)Field(context, "systemEventsSubscribed"), "Power handler retained after exit.");
            Assert(presentationTimerDisposed && speechTimerDisposed && iconDisposed, "Exit retained a timer or tray icon.");
            Assert(((Control)Field(context, "panel")).IsDisposed && ((Control)Field(context, "dispatcher")).IsDisposed
                && ((Control)Field(Field(context, "panel"), "indicator")).IsDisposed, "Exit retained a UI dispatcher/panel/indicator.");
            Application.DoEvents();
            Assert(queued == 0, "Queued UI callback acted after exit.");
            Call(context, "ExitApplication", null, EventArgs.Empty);
            Assert(!(bool)Field(context, "systemEventsSubscribed"), "Repeated exit restored power callbacks.");
            Call(context, "OnPowerModeChanged", null, new PowerModeChangedEventArgs(PowerModes.Resume));
            Call(context, "PollSpeechHealth");
            Call(context, "PrepareSpeechStart");
            Call(context, "OnRecognitionCompleted", null, null);
            Call(context, "OnSpeechRecognized", null, null);
            Assert(!(bool)recovery.GetType().GetProperty("Requested").GetValue(recovery, null)
                && Field(context, "recognizer") == null && Field(context, "microphoneInput") == null,
                "Late resume, initialization, or callbacks restored capture after exit.");
            Assert((bool)Field(Field(context, "session"), "disposed"), "Exit retained its presentation session.");
            context = null;
            CheckSessionExitFailure(type);
            CheckProcessCaptureBlock(type, assembly);
            Console.WriteLine("Passed " + assertions + " no-capture STA/lifecycle assertions; no PowerPoint operations or microphone opens.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally { if (context != null && !(bool)Field(context, "exiting")) Call(context, "ExitApplication", null, EventArgs.Empty); }
    }
}
'@
try {
    New-Item -ItemType Directory -Path $output | Out-Null
    $source = Join-Path $output "SpeechLifecycleTests.cs"
    $exe = Join-Path $output "SpeechLifecycleTests.exe"
    [IO.File]::WriteAllText($source, $code, [Text.UTF8Encoding]::new($false))
    & $compiler /nologo /langversion:5 /warnaserror+ /target:exe "/out:$exe" `
        /reference:System.dll /reference:System.Windows.Forms.dll `
        ("/reference:" + (Join-Path (Split-Path -Parent $compiler) "WPF\System.Speech.dll")) $source
    if ($LASTEXITCODE -ne 0) { throw "Speech lifecycle compilation failed." }
    & $exe ([IO.Path]::GetFullPath($Executable))
    if ($LASTEXITCODE -ne 0) { throw "Speech lifecycle tests failed." }
    $integrationFailures = @()
    foreach ($scenario in @("--pause-cleanup", "--setup-close-cleanup")) {
        foreach ($culture in @("fr-FR", "en-US")) {
            $probe = New-Object Diagnostics.Process
            $probe.StartInfo.FileName = $exe
            $probe.StartInfo.Arguments = '"' + [IO.Path]::GetFullPath($Executable) + '" ' + $scenario + ' ' + $culture
            $probe.StartInfo.UseShellExecute = $false
            $probe.StartInfo.CreateNoWindow = $true
            $probe.StartInfo.RedirectStandardOutput = $true
            $probe.StartInfo.RedirectStandardError = $true
            try {
                if (-not $probe.Start()) { throw "Could not start the no-capture integration fixture." }
                if (-not $probe.WaitForExit(30000)) {
                    Stop-Process -Id $probe.Id -Force
                    [void]$probe.WaitForExit(5000)
                    throw "Speech integration fixture timed out: $scenario $culture."
                }
                $stdout = $probe.StandardOutput.ReadToEnd()
                $stderr = $probe.StandardError.ReadToEnd()
                if ($stdout) { Write-Host $stdout.TrimEnd() }
                if ($stderr) { Write-Host $stderr.TrimEnd() }
                if ($probe.ExitCode -ne 0) { $integrationFailures += "$scenario $culture" }
            } finally { $probe.Dispose() }
        }
    }
    if ($integrationFailures.Count -ne 0) {
        throw ("Speech integration fixtures failed: " + ($integrationFailures -join "; "))
    }
} finally {
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Recurse -Force }
}
