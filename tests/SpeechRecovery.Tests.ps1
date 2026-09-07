[CmdletBinding()]
param()
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$output = Join-Path $PSScriptRoot ("speech-recovery-" + [Guid]::NewGuid().ToString("N"))
$compiler = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (!$compiler) { throw "Windows .NET Framework compiler required." }
$code = @'
using System;
using JarvisPowerPoint;
internal static class RecoveryTests
{
    private static int assertions;
    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
        assertions++;
    }
    private static TimeSpan Seconds(double value) { return TimeSpan.FromSeconds(value); }
    private static void Main()
    {
        var recovery = new SpeechRecovery();
        Assert(!recovery.CanStart, "Unrequested capture was permitted.");
        recovery.RequestStart();
        Assert(recovery.CanStart && !recovery.Succeeded(), "Initial start incorrectly marked recovered.");
        int first = recovery.Generation;
        Assert(recovery.AcceptsCallback(first), "Current callback rejected.");
        recovery.Fail(Seconds(0), true);
        Assert(!recovery.AcceptsCallback(first), "Fault allowed a queued stale callback.");
        int[] due = { 1, 3, 8, 18 };
        for (int i = 0; i < due.Length; i++)
        {
            Assert(recovery.State == SpeechRecoveryState.Waiting, "Missing waiting state.");
            Assert(!recovery.TryBeginRetry(Seconds(due[i] - 0.001)), "Retry ran before backoff elapsed.");
            Assert(recovery.TryBeginRetry(Seconds(due[i])), "Due retry did not run.");
            Assert(!recovery.TryBeginRetry(Seconds(due[i])), "One attempt dispatched twice.");
            Assert(recovery.Attempts == i + 1 && recovery.State == SpeechRecoveryState.Recovering, "Retry accounting wrong.");
            if (i == 1)
            {
                Assert(recovery.Succeeded(), "Recovery not explicitly reported.");
                Assert(recovery.AcceptsCallback(recovery.Generation), "Recovered callback blocked.");
            }
            recovery.Fail(Seconds(due[i]), true);
        }
        Assert(recovery.State == SpeechRecoveryState.Error && !recovery.TryBeginRetry(Seconds(9999)),
            "Retry budget was reset by brief success or exhaustion kept polling.");
        recovery.RequestStart();
        Assert(recovery.Attempts == 0, "Explicit restart did not renew budget.");
        recovery.Fail(Seconds(0), false);
        Assert(recovery.State == SpeechRecoveryState.Error && !recovery.TryBeginRetry(Seconds(10)),
            "Unpinned Windows default was automatically rebound.");
        recovery.RequestStart();
        recovery.Succeeded();
        first = recovery.Generation;
        recovery.Suspend();
        Assert(!recovery.CanStart && !recovery.AcceptsCallback(first), "Suspension did not revoke callbacks.");
        recovery.Resume(Seconds(100), true);
        Assert(!recovery.TryBeginRetry(Seconds(100)) && recovery.TryBeginRetry(Seconds(101)),
            "Resume must wait for driver settling and preserve requested selected mic.");
        recovery.Succeeded();
        recovery.Suspend();
        recovery.Resume(Seconds(200), false);
        Assert(recovery.State == SpeechRecoveryState.Error, "Default changed after sleep without explicit selection.");
        recovery.Pause();
        recovery.Suspend();
        recovery.Resume(Seconds(300), true);
        Assert(!recovery.CanStart && !recovery.TryBeginRetry(Seconds(400)), "Resume bypassed pause.");
        recovery.RequestStart();
        recovery.Setup(true);
        recovery.Fail(Seconds(0), true);
        recovery.Suspend();
        recovery.Resume(Seconds(1), true);
        Assert(!recovery.CanStart && !recovery.TryBeginRetry(Seconds(10)), "Recovery bypassed setup.");
        recovery.Setup(false);
        Assert(recovery.CanStart && !recovery.TryBeginRetry(Seconds(10)), "Closing setup triggered an implicit retry.");
        recovery.RequestStart();
        recovery.Succeeded();
        first = recovery.Generation;
        recovery.Pause();
        recovery.RequestStart();
        recovery.Succeeded();
        Assert(!recovery.AcceptsCallback(first), "Queued callback survived explicit restart.");
        recovery.Exit();
        recovery.RequestStart(); recovery.Setup(false); recovery.Resume(Seconds(0), true);
        recovery.Fail(Seconds(0), true); recovery.Pause(); recovery.Suspend();
        Assert(recovery.State == SpeechRecoveryState.Exiting && !recovery.CanStart, "Exit was reversible.");

        var blocked = new SpeechRecovery();
        blocked.RequestStart();
        blocked.Succeeded();
        int blockedGeneration = blocked.Generation;
        blocked.RequireRestart();
        Assert(blocked.RestartRequired && !blocked.CanStart && blocked.State == SpeechRecoveryState.Error
            && !blocked.AcceptsCallback(blockedGeneration), "Unsafe ownership retained capture or queued commands.");
        blocked.RequestStart();
        blocked.Fail(Seconds(0), true);
        Assert(!blocked.TryBeginRetry(Seconds(999)) && !blocked.Succeeded()
            && blocked.Attempts == 0, "Explicit start or automatic retry bypassed restart-required ownership.");
        blocked.Pause();
        blocked.RequestStart();
        blocked.Setup(true);
        blocked.Setup(false);
        blocked.Suspend();
        blocked.Resume(Seconds(1000), true);
        Assert(blocked.RestartRequired && !blocked.Requested && !blocked.CanStart
            && !blocked.InSetup && !blocked.IsSuspended && blocked.State == SpeechRecoveryState.Error,
            "Pause, setup, or Windows resume cleared the restart requirement.");
        blockedGeneration = blocked.Generation;
        blocked.RequireRestart();
        Assert(blocked.Generation == blockedGeneration, "Repeated restart status invalidated state needlessly.");
        blocked.Exit();
        blocked.RequireRestart();
        blocked.RequestStart();
        Assert(blocked.State == SpeechRecoveryState.Exiting && !blocked.CanStart,
            "Restart-required reporting resurrected an exited recovery.");

        foreach (string grammar in new[] { "JarvisNext", "JarvisPrevious", "JarvisActions", "JarvisGoToSlide" })
        {
            Assert(!SpeechConfidence.Accept(grammar, 0.749f), "Short confidence threshold lowered.");
            Assert(SpeechConfidence.Accept(grammar, 0.75f), "Short threshold boundary rejected.");
        }
        foreach (string grammar in new[] { "JarvisSearch", "JarvisAlias" })
        {
            Assert(!SpeechConfidence.Accept(grammar, 0.849f), "Dictation threshold lowered.");
            Assert(SpeechConfidence.Accept(grammar, 0.85f), "Dictation boundary rejected.");
        }
        foreach (float value in new[] { float.NaN, float.PositiveInfinity, -1f, 1.01f })
            Assert(!SpeechConfidence.Accept("JarvisNext", value), "Invalid confidence accepted.");
        Assert(!SpeechConfidence.Accept("Unknown", 1f), "Unknown grammar accepted.");
        Assert(IdlePolling.PresentationInterval(false, false, false, false, false) == 0, "Idle COM polling still runs.");
        Assert(IdlePolling.PresentationInterval(false, false, true, false, false) == 5000, "Return-point idle interval incorrect.");
        Assert(IdlePolling.PresentationInterval(true, false, false, false, false) == 500, "Visible panel not responsive.");
        Assert(IdlePolling.PresentationInterval(false, true, false, false, false) == 500, "Rehearsal not responsive.");
        Assert(IdlePolling.PresentationInterval(true, true, true, true, false) == 0, "Suspended timer still runs.");
        Assert(IdlePolling.SpeechInterval(true, false, false, false, false) == 1000, "Native fault detection not timely.");
        Assert(IdlePolling.SpeechInterval(true, false, true, false, false) == 500, "Visible meter not responsive.");
        Assert(IdlePolling.SpeechInterval(false, true, false, false, false) == 1000, "Waiting retry timer absent.");
        Assert(IdlePolling.SpeechInterval(false, false, false, false, false) == 0, "Paused/error capture polling continues.");
        Assert(IdlePolling.SpeechInterval(true, true, true, false, true) == 0, "Exiting health timer continues.");
        Assert(IdlePolling.SpeechInterval(true, true, true, true, false) == 0, "Suspended health timer continues.");
        int presentationTicks = 0, healthTicks = 0, returnPointTicks = 0;
        for (int ms = 1; ms <= 60000; ms++)
        {
            int presentation = IdlePolling.PresentationInterval(false, false, false, false, false);
            if (presentation > 0 && ms % presentation == 0) presentationTicks++;
            if (ms % IdlePolling.SpeechInterval(true, false, false, false, false) == 0) healthTicks++;
            if (ms % IdlePolling.PresentationInterval(false, false, true, false, false) == 0) returnPointTicks++;
        }
        Assert(presentationTicks == 0 && healthTicks == 60 && returnPointTicks == 12, "Deterministic idle schedule regressed.");
        Console.WriteLine("Passed " + assertions + " recovery/confidence/polling assertions; no microphone, UI or PowerPoint opened.");
        Console.WriteLine("Simulated 60s: tray-only 0 presentation + 60 health ticks (v1.4: 120 combined); return-point-only 12 presentation ticks. Paused tray: 0 ticks. Native capture cadence unchanged at 20ms.");
    }
}
'@
try {
    New-Item -ItemType Directory -Path $output | Out-Null
    $source = Join-Path $output "RecoveryTests.cs"
    $exe = Join-Path $output "RecoveryTests.exe"
    [IO.File]::WriteAllText($source, $code, [Text.UTF8Encoding]::new($false))
    & $compiler /nologo /langversion:5 /warnaserror+ /target:exe "/out:$exe" /reference:System.dll `
        (Join-Path $root "SpeechRecovery.cs") $source
    if ($LASTEXITCODE -ne 0) { throw "Recovery test compilation failed." }
    & $exe
    if ($LASTEXITCODE -ne 0) { throw "Recovery tests failed." }
} finally {
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Recurse -Force }
}
