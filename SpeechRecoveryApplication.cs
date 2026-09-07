using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace JarvisPowerPoint
{
    internal sealed partial class JarvisApplicationContext
    {
        private readonly SpeechRecovery speechRecovery = new SpeechRecovery();
        private readonly System.Windows.Forms.Timer speechHealthTimer = new System.Windows.Forms.Timer();
        private readonly object audioEventGate = new object();
        private object latestAudioSender;
        private int latestAudioLevel;
        private bool audioEventPending;
        private bool cleanupBlocked;
        private bool speechFailureNotified;
        private bool systemEventsSubscribed;
        private string previousPanelStatus;
        private string previousPanelHeard;
        private string previousPanelOutcome;
        private int previousPanelLevel = -1;

        private void InitializeSpeechRecovery()
        {
            speechRecovery.RequestStart();
            speechHealthTimer.Tick += delegate { PollSpeechHealth(); };
            panel.VisibleChanged += delegate { UpdatePollingSchedule(); UpdateStatus(); };
            SystemEvents.PowerModeChanged += OnPowerModeChanged;
            systemEventsSubscribed = true;
            UpdatePollingSchedule();
        }

        private void OnPowerModeChanged(object sender, PowerModeChangedEventArgs args)
        {
            OnUi(delegate
            {
                if (exiting) return;
                if (args.Mode == PowerModes.Suspend)
                {
                    clock.Stop();
                    speechRecovery.Suspend();
                    Exception error = SpeechInputErrors.CaptureExpectedFailure(DisposeRecognizer);
                    if (error != null)
                    {
                        cleanupBlocked = true;
                        diagnostics.Record(DiagnosticCategory.Speech, false, error);
                        lastOutcome = error.Message;
                    }
                    listeningStatus = cleanupBlocked || SpeechInput.CaptureBlocked
                        ? SpeechInput.RestartRequiredMessage(English)
                        : (English ? "Suspended; microphone released" : "Veille ; microphone libéré");
                }
                else if (args.Mode == PowerModes.Resume)
                {
                    clock.Start();
                    speechRecovery.Resume(clock.Elapsed, AutomaticRecoveryAllowed);
                    SetRecoveryStatus();
                }
                UpdatePollingSchedule();
                UpdateStatus();
            });
        }

        private bool AutomaticRecoveryAllowed
        {
            // SAPI's "default" selection has no pinned endpoint identity here. Never rebind it
            // after a fault/sleep without an explicit user restart; Windows may have changed it.
            get { return microphoneId != "default" && !cleanupBlocked && !SpeechInput.CaptureBlocked
                && !speechRecovery.RestartRequired; }
        }

        private void PrepareSpeechStart()
        {
            if (exiting) return;
            if (cleanupBlocked || SpeechInput.CaptureBlocked || speechRecovery.RestartRequired)
            {
                SetRecoveryStatus();
                UpdatePollingSchedule();
                UpdateStatus();
                return;
            }
            speechFailureNotified = false;
            speechRecovery.RequestStart();
            InitializeSpeechRecognition();
        }

        private void SpeechFailed(Exception failure)
        {
            Exception cleanup = SpeechInputErrors.CaptureExpectedFailure(DisposeRecognizer);
            if (cleanup != null)
            {
                cleanupBlocked = true;
                diagnostics.Record(DiagnosticCategory.Speech, false, cleanup);
            }
            if (SpeechInputErrors.HasUnsafeCleanup(failure) || SpeechInput.CaptureBlocked) cleanupBlocked = true;
            diagnostics.Record(DiagnosticCategory.Speech, false, failure);
            lastOutcome = failure.Message + (cleanup == null ? "" : "\n" + cleanup.Message);
            speechRecovery.Fail(clock.Elapsed, AutomaticRecoveryAllowed);
            SetRecoveryStatus();
            UpdatePollingSchedule();
            UpdateStatus();
            if (!speechFailureNotified && !exiting)
            {
                speechFailureNotified = true;
                ShowNotification(English ? "Microphone interrupted" : "Microphone interrompu",
                    cleanupBlocked ? SpeechInput.RestartRequiredMessage(English)
                        : English ? "Open Presenter tools for recovery status. No other microphone was selected."
                            : "Ouvrez les outils présentateur pour l’état de reprise. Aucun autre microphone sélectionné.");
            }
        }

        private void SetRecoveryStatus()
        {
            if (cleanupBlocked || SpeechInput.CaptureBlocked || speechRecovery.RestartRequired)
            {
                cleanupBlocked = true;
                speechRecovery.RequireRestart();
                listeningStatus = SpeechInput.RestartRequiredMessage(English);
            }
            else if (speechRecovery.State == SpeechRecoveryState.Waiting)
                listeningStatus = English ? "Waiting to retry the same microphone" : "En attente de reprise du même microphone";
            else if (speechRecovery.State == SpeechRecoveryState.Error)
                listeningStatus = microphoneId == "default"
                        ? (English ? "Windows default may have changed; pause/resume or open setup to explicitly select it again"
                            : "Le micro Windows par défaut peut avoir changé ; pause/reprise ou réglages pour le sélectionner explicitement")
                        : (English ? "Recovery stopped after four retries; reconnect and pause/resume or open setup"
                            : "Reprise arrêtée après quatre essais ; reconnectez puis pause/reprise ou réglages");
            else if (!listeningRequested) listeningStatus = "";
        }

        private void PollSpeechHealth()
        {
            if (exiting || speechRecovery.IsSuspended || activeSetup != null) return;
            if (cleanupBlocked || SpeechInput.CaptureBlocked || speechRecovery.RestartRequired)
            {
                if (recognizer != null || microphoneInput != null)
                    SpeechFailed(new SpeechCleanupException(SpeechInput.RestartRequiredMessage(English)));
                else
                {
                    SetRecoveryStatus();
                    UpdatePollingSchedule();
                    UpdateStatus();
                }
                return;
            }
            Exception error = SpeechInput.GetError(microphoneInput);
            if (error != null) { SpeechFailed(error); return; }
            if (speechRecovery.TryBeginRetry(clock.Elapsed))
            {
                listeningStatus = English ? "Recovering the selected microphone" : "Reprise du microphone sélectionné";
                UpdateStatus();
                InitializeSpeechRecognition();
                return;
            }
            if (recognitionRunning)
            {
                if (microphoneInput != null) audioLevel = SpeechInput.GetLevel(microphoneInput);
                else
                {
                    lock (audioEventGate)
                    {
                        if (audioEventPending && latestAudioSender == recognizer)
                            audioLevel = latestAudioLevel;
                        audioEventPending = false;
                    }
                }
                if (audioLevel > 3) lastAudioUtc = DateTime.UtcNow;
                if (panel.Visible || panel.IndicatorEnabled) UpdateStatus();
            }
            UpdatePollingSchedule();
        }

        private void UpdatePollingSchedule()
        {
            if (exiting || speechHealthTimer == null || panel == null) return;
            ApplyTimer(presentationTimer, IdlePolling.PresentationInterval(panel.Visible, rehearsal.IsRunning,
                session.HasReturnPoint, speechRecovery.IsSuspended, exiting));
            ApplyTimer(speechHealthTimer, IdlePolling.SpeechInterval(recognitionRunning,
                speechRecovery.State == SpeechRecoveryState.Waiting, panel.Visible, speechRecovery.IsSuspended, exiting));
        }

        private static void ApplyTimer(System.Windows.Forms.Timer timer, int interval)
        {
            if (interval == 0) { if (timer.Enabled) timer.Stop(); return; }
            if (timer.Interval != interval) timer.Interval = interval;
            if (!timer.Enabled) timer.Start();
        }

        private void StopSpeechRecovery()
        {
            speechRecovery.Exit();
            if (systemEventsSubscribed)
            {
                SystemEvents.PowerModeChanged -= OnPowerModeChanged;
                systemEventsSubscribed = false;
            }
            speechHealthTimer.Stop();
            speechHealthTimer.Dispose();
        }
    }
}
