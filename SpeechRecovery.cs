using System;

namespace JarvisPowerPoint
{
    internal enum SpeechRecoveryState { Stopped, Listening, Suspended, Waiting, Recovering, Error, Exiting }

    // UI-STA owned; callers supply monotonic time so sleep and wall-clock changes cannot create retry storms.
    internal sealed class SpeechRecovery
    {
        private static readonly int[] RetrySeconds = { 1, 2, 5, 10 };
        public SpeechRecoveryState State { get; private set; }
        public int Attempts { get; private set; }
        public TimeSpan NextAttempt { get; private set; }
        public bool Requested { get; private set; }
        public bool InSetup { get; private set; }
        public bool IsSuspended { get; private set; }
        public bool RestartRequired { get; private set; }
        public int Generation { get; private set; }
        public bool CanStart { get { return Requested && !RestartRequired && !InSetup && !IsSuspended && State != SpeechRecoveryState.Exiting; } }
        private SpeechRecoveryState InactiveState
        {
            get { return IsSuspended ? SpeechRecoveryState.Suspended
                : RestartRequired ? SpeechRecoveryState.Error : SpeechRecoveryState.Stopped; }
        }

        public void RequireRestart()
        {
            if (RestartRequired || State == SpeechRecoveryState.Exiting) return;
            RestartRequired = true;
            Generation++;
            State = InactiveState;
        }

        public void RequestStart()
        {
            if (RestartRequired || State == SpeechRecoveryState.Exiting) return;
            Requested = true;
            Attempts = 0;
            Generation++;
            State = IsSuspended ? SpeechRecoveryState.Suspended : SpeechRecoveryState.Stopped;
        }

        public void Pause()
        {
            if (State == SpeechRecoveryState.Exiting) return;
            Requested = false;
            Generation++;
            State = InactiveState;
        }

        public void Setup(bool open)
        {
            if (State == SpeechRecoveryState.Exiting) return;
            InSetup = open;
            Generation++;
            State = InactiveState;
        }

        public void Suspend()
        {
            if (State == SpeechRecoveryState.Exiting) return;
            IsSuspended = true;
            Generation++;
            State = SpeechRecoveryState.Suspended;
        }

        public void Resume(TimeSpan now, bool automaticAllowed)
        {
            if (!IsSuspended || State == SpeechRecoveryState.Exiting) return;
            IsSuspended = false;
            Attempts = 0;
            Fail(now, automaticAllowed);
        }

        public void Fail(TimeSpan now, bool automaticAllowed)
        {
            if (State == SpeechRecoveryState.Exiting) return;
            Generation++;
            if (!CanStart) { State = InactiveState; return; }
            if (!automaticAllowed || Attempts >= RetrySeconds.Length) { State = SpeechRecoveryState.Error; return; }
            NextAttempt = now + TimeSpan.FromSeconds(RetrySeconds[Attempts]);
            State = SpeechRecoveryState.Waiting;
        }

        public bool TryBeginRetry(TimeSpan now)
        {
            if (!CanStart || State != SpeechRecoveryState.Waiting || now < NextAttempt) return false;
            Attempts++;
            State = SpeechRecoveryState.Recovering;
            return true;
        }

        public bool Succeeded()
        {
            if (!CanStart) return false;
            bool recovered = Attempts > 0;
            // Keep the retry budget across brief successes; only an explicit start or resume renews it.
            State = SpeechRecoveryState.Listening;
            return recovered;
        }

        public bool AcceptsCallback(int generation)
        {
            return CanStart && State == SpeechRecoveryState.Listening && generation == Generation;
        }

        public void Exit() { Generation++; Requested = false; State = SpeechRecoveryState.Exiting; }
    }

    internal static class SpeechConfidence
    {
        // SAPI confidence is a heuristic score, not a calibrated probability of correctness.
        public const float ShortCommand = 0.75f;
        public const float FreeText = 0.85f;
        public static bool Accept(string grammar, float confidence)
        {
            if (float.IsNaN(confidence) || float.IsInfinity(confidence) || confidence < 0 || confidence > 1) return false;
            switch (grammar)
            {
                case "JarvisSearch":
                case "JarvisAlias": return confidence >= FreeText;
                case "JarvisNext":
                case "JarvisPrevious":
                case "JarvisGoToSlide":
                case "JarvisActions": return confidence >= ShortCommand;
                default: return false;
            }
        }
    }

    internal static class IdlePolling
    {
        public static int PresentationInterval(bool visible, bool rehearsing, bool returnPoint, bool suspended, bool exiting)
        {
            if (suspended || exiting) return 0;
            return visible || rehearsing ? 500 : returnPoint ? 5000 : 0;
        }

        public static int SpeechInterval(bool running, bool waiting, bool visible, bool suspended, bool exiting)
        {
            if (suspended || exiting || (!running && !waiting)) return 0;
            return visible && running ? 500 : 1000;
        }
    }
}
