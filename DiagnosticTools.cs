using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace JarvisPowerPoint
{
    internal enum DiagnosticCategory { Speech, Navigation, Settings, Profiles, Hotkeys, Updates, Preflight }
    internal enum ListeningState { Listening, Paused, Unavailable }

    internal sealed class DiagnosticSnapshot
    {
        public DiagnosticSnapshot() { State = ListeningState.Unavailable; }

        public string CultureName { get; set; }
        public bool UsesDefaultMicrophone { get; set; }
        public ListeningState State { get; set; }
        public bool SlideshowAvailable { get; set; }
        public int ScreenCount { get; set; }
        public bool HotkeysEnabled { get; set; }
    }

    internal sealed class DiagnosticRecorder
    {
        private readonly object gate = new object();
        private readonly Queue<DiagnosticEvent> events = new Queue<DiagnosticEvent>();
        private const int EventLimit = 100;
        private static readonly HashSet<Type> SafeErrorTypes = new HashSet<Type> {
            typeof(Exception), typeof(IOException), typeof(FileNotFoundException),
            typeof(DirectoryNotFoundException), typeof(PathTooLongException),
            typeof(UnauthorizedAccessException), typeof(SecurityException),
            typeof(InvalidOperationException), typeof(ObjectDisposedException),
            typeof(ArgumentException), typeof(ArgumentNullException),
            typeof(ArgumentOutOfRangeException), typeof(TimeoutException),
            typeof(NotSupportedException), typeof(COMException),
            typeof(System.ComponentModel.Win32Exception), typeof(AggregateException),
            typeof(DllNotFoundException), typeof(EntryPointNotFoundException),
            typeof(BadImageFormatException)
        };

        public DiagnosticRecorder() { }

        public void Record(DiagnosticCategory category, bool success, Exception error = null)
        {
            if (!Enum.IsDefined(typeof(DiagnosticCategory), category))
                throw new ArgumentOutOfRangeException("category");
            // Never retain an Exception: messages, stack traces, Data and inner errors may contain secrets.
            var entry = new DiagnosticEvent {
                Timestamp = DateTime.UtcNow,
                Category = category,
                Success = success,
                ErrorType = error == null ? "none"
                    : SafeErrorTypes.Contains(error.GetType()) ? error.GetType().FullName : "System.Exception",
                HResult = error == null ? 0 : error.HResult
            };
            lock (gate)
            {
                if (events.Count == EventLimit) events.Dequeue();
                events.Enqueue(entry);
            }
        }

        public string BuildReport(DiagnosticSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            DiagnosticEvent[] recent;
            lock (gate) recent = events.ToArray();
            var report = new StringBuilder();
            report.AppendLine("Jarvis PowerPoint diagnostic format 1");
            report.AppendLine("Local export; no audio, phrases, presentation content, identifiers or file paths.");
            report.AppendLine("AppVersion: " + typeof(DiagnosticRecorder).Assembly.GetName().Version);
            report.AppendLine("FrameworkVersion: " + Environment.Version);
            report.AppendLine("OSReportedVersion: " + Environment.OSVersion.Version);
            report.AppendLine("ProcessBits: " + (Environment.Is64BitProcess ? "64" : "32"));
            report.AppendLine("Culture: " + SafeCulture(snapshot.CultureName));
            report.AppendLine("DefaultMicrophone: " + YesNo(snapshot.UsesDefaultMicrophone));
            report.AppendLine("ListeningState: " + SafeState(snapshot.State));
            report.AppendLine("SlideshowAvailable: " + YesNo(snapshot.SlideshowAvailable));
            report.AppendLine("ScreenCount: " + Math.Max(0, Math.Min(64, snapshot.ScreenCount))
                .ToString(CultureInfo.InvariantCulture));
            report.AppendLine("HotkeysEnabled: " + YesNo(snapshot.HotkeysEnabled));
            report.AppendLine("RecentEvents: " + recent.Length.ToString(CultureInfo.InvariantCulture));
            report.AppendLine("UTC | Category | Outcome | ErrorType | HRESULT");
            foreach (DiagnosticEvent entry in recent)
            {
                report.Append(entry.Timestamp.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture));
                report.Append(" | ").Append(entry.Category.ToString());
                report.Append(entry.Success ? " | Success | " : " | Failure | ");
                report.Append(entry.ErrorType).Append(" | 0x");
                report.AppendLine(unchecked((uint)entry.HResult).ToString("X8", CultureInfo.InvariantCulture));
            }
            return report.ToString();
        }

        public void Export(string path, DiagnosticSnapshot snapshot)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Choose a local export file.", "path");
            string fullPath = Path.GetFullPath(path);
            if (fullPath.StartsWith(@"\\", StringComparison.Ordinal)
                || new DriveInfo(Path.GetPathRoot(fullPath)).DriveType == DriveType.Network)
                throw new ArgumentException("Choose a local drive, not a network location.", "path");
            File.WriteAllText(fullPath, BuildReport(snapshot), new UTF8Encoding(false));
        }

        internal static string SafeCulture(string culture)
        {
            if (string.Equals(culture, "fr-FR", StringComparison.OrdinalIgnoreCase)) return "fr-FR";
            if (string.Equals(culture, "en-US", StringComparison.OrdinalIgnoreCase)) return "en-US";
            return "unknown";
        }

        private static string SafeState(ListeningState state)
        {
            switch (state)
            {
                case ListeningState.Listening: return "Listening";
                case ListeningState.Paused: return "Paused";
                case ListeningState.Unavailable: return "Unavailable";
                default: return "unknown";
            }
        }

        private static string YesNo(bool value) { return value ? "yes" : "no"; }

        private sealed class DiagnosticEvent
        {
            public DateTime Timestamp;
            public DiagnosticCategory Category;
            public bool Success;
            public string ErrorType;
            public int HResult;
        }
    }
}
