using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Speech.Recognition;
using System.Windows.Forms;

namespace JarvisPowerPoint
{
    internal sealed class PreflightContext
    {
        public string CultureName { get; set; }
        public string MicrophoneId { get; set; }
        public bool IsListening { get; set; }
        public bool IndicatorEnabled { get; set; }
        public string IndicatorDisplay { get; set; }
        public bool HasRecentAudio { get; set; }
    }

    internal enum PreflightSeverity { Pass, Warning, Fail }

    internal sealed class PreflightItem
    {
        public string CheckKey { get; set; }
        public PreflightSeverity Severity { get; set; }
        public string Detail { get; set; }
    }

    internal enum PreflightPresentationState { Unavailable, Available, NoShow, MultipleShows }

    internal sealed class PreflightDisplays
    {
        public IList<string> ScreenIds { get; set; }
        public string PrimaryScreenId { get; set; }
        public string PresentationScreenId { get; set; }
    }

    internal sealed class PreflightEvidence
    {
        public bool RecognizersRead { get; set; }
        public IList<string> InstalledCultures { get; set; }
        public bool MicrophonesRead { get; set; }
        public IList<string> MicrophoneIds { get; set; }
        public PreflightPresentationState PresentationState { get; set; }
        public PreflightDisplays Displays { get; set; }
    }

    internal interface IPreflightEnvironment
    {
        IList<string> GetRecognizerCultures();
        IList<string> GetMicrophoneIds();
        PreflightDisplays GetDisplays(IntPtr presentationWindow);
    }

    internal static class PreflightService
    {
        public static List<PreflightItem> Run(PreflightContext context,
            Func<PresentationSnapshot> snapshot, bool english)
        {
            return Run(context, snapshot, english, new WindowsPreflightEnvironment());
        }

        internal static List<PreflightItem> Run(PreflightContext context,
            Func<PresentationSnapshot> snapshot, bool english, IPreflightEnvironment environment)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            if (environment == null) throw new ArgumentNullException("environment");
            if (context == null)
                return new List<PreflightItem> { Item("Context", PreflightSeverity.Fail, english,
                    "Current settings could not be read. Close this dialog and try again.",
                    "Les r\u00e9glages actuels sont indisponibles. Fermez cette fen\u00eatre et r\u00e9essayez.") };
            var evidence = new PreflightEvidence { PresentationState = PreflightPresentationState.Unavailable };
            evidence.RecognizersRead = ReadExpected(delegate {
                evidence.InstalledCultures = environment.GetRecognizerCultures();
            });
            evidence.MicrophonesRead = ReadExpected(delegate {
                evidence.MicrophoneIds = environment.GetMicrophoneIds();
            });
            IntPtr window = IntPtr.Zero;
            try
            {
                PresentationSnapshot current = snapshot();
                evidence.PresentationState = current == null
                    ? PreflightPresentationState.NoShow : PreflightPresentationState.Available;
                if (current != null) window = current.WindowHandle;
            }
            catch (PresentationUnavailableException error)
            {
                switch (error.Reason)
                {
                    case PresentationUnavailableReason.NoSlideShow:
                        evidence.PresentationState = PreflightPresentationState.NoShow;
                        break;
                    case PresentationUnavailableReason.MultipleSlideShows:
                        evidence.PresentationState = PreflightPresentationState.MultipleShows;
                        break;
                }
            }
            catch (Exception error)
            {
                if (!IsExpected(error)) throw;
            }
            ReadExpected(delegate { evidence.Displays = environment.GetDisplays(window); });
            return Evaluate(context, evidence, english);
        }

        internal static List<PreflightItem> Evaluate(PreflightContext context,
            PreflightEvidence evidence, bool english)
        {
            if (context == null) throw new ArgumentNullException("context");
            if (evidence == null) throw new ArgumentNullException("evidence");
            var items = new List<PreflightItem>();
            string culture = DiagnosticRecorder.SafeCulture(context.CultureName);
            if (culture == "unknown")
                items.Add(Item("Language", PreflightSeverity.Fail, english,
                    "Choose French (France) or English (United States) in setup.",
                    "Choisissez le fran\u00e7ais (France) ou l'anglais (\u00c9tats-Unis) dans la configuration."));
            else if (!evidence.RecognizersRead || evidence.InstalledCultures == null)
                items.Add(Item("Language", PreflightSeverity.Fail, english,
                    "Installed speech recognizers could not be enumerated. Review Windows speech settings.",
                    "Les moteurs vocaux install\u00e9s sont indisponibles. V\u00e9rifiez les param\u00e8tres vocaux Windows."));
            else if (!Contains(evidence.InstalledCultures, culture, StringComparison.OrdinalIgnoreCase))
                items.Add(Item("Language", PreflightSeverity.Fail, english,
                    "The selected speech language is not installed. Install it in Windows or change the language.",
                    "La langue vocale choisie n'est pas install\u00e9e. Installez-la dans Windows ou changez de langue."));
            else
                items.Add(Item("Language", PreflightSeverity.Pass, english,
                    "A recognizer for the selected FR/EN language is installed.",
                    "Un moteur vocal pour la langue FR/EN choisie est install\u00e9."));

            if (!evidence.MicrophonesRead || evidence.MicrophoneIds == null)
                items.Add(Item("Microphone", PreflightSeverity.Fail, english,
                    "Microphone availability could not be read. Open setup and refresh devices.",
                    "La disponibilit\u00e9 des microphones est inconnue. Ouvrez la configuration et actualisez les appareils."));
            else
            {
                bool available = !string.IsNullOrEmpty(context.MicrophoneId)
                    && Contains(evidence.MicrophoneIds, context.MicrophoneId, StringComparison.Ordinal);
                if (context.MicrophoneId == "default")
                {
                    bool physicalDevice = false;
                    foreach (string id in evidence.MicrophoneIds)
                        if (!string.IsNullOrEmpty(id) && id != "default") physicalDevice = true;
                    available = available && physicalDevice;
                }
                items.Add(available
                    ? Item("Microphone", PreflightSeverity.Pass, english,
                        "The selected input is listed by Windows. This check did not open or test a microphone.",
                        "L'entr\u00e9e choisie figure dans Windows. Ce contr\u00f4le n'a ni ouvert ni test\u00e9 de microphone.")
                    : Item("Microphone", PreflightSeverity.Fail, english,
                        "The selected input is unavailable. Reconnect it or choose an input in setup; no fallback was selected.",
                        "L'entr\u00e9e choisie est indisponible. Rebranchez-la ou choisissez une entr\u00e9e ; aucun remplacement automatique."));
            }

            items.Add(context.IsListening
                ? Item("Listening", PreflightSeverity.Pass, english,
                    "Speech recognition is reported active.", "La reconnaissance vocale est active.")
                : Item("Listening", PreflightSeverity.Warning, english,
                    "Speech recognition is paused or unavailable. Enable listening when ready, or use keyboard controls.",
                    "La reconnaissance est en pause ou indisponible. Activez l'\u00e9coute au moment voulu ou utilisez le clavier."));
            items.Add(context.IsListening && context.HasRecentAudio
                ? Item("Audio", PreflightSeverity.Pass, english,
                    "Recent input activity was reported; this does not verify recognition accuracy.",
                    "Une activit\u00e9 audio r\u00e9cente a \u00e9t\u00e9 signal\u00e9e ; la pr\u00e9cision vocale n'est pas v\u00e9rifi\u00e9e.")
                : Item("Audio", PreflightSeverity.Warning, english,
                    "No recent audio activity was verified. Silence does not mean a broken microphone; use setup for an explicit test.",
                    "Aucune activit\u00e9 audio r\u00e9cente v\u00e9rifi\u00e9e. Le silence ne prouve pas une panne ; utilisez le test explicite de configuration."));

            switch (evidence.PresentationState)
            {
                case PreflightPresentationState.Available:
                    items.Add(Item("Slideshow", PreflightSeverity.Pass, english,
                        "One active PowerPoint slide show was verified by the presentation connection.",
                        "Un seul diaporama PowerPoint actif a \u00e9t\u00e9 v\u00e9rifi\u00e9 par la connexion."));
                    break;
                case PreflightPresentationState.NoShow:
                    items.Add(Item("Slideshow", PreflightSeverity.Fail, english,
                        "No active PowerPoint slide show. Start exactly one slide show, then refresh.",
                        "Aucun diaporama PowerPoint actif. D\u00e9marrez un seul diaporama, puis actualisez."));
                    break;
                case PreflightPresentationState.MultipleShows:
                    items.Add(Item("Slideshow", PreflightSeverity.Fail, english,
                        "More than one PowerPoint slide show is running. Keep exactly one, then refresh.",
                        "Plusieurs diaporamas PowerPoint sont actifs. Gardez-en un seul, puis actualisez."));
                    break;
                default:
                    items.Add(Item("Slideshow", PreflightSeverity.Fail, english,
                        "The active slide show could not be verified. Check PowerPoint and refresh; no slide or focus was changed.",
                        "Le diaporama actif n'a pas pu \u00eatre v\u00e9rifi\u00e9. V\u00e9rifiez PowerPoint et actualisez ; aucune diapositive ni activation modifi\u00e9e."));
                    break;
            }
            items.Add(EvaluateIndicator(context, evidence.Displays, english));
            return items;
        }

        private static PreflightItem EvaluateIndicator(PreflightContext context, PreflightDisplays displays, bool english)
        {
            if (!context.IndicatorEnabled)
                return Item("Indicator", PreflightSeverity.Warning, english,
                    "The listening indicator is disabled; no on-screen listening cue will be shown.",
                    "L'indicateur d'\u00e9coute est d\u00e9sactiv\u00e9 ; aucun t\u00e9moin d'\u00e9coute ne sera affich\u00e9.");
            if (displays == null || displays.ScreenIds == null || displays.ScreenIds.Count == 0)
                return Item("Indicator", PreflightSeverity.Warning, english,
                    "Display placement is unknown. Verify indicator placement and screen-sharing settings yourself.",
                    "La position sur les \u00e9crans est inconnue. V\u00e9rifiez l'indicateur et le partage d'\u00e9cran vous-m\u00eame.");
            string selected = context.IndicatorDisplay;
            if (string.IsNullOrEmpty(selected))
                return Item("Indicator", PreflightSeverity.Warning, english,
                    "The selected indicator screen is unknown. Verify the selected screen; no primary-screen fallback is assumed.",
                    "L'\u00e9cran choisi pour l'indicateur est inconnu. V\u00e9rifiez la s\u00e9lection ; aucun remplacement par l'\u00e9cran principal n'est suppos\u00e9.");
            if (!Contains(displays.ScreenIds, selected, StringComparison.Ordinal))
                return Item("Indicator", PreflightSeverity.Fail, english,
                    "The selected indicator screen is unavailable. Select a connected screen or hide the indicator.",
                    "L'\u00e9cran choisi pour l'indicateur est indisponible. Choisissez un \u00e9cran connect\u00e9 ou masquez l'indicateur.");
            if (displays.ScreenIds.Count == 1)
                return Item("Indicator", PreflightSeverity.Warning, english,
                    "Only one display is available. The indicator may appear with the slides or in captures; hide it if needed.",
                    "Un seul \u00e9cran est disponible. L'indicateur peut appara\u00eetre avec les diapositives ou dans les captures ; masquez-le au besoin.");
            if (string.IsNullOrEmpty(displays.PresentationScreenId)
                || !Contains(displays.ScreenIds, displays.PresentationScreenId, StringComparison.Ordinal))
                return Item("Indicator", PreflightSeverity.Warning, english,
                    "The slide-show screen is unknown. Do not assume the indicator is hidden from attendees or captures.",
                    "L'\u00e9cran du diaporama est inconnu. Ne supposez pas que l'indicateur est masqu\u00e9 aux participants ou aux captures.");
            if (selected == displays.PresentationScreenId)
                return Item("Indicator", PreflightSeverity.Warning, english,
                    "The indicator and slide show share a screen. Attendees or captures may see the indicator.",
                    "L'indicateur et le diaporama partagent un \u00e9cran. Les participants ou les captures peuvent voir l'indicateur.");
            return Item("Indicator", PreflightSeverity.Pass, english,
                "The indicator is on a different screen. Screen sharing or capture may still expose it; verify what you share.",
                "L'indicateur est sur un autre \u00e9cran. Le partage ou la capture peut encore l'exposer ; v\u00e9rifiez ce que vous partagez.");
        }

        internal static bool IsExpected(Exception error)
        {
            return error is IOException || error is InvalidOperationException || error is ArgumentException
                || error is COMException || error is SecurityException || error is UnauthorizedAccessException
                || error is Win32Exception || error is TimeoutException || error is NotSupportedException
                || error is DllNotFoundException || error is EntryPointNotFoundException || error is BadImageFormatException;
        }

        private static bool ReadExpected(Action action)
        {
            try { action(); return true; }
            catch (Exception error)
            {
                if (!IsExpected(error)) throw;
                return false;
            }
        }

        private static bool Contains(IList<string> values, string target, StringComparison comparison)
        {
            foreach (string value in values)
                if (string.Equals(value, target, comparison)) return true;
            return false;
        }

        internal static PreflightItem Item(string key, PreflightSeverity severity, bool english, string en, string fr)
        {
            return new PreflightItem { CheckKey = key, Severity = severity, Detail = english ? en : fr };
        }

        private sealed class WindowsPreflightEnvironment : IPreflightEnvironment
        {
            public IList<string> GetRecognizerCultures()
            {
                var cultures = new List<string>();
                foreach (RecognizerInfo recognizer in SpeechRecognitionEngine.InstalledRecognizers())
                    cultures.Add(recognizer.Culture.Name);
                return cultures;
            }

            public IList<string> GetMicrophoneIds()
            {
                var ids = new List<string>();
                foreach (AudioInputDevice device in SpeechInput.GetDevices()) ids.Add(device.Id);
                return ids;
            }

            public PreflightDisplays GetDisplays(IntPtr presentationWindow)
            {
                var displays = new PreflightDisplays { ScreenIds = new List<string>() };
                foreach (Screen screen in Screen.AllScreens)
                {
                    displays.ScreenIds.Add(screen.DeviceName);
                    if (screen.Primary) displays.PrimaryScreenId = screen.DeviceName;
                }
                if (presentationWindow != IntPtr.Zero && IsWindow(presentationWindow))
                    displays.PresentationScreenId = Screen.FromHandle(presentationWindow).DeviceName;
                return displays;
            }

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool IsWindow(IntPtr window);
        }
    }

    internal sealed class PreflightDialog : Form
    {
        private readonly bool english;
        private readonly Func<PreflightContext> getContext;
        private readonly Func<PresentationSnapshot> getSnapshot;
        private readonly Action openSetup;
        private readonly IPreflightEnvironment environment;
        private readonly DataGridView results;

        public PreflightDialog(bool english, Func<PreflightContext> getContext,
            Func<PresentationSnapshot> getSnapshot, Action openSetup)
            : this(english, getContext, getSnapshot, openSetup, null) { }

        internal PreflightDialog(bool english, Func<PreflightContext> getContext,
            Func<PresentationSnapshot> getSnapshot, Action openSetup, IPreflightEnvironment environment)
        {
            if (getContext == null) throw new ArgumentNullException("getContext");
            if (getSnapshot == null) throw new ArgumentNullException("getSnapshot");
            this.english = english;
            this.getContext = getContext;
            this.getSnapshot = getSnapshot;
            this.openSetup = openSetup;
            this.environment = environment;
            Text = english ? "Pre-presentation diagnostic" : "Diagnostic avant pr\u00e9sentation";
            AccessibleName = Text;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = true;
            MinimumSize = new Size(720, 480);
            ClientSize = new Size(940, 550);
            Font = SystemFonts.MessageBoxFont;

            results = new DataGridView {
                Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false,
                AllowUserToDeleteRows = false, AllowUserToResizeRows = false,
                RowHeadersVisible = false, AutoGenerateColumns = false, MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                BackgroundColor = SystemColors.Window, BorderStyle = BorderStyle.FixedSingle,
                AccessibleName = english ? "Diagnostic results" : "R\u00e9sultats du diagnostic"
            };
            results.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            results.DefaultCellStyle.ForeColor = SystemColors.WindowText;
            results.DefaultCellStyle.BackColor = SystemColors.Window;
            results.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
            results.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
            results.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            results.StandardTab = true;
            results.TabIndex = 0;
            results.Columns.Add(new DataGridViewTextBoxColumn {
                Name = "Status", HeaderText = english ? "Status" : "\u00c9tat",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            results.Columns.Add(new DataGridViewTextBoxColumn {
                Name = "Check", HeaderText = english ? "Check" : "Contr\u00f4le",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            results.Columns.Add(new DataGridViewTextBoxColumn {
                Name = "Detail", HeaderText = english ? "Details" : "D\u00e9tails",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            var notice = new Label {
                Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(0, 8, 0, 8), UseMnemonic = false,
                Text = english
                    ? "Read-only checks: no extra microphone capture, slide changes or preference changes. "
                        + "Results are a snapshot, not a privacy guarantee. Use setup only when you want an explicit audio test."
                    : "Contr\u00f4les en lecture seule : aucune capture micro suppl\u00e9mentaire, aucune modification "
                        + "des diapositives ou r\u00e9glages. R\u00e9sultats instantan\u00e9s, sans garantie de confidentialit\u00e9. "
                        + "Utilisez la configuration pour un test audio explicite."
            };
            var refresh = new Button { Text = english ? "Refresh" : "Actualiser", AutoSize = true };
            var setup = new Button {
                Text = english ? "Open setup..." : "Configuration...", AutoSize = true, Enabled = openSetup != null
            };
            var close = new Button {
                Text = english ? "Close" : "Fermer", AutoSize = true, DialogResult = DialogResult.Cancel
            };
            UiAccessibility.Caption(refresh, english ? "&Refresh" : "&Actualiser");
            UiAccessibility.Caption(setup, english ? "&Open setup..." : "&Configuration...");
            UiAccessibility.Caption(close, english ? "&Close" : "&Fermer");
            refresh.TabIndex = 0;
            setup.TabIndex = 1;
            close.TabIndex = 2;
            refresh.Click += delegate { RefreshResults(); };
            setup.Click += delegate { OpenSetupAndRefresh(); };
            var buttons = new FlowLayoutPanel {
                Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, TabIndex = 1
            };
            buttons.Controls.Add(close);
            buttons.Controls.Add(setup);
            buttons.Controls.Add(refresh);
            var layout = new TableLayoutPanel {
                Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 3
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(results, 0, 0);
            layout.Controls.Add(notice, 0, 1);
            layout.Controls.Add(buttons, 0, 2);
            Controls.Add(layout);
            AcceptButton = refresh;
            CancelButton = close;
            Shown += delegate { RefreshResults(); };
            UiAccessibility.Initialize(this);
        }

        internal void RefreshResults()
        {
            List<PreflightItem> items;
            try
            {
                PreflightContext context = getContext();
                items = environment == null ? PreflightService.Run(context, getSnapshot, english)
                    : PreflightService.Run(context, getSnapshot, english, environment);
            }
            catch (Exception error)
            {
                if (!PreflightService.IsExpected(error)) throw;
                items = new List<PreflightItem> { PreflightService.Item("Context", PreflightSeverity.Fail, english,
                    "The diagnostic could not read current state. Close this dialog and try again.",
                    "L'\u00e9tat actuel est indisponible. Fermez cette fen\u00eatre et r\u00e9essayez.") };
            }
            results.Rows.Clear();
            foreach (PreflightItem item in items) AddResult(item);
        }

        internal void OpenSetupAndRefresh()
        {
            if (openSetup == null) return;
            bool failed = false;
            try { openSetup(); }
            catch (Exception error)
            {
                if (!PreflightService.IsExpected(error)) throw;
                failed = true;
            }
            RefreshResults();
            if (failed) AddResult(PreflightService.Item("Setup", PreflightSeverity.Fail, english,
                "Setup could not be completed. Close this dialog and reopen setup.",
                "La configuration n'a pas abouti. Fermez cette fen\u00eatre et rouvrez la configuration."));
        }

        private void AddResult(PreflightItem item)
        {
            string status = item.Severity == PreflightSeverity.Pass ? (english ? "Pass" : "OK")
                : item.Severity == PreflightSeverity.Warning ? (english ? "Warning" : "Attention")
                : (english ? "Fail" : "\u00c9chec");
            int row = results.Rows.Add(status, CheckLabel(item.CheckKey), item.Detail);
            results.Rows[row].Cells[0].Style.ForeColor = SystemColors.WindowText;
        }

        private string CheckLabel(string key)
        {
            switch (key)
            {
                case "Language": return english ? "Speech language" : "Langue vocale";
                case "Microphone": return "Microphone";
                case "Listening": return english ? "Listening" : "\u00c9coute";
                case "Audio": return english ? "Audio activity" : "Activit\u00e9 audio";
                case "Slideshow": return english ? "Slide show" : "Diaporama";
                case "Indicator": return english ? "Indicator screen" : "\u00c9cran de l'indicateur";
                case "Setup": return english ? "Setup" : "Configuration";
                default: return english ? "Current state" : "\u00c9tat actuel";
            }
        }
    }
}
