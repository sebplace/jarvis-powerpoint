using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Speech.Recognition;
using System.Threading;
using System.Windows.Forms;

namespace JarvisPowerPoint
{
    internal sealed class SetupDialog : Form
    {
        private readonly ComboBox languages = new ComboBox();
        private readonly ComboBox microphones = new ComboBox();
        private readonly Label heading = new Label();
        private readonly Label instructions = new Label();
        private readonly Label languageLabel = new Label();
        private readonly Label microphoneLabel = new Label();
        private readonly Label availability = new Label();
        private readonly Label levelLabel = new Label();
        private readonly Label heardLabel = new Label();
        private readonly Label status = new Label();
        private readonly Label privacy = new Label();
        private readonly ProgressBar level = new ProgressBar();
        private readonly CheckBox skip = new CheckBox();
        private readonly Button refresh = new Button();
        private readonly Button test = new Button();
        private readonly Button finish = new Button();
        private readonly Button cancel = new Button();
        private readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        private readonly Font uiFont;
        private readonly Font headingFont;
        private readonly Dictionary<string, RecognizerInfo> recognizers =
            new Dictionary<string, RecognizerInfo>(StringComparer.OrdinalIgnoreCase);
        private IList<AudioInputDevice> devices = new List<AudioInputDevice>();
        private TestSession session;
        private bool updating;
        private bool closing;
        private bool tested;
        private string heard = "";
        private string failure;
        private string enumerationFailure;
        private string recognizerFailure;

        public SetupDialog(string initialCulture, string microphoneId)
        {
            Text = "Jarvis PowerPoint";
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = true;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(720, 670);
            MinimumSize = new Size(650, 660);
            uiFont = new Font(SystemFonts.MessageBoxFont.FontFamily, 10);
            headingFont = new Font(uiFont, FontStyle.Bold);
            Font = uiFont;
            BuildLayout();

            updating = true;
            languages.Items.Add(new LanguageChoice("fr-FR", "Français (France)"));
            languages.Items.Add(new LanguageChoice("en-US", "English (United States)"));
            languages.SelectedIndex = string.Equals(initialCulture, "en-US",
                StringComparison.OrdinalIgnoreCase) ? 1 : 0;
            updating = false;
            LoadChoices(string.IsNullOrEmpty(microphoneId) ? "default" : microphoneId);

            languages.SelectedIndexChanged += SelectionChanged;
            microphones.SelectedIndexChanged += SelectionChanged;
            refresh.Click += delegate { StopTest(); LoadChoices(SelectedMicrophoneId); };
            test.Click += delegate
            {
                if (session != null) StopTest();
                else StartTest();
            };
            skip.CheckedChanged += delegate
            {
                if (skip.Checked) StopTest();
                UpdateText();
            };
            finish.Click += Finish;
            cancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
            AcceptButton = finish;
            CancelButton = cancel;
            timer.Interval = 100;
            timer.Tick += PollTest;
            timer.Start();
            UpdateText();
        }

        public string SelectedCulture
        {
            get
            {
                var language = languages.SelectedItem as LanguageChoice;
                return language == null ? "fr-FR" : language.Culture;
            }
        }

        public string SelectedMicrophoneId
        {
            get
            {
                var microphone = microphones.SelectedItem as AudioInputDevice;
                return microphone == null ? "default" : microphone.Id;
            }
        }

        private bool English { get { return SelectedCulture == "en-US"; } }
        private string T(string french, string english) { return English ? english : french; }

        private void BuildLayout()
        {
            var layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.AutoScroll = true;
            layout.Padding = new Padding(18);
            layout.ColumnCount = 1;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(layout);
            heading.Font = headingFont;
            AddRow(layout, heading);
            AddRow(layout, instructions);
            AddRow(layout, languageLabel);
            languages.DropDownStyle = ComboBoxStyle.DropDownList;
            languages.AccessibleName = "Language";
            AddRow(layout, languages);
            AddRow(layout, microphoneLabel);
            microphones.DropDownStyle = ComboBoxStyle.DropDownList;
            microphones.DropDownWidth = 640;
            AddRow(layout, microphones);
            refresh.AutoSize = true;
            refresh.Anchor = AnchorStyles.Left;
            AddRow(layout, refresh);
            availability.ForeColor = Color.DarkRed;
            AddRow(layout, availability);
            test.AutoSize = true;
            test.Anchor = AnchorStyles.Left;
            AddRow(layout, test);
            AddRow(layout, levelLabel);
            level.Minimum = 0;
            level.Maximum = 100;
            level.Height = 18;
            AddRow(layout, level);
            AddRow(layout, heardLabel);
            AddRow(layout, status);
            skip.AutoSize = true;
            AddRow(layout, skip);
            privacy.ForeColor = SystemColors.GrayText;
            AddRow(layout, privacy);
            var buttons = new FlowLayoutPanel();
            buttons.AutoSize = true;
            buttons.FlowDirection = FlowDirection.RightToLeft;
            finish.AutoSize = true;
            cancel.AutoSize = true;
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(finish);
            AddRow(layout, buttons);
        }

        private static void AddRow(TableLayoutPanel layout, Control control)
        {
            int row = layout.RowCount++;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            control.Margin = new Padding(0, 0, 0, 9);
            control.Dock = DockStyle.Top;
            var label = control as Label;
            if (label != null)
            {
                label.AutoSize = true;
                label.MaximumSize = new Size(650, 0);
            }
            layout.Controls.Add(control, 0, row);
        }

        private void LoadChoices(string selectedId)
        {
            updating = true;
            tested = false;
            heard = "";
            failure = null;
            enumerationFailure = null;
            recognizerFailure = null;
            recognizers.Clear();
            Exception error = SpeechInputErrors.CaptureExpectedFailure(delegate
            {
                foreach (RecognizerInfo recognizer in SpeechRecognitionEngine.InstalledRecognizers())
                {
                    string culture = recognizer.Culture.Name;
                    if ((culture == "fr-FR" || culture == "en-US") && !recognizers.ContainsKey(culture))
                        recognizers.Add(culture, recognizer);
                }
            });
            if (error != null)
            {
                recognizers.Clear();
                recognizerFailure = Describe(error);
            }

            error = SpeechInputErrors.CaptureExpectedFailure(delegate { devices = SpeechInput.GetDevices(); });
            if (error != null)
            {
                enumerationFailure = Describe(error);
                devices = new List<AudioInputDevice>();
            }
            microphones.Enabled = enumerationFailure == null;
            microphones.Items.Clear();
            int selectedIndex = -1;
            foreach (AudioInputDevice device in devices)
            {
                if (device.Id == selectedId) selectedIndex = microphones.Items.Count;
                microphones.Items.Add(device.Id == "default"
                    ? new AudioInputDevice("default", T("Par défaut (Windows)", "Default (Windows)"))
                    : device);
            }
            if (selectedIndex < 0)
            {
                selectedIndex = microphones.Items.Count;
                microphones.Items.Add(new AudioInputDevice(selectedId,
                    UnavailableMicrophoneLabel()));
            }
            microphones.SelectedIndex = selectedIndex;
            updating = false;
            UpdateText();
        }

        private bool DeviceAvailable()
        {
            if (enumerationFailure != null || devices.Count <= 1) return false;
            foreach (AudioInputDevice device in devices)
                if (device.Id == SelectedMicrophoneId) return true;
            return false;
        }

        private bool CanUseSelection()
        {
            return recognizerFailure == null && recognizers.ContainsKey(SelectedCulture) && DeviceAvailable();
        }

        private string UnavailableMicrophoneLabel()
        {
            return enumerationFailure == null
                ? T("Microphone enregistré indisponible — sélectionnez à nouveau",
                    "Saved microphone unavailable — select again")
                : T("Impossible de lister les microphones — actualisez",
                    "Microphone enumeration failed — refresh");
        }

        private void SelectionChanged(object sender, EventArgs e)
        {
            if (updating) return;
            StopTest();
            tested = false;
            heard = "";
            failure = null;
            skip.Checked = false;
            updating = true;
            int selected = microphones.SelectedIndex;
            for (int i = 0; i < microphones.Items.Count; i++)
            {
                var device = (AudioInputDevice)microphones.Items[i];
                if (!ContainsDevice(device.Id))
                    microphones.Items[i] = new AudioInputDevice(device.Id, UnavailableMicrophoneLabel());
                else if (device.Id == "default")
                    microphones.Items[i] = new AudioInputDevice("default",
                        T("Par défaut (Windows)", "Default (Windows)"));
            }
            microphones.SelectedIndex = selected;
            updating = false;
            UpdateText();
        }

        private bool ContainsDevice(string id)
        {
            foreach (AudioInputDevice device in devices)
                if (device.Id == id) return true;
            return false;
        }

        private void UpdateText()
        {
            if (closing || IsDisposed) return;
            Text = T("Configuration de Jarvis PowerPoint", "Jarvis PowerPoint setup");
            heading.Text = T("Bienvenue — langue et microphone", "Welcome — language and microphone");
            instructions.Text = T(
                "Choisissez une langue et un microphone. Le test est facultatif : démarrez-le, puis dites « Jarvis test ».",
                "Choose a language and microphone. Testing is optional: start the test, then say “Jarvis test”.");
            languageLabel.Text = T("1. Langue de reconnaissance", "1. Recognition language");
            microphoneLabel.Text = T("2. Microphone (uniquement pour Jarvis)", "2. Microphone (Jarvis only)");
            languages.AccessibleName = languageLabel.Text;
            microphones.AccessibleName = microphoneLabel.Text;
            refresh.Text = T("Actualiser les langues et microphones", "Refresh languages and microphones");
            test.Text = session == null ? T("3. Démarrer le test local", "3. Start local test")
                : T("Arrêter le test", "Stop test");
            test.Enabled = session != null || CanUseSelection();
            skip.Text = T("Ignorer le test — enregistrer sans test vocal réussi",
                "Skip test — save without a successful voice test");
            finish.Text = T("Enregistrer / Terminer", "Save / Finish");
            cancel.Text = T("Annuler", "Cancel");
            finish.Enabled = CanUseSelection() && (tested || skip.Checked);
            privacy.Text = T("L'audio reste local ; le test ne contrôle pas PowerPoint. "
                + "Aucun enregistrement audio sur disque, aucun envoi réseau.",
                "Audio stays local; test does not control PowerPoint. "
                + "No audio is recorded to disk or sent over the network.");
            levelLabel.Text = T("Niveau d'entrée : ", "Input level: ") + level.Value + "%";
            level.AccessibleName = levelLabel.Text;
            heardLabel.Text = T("Phrase entendue : ", "Heard phrase: ")
                + (heard.Length == 0 ? T("(aucune)", "(none)") : heard);
            var problems = new List<string>();
            if (!recognizers.ContainsKey(SelectedCulture))
                problems.Add(T("Le moteur vocal ", "The speech recognizer ") + SelectedCulture
                    + T(" n'est pas installé pour cette application .NET. Dans Paramètres Windows > "
                        + "Heure et langue > Langue et région > Options linguistiques, installez la reconnaissance vocale "
                        + "Français (France) ou Anglais (États-Unis). Si nécessaire, utilisez Panneau de configuration > "
                        + "Reconnaissance vocale > Options vocales avancées. Actualisez, ou annulez puis relancez Jarvis.",
                        " is not installed for this .NET application. In Windows Settings > Time & language > "
                        + "Language & region > Language options, install French (France) or English (United States) "
                        + "speech recognition. If needed, use Control Panel > Speech Recognition > Advanced speech "
                        + "options. Refresh, or cancel and restart Jarvis."));
            if (recognizerFailure != null) problems.Add(recognizerFailure);
            if (!DeviceAvailable())
                problems.Add(T("Microphone indisponible. Branchez-le, actualisez et sélectionnez-le à nouveau. "
                    + "Vérifiez les autorisations du microphone pour les applications de bureau dans Paramètres Windows.",
                    "Microphone unavailable. Connect it, refresh and select it again. "
                    + "Check Windows Settings microphone permissions for desktop apps."));
            if (enumerationFailure != null) problems.Add(enumerationFailure);
            availability.Text = string.Join(Environment.NewLine, problems.ToArray());
            availability.Visible = problems.Count != 0;
            status.ForeColor = failure == null ? SystemColors.ControlText : Color.DarkRed;
            if (failure != null)
                status.Text = T("Échec du test : ", "Test failed: ") + failure;
            else if (session != null)
                status.Text = tested
                    ? T("Phrase reconnue. Vous pouvez arrêter et enregistrer.", "Phrase recognized. You can stop and save.")
                    : T("Écoute locale… dites « Jarvis test ».", "Listening locally… say “Jarvis test”.");
            else
                status.Text = tested ? T("Test réussi. Vous pouvez enregistrer.", "Test passed. You can save.")
                    : skip.Checked ? T("Test ignoré. Vous pouvez enregistrer.", "Test skipped. You can save.")
                    : T("Microphone fermé. Lancez le test ou cochez « Ignorer le test ».",
                        "Microphone closed. Start the test or check “Skip test”.");
        }

        private void StartTest()
        {
            if (!CanUseSelection() || closing) return;
            skip.Checked = false;
            tested = false;
            heard = "";
            failure = null;
            Exception error = SpeechInputErrors.CaptureExpectedFailure(delegate
            {
                session = new TestSession(recognizers[SelectedCulture], SelectedMicrophoneId);
            });
            if (error != null) failure = Describe(error);
            UpdateText();
        }

        private void PollTest(object sender, EventArgs e)
        {
            if (closing || session == null) return;
            level.Value = Math.Max(0, Math.Min(100, session.Level));
            string phrase = session.Heard;
            if (!string.IsNullOrEmpty(phrase))
            {
                heard = phrase;
                tested = true;
            }
            bool completed = session.Completed;
            Exception error = session.Error;
            if (error != null || completed)
            {
                failure = error == null ? T("L'écoute s'est arrêtée. Réessayez.",
                    "Listening stopped. Please try again.") : Describe(error);
                tested = false;
                StopTest();
            }
            UpdateText();
        }

        private void StopTest()
        {
            TestSession old = session;
            session = null;
            if (old != null)
            {
                Exception error = SpeechInputErrors.CaptureExpectedFailure(old.Dispose);
                if (error != null) { failure = Describe(error); tested = false; }
            }
            if (!level.IsDisposed) level.Value = 0;
            UpdateText();
        }

        private void Finish(object sender, EventArgs e)
        {
            if (!CanUseSelection() || (!tested && !skip.Checked)) return;
            // Revalidate without opening the microphone; never silently save a reordered device.
            Exception error = SpeechInputErrors.CaptureExpectedFailure(delegate
            {
                IList<AudioInputDevice> current = SpeechInput.GetDevices();
                if (current.Count <= 1)
                    throw new InvalidOperationException(T("Aucun microphone disponible.", "No microphone is available."));
                if (SelectedMicrophoneId != "default")
                    SpeechInput.ResolveDeviceIndex(SelectedMicrophoneId, current);
            });
            if (error != null)
            {
                StopTest();
                failure = Describe(error);
                UpdateText();
                return;
            }
            StopTest();
            DialogResult = DialogResult.OK;
            Close();
        }

        private static string Describe(Exception error)
        {
            string message = error.Message;
            if (error.InnerException != null) message += " " + error.InnerException.Message;
            return message;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (e.Cancel) return;
            closing = true;
            timer.Stop();
            StopTest();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                closing = true;
                timer.Dispose();
                StopTest();
            }
            base.Dispose(disposing);
            if (disposing)
            {
                if (headingFont != null) headingFont.Dispose();
                if (uiFont != null) uiFont.Dispose();
            }
        }

        private sealed class LanguageChoice
        {
            public readonly string Culture;
            private readonly string name;
            public LanguageChoice(string culture, string name) { Culture = culture; this.name = name; }
            public override string ToString() { return name; }
        }

        private sealed class TestSession : IDisposable
        {
            private SpeechRecognitionEngine engine;
            private IDisposable input;
            private int level;
            private string heard;
            private Exception error;
            private volatile bool completed;
            private volatile bool disposed;

            public TestSession(RecognizerInfo recognizer, string microphone)
            {
                try
                {
                    engine = new SpeechRecognitionEngine(recognizer);
                    var builder = new GrammarBuilder();
                    builder.Culture = new CultureInfo(recognizer.Culture.Name);
                    builder.Append("Jarvis test");
                    engine.LoadGrammar(new Grammar(builder));
                    engine.AudioLevelUpdated += OnLevel;
                    engine.SpeechRecognized += OnRecognized;
                    engine.RecognizeCompleted += OnCompleted;
                    input = SpeechInput.Attach(engine, microphone);
                    engine.RecognizeAsync(RecognizeMode.Multiple);
                }
                catch
                {
                    SpeechInputErrors.CleanupPreservingFailure(Dispose);
                    throw;
                }
            }

            // Events only publish state. The WinForms timer reads it on the UI thread, so queued
            // recognizer events never capture controls, invoke a closed window, or control PowerPoint.
            public int Level
            {
                get
                {
                    return input == null ? Interlocked.CompareExchange(ref level, 0, 0)
                        : SpeechInput.GetLevel(input);
                }
            }
            public string Heard { get { return Interlocked.CompareExchange(ref heard, null, null); } }
            public Exception Error
            {
                get { return Interlocked.CompareExchange(ref error, null, null) ?? SpeechInput.GetError(input); }
            }
            public bool Completed { get { return completed; } }

            private void OnLevel(object sender, AudioLevelUpdatedEventArgs e)
            {
                if (!disposed) Interlocked.Exchange(ref level, e.AudioLevel);
            }
            private void OnRecognized(object sender, SpeechRecognizedEventArgs e)
            {
                if (!disposed && string.Equals(e.Result.Text, "Jarvis test", StringComparison.OrdinalIgnoreCase))
                    Interlocked.Exchange(ref heard, e.Result.Text);
            }
            private void OnCompleted(object sender, RecognizeCompletedEventArgs e)
            {
                if (disposed) return;
                if (e.Error != null) Interlocked.CompareExchange(ref error, e.Error, null);
                completed = true;
            }

            public void Dispose()
            {
                if (disposed) return;
                disposed = true;
                if (engine == null) return;
                SpeechRecognitionEngine old = engine;
                engine = null;
                old.AudioLevelUpdated -= OnLevel;
                old.SpeechRecognized -= OnRecognized;
                old.RecognizeCompleted -= OnCompleted;
                var failures = new List<Exception>();
                Exception failure;
                try
                {
                    failure = SpeechInputErrors.CaptureExpectedFailure(delegate
                    {
                        if (input != null) input.Dispose();
                    });
                    if (failure != null) failures.Add(failure);
                }
                finally
                {
                    input = null;
                    try
                    {
                        failure = SpeechInputErrors.CaptureExpectedFailure(delegate
                        {
                            try { old.RecognizeAsyncCancel(); }
                            catch (InvalidOperationException) { } // Startup may have failed before listening began.
                        });
                        if (failure != null) failures.Add(failure);
                    }
                    finally
                    {
                        failure = SpeechInputErrors.CaptureExpectedFailure(old.Dispose);
                        if (failure != null) failures.Add(failure);
                    }
                }
                if (failures.Count != 0)
                    throw new IOException("Local speech test cleanup failed.", new AggregateException(failures));
            }
        }
    }
}
