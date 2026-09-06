using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Speech.Recognition;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

[assembly: AssemblyTitle("Jarvis PowerPoint")]
[assembly: AssemblyDescription("Commande vocale locale pour avancer un diaporama PowerPoint")]
[assembly: AssemblyCompany("Jarvis PowerPoint")]
[assembly: AssemblyProduct("Jarvis PowerPoint")]
[assembly: AssemblyVersion("1.3.0.0")]

namespace JarvisPowerPoint
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length == 1 &&
                (args[0].Equals("--next", StringComparison.OrdinalIgnoreCase) ||
                 args[0].Equals("--previous", StringComparison.OrdinalIgnoreCase)))
            {
                string message;
                bool succeeded = args[0].Equals("--next", StringComparison.OrdinalIgnoreCase)
                    ? PowerPointController.TryAdvanceSlide(out message)
                    : PowerPointController.TryReturnToPreviousSlide(out message);
                Console.WriteLine(message);
                Environment.ExitCode = succeeded ? 0 : 1;
                return;
            }

            if (args.Length == 2 && args[0].Equals("--goto", StringComparison.OrdinalIgnoreCase))
            {
                int slideNumber;
                string message = "Numéro de diapositive invalide.";
                bool succeeded = int.TryParse(args[1], out slideNumber) &&
                    PowerPointController.TryGoToSlide(slideNumber, out message);
                Console.WriteLine(message);
                Environment.ExitCode = succeeded ? 0 : 1;
                return;
            }

            if (args.Length >= 2 && args[0].Equals("--search", StringComparison.OrdinalIgnoreCase))
            {
                string message;
                bool succeeded = PowerPointController.TryFindAndGoToSlide(
                    string.Join(" ", args.Skip(1).ToArray()),
                    false,
                    out message);
                Console.WriteLine(message);
                Environment.ExitCode = succeeded ? 0 : 1;
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new JarvisApplicationContext());
        }
    }

    internal sealed class JarvisApplicationContext : ApplicationContext
    {
        private const float MinimumConfidence = 0.70f;
        private static readonly TimeSpan CommandCooldown = TimeSpan.FromMilliseconds(1200);

        private readonly NotifyIcon notifyIcon;
        private readonly ToolStripMenuItem statusItem;
        private readonly ToolStripMenuItem toggleItem;
        private readonly ToolStripMenuItem frenchLanguageItem;
        private readonly ToolStripMenuItem englishLanguageItem;
        private readonly ToolStripMenuItem shortcutItem;
        private readonly ToolStripMenuItem toolsItem;
        private readonly ToolStripMenuItem setupItem;
        private readonly Control dispatcher = new Control();
        private readonly Stopwatch clock = Stopwatch.StartNew();
        private readonly Timer presentationTimer = new Timer { Interval = 500 };
        private readonly RehearsalTracker rehearsal = new RehearsalTracker();
        private readonly PresenterPanel panel;
        private readonly PresentationSession session;
        private SpeechRecognitionEngine recognizer;
        private IDisposable microphoneInput;
        private SetupDialog activeSetup;
        private DateTime lastCommandUtc = DateTime.MinValue;
        private string currentCultureName;
        private string microphoneId = "default";
        private string lastHeard = "";
        private string lastOutcome = "";
        private string listeningStatus = "";
        private string rehearsalHint = "";
        private string displayedSessionKey;
        private int audioLevel;
        private bool listeningRequested = true;
        private bool acceptsCommands;
        private bool recognitionRunning;
        private bool exiting;
        private bool English { get { return currentCultureName == "en-US"; } }

        public JarvisApplicationContext()
        {
            currentCultureName = LoadPreferredCulture();
            IntPtr handle = dispatcher.Handle;
            session = new PresentationSession(English, Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "JarvisPowerPoint"));
            panel = new PresenterPanel(English);
            WirePresenterPanel();
            statusItem = new ToolStripMenuItem("Initialisation...") { Enabled = false };
            toggleItem = new ToolStripMenuItem("Mettre en pause", null, ToggleListening);
            frenchLanguageItem = new ToolStripMenuItem("Français", null, ChangeLanguage)
            {
                Tag = "fr-FR"
            };
            englishLanguageItem = new ToolStripMenuItem("English", null, ChangeLanguage)
            {
                Tag = "en-US"
            };
            shortcutItem = new ToolStripMenuItem(
                currentCultureName == "en-US"
                    ? "Create / update Start menu shortcut..."
                    : "Créer / actualiser le raccourci Démarrer...",
                null,
                delegate { ConfigureStartMenuShortcut(true); });
            toolsItem = new ToolStripMenuItem("Outils présentateur...", null, delegate { ShowPresenterPanel(); });
            setupItem = new ToolStripMenuItem("Langue et microphone...", null, delegate { OpenSetup(false); });

            var languageMenu = new ToolStripMenuItem("Langue");
            languageMenu.DropDownItems.Add(frenchLanguageItem);
            languageMenu.DropDownItems.Add(englishLanguageItem);

            var menu = new ContextMenuStrip();
            menu.Items.Add(statusItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(toggleItem);
            menu.Items.Add(toolsItem);
            menu.Items.Add(setupItem);
            menu.Items.Add(languageMenu);
            menu.Items.Add(new ToolStripMenuItem("Tester « suivant »", null, TestNextSlide));
            menu.Items.Add(new ToolStripMenuItem("Tester « précédent »", null, TestPreviousSlide));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(shortcutItem);
            menu.Items.Add(new ToolStripMenuItem("Quitter", null, ExitApplication));

            notifyIcon = new NotifyIcon
            {
                ContextMenuStrip = menu,
                Icon = SystemIcons.Information,
                Text = "Jarvis PowerPoint",
                Visible = true
            };
            notifyIcon.DoubleClick += ToggleListening;

            presentationTimer.Tick += delegate { PollPresentation(); };
            presentationTimer.Start();
            Application.Idle += OnFirstIdle;
        }

        private void OnFirstIdle(object sender, EventArgs args)
        {
            Application.Idle -= OnFirstIdle;
            RunUiAction(delegate
            {
                object storedMicrophone = Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\JarvisPowerPoint", "MicrophoneId", "default");
                if (!(storedMicrophone is string))
                {
                    throw new InvalidDataException(English ? "Invalid microphone preference. Open microphone setup." : "Préférence de microphone invalide. Ouvrez les réglages.");
                }
                microphoneId = (string)storedMicrophone;
                object completed = Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\JarvisPowerPoint", "SetupCompleted", 0);
                if (!(completed is int) || (int)completed != 1) { OpenSetup(true); }
                if (exiting) { return; }
                ConfigureStartMenuShortcut(false);
                InitializeSpeechRecognition();
            });
        }

        private void OpenSetup(bool firstRun)
        {
            if (activeSetup != null) { activeSetup.Activate(); return; }
            DisposeRecognizer();
            UpdateStatus();
            RunUiAction(delegate
            {
                using (var dialog = new SetupDialog(currentCultureName, microphoneId))
                {
                    activeSetup = dialog;
                    try
                    {
                        if (dialog.ShowDialog() == DialogResult.OK && !exiting)
                        {
                            currentCultureName = dialog.SelectedCulture;
                            microphoneId = dialog.SelectedMicrophoneId;
                            Registry.SetValue(@"HKEY_CURRENT_USER\Software\JarvisPowerPoint", "Language", currentCultureName);
                            Registry.SetValue(@"HKEY_CURRENT_USER\Software\JarvisPowerPoint", "MicrophoneId", microphoneId);
                            Registry.SetValue(@"HKEY_CURRENT_USER\Software\JarvisPowerPoint", "SetupCompleted", 1, RegistryValueKind.DWord);
                        }
                    }
                    finally { activeSetup = null; }
                }
            });
            if (exiting) { return; }
            session.English = English;
            UpdateLanguageMenu();
            if (!exiting && !firstRun) { InitializeSpeechRecognition(); }
        }

        private void WirePresenterPanel()
        {
            panel.PauseRequested += ToggleListening;
            panel.SettingsRequested += delegate { OpenSetup(false); };
            panel.QuestionsRequested += delegate { ExecuteCommand(session.BeginQuestions); };
            panel.ResumeRequested += delegate { ExecuteCommand(session.Resume); };
            panel.NextResultRequested += delegate { ExecuteCommand(session.NextResult); };
            panel.BlackRequested += delegate { ExecuteCommand(delegate { return session.SetBlackScreen(true); }); };
            panel.DisplayRequested += delegate { ExecuteCommand(delegate { return session.SetBlackScreen(false); }); };
            panel.StartRehearsalRequested += delegate { ExecuteCommand(StartRehearsal); };
            panel.StopRehearsalRequested += delegate { ExecuteCommand(StopRehearsal); };
            panel.ExportRehearsalRequested += delegate
            {
                RunUiAction(delegate
                {
                    if (rehearsal.Entries.Count == 0) { throw new InvalidOperationException(English ? "No rehearsal to export." : "Aucune répétition à exporter."); }
                    using (var dialog = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "Jarvis-rehearsal.csv", AddExtension = true })
                    {
                        if (dialog.ShowDialog(panel) == DialogResult.OK)
                        {
                            rehearsal.Export(dialog.FileName);
                            lastOutcome = English ? "Report exported locally." : "Rapport exporté localement.";
                        }
                    }
                });
            };
            panel.BudgetRequested += delegate
            {
                RunUiAction(delegate
                {
                    if (panel.SelectedRehearsalSlideId == 0) { throw new InvalidOperationException(English ? "Select a recorded slide." : "Sélectionnez une diapositive chronométrée."); }
                    rehearsal.SetBudget(panel.SelectedRehearsalSlideId, panel.BudgetSeconds);
                    lastOutcome = English ? "Slide budget updated." : "Budget du slide actualisé.";
                    panel.SetRehearsal(rehearsal);
                });
            };
            panel.RefreshAliasesRequested += delegate { RefreshAliases(); };
            panel.SaveAliasRequested += delegate
            {
                RunUiAction(delegate
                {
                    if (string.IsNullOrWhiteSpace(panel.AliasNameInput)) { throw new InvalidOperationException(English ? "Enter an alias." : "Saisissez un alias."); }
                    session.SaveAlias(panel.AliasNameInput);
                    lastOutcome = English ? "Alias saved for this presentation." : "Alias enregistré pour cette présentation.";
                    panel.SetAliases(session.GetAliases());
                });
            };
            panel.DeleteAliasRequested += delegate
            {
                RunUiAction(delegate
                {
                    string alias = RequireSelectedAlias();
                    string key = session.Snapshot().PresentationKey;
                    if (MessageBox.Show(panel, (English ? "Delete alias: " : "Supprimer l’alias : ") + alias + " ?",
                        "Jarvis", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) { return; }
                    if (session.Snapshot().PresentationKey != key)
                    {
                        throw new InvalidOperationException(English ? "The presentation changed. Refresh the aliases." : "La présentation a changé. Actualisez les alias.");
                    }
                    session.RemoveAlias(alias);
                    lastOutcome = English ? "Alias deleted." : "Alias supprimé.";
                    panel.SetAliases(session.GetAliases());
                });
            };
            panel.NavigateAliasRequested += delegate { ExecuteCommand(delegate { return session.GoToAlias(RequireSelectedAlias()); }); };
        }

        private string RequireSelectedAlias()
        {
            if (string.IsNullOrWhiteSpace(panel.SelectedAliasName))
            {
                throw new InvalidOperationException(English ? "Select an alias first." : "Sélectionnez un alias.");
            }
            return panel.SelectedAliasName;
        }

        private void ShowPresenterPanel()
        {
            panel.Show();
            panel.Activate();
            UpdateStatus();
            PollPresentation();
        }

        private void RefreshAliases()
        {
            // Never leave another deck's aliases displayed after switching presentations.
            panel.SetAliases(new List<SlideAlias>());
            RunUiAction(delegate { panel.SetAliases(session.GetAliases()); });
        }

        private string StartRehearsal()
        {
            if (rehearsal.IsRunning) { throw new InvalidOperationException(English ? "Rehearsal is already running." : "La répétition est déjà en cours."); }
            rehearsal.DefaultBudgetSeconds = panel.BudgetSeconds;
            rehearsal.Start(session.Snapshot(), clock.Elapsed);
            panel.SetRehearsal(rehearsal);
            return English ? "Rehearsal started. No audio is recorded." : "Répétition démarrée. Aucun audio enregistré.";
        }

        private string StopRehearsal()
        {
            if (!rehearsal.IsRunning) { throw new InvalidOperationException(English ? "No rehearsal is running." : "Aucune répétition en cours."); }
            rehearsal.Stop(clock.Elapsed);
            panel.SetRehearsal(rehearsal);
            return English ? "Rehearsal stopped; report is in Presenter tools." : "Répétition arrêtée ; bilan dans les outils présentateur.";
        }

        private void PollPresentation()
        {
            if (exiting) { return; }
            Exception microphoneError = SpeechInput.GetError(microphoneInput);
            if (microphoneError != null)
            {
                string message = microphoneError.Message;
                try { DisposeRecognizer(); }
                catch (IOException exception) { message += "\n" + exception.Message; }
                SetUnavailable(English ? "Microphone stopped" : "Microphone arrêté", message);
            }
            if (microphoneInput != null)
            {
                audioLevel = SpeechInput.GetLevel(microphoneInput);
                UpdateStatus();
            }
            if (!panel.Visible && !rehearsal.IsRunning && !session.HasReturnPoint) { return; }
            try
            {
                PresentationSnapshot snapshot = session.Snapshot();
                if (snapshot.SessionKey != displayedSessionKey)
                {
                    displayedSessionKey = snapshot.SessionKey;
                    panel.SetAliases(new List<SlideAlias>());
                    if (panel.AliasesSelected) { RefreshAliases(); }
                }
                rehearsal.Observe(snapshot, clock.Elapsed);
                RehearsalEntry current = rehearsal.Entries.FirstOrDefault(item => item.SlideId == snapshot.SlideId);
                rehearsalHint = rehearsal.IsRunning && current != null && current.OverBudget
                    ? (English ? "Over slide budget" : "Budget du slide dépassé") : "";
                panel.SetPresentation(snapshot.PresentationName + " · " + snapshot.SlideNumber + "/" + snapshot.SlideCount
                    + (session.QuestionsMode ? (English ? " · Questions mode" : " · Mode questions") : ""));
            }
            catch (PresentationUnavailableException exception)
            {
                displayedSessionKey = null;
                panel.SetAliases(new List<SlideAlias>());
                rehearsal.Observe(null, clock.Elapsed);
                rehearsalHint = "";
                panel.SetPresentation(exception.Message);
            }
            catch (COMException exception)
            {
                displayedSessionKey = null;
                panel.SetAliases(new List<SlideAlias>());
                rehearsal.Observe(null, clock.Elapsed);
                rehearsalHint = "";
                panel.SetPresentation(exception.Message);
            }
            catch (InvalidOperationException exception)
            {
                displayedSessionKey = null;
                panel.SetAliases(new List<SlideAlias>());
                rehearsal.Observe(null, clock.Elapsed);
                rehearsalHint = "";
                panel.SetPresentation(exception.Message);
            }
            catch (System.ComponentModel.Win32Exception exception)
            {
                rehearsal.Observe(null, clock.Elapsed);
                rehearsalHint = "";
                panel.SetPresentation(exception.Message);
            }
            panel.SetRehearsal(rehearsal);
            UpdateStatus();
        }

        private void ExecuteCommand(Func<string> command)
        {
            RunUiAction(delegate
            {
                if (rehearsal.IsRunning) { PollPresentation(); }
                lastOutcome = command();
                PollPresentation();
            });
        }

        private void RunUiAction(Action action)
        {
            bool failed = true;
            try { action(); failed = false; }
            catch (COMException exception) { lastOutcome = exception.Message; }
            catch (InvalidOperationException exception) { lastOutcome = exception.Message; }
            catch (IOException exception) { lastOutcome = exception.Message; }
            catch (UnauthorizedAccessException exception) { lastOutcome = exception.Message; }
            catch (SecurityException exception) { lastOutcome = exception.Message; }
            catch (ArgumentException exception) { lastOutcome = exception.Message; }
            catch (System.Xml.XmlException exception) { lastOutcome = exception.Message; }
            catch (System.ComponentModel.Win32Exception exception) { lastOutcome = exception.Message; }
            if (exiting) { return; }
            UpdateStatus();
            if (failed && !panel.Visible)
            {
                ShowNotification(English ? "Action not completed" : "Action non exécutée",
                    English ? "Open Presenter tools from the Jarvis tray menu for details."
                        : "Ouvrez les outils présentateur depuis l’icône Jarvis pour voir les détails.");
            }
        }

        private void ConfigureStartMenuShortcut(bool fromMenu)
        {
            bool english = currentCultureName == "en-US";
            try
            {
                var shortcut = new StartMenuShortcut(
                    Environment.GetFolderPath(Environment.SpecialFolder.Programs),
                    Application.ExecutablePath,
                    @"HKEY_CURRENT_USER\Software\JarvisPowerPoint");

                ShortcutResult result = shortcut.Configure(
                    delegate
                    {
                        string question = english
                            ? "Create or update a Start menu shortcut for Jarvis PowerPoint?"
                                + "\n\nYou can then find it by pressing Windows and typing Jarvis PowerPoint."
                                + "\nNo administrator rights are needed. This does not start Jarvis with Windows."
                                + "\n\nKeep the executable in this location:"
                            : "Créer ou actualiser un raccourci Jarvis PowerPoint dans le menu Démarrer ?"
                                + "\n\nVous pourrez le retrouver avec la touche Windows en tapant Jarvis PowerPoint."
                                + "\nAucun droit administrateur requis. Jarvis ne démarrera pas avec Windows."
                                + "\n\nConservez l’exécutable à cet emplacement :";
                        question += "\n" + Application.ExecutablePath;
                        question += english
                            ? "\n\nYou can do this later from the Jarvis tray menu."
                            : "\n\nVous pourrez le faire plus tard depuis le menu de l’icône Jarvis.";
                        return MessageBox.Show(
                            question, "Jarvis PowerPoint", MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
                    },
                    fromMenu);

                if (result == ShortcutResult.Created)
                {
                    MessageBox.Show(
                        english
                            ? "Shortcut saved. Press Windows and search for Jarvis PowerPoint."
                            : "Raccourci enregistré. Appuyez sur Windows et cherchez Jarvis PowerPoint.",
                        "Jarvis PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (IOException exception)
            {
                ShowShortcutError(english, exception.Message);
            }
            catch (UnauthorizedAccessException exception)
            {
                ShowShortcutError(english, exception.Message);
            }
            catch (SecurityException exception)
            {
                ShowShortcutError(english, exception.Message);
            }
            catch (COMException exception)
            {
                ShowShortcutError(english, exception.Message);
            }
        }

        private static void ShowShortcutError(bool english, string details)
        {
            MessageBox.Show(
                (english
                    ? "Could not complete Start menu setup. Jarvis will continue to run."
                        + "\nRetry from the Jarvis tray menu.\n\n"
                    : "La configuration du raccourci n’a pas pu être terminée. Jarvis reste utilisable."
                        + "\nRéessayez depuis le menu de l’icône Jarvis.\n\n") + details,
                "Jarvis PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void InitializeSpeechRecognition()
        {
            DisposeRecognizer();
            if (!listeningRequested) { UpdateStatus(); return; }
            try
            {
                RecognizerInfo recognizerInfo = SpeechRecognitionEngine.InstalledRecognizers()
                    .FirstOrDefault(info => info.Culture.Name.Equals(currentCultureName, StringComparison.OrdinalIgnoreCase));
                if (recognizerInfo == null)
                {
                    SetUnavailable(
                        English ? "Speech language missing" : "Langue vocale absente",
                        English ? "Install English (United States) speech recognition in Windows settings."
                            : "Installez la reconnaissance vocale Français (France) dans les paramètres Windows.");
                    UpdateLanguageMenu();
                    return;
                }
                recognizer = new SpeechRecognitionEngine(recognizerInfo);

                var nextCommand = new GrammarBuilder { Culture = recognizerInfo.Culture };
                nextCommand.Append("Jarvis");
                nextCommand.Append(currentCultureName == "en-US" ? "next" : "suivant");

                var previousCommand = new GrammarBuilder { Culture = recognizerInfo.Culture };
                previousCommand.Append("Jarvis");
                previousCommand.Append(currentCultureName == "en-US" ? "previous" : "précédent");

                Grammar goToGrammar = CreateGoToGrammar(recognizerInfo.Culture);
                Grammar searchGrammar = CreateSearchGrammar(recognizerInfo.Culture);
                recognizer.LoadGrammar(new Grammar(nextCommand) { Name = "JarvisNext" });
                recognizer.LoadGrammar(new Grammar(previousCommand) { Name = "JarvisPrevious" });
                recognizer.LoadGrammar(goToGrammar);
                recognizer.LoadGrammar(searchGrammar);
                recognizer.LoadGrammar(CreateActionGrammar(recognizerInfo.Culture));
                recognizer.LoadGrammar(CreateAliasGrammar(recognizerInfo.Culture));
                microphoneInput = SpeechInput.Attach(recognizer, microphoneId);
                recognizer.SpeechRecognized += OnSpeechRecognized;
                recognizer.RecognizeCompleted += OnRecognitionCompleted;
                recognizer.AudioLevelUpdated += OnAudioLevelUpdated;

                toggleItem.Enabled = true;
                StartRecognition();
                UpdateLanguageMenu();
            }
            catch (InvalidOperationException exception)
            {
                DisposeRecognizer();
                SetUnavailable("Microphone indisponible", exception.Message);
            }
            catch (IOException exception)
            {
                DisposeRecognizer();
                SetUnavailable("Microphone indisponible", exception.Message);
            }
            catch (COMException exception)
            {
                DisposeRecognizer();
                SetUnavailable("Microphone indisponible", exception.Message);
            }
            catch (ArgumentException exception)
            {
                DisposeRecognizer();
                SetUnavailable("Reconnaissance indisponible", exception.Message);
            }
            catch (UnauthorizedAccessException exception)
            {
                DisposeRecognizer();
                SetUnavailable(English ? "Microphone access denied" : "Accès microphone refusé", exception.Message);
            }
            catch (SecurityException exception)
            {
                DisposeRecognizer();
                SetUnavailable(English ? "Microphone access denied" : "Accès microphone refusé", exception.Message);
            }
        }

        private Grammar CreateGoToGrammar(CultureInfo culture)
        {
            var command = new GrammarBuilder { Culture = culture };
            command.Append("Jarvis");

            if (culture.Name == "en-US")
            {
                command.Append("go to slide");
            }
            else
            {
                command.Append(new Choices("va au slide", "vas au slide"));
            }

            var slideNumbers = new List<GrammarBuilder>();
            for (int number = 1; number <= 999; number++)
            {
                string spokenNumber = culture.Name == "en-US"
                    ? NumberWords.ToEnglish(number)
                    : NumberWords.ToFrench(number);
                slideNumbers.Add(new SemanticResultValue(spokenNumber, number));
            }

            command.Append(
                new SemanticResultKey(
                    "slideNumber",
                    new Choices(slideNumbers.ToArray())));

            return new Grammar(command) { Name = "JarvisGoToSlide" };
        }

        private Grammar CreateSearchGrammar(CultureInfo culture)
        {
            var command = new GrammarBuilder { Culture = culture };
            command.Append("Jarvis");
            command.Append(
                culture.Name == "en-US"
                    ? new Choices("search for", "find", "go to the slide about")
                    : new Choices("cherche", "trouve", "va au slide sur", "vas au slide sur"));
            command.AppendDictation();
            return new Grammar(command) { Name = "JarvisSearch" };
        }

        private Grammar CreateAliasGrammar(CultureInfo culture)
        {
            var command = new GrammarBuilder { Culture = culture };
            command.Append("Jarvis");
            command.Append(culture.Name == "en-US" ? "shortcut" : "raccourci");
            command.AppendDictation();
            return new Grammar(command) { Name = "JarvisAlias" };
        }

        private Grammar CreateActionGrammar(CultureInfo culture)
        {
            bool english = culture.Name == "en-US";
            var alternatives = new List<GrammarBuilder>();
            AddActions(alternatives, "resume", english
                ? new[] { "resume", "resume the presentation", "go back to where I was" }
                : new[] { "reprends", "reprends la présentation", "reviens au point de départ", "reviens où j'étais" });
            AddActions(alternatives, "results", english ? new[] { "next match", "another result" } : new[] { "autre résultat", "résultat suivant" });
            AddActions(alternatives, "questions", english ? new[] { "questions mode" } : new[] { "mode questions" });
            AddActions(alternatives, "black", english ? new[] { "black screen" } : new[] { "écran noir" });
            AddActions(alternatives, "display", english ? new[] { "restore slides", "show slides" } : new[] { "affiche", "affiche les slides" });
            AddActions(alternatives, "rehearse", english ? new[] { "start rehearsal" } : new[] { "démarre la répétition" });
            AddActions(alternatives, "stopRehearsal", english ? new[] { "stop rehearsal" } : new[] { "arrête la répétition" });
            AddActions(alternatives, "test", new[] { "test" });
            var command = new GrammarBuilder { Culture = culture };
            command.Append("Jarvis");
            command.Append(new SemanticResultKey("action", new Choices(alternatives.ToArray())));
            return new Grammar(command) { Name = "JarvisActions" };
        }

        private static void AddActions(List<GrammarBuilder> alternatives, string action, string[] phrases)
        {
            foreach (string phrase in phrases) { alternatives.Add(new SemanticResultValue(phrase, action)); }
        }

        private void ChangeLanguage(object sender, EventArgs eventArgs)
        {
            var item = sender as ToolStripMenuItem;
            string requestedCulture = item == null ? null : item.Tag as string;

            if (string.IsNullOrEmpty(requestedCulture) ||
                requestedCulture.Equals(currentCultureName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            currentCultureName = requestedCulture;
            session.English = English;
            UpdateLanguageMenu();
            InitializeSpeechRecognition();
            SavePreferredCulture();
        }

        private void StartRecognition()
        {
            if (recognizer == null || recognitionRunning || exiting)
            {
                return;
            }

            try
            {
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
                recognitionRunning = true;
                acceptsCommands = true;
                listeningStatus = "";
                UpdateStatus();
            }
            catch (InvalidOperationException exception)
            {
                acceptsCommands = false;
                SetUnavailable("Écoute impossible", exception.Message);
            }
        }

        private void OnSpeechRecognized(object sender, SpeechRecognizedEventArgs eventArgs)
        {
            OnUi(delegate
            {
                if (sender != recognizer || !acceptsCommands) { return; }
                lastHeard = eventArgs.Result.Text;
                if (eventArgs.Result.Confidence < MinimumConfidence)
                {
                    lastOutcome = English ? "Not executed: please repeat." : "Non exécuté : veuillez répéter.";
                    UpdateStatus();
                    return;
                }
                if (DateTime.UtcNow - lastCommandUtc < CommandCooldown) { return; }
                lastCommandUtc = DateTime.UtcNow;
                ExecuteCommand(delegate { return DispatchRecognition(eventArgs.Result); });
            });
        }

        private string DispatchRecognition(RecognitionResult result)
        {
            switch (result.Grammar.Name)
            {
                case "JarvisNext": return session.Next();
                case "JarvisPrevious": return session.Previous();
                case "JarvisGoToSlide":
                    return session.GoTo(Convert.ToInt32(result.Semantics["slideNumber"].Value, CultureInfo.InvariantCulture));
                case "JarvisSearch": return session.Search(ExtractSearchQuery(result.Text));
                case "JarvisAlias":
                    string prefix = English ? "Jarvis shortcut " : "Jarvis raccourci ";
                    if (!result.Text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException(English ? "Alias not understood." : "Alias non compris.");
                    }
                    return session.GoToAlias(result.Text.Substring(prefix.Length).Trim());
                case "JarvisActions":
                    switch ((string)result.Semantics["action"].Value)
                    {
                        case "resume": return session.Resume();
                        case "results": return session.NextResult();
                        case "questions": return session.BeginQuestions();
                        case "black": return session.SetBlackScreen(true);
                        case "display": return session.SetBlackScreen(false);
                        case "rehearse": return StartRehearsal();
                        case "stopRehearsal": return StopRehearsal();
                        case "test": return English ? "Voice command understood; no slide changed." : "Commande comprise ; aucune diapositive modifiée.";
                    }
                    break;
            }
            throw new InvalidOperationException(English ? "Unknown voice command." : "Commande vocale inconnue.");
        }

        private void OnUi(Action action)
        {
            if (exiting || dispatcher.IsDisposed) { return; }
            if (dispatcher.InvokeRequired)
            {
                try { dispatcher.BeginInvoke(action); }
                catch (InvalidOperationException)
                {
                    if (!exiting && !dispatcher.IsDisposed) { throw; }
                }
            }
            else { action(); }
        }

        private void OnAudioLevelUpdated(object sender, AudioLevelUpdatedEventArgs args)
        {
            OnUi(delegate
            {
                if (exiting || sender != recognizer) { return; }
                audioLevel = args.AudioLevel;
                UpdateStatus();
            });
        }

        private string ExtractSearchQuery(string recognizedText)
        {
            string[] prefixes = currentCultureName == "en-US"
                ? new[] { "Jarvis go to the slide about ", "Jarvis search for ", "Jarvis find " }
                : new[]
                {
                    "Jarvis va au slide sur ",
                    "Jarvis vas au slide sur ",
                    "Jarvis cherche ",
                    "Jarvis trouve "
                };

            string prefix = prefixes.FirstOrDefault(
                value => recognizedText.StartsWith(value, StringComparison.CurrentCultureIgnoreCase));
            return prefix == null ? string.Empty : recognizedText.Substring(prefix.Length).Trim();
        }

        private void ToggleListening(object sender, EventArgs eventArgs)
        {
            RunUiAction(delegate
            {
                listeningRequested = !listeningRequested;
                if (listeningRequested) { InitializeSpeechRecognition(); }
                else { DisposeRecognizer(); listeningStatus = ""; }
            });
        }

        private void OnRecognitionCompleted(object sender, RecognizeCompletedEventArgs eventArgs)
        {
            OnUi(delegate
            {
                if (exiting || sender != recognizer) { return; }
                recognitionRunning = false;
                acceptsCommands = false;
                DisposeRecognizer();
                SetUnavailable(English ? "Recognition stopped" : "Reconnaissance arrêtée",
                    eventArgs.Error == null
                        ? (English ? "Restart listening or check microphone setup." : "Relancez l’écoute ou vérifiez le microphone.")
                        : eventArgs.Error.Message);
            });
        }

        private void TestNextSlide(object sender, EventArgs eventArgs)
        {
            ExecuteCommand(session.Next);
        }

        private void TestPreviousSlide(object sender, EventArgs eventArgs)
        {
            ExecuteCommand(session.Previous);
        }

        private void UpdateStatus()
        {
            if (exiting) { return; }
            string language = currentCultureName == "en-US" ? "English" : "Français";
            string state = acceptsCommands ? (English ? "Listening" : "À l’écoute")
                : listeningRequested ? (English ? "Unavailable" : "Indisponible") : (English ? "Paused" : "En pause");
            statusItem.Text = state + " · " + language;
            toggleItem.Text = listeningRequested ? (English ? "Pause listening" : "Mettre en pause") : (English ? "Resume listening" : "Reprendre l’écoute");
            notifyIcon.Text = "Jarvis PowerPoint · " + state;
            notifyIcon.Icon = acceptsCommands ? SystemIcons.Information : SystemIcons.Warning;
            panel.SetStatus(state + " · " + language + (listeningStatus.Length == 0 ? "" : " · " + listeningStatus),
                lastHeard, rehearsalHint.Length == 0 ? lastOutcome : rehearsalHint + " · " + lastOutcome, audioLevel);
        }

        private void UpdateLanguageMenu()
        {
            frenchLanguageItem.Checked = currentCultureName == "fr-FR";
            englishLanguageItem.Checked = currentCultureName == "en-US";
            shortcutItem.Text = currentCultureName == "en-US"
                ? "Create / update Start menu shortcut..."
                : "Créer / actualiser le raccourci Démarrer...";
            toolsItem.Text = English ? "Presenter tools..." : "Outils présentateur...";
            setupItem.Text = English ? "Language and microphone..." : "Langue et microphone...";
            panel.SetLanguage(English);
        }

        private string GetCommandHelp()
        {
            return currentCultureName == "en-US"
                ? "Say “Jarvis, next!”, “go to slide X!” or “search for X!”."
                : "Dites « Jarvis, suivant ! », « va au slide X ! » ou « cherche X ! ».";
        }

        private static string LoadPreferredCulture()
        {
            try
            {
                object value = Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\JarvisPowerPoint",
                    "Language",
                    "fr-FR");
                string culture = value as string;
                return culture == "en-US" ? "en-US" : "fr-FR";
            }
            catch (SecurityException)
            {
                return "fr-FR";
            }
            catch (UnauthorizedAccessException)
            {
                return "fr-FR";
            }
        }

        private void SavePreferredCulture()
        {
            try
            {
                Registry.SetValue(
                    @"HKEY_CURRENT_USER\Software\JarvisPowerPoint",
                    "Language",
                    currentCultureName,
                    RegistryValueKind.String);
            }
            catch (SecurityException exception)
            {
                ShowNotification("Préférence non enregistrée", exception.Message);
            }
            catch (UnauthorizedAccessException exception)
            {
                ShowNotification("Préférence non enregistrée", exception.Message);
            }
        }

        private void SetUnavailable(string status, string details)
        {
            acceptsCommands = false;
            listeningStatus = status;
            lastOutcome = details;
            toggleItem.Enabled = true;
            UpdateStatus();
            ShowNotification(English ? "Check microphone setup" : "Vérifiez les réglages du microphone",
                English ? "Open Language and microphone from the Jarvis tray menu."
                    : "Ouvrez Langue et microphone depuis le menu de l’icône Jarvis.");
        }

        private void ShowNotification(string title, string message)
        {
            notifyIcon.BalloonTipTitle = title;
            notifyIcon.BalloonTipText = message;
            notifyIcon.ShowBalloonTip(3500);
        }

        private void ExitApplication(object sender, EventArgs eventArgs)
        {
            exiting = true;
            if (activeSetup != null) { activeSetup.Close(); }
            Application.Idle -= OnFirstIdle;
            presentationTimer.Stop();
            presentationTimer.Dispose();
            notifyIcon.Visible = false;
            DisposeRecognizer();
            panel.Dispose();
            dispatcher.Dispose();
            notifyIcon.Dispose();
            ExitThread();
        }

        private void DisposeRecognizer()
        {
            acceptsCommands = false;
            audioLevel = 0;
            SpeechRecognitionEngine oldRecognizer = recognizer;
            recognizer = null;
            if (oldRecognizer == null)
            {
                return;
            }

            oldRecognizer.SpeechRecognized -= OnSpeechRecognized;
            oldRecognizer.RecognizeCompleted -= OnRecognitionCompleted;
            oldRecognizer.AudioLevelUpdated -= OnAudioLevelUpdated;
            try
            {
                IDisposable oldInput = microphoneInput;
                microphoneInput = null;
                if (oldInput != null) { oldInput.Dispose(); }
            }
            finally
            {
                oldRecognizer.Dispose();
                recognitionRunning = false;
            }
        }
    }

    internal static class NumberWords
    {
        private static readonly string[] FrenchUnits =
        {
            "zéro", "un", "deux", "trois", "quatre", "cinq", "six", "sept",
            "huit", "neuf", "dix", "onze", "douze", "treize", "quatorze",
            "quinze", "seize"
        };

        private static readonly string[] EnglishUnits =
        {
            "zero", "one", "two", "three", "four", "five", "six", "seven",
            "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen",
            "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"
        };

        private static readonly string[] EnglishTens =
        {
            "", "", "twenty", "thirty", "forty", "fifty", "sixty",
            "seventy", "eighty", "ninety"
        };

        public static string ToFrench(int number)
        {
            if (number < 1 || number > 999)
            {
                throw new ArgumentOutOfRangeException("number");
            }

            if (number < 100)
            {
                return FrenchUnderOneHundred(number);
            }

            int hundreds = number / 100;
            int remainder = number % 100;
            string prefix = hundreds == 1 ? "cent" : FrenchUnits[hundreds] + " cent";
            return remainder == 0 ? prefix : prefix + " " + FrenchUnderOneHundred(remainder);
        }

        public static string ToEnglish(int number)
        {
            if (number < 1 || number > 999)
            {
                throw new ArgumentOutOfRangeException("number");
            }

            if (number < 100)
            {
                return EnglishUnderOneHundred(number);
            }

            int hundreds = number / 100;
            int remainder = number % 100;
            string prefix = EnglishUnits[hundreds] + " hundred";
            return remainder == 0 ? prefix : prefix + " " + EnglishUnderOneHundred(remainder);
        }

        private static string FrenchUnderOneHundred(int number)
        {
            if (number <= 16)
            {
                return FrenchUnits[number];
            }

            if (number < 20)
            {
                return "dix " + FrenchUnits[number - 10];
            }

            if (number < 70)
            {
                string[] tens = { "", "", "vingt", "trente", "quarante", "cinquante", "soixante" };
                int ten = number / 10;
                int unit = number % 10;
                if (unit == 0)
                {
                    return tens[ten];
                }

                return tens[ten] + (unit == 1 ? " et " : " ") + FrenchUnits[unit];
            }

            if (number < 80)
            {
                int remainder = number - 60;
                return "soixante" +
                    (remainder == 11 ? " et " : " ") +
                    FrenchUnderOneHundred(remainder);
            }

            int lastPart = number - 80;
            return lastPart == 0
                ? "quatre vingts"
                : "quatre vingt " + FrenchUnderOneHundred(lastPart);
        }

        private static string EnglishUnderOneHundred(int number)
        {
            if (number < 20)
            {
                return EnglishUnits[number];
            }

            int ten = number / 10;
            int unit = number % 10;
            return unit == 0 ? EnglishTens[ten] : EnglishTens[ten] + " " + EnglishUnits[unit];
        }
    }

    internal static class PowerPointController
    {
        private static readonly HashSet<string> SearchStopWords = new HashSet<string>(
            new[]
            {
                "a", "about", "an", "and", "au", "aux", "de", "des", "du", "et",
                "for", "la", "le", "les", "of", "on", "sur", "the", "un", "une"
            },
            StringComparer.Ordinal);

        public static bool TryAdvanceSlide(out string message)
        {
            return TryMoveSlide(true, out message);
        }

        public static bool TryReturnToPreviousSlide(out string message)
        {
            return TryMoveSlide(false, out message);
        }

        public static bool TryGoToSlide(int slideNumber, out string message)
        {
            PowerPoint.Application application = null;
            PowerPoint.SlideShowWindows slideShowWindows = null;
            PowerPoint.SlideShowWindow slideShowWindow = null;
            PowerPoint.SlideShowView view = null;
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;

            try
            {
                application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                slideShowWindows = application.SlideShowWindows;

                if (slideShowWindows.Count == 0)
                {
                    message = "PowerPoint est ouvert, mais aucun diaporama n’est en cours.";
                    return false;
                }

                slideShowWindow = slideShowWindows[1];
                presentation = slideShowWindow.Presentation;
                slides = presentation.Slides;
                if (slideNumber < 1 || slideNumber > slides.Count)
                {
                    message = string.Format(
                        CultureInfo.CurrentCulture,
                        "Le diaporama contient {0} diapositives. Le numéro {1} est invalide.",
                        slides.Count,
                        slideNumber);
                    return false;
                }

                view = slideShowWindow.View;
                view.GotoSlide(slideNumber, Office.MsoTriState.msoFalse);
                message = string.Format(
                    CultureInfo.CurrentCulture,
                    "Diapositive {0}.",
                    slideNumber);
                return true;
            }
            catch (COMException)
            {
                message = "Ouvrez PowerPoint et démarrez le diaporama.";
                return false;
            }
            finally
            {
                ReleaseComObject(slides);
                ReleaseComObject(presentation);
                ReleaseComObject(view);
                ReleaseComObject(slideShowWindow);
                ReleaseComObject(slideShowWindows);
                ReleaseComObject(application);
            }
        }

        public static bool TryFindAndGoToSlide(string query, bool useEnglish, out string message)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                message = useEnglish ? "I did not understand what to search for." : "Je n’ai pas compris quoi chercher.";
                return false;
            }

            PowerPoint.Application application = null;
            PowerPoint.SlideShowWindows slideShowWindows = null;
            PowerPoint.SlideShowWindow slideShowWindow = null;
            PowerPoint.SlideShowView view = null;
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;

            try
            {
                application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                slideShowWindows = application.SlideShowWindows;
                if (slideShowWindows.Count == 0)
                {
                    message = useEnglish
                        ? "PowerPoint is open, but no slide show is running."
                        : "PowerPoint est ouvert, mais aucun diaporama n’est en cours.";
                    return false;
                }

                slideShowWindow = slideShowWindows[1];
                presentation = slideShowWindow.Presentation;
                slides = presentation.Slides;
                List<SlideSearchResult> results = FindSlides(slides, query);
                if (results.Count == 0)
                {
                    message = useEnglish
                        ? string.Format(CultureInfo.CurrentCulture, "No slide contains “{0}”.", query)
                        : string.Format(CultureInfo.CurrentCulture, "Aucune diapositive ne contient « {0} ».", query);
                    return false;
                }

                SlideSearchResult bestResult = results[0];
                view = slideShowWindow.View;
                view.GotoSlide(bestResult.SlideNumber, Office.MsoTriState.msoFalse);

                string title = string.IsNullOrWhiteSpace(bestResult.Title)
                    ? string.Empty
                    : " — " + bestResult.Title;
                message = useEnglish
                    ? string.Format(
                        CultureInfo.CurrentCulture,
                        "Slide {0}{1}{2}",
                        bestResult.SlideNumber,
                        title,
                        results.Count > 1
                            ? string.Format(CultureInfo.CurrentCulture, " ({0} matches)", results.Count)
                            : string.Empty)
                    : string.Format(
                        CultureInfo.CurrentCulture,
                        "Diapositive {0}{1}{2}",
                        bestResult.SlideNumber,
                        title,
                        results.Count > 1
                            ? string.Format(CultureInfo.CurrentCulture, " ({0} résultats)", results.Count)
                            : string.Empty);
                return true;
            }
            catch (COMException)
            {
                message = useEnglish
                    ? "Open PowerPoint and start the slide show."
                    : "Ouvrez PowerPoint et démarrez le diaporama.";
                return false;
            }
            finally
            {
                ReleaseComObject(slides);
                ReleaseComObject(presentation);
                ReleaseComObject(view);
                ReleaseComObject(slideShowWindow);
                ReleaseComObject(slideShowWindows);
                ReleaseComObject(application);
            }
        }

        internal static List<SlideSearchResult> FindSlides(PowerPoint.Slides slides, string query)
        {
            string normalizedQuery = NormalizeForSearch(query);
            string[] queryWords = normalizedQuery
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(word => !SearchStopWords.Contains(word))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            if (queryWords.Length == 0)
            {
                queryWords = normalizedQuery
                    .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            }
            var results = new List<SlideSearchResult>();

            for (int index = 1; index <= slides.Count; index++)
            {
                PowerPoint.Slide slide = null;
                try
                {
                    slide = slides[index];
                    string title = GetSlideTitle(slide);
                    string content = GetSlideText(slide);
                    int score = ScoreSlide(title, content, normalizedQuery, queryWords);
                    if (score > 0)
                    {
                        results.Add(new SlideSearchResult(index, title, score));
                    }
                }
                finally
                {
                    ReleaseComObject(slide);
                }
            }

            return results
                .OrderByDescending(result => result.Score)
                .ThenBy(result => result.SlideNumber)
                .ToList();
        }

        internal static string GetSlideTitle(PowerPoint.Slide slide)
        {
            PowerPoint.Shapes shapes = null;
            PowerPoint.Shape titleShape = null;
            PowerPoint.TextFrame textFrame = null;
            PowerPoint.TextRange textRange = null;

            try
            {
                shapes = slide.Shapes;
                if (shapes.HasTitle != Office.MsoTriState.msoTrue)
                {
                    return string.Empty;
                }

                titleShape = shapes.Title;
                if (titleShape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return string.Empty;
                }

                textFrame = titleShape.TextFrame;
                if (textFrame.HasText != Office.MsoTriState.msoTrue)
                {
                    return string.Empty;
                }

                textRange = textFrame.TextRange;
                string title = CleanDisplayText(textRange.Text);
                return title.Length <= 80 ? title : title.Substring(0, 77) + "...";
            }
            finally
            {
                ReleaseComObject(textRange);
                ReleaseComObject(textFrame);
                ReleaseComObject(titleShape);
                ReleaseComObject(shapes);
            }
        }

        private static string GetSlideText(PowerPoint.Slide slide)
        {
            PowerPoint.Shapes shapes = null;
            var text = new StringBuilder();

            try
            {
                shapes = slide.Shapes;
                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[index];
                        AppendShapeText(shape, text);
                    }
                    finally
                    {
                        ReleaseComObject(shape);
                    }
                }
            }
            finally
            {
                ReleaseComObject(shapes);
            }

            return text.ToString();
        }

        private static void AppendShapeText(PowerPoint.Shape shape, StringBuilder text)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                PowerPoint.GroupShapes groupItems = null;
                try
                {
                    groupItems = shape.GroupItems;
                    for (int index = 1; index <= groupItems.Count; index++)
                    {
                        PowerPoint.Shape groupShape = null;
                        try
                        {
                            groupShape = (PowerPoint.Shape)groupItems[index];
                            AppendShapeText(groupShape, text);
                        }
                        finally
                        {
                            ReleaseComObject(groupShape);
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(groupItems);
                }
            }

            if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
            {
                PowerPoint.TextFrame textFrame = null;
                PowerPoint.TextRange textRange = null;
                try
                {
                    textFrame = shape.TextFrame;
                    if (textFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        textRange = textFrame.TextRange;
                        text.Append(' ').Append(textRange.Text);
                    }
                }
                finally
                {
                    ReleaseComObject(textRange);
                    ReleaseComObject(textFrame);
                }
            }

            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                AppendTableText(shape, text);
            }

            if (!string.IsNullOrWhiteSpace(shape.AlternativeText))
            {
                text.Append(' ').Append(shape.AlternativeText);
            }
        }

        private static void AppendTableText(PowerPoint.Shape shape, StringBuilder text)
        {
            PowerPoint.Table table = null;
            PowerPoint.Rows rows = null;
            PowerPoint.Columns columns = null;

            try
            {
                table = shape.Table;
                rows = table.Rows;
                columns = table.Columns;
                for (int row = 1; row <= rows.Count; row++)
                {
                    for (int column = 1; column <= columns.Count; column++)
                    {
                        PowerPoint.Cell cell = null;
                        PowerPoint.Shape cellShape = null;
                        PowerPoint.TextFrame textFrame = null;
                        PowerPoint.TextRange textRange = null;
                        try
                        {
                            cell = table.Cell(row, column);
                            cellShape = cell.Shape;
                            textFrame = cellShape.TextFrame;
                            if (textFrame.HasText == Office.MsoTriState.msoTrue)
                            {
                                textRange = textFrame.TextRange;
                                text.Append(' ').Append(textRange.Text);
                            }
                        }
                        finally
                        {
                            ReleaseComObject(textRange);
                            ReleaseComObject(textFrame);
                            ReleaseComObject(cellShape);
                            ReleaseComObject(cell);
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(columns);
                ReleaseComObject(rows);
                ReleaseComObject(table);
            }
        }

        private static int ScoreSlide(
            string title,
            string content,
            string normalizedQuery,
            string[] queryWords)
        {
            string normalizedTitle = NormalizeForSearch(title);
            string normalizedContent = NormalizeForSearch(content);

            if (normalizedTitle == normalizedQuery)
            {
                return 1000;
            }

            if (ContainsPhrase(normalizedTitle, normalizedQuery))
            {
                return 800;
            }

            if (ContainsPhrase(normalizedContent, normalizedQuery))
            {
                return 500;
            }

            int titleMatches = queryWords.Count(word => ContainsWord(normalizedTitle, word));
            int contentMatches = queryWords.Count(word => ContainsWord(normalizedContent, word));
            if (titleMatches == 0 && contentMatches == 0)
            {
                return 0;
            }

            return (titleMatches * 100) + (contentMatches * 20);
        }

        private static bool ContainsPhrase(string text, string phrase)
        {
            return !string.IsNullOrEmpty(phrase) &&
                (" " + text + " ").Contains(" " + phrase + " ");
        }

        private static bool ContainsWord(string text, string word)
        {
            return !string.IsNullOrEmpty(word) &&
                (" " + text + " ").Contains(" " + word + " ");
        }

        internal static string NormalizeForSearch(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string decomposed = value.Normalize(NormalizationForm.FormD);
            var normalized = new StringBuilder(decomposed.Length);
            bool previousWasSpace = true;

            foreach (char character in decomposed)
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(character);
                if (category == UnicodeCategory.NonSpacingMark)
                {
                    continue;
                }

                if (char.IsLetterOrDigit(character))
                {
                    normalized.Append(char.ToLowerInvariant(character));
                    previousWasSpace = false;
                }
                else if (!previousWasSpace)
                {
                    normalized.Append(' ');
                    previousWasSpace = true;
                }
            }

            return normalized.ToString().Trim();
        }

        private static string CleanDisplayText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return string.Join(
                " ",
                value.Split((char[])null, StringSplitOptions.RemoveEmptyEntries));
        }

        private static bool TryMoveSlide(bool moveForward, out string message)
        {
            PowerPoint.Application application = null;
            PowerPoint.SlideShowWindows slideShowWindows = null;
            PowerPoint.SlideShowWindow slideShowWindow = null;
            PowerPoint.SlideShowView view = null;

            try
            {
                application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                slideShowWindows = application.SlideShowWindows;

                if (slideShowWindows.Count == 0)
                {
                    message = "PowerPoint est ouvert, mais aucun diaporama n’est en cours.";
                    return false;
                }

                slideShowWindow = slideShowWindows[1];
                view = slideShowWindow.View;
                if (moveForward)
                {
                    view.Next();
                }
                else
                {
                    view.Previous();
                }

                message = moveForward ? "Diapositive suivante." : "Diapositive précédente.";
                return true;
            }
            catch (COMException)
            {
                message = "Ouvrez PowerPoint et démarrez le diaporama.";
                return false;
            }
            finally
            {
                ReleaseComObject(view);
                ReleaseComObject(slideShowWindow);
                ReleaseComObject(slideShowWindows);
                ReleaseComObject(application);
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }

        internal sealed class SlideSearchResult
        {
            public SlideSearchResult(int slideNumber, string title, int score)
            {
                SlideNumber = slideNumber;
                Title = title;
                Score = score;
            }

            public int SlideNumber { get; private set; }

            public string Title { get; private set; }

            public int Score { get; private set; }
        }
    }
}
