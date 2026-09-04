using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Speech.Recognition;
using System.Windows.Forms;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

[assembly: AssemblyTitle("Jarvis PowerPoint")]
[assembly: AssemblyDescription("Commande vocale locale pour avancer un diaporama PowerPoint")]
[assembly: AssemblyCompany("Jarvis PowerPoint")]
[assembly: AssemblyProduct("Jarvis PowerPoint")]
[assembly: AssemblyVersion("1.1.0.0")]

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
        private SpeechRecognitionEngine recognizer;
        private DateTime lastCommandUtc = DateTime.MinValue;
        private string currentCultureName;
        private bool acceptsCommands;
        private bool recognitionRunning;
        private bool restartAfterCancel;
        private bool exiting;

        public JarvisApplicationContext()
        {
            currentCultureName = LoadPreferredCulture();
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

            var languageMenu = new ToolStripMenuItem("Langue");
            languageMenu.DropDownItems.Add(frenchLanguageItem);
            languageMenu.DropDownItems.Add(englishLanguageItem);

            var menu = new ContextMenuStrip();
            menu.Items.Add(statusItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(toggleItem);
            menu.Items.Add(languageMenu);
            menu.Items.Add(new ToolStripMenuItem("Tester « suivant »", null, TestNextSlide));
            menu.Items.Add(new ToolStripMenuItem("Tester « précédent »", null, TestPreviousSlide));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(new ToolStripMenuItem("Quitter", null, ExitApplication));

            notifyIcon = new NotifyIcon
            {
                ContextMenuStrip = menu,
                Icon = SystemIcons.Information,
                Text = "Jarvis PowerPoint",
                Visible = true
            };
            notifyIcon.DoubleClick += ToggleListening;

            InitializeSpeechRecognition();
        }

        private void InitializeSpeechRecognition()
        {
            RecognizerInfo recognizerInfo = SpeechRecognitionEngine.InstalledRecognizers()
                .FirstOrDefault(info => info.Culture.Name.Equals(currentCultureName, StringComparison.OrdinalIgnoreCase));

            if (recognizerInfo == null)
            {
                SetUnavailable(
                    "Langue vocale absente",
                    "Installez le module de reconnaissance vocale " +
                    (currentCultureName == "en-US" ? "English (United States)" : "Français (France)") +
                    " dans les paramètres Windows.");
                UpdateLanguageMenu();
                return;
            }

            try
            {
                recognizer = new SpeechRecognitionEngine(recognizerInfo);

                var nextCommand = new GrammarBuilder { Culture = recognizerInfo.Culture };
                nextCommand.Append("Jarvis");
                nextCommand.Append(currentCultureName == "en-US" ? "next" : "suivant");

                var previousCommand = new GrammarBuilder { Culture = recognizerInfo.Culture };
                previousCommand.Append("Jarvis");
                previousCommand.Append(currentCultureName == "en-US" ? "previous" : "précédent");

                Grammar goToGrammar = CreateGoToGrammar(recognizerInfo.Culture);
                recognizer.LoadGrammar(new Grammar(nextCommand) { Name = "JarvisNext" });
                recognizer.LoadGrammar(new Grammar(previousCommand) { Name = "JarvisPrevious" });
                recognizer.LoadGrammar(goToGrammar);
                recognizer.SetInputToDefaultAudioDevice();
                recognizer.SpeechRecognized += OnSpeechRecognized;
                recognizer.RecognizeCompleted += OnRecognitionCompleted;

                acceptsCommands = true;
                toggleItem.Enabled = true;
                StartRecognition();
                UpdateLanguageMenu();
                ShowNotification("Jarvis est prêt", GetCommandHelp());
            }
            catch (InvalidOperationException exception)
            {
                SetUnavailable("Microphone indisponible", exception.Message);
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

        private void ChangeLanguage(object sender, EventArgs eventArgs)
        {
            var item = sender as ToolStripMenuItem;
            string requestedCulture = item == null ? null : item.Tag as string;

            if (string.IsNullOrEmpty(requestedCulture) ||
                requestedCulture.Equals(currentCultureName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            DisposeRecognizer();
            currentCultureName = requestedCulture;
            acceptsCommands = false;
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
            if (!acceptsCommands ||
                eventArgs.Result.Confidence < MinimumConfidence ||
                DateTime.UtcNow - lastCommandUtc < CommandCooldown)
            {
                return;
            }

            lastCommandUtc = DateTime.UtcNow;

            string message;
            bool succeeded;
            if (eventArgs.Result.Grammar.Name == "JarvisGoToSlide")
            {
                int slideNumber = Convert.ToInt32(
                    eventArgs.Result.Semantics["slideNumber"].Value,
                    CultureInfo.InvariantCulture);
                succeeded = PowerPointController.TryGoToSlide(slideNumber, out message);
            }
            else
            {
                succeeded = eventArgs.Result.Grammar.Name == "JarvisPrevious"
                    ? PowerPointController.TryReturnToPreviousSlide(out message)
                    : PowerPointController.TryAdvanceSlide(out message);
            }

            if (!succeeded)
            {
                ShowNotification("Commande entendue", message);
            }
        }

        private void ToggleListening(object sender, EventArgs eventArgs)
        {
            if (recognizer == null)
            {
                return;
            }

            if (acceptsCommands)
            {
                acceptsCommands = false;
                restartAfterCancel = false;
                if (recognitionRunning)
                {
                    recognizer.RecognizeAsyncCancel();
                }
            }
            else
            {
                acceptsCommands = true;
                if (recognitionRunning)
                {
                    restartAfterCancel = true;
                }
                else
                {
                    StartRecognition();
                }
            }

            UpdateStatus();
        }

        private void OnRecognitionCompleted(object sender, RecognizeCompletedEventArgs eventArgs)
        {
            recognitionRunning = false;

            if (exiting)
            {
                return;
            }

            if (eventArgs.Error != null)
            {
                acceptsCommands = false;
                SetUnavailable("Reconnaissance vocale arrêtée", eventArgs.Error.Message);
                return;
            }

            if (restartAfterCancel && acceptsCommands)
            {
                restartAfterCancel = false;
                StartRecognition();
            }
            else
            {
                UpdateStatus();
            }
        }

        private void TestNextSlide(object sender, EventArgs eventArgs)
        {
            string message;
            if (PowerPointController.TryAdvanceSlide(out message))
            {
                ShowNotification("Test réussi", "Le diaporama est passé à la diapositive suivante.");
            }
            else
            {
                ShowNotification("Test impossible", message);
            }
        }

        private void TestPreviousSlide(object sender, EventArgs eventArgs)
        {
            string message;
            if (PowerPointController.TryReturnToPreviousSlide(out message))
            {
                ShowNotification("Test réussi", "Le diaporama est revenu à la diapositive précédente.");
            }
            else
            {
                ShowNotification("Test impossible", message);
            }
        }

        private void UpdateStatus()
        {
            if (recognizer == null)
            {
                return;
            }

            string language = currentCultureName == "en-US" ? "English" : "Français";
            statusItem.Text = (acceptsCommands ? "État : à l’écoute — " : "État : en pause — ") + language;
            toggleItem.Text = acceptsCommands ? "Mettre en pause" : "Reprendre l’écoute";
            notifyIcon.Text = acceptsCommands ? "Jarvis PowerPoint — à l’écoute" : "Jarvis PowerPoint — en pause";
            notifyIcon.Icon = acceptsCommands ? SystemIcons.Information : SystemIcons.Warning;
        }

        private void UpdateLanguageMenu()
        {
            frenchLanguageItem.Checked = currentCultureName == "fr-FR";
            englishLanguageItem.Checked = currentCultureName == "en-US";
        }

        private string GetCommandHelp()
        {
            return currentCultureName == "en-US"
                ? "Say “Jarvis, next!”, “previous!” or “go to slide X!” during your slide show."
                : "Dites « Jarvis, suivant ! », « précédent ! » ou « va au slide X ! ».";
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
            statusItem.Text = "État : " + status;
            toggleItem.Enabled = false;
            notifyIcon.Text = "Jarvis PowerPoint — indisponible";
            notifyIcon.Icon = SystemIcons.Error;
            ShowNotification(status, details);
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
            notifyIcon.Visible = false;
            DisposeRecognizer();

            notifyIcon.Dispose();
            ExitThread();
        }

        private void DisposeRecognizer()
        {
            if (recognizer == null)
            {
                return;
            }

            recognizer.SpeechRecognized -= OnSpeechRecognized;
            recognizer.RecognizeCompleted -= OnRecognitionCompleted;
            if (recognitionRunning)
            {
                recognizer.RecognizeAsyncCancel();
            }

            recognizer.Dispose();
            recognizer = null;
            recognitionRunning = false;
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
    }
}
