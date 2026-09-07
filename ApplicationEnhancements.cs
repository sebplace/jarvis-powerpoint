using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Forms;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace JarvisPowerPoint
{
    internal sealed partial class JarvisApplicationContext
    {
        private const string SettingsKey = @"HKEY_CURRENT_USER\Software\JarvisPowerPoint";
        private readonly PresentationProfileStore profiles = new PresentationProfileStore(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "JarvisPowerPoint"));
        private readonly DiagnosticRecorder diagnostics = new DiagnosticRecorder();
        private readonly GlobalHotkeys hotkeys = new GlobalHotkeys();
        private ToolStripMenuItem preflightItem;
        private ToolStripMenuItem profilesItem;
        private ToolStripMenuItem updatesItem;
        private ToolStripMenuItem refreshSearchItem;
        private SearchProgressDialog activeSearch;
        private bool applyingPreferences;
        private int modalDepth;
        private DateTime lastAudioUtc = DateTime.MinValue;
        private string historyHandledRunId;

        private void InitializeEnhancements()
        {
            preflightItem = new ToolStripMenuItem("", null, delegate { ShowPreflight(); });
            profilesItem = new ToolStripMenuItem("", null, delegate { ShowProfiles(); });
            updatesItem = new ToolStripMenuItem("", null, delegate { ShowUpdates(); });
            refreshSearchItem = new ToolStripMenuItem("", null, delegate
            {
                RunUiAction(delegate
                {
                    session.InvalidateSearchIndex();
                    lastOutcome = English ? "Search index cleared; the next search will rebuild it."
                        : "Index effacé ; la prochaine recherche le reconstruira.";
                });
            });
            var menu = notifyIcon.ContextMenuStrip;
            int index = menu.Items.Count - 1;
            menu.Items.Insert(index++, preflightItem);
            menu.Items.Insert(index++, profilesItem);
            menu.Items.Insert(index++, updatesItem);
            menu.Items.Insert(index, refreshSearchItem);
            panel.PreflightRequested += delegate { ShowPreflight(); };
            panel.ProfilesRequested += delegate { ShowProfiles(); };
            panel.DiagnosticsRequested += delegate { ShowDiagnostics(); };
            panel.UpdatesRequested += delegate { ShowUpdates(); };
            panel.SaveRehearsalRequested += delegate
            {
                RunUiAction(delegate
                {
                    if (rehearsal.IsRunning) { throw new InvalidOperationException(English ? "Stop rehearsal before saving the final report." : "Arrêtez la répétition avant d’enregistrer le bilan final."); }
                    lastOutcome = PersistRehearsal();
                }, DiagnosticCategory.Profiles);
            };
            panel.CautiousSearchChanged += delegate
            {
                if (!applyingPreferences)
                {
                    RunUiAction(delegate { Registry.SetValue(SettingsKey, "CautiousSearch", panel.CautiousSearch ? 1 : 0, RegistryValueKind.DWord); });
                }
            };
            panel.HotkeysChanged += delegate
            {
                if (!applyingPreferences)
                {
                    RunUiAction(delegate
                    {
                        ApplyHotkeys(panel.HotkeysEnabled);
                        Registry.SetValue(SettingsKey, "KeyboardFallback", hotkeys.Enabled ? 1 : 0, RegistryValueKind.DWord);
                    }, DiagnosticCategory.Hotkeys);
                }
            };
            hotkeys.CommandPressed += delegate(object sender, HotkeyEventArgs args)
            {
                if (exiting) { return; }
                switch (args.Command)
                {
                    case PresentationHotkey.Next: ExecuteCommand(session.Next); break;
                    case PresentationHotkey.Previous: ExecuteCommand(session.Previous); break;
                    case PresentationHotkey.Resume: ExecuteCommand(session.Resume); break;
                    case PresentationHotkey.NextResult: ExecuteCommand(session.NextResult); break;
                    case PresentationHotkey.BlackScreen: ExecuteCommand(delegate { return session.SetBlackScreen(true); }); break;
                    case PresentationHotkey.RestoreSlides: ExecuteCommand(delegate { return session.SetBlackScreen(false); }); break;
                }
            };
            UpdateEnhancementLanguage();
        }

        private void UpdateEnhancementLanguage()
        {
            if (preflightItem == null) { return; }
            preflightItem.Text = English ? "Pre-presentation check..." : "Vérifier avant présentation...";
            profilesItem.Text = English ? "Routes, budgets and history..." : "Parcours, budgets et historique...";
            updatesItem.Text = English ? "Updates..." : "Mises à jour...";
            refreshSearchItem.Text = English ? "Refresh search index" : "Actualiser l’index de recherche";
            bool managed = ManagedDeployment.IsManaged;
            shortcutItem.Enabled = !managed;
            shortcutItem.ToolTipText = managed ? ManagedDeployment.GetMessage(English) : "";
        }

        private static bool ReadBooleanSetting(string name, bool defaultValue)
        {
            object value = Registry.GetValue(SettingsKey, name, defaultValue ? 1 : 0);
            if (!(value is int) || ((int)value != 0 && (int)value != 1)) { throw new InvalidDataException("Invalid Jarvis preference: " + name); }
            return (int)value == 1;
        }

        private void LoadEnhancementSettings()
        {
            applyingPreferences = true;
            try
            {
                panel.CautiousSearch = ReadBooleanSetting("CautiousSearch", true);
                panel.HotkeysEnabled = ReadBooleanSetting("KeyboardFallback", false);
                try { ApplyHotkeys(panel.HotkeysEnabled); }
                catch (InvalidOperationException error)
                {
                    Registry.SetValue(SettingsKey, "KeyboardFallback", 0, RegistryValueKind.DWord);
                    throw new InvalidOperationException(
                        (English ? "Keyboard fallback was disabled in settings. " : "Les raccourcis ont été désactivés dans les réglages. ") + error.Message, error);
                }
            }
            finally { applyingPreferences = false; }
        }

        private void ApplyHotkeys(bool enable)
        {
            try
            {
                if (enable) { hotkeys.Enable(); }
                else { hotkeys.Disable(); }
            }
            finally
            {
                bool previous = applyingPreferences;
                applyingPreferences = true;
                try { panel.HotkeysEnabled = hotkeys.Enabled; }
                finally { applyingPreferences = previous; }
            }
        }

        private DialogResult RunModal(Form dialog)
        {
            using (dialog)
            {
                modalDepth++;
                try { return dialog.ShowDialog(panel); }
                finally { modalDepth--; }
            }
        }

        private string SearchWithConfirmation(string query)
        {
            SearchProposal proposal;
            using (SearchOperation operation = session.BeginSearch(query))
            {
                if (!RunSearchOperation(operation)) { return SearchCancelled(); }
                proposal = operation.Proposal;
            }
            int target = proposal.Candidates[0].SlideId;
            if (panel.CautiousSearch && proposal.IsAmbiguous)
            {
                using (var dialog = new SearchChoiceDialog(proposal, English))
                {
                    modalDepth++;
                    try
                    {
                        if (dialog.ShowDialog(panel) != DialogResult.OK)
                        {
                            return SearchCancelled();
                        }
                        target = dialog.SelectedSlideId;
                    }
                    finally { modalDepth--; }
                }
            }
            if (exiting) { return SearchCancelled(); }
            using (SearchOperation operation = session.BeginAcceptSearch(proposal, target))
            {
                return RunSearchOperation(operation) ? operation.Result : SearchCancelled();
            }
        }

        private string SearchCancelled()
        {
            return English ? "Search cancelled; no slide changed." : "Recherche annulée ; aucune diapositive modifiée.";
        }

        private bool RunSearchOperation(SearchOperation operation)
        {
            using (var dialog = new SearchProgressDialog(English, delegate { return operation.Step(15); }, delegate
            {
                string phase;
                switch (operation.Phase)
                {
                    case SearchPhase.Reading: phase = English ? "Reading slides" : "Lecture des slides"; break;
                    case SearchPhase.Validating:
                        phase = operation.RequiresTextValidation
                            ? (English ? "Final content check next; cancellation is unavailable during this check"
                                : "Contrôle final à venir ; annulation indisponible pendant ce contrôle")
                            : (English ? "Checking for changes" : "Contrôle des modifications");
                        break;
                    case SearchPhase.Scoring: phase = English ? "Ranking results" : "Classement des résultats"; break;
                    case SearchPhase.Complete: phase = English ? "Complete" : "Terminé"; break;
                    default: phase = English ? "Preparing search" : "Préparation de la recherche"; break;
                }
                return phase + " : " + operation.CompletedSlides + " / " + operation.TotalSlides;
            }))
            {
                activeSearch = dialog;
                modalDepth++;
                try
                {
                    DialogResult result = dialog.ShowDialog(panel);
                    dialog.ThrowIfFailed();
                    if (result == DialogResult.OK && !exiting) { return true; }
                    operation.Cancel();
                    return false;
                }
                finally { activeSearch = null; modalDepth--; }
            }
        }

        private void ShowProfiles()
        {
            RunUiAction(delegate
            {
                RunModal(new PresentationProfilesDialog(English, session, profiles, panel.BudgetSeconds));
                PollPresentation();
            }, DiagnosticCategory.Profiles);
        }

        private void SaveSelectedRehearsalBudget()
        {
            int slideId = panel.SelectedRehearsalSlideId;
            if (!string.IsNullOrEmpty(rehearsal.SavedPath))
            {
                Dictionary<int, int> values = profiles.LoadBudgets(rehearsal.SavedPath);
                values[slideId] = panel.BudgetSeconds;
                profiles.SaveBudgets(rehearsal.SavedPath, values);
            }
            rehearsal.SetBudget(slideId, panel.BudgetSeconds);
            historyHandledRunId = null;
            if (!rehearsal.IsRunning) { PersistRehearsal(); }
        }

        private string PersistRehearsal()
        {
            if (string.IsNullOrEmpty(rehearsal.RunId) || rehearsal.Entries.Count == 0)
            {
                return English ? "No rehearsal to save." : "Aucune répétition à enregistrer.";
            }
            if (historyHandledRunId == rehearsal.RunId)
            {
                return English ? "Rehearsal stopped; report is available." : "Répétition arrêtée ; bilan disponible.";
            }
            if (string.IsNullOrEmpty(rehearsal.SavedPath))
            {
                historyHandledRunId = rehearsal.RunId;
                return English
                    ? "Rehearsal kept in memory only: the presentation has not been saved. Use CSV export before exiting."
                    : "Bilan conservé en mémoire uniquement : la présentation n’est pas enregistrée. Exportez le CSV avant de quitter.";
            }
            profiles.SaveRun(rehearsal.SavedPath, new RehearsalRecord
            {
                Id = rehearsal.RunId,
                StartedUtc = rehearsal.StartedUtc,
                Slides = rehearsal.Entries.Select(item => new SlideTiming
                {
                    SlideId = item.SlideId, Seconds = item.Seconds, BudgetSeconds = item.BudgetSeconds
                }).ToList()
            });
            historyHandledRunId = rehearsal.RunId;
            return English ? "Rehearsal saved locally; compare it in Routes and history." : "Répétition enregistrée localement ; comparez-la dans Parcours et historique.";
        }

        private PreflightContext PreflightContext()
        {
            return new PreflightContext
            {
                CultureName = currentCultureName,
                MicrophoneId = microphoneId,
                IsListening = acceptsCommands,
                IndicatorEnabled = panel.IndicatorEnabled,
                IndicatorDisplay = panel.IndicatorDisplay,
                HasRecentAudio = DateTime.UtcNow - lastAudioUtc < TimeSpan.FromSeconds(15)
            };
        }

        private void ShowPreflight()
        {
            RunUiAction(delegate
            {
                RunModal(new PreflightDialog(English, PreflightContext, SnapshotForPreflight, delegate { OpenSetup(false); }));
            }, DiagnosticCategory.Preflight);
        }

        private PresentationSnapshot SnapshotForPreflight()
        {
            try { return session.Snapshot(); }
            catch (PresentationUnavailableException error)
            {
                if (error.Reason == PresentationUnavailableReason.NoSlideShow) { return null; }
                throw;
            }
        }

        private DiagnosticSnapshot DiagnosticSnapshot()
        {
            bool hasShow = false;
            try { session.Snapshot(); hasShow = true; }
            catch (PresentationUnavailableException) { }
            catch (COMException error) { diagnostics.Record(DiagnosticCategory.Navigation, false, error); }
            catch (Win32Exception error) { diagnostics.Record(DiagnosticCategory.Navigation, false, error); }
            return new DiagnosticSnapshot
            {
                CultureName = currentCultureName,
                UsesDefaultMicrophone = microphoneId == "default",
                State = acceptsCommands ? ListeningState.Listening : listeningRequested ? ListeningState.Unavailable : ListeningState.Paused,
                SlideshowAvailable = hasShow,
                ScreenCount = Screen.AllScreens.Length,
                HotkeysEnabled = hotkeys.Enabled
            };
        }

        private void ShowDiagnostics()
        {
            RunUiAction(delegate
            {
                DiagnosticSnapshot snapshot = DiagnosticSnapshot();
                using (var form = new Form
                {
                    Text = English ? "Local diagnostic report" : "Diagnostic local",
                    Size = new Size(740, 550), StartPosition = FormStartPosition.CenterParent
                })
                {
                    form.MinimumSize = new Size(420, 300);
                    var text = new TextBox
                    {
                        Dock = DockStyle.Fill, ReadOnly = true, Multiline = true,
                        ScrollBars = ScrollBars.Both, WordWrap = false, Text = diagnostics.BuildReport(snapshot)
                    };
                    UiAccessibility.Name(text, English ? "Diagnostic report" : "Rapport de diagnostic", 0);
                    var export = new Button { AutoSize = true, Dock = DockStyle.Bottom, MinimumSize = new Size(0, 40) };
                    UiAccessibility.Caption(export, English ? "&Save report locally..." : "&Enregistrer le diagnostic...");
                    export.TabIndex = 1;
                    export.Click += delegate
                    {
                        RunUiAction(delegate
                        {
                            using (var save = new SaveFileDialog { Filter = "Text (*.txt)|*.txt", FileName = "Jarvis-diagnostic.txt" })
                            {
                                if (save.ShowDialog(form) == DialogResult.OK) { diagnostics.Export(save.FileName, snapshot); }
                            }
                        });
                    };
                    form.Controls.Add(text);
                    form.Controls.Add(export);
                    UiAccessibility.Initialize(form);
                    RunModal(form);
                }
            });
        }

        private bool CanInstallUpdate()
        {
            if (ManagedDeployment.IsManaged) { return false; }
            if (rehearsal.IsRunning || activeSetup != null) { return false; }
            if (!string.IsNullOrEmpty(rehearsal.RunId) && historyHandledRunId != rehearsal.RunId) { return false; }
            try
            {
                return !AnyRunningSlideShow();
            }
            catch (COMException error)
            {
                diagnostics.Record(DiagnosticCategory.Updates, false, error);
                return false;
            }
        }

        private static bool AnyRunningSlideShow()
        {
            PowerPoint.Application application = null;
            PowerPoint.SlideShowWindows windows = null;
            try
            {
                application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                windows = application.SlideShowWindows;
                return windows.Count != 0;
            }
            catch (COMException error)
            {
                if (error.ErrorCode == unchecked((int)0x800401E3)) { return false; }
                throw;
            }
            finally
            {
                if (windows != null) { Marshal.ReleaseComObject(windows); }
                if (application != null) { Marshal.ReleaseComObject(application); }
            }
        }

        private void ShowUpdates()
        {
            RunUiAction(delegate
            {
                RunModal(new UpdateDialog(English, CanInstallUpdate, delegate
                {
                    ExitApplication(this, EventArgs.Empty);
                    if (!exiting)
                        throw new InvalidOperationException(English
                            ? "Application shutdown was cancelled; the update was not authorized."
                            : "La fermeture de l'application a été annulée ; la mise à jour n'est pas autorisée.");
                }));
            }, DiagnosticCategory.Updates);
        }

        private bool PrepareToExit()
        {
            if (activeSearch != null) { activeSearch.DialogResult = DialogResult.Cancel; activeSearch.Close(); }
            Exception failure = null;
            try
            {
                if (rehearsal.IsRunning) { rehearsal.Stop(clock.Elapsed); }
                PersistRehearsal();
            }
            catch (IOException error) { failure = error; }
            catch (System.Xml.XmlException error) { failure = error; }
            catch (UnauthorizedAccessException error) { failure = error; }
            catch (SecurityException error) { failure = error; }
            catch (InvalidOperationException error) { failure = error; }
            catch (ArgumentException error) { failure = error; }
            if (failure == null) { return true; }
            diagnostics.Record(DiagnosticCategory.Profiles, false, failure);
            return MessageBox.Show(
                (English ? "Could not save rehearsal history. Exit anyway?\n\n" : "Impossible d’enregistrer l’historique. Quitter quand même ?\n\n") + failure.Message,
                "Jarvis PowerPoint", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2) == DialogResult.Yes;
        }
    }
}
