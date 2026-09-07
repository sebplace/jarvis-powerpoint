using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace JarvisPowerPoint
{
    internal sealed class PresenterPanel : Form
    {
        private readonly TabControl tabs = new TabControl { Dock = DockStyle.Fill };
        private readonly Label status = new Label { AutoSize = true };
        private readonly Label heard = new Label { AutoSize = true };
        private readonly Label outcome = new Label { AutoSize = true };
        private readonly Label presentation = new Label { AutoSize = true };
        private readonly ProgressBar audioLevel = new ProgressBar { Width = 300 };
        private readonly ComboBox displays = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 470 };
        private readonly CheckBox indicatorEnabled = new CheckBox { AutoSize = true };
        private readonly CheckBox cautiousSearch = new CheckBox { AutoSize = true, Checked = true };
        private readonly CheckBox hotkeysEnabled = new CheckBox { AutoSize = true };
        private readonly Label privacy = new Label { AutoSize = true };
        private readonly ListView aliasList = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false, ShowItemToolTips = true };
        private readonly TextBox aliasInput = new TextBox { Width = 240, MaxLength = 80 };
        private readonly ListView rehearsalList = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false, ShowItemToolTips = true };
        private readonly NumericUpDown budget = new NumericUpDown { Minimum = 1, Maximum = 3600, Value = 90, Width = 80 };
        private readonly Label rehearsalSummary = new Label { AutoSize = true };
        private readonly Dictionary<Button, string[]> buttonTexts = new Dictionary<Button, string[]>();
        private readonly ToolTip tooltips = new ToolTip();
        private readonly ListeningIndicator indicator = new ListeningIndicator();
        private readonly Label aliasHelp = new Label { AutoSize = true };
        private readonly Label budgetHelp = new Label { AutoSize = true };
        private readonly Button startRehearsal;
        private readonly Button stopRehearsal;
        private readonly Button listenButton;
        private readonly Button microphoneSetupButton;
        private bool english;

        public event EventHandler PauseRequested;
        public event EventHandler SettingsRequested;
        public event EventHandler QuestionsRequested;
        public event EventHandler ResumeRequested;
        public event EventHandler NextResultRequested;
        public event EventHandler BlackRequested;
        public event EventHandler DisplayRequested;
        public event EventHandler StartRehearsalRequested;
        public event EventHandler StopRehearsalRequested;
        public event EventHandler ExportRehearsalRequested;
        public event EventHandler BudgetRequested;
        public event EventHandler RefreshAliasesRequested;
        public event EventHandler SaveAliasRequested;
        public event EventHandler DeleteAliasRequested;
        public event EventHandler NavigateAliasRequested;
        public event EventHandler ProfilesRequested;
        public event EventHandler PreflightRequested;
        public event EventHandler DiagnosticsRequested;
        public event EventHandler UpdatesRequested;
        public event EventHandler SaveRehearsalRequested;
        public event EventHandler CautiousSearchChanged;
        public event EventHandler HotkeysChanged;

        public string AliasNameInput { get { return aliasInput.Text.Trim(); } }
        public string SelectedAliasName { get { return aliasList.SelectedItems.Count == 0 ? null : aliasList.SelectedItems[0].Text; } }
        public int BudgetSeconds { get { return (int)budget.Value; } }
        public bool AliasesSelected { get { return tabs.SelectedIndex == 2; } }
        public bool IndicatorEnabled { get { return indicatorEnabled.Checked; } }
        public string IndicatorDisplay { get { return displays.SelectedItem as string; } }
        public bool CautiousSearch { get { return cautiousSearch.Checked; } set { cautiousSearch.Checked = value; } }
        public bool HotkeysEnabled { get { return hotkeysEnabled.Checked; } set { hotkeysEnabled.Checked = value; } }
        public int SelectedRehearsalSlideId
        {
            get { return rehearsalList.SelectedItems.Count == 0 ? 0 : (int)rehearsalList.SelectedItems[0].Tag; }
        }

        public PresenterPanel(bool useEnglish)
        {
            Text = "Jarvis PowerPoint";
            Size = new Size(760, 590);
            MinimumSize = new Size(660, 520);
            StartPosition = FormStartPosition.CenterScreen;
            Icon = SystemIcons.Information;
            Font = new Font("Segoe UI", 10);
            Controls.Add(tabs);
            var statusTab = new TabPage();
            var rehearsalTab = new TabPage();
            var aliasesTab = new TabPage();
            tabs.TabPages.AddRange(new[] { statusTab, rehearsalTab, aliasesTab });

            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 0,
                AutoScroll = true, Padding = new Padding(12)
            };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            foreach (Control control in new Control[] { presentation, status, audioLevel, heard, outcome })
                UiAccessibility.AddRow(main, control);
            var commands = Row();
            listenButton = MakeButton("&Écoute / pause", "&Listen / pause", delegate { Raise(PauseRequested); });
            microphoneSetupButton = MakeButton("Réglages &micro...", "&Microphone setup...", delegate { Raise(SettingsRequested); });
            commands.Controls.Add(listenButton);
            commands.Controls.Add(microphoneSetupButton);
            commands.Controls.Add(MakeButton("Mode &questions", "&Questions mode", delegate { Raise(QuestionsRequested); }));
            commands.Controls.Add(MakeButton("&Reprendre", "&Resume", delegate { Raise(ResumeRequested); }));
            commands.Controls.Add(MakeButton("&Autre résultat", "&Next match", delegate { Raise(NextResultRequested); }));
            commands.Controls.Add(MakeButton("Écran &noir", "&Black screen", delegate { Raise(BlackRequested); }));
            commands.Controls.Add(MakeButton("A&fficher", "Restore &slides", delegate { Raise(DisplayRequested); }));
            UiAccessibility.AddRow(main, commands);
            var utilities = Row();
            utilities.Controls.Add(MakeButton("&Vérifier avant présentation", "Pre-presentation &check", delegate { Raise(PreflightRequested); }));
            utilities.Controls.Add(MakeButton("&Parcours et historique...", "Routes and &history...", delegate { Raise(ProfilesRequested); }));
            utilities.Controls.Add(MakeButton("&Diagnostic local...", "Local &diagnostics...", delegate { Raise(DiagnosticsRequested); }));
            utilities.Controls.Add(MakeButton("Mises à &jour...", "&Updates...", delegate { Raise(UpdatesRequested); }));
            UiAccessibility.AddRow(main, utilities);
            UiAccessibility.AddRow(main, cautiousSearch);
            UiAccessibility.AddRow(main, hotkeysEnabled);
            cautiousSearch.CheckedChanged += delegate { Raise(CautiousSearchChanged); };
            hotkeysEnabled.CheckedChanged += delegate { Raise(HotkeysChanged); };
            UiAccessibility.AddRow(main, indicatorEnabled);
            UiAccessibility.AddRow(main, displays);
            UiAccessibility.AddRow(main, privacy);
            statusTab.Controls.Add(main);

            var rehearsalTop = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8), TabIndex = 0 };
            rehearsalList.TabIndex = 1;
            rehearsalTop.Controls.Add(budgetHelp);
            rehearsalTop.Controls.Add(budget);
            rehearsalTop.Controls.Add(MakeButton("Appliquer à la sélection", "Apply to selected slide", delegate { Raise(BudgetRequested); }));
            startRehearsal = MakeButton("Nouvelle répétition", "New rehearsal", delegate { Raise(StartRehearsalRequested); });
            stopRehearsal = MakeButton("Arrêter", "Stop", delegate { Raise(StopRehearsalRequested); });
            rehearsalTop.Controls.Add(startRehearsal);
            rehearsalTop.Controls.Add(stopRehearsal);
            rehearsalTop.Controls.Add(MakeButton("Exporter CSV...", "Export CSV...", delegate { Raise(ExportRehearsalRequested); }));
            rehearsalTop.Controls.Add(MakeButton("Enregistrer le bilan", "Save rehearsal", delegate { Raise(SaveRehearsalRequested); }));
            rehearsalTop.Controls.Add(MakeButton("Comparer...", "Compare...", delegate { Raise(ProfilesRequested); }));
            rehearsalTop.Controls.Add(rehearsalSummary);
            rehearsalTop.SetFlowBreak(rehearsalSummary, true);
            UiAccessibility.WrapLabel(rehearsalSummary, rehearsalTop);
            rehearsalList.Columns.Add("Slide", 65);
            rehearsalList.Columns.Add("Titre", 265);
            rehearsalList.Columns.Add("Temps", 90);
            rehearsalList.Columns.Add("Budget", 90);
            rehearsalList.Columns.Add("Dépassement", 100);
            UiAccessibility.ScrollableSection(rehearsalTab, rehearsalTop, rehearsalList, 160);

            var aliasesTop = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8), TabIndex = 0 };
            aliasList.TabIndex = 1;
            UiAccessibility.WrapLabel(aliasHelp, aliasesTop);
            aliasesTop.Controls.Add(aliasHelp);
            aliasesTop.SetFlowBreak(aliasHelp, true);
            aliasesTop.Controls.Add(aliasInput);
            aliasesTop.Controls.Add(MakeButton("Associer au slide courant", "Assign to current slide", delegate { Raise(SaveAliasRequested); }));
            aliasesTop.Controls.Add(MakeButton("Actualiser", "Refresh", delegate { Raise(RefreshAliasesRequested); }));
            aliasesTop.Controls.Add(MakeButton("Ouvrir", "Go to", delegate { Raise(NavigateAliasRequested); }));
            aliasesTop.Controls.Add(MakeButton("Supprimer", "Delete", delegate { Raise(DeleteAliasRequested); }));
            aliasList.Columns.Add("Alias", 200);
            aliasList.Columns.Add("Slide", 65);
            aliasList.Columns.Add("Titre", 355);
            UiAccessibility.ScrollableSection(aliasesTab, aliasesTop, aliasList, 160);
            tabs.SelectedIndexChanged += delegate { if (tabs.SelectedIndex == 2) { Raise(RefreshAliasesRequested); } };

            foreach (Screen screen in Screen.AllScreens) { displays.Items.Add(screen.DeviceName); }
            if (displays.Items.Count > 0) { displays.SelectedIndex = 0; }
            indicatorEnabled.CheckedChanged += delegate { UpdateIndicatorVisibility(); };
            displays.SelectedIndexChanged += delegate { UpdateIndicatorVisibility(); };
            FormClosing += delegate(object sender, FormClosingEventArgs args)
            {
                if (args.CloseReason == CloseReason.UserClosing) { args.Cancel = true; Hide(); }
            };
            SetLanguage(useEnglish);
            UiAccessibility.Initialize(this);
        }

        private static FlowLayoutPanel Row()
        {
            return new FlowLayoutPanel { AutoSize = true, WrapContents = true };
        }

        private Button MakeButton(string french, string en, EventHandler action)
        {
            var button = new Button { AutoSize = true, Padding = new Padding(4), Margin = new Padding(4) };
            buttonTexts.Add(button, new[] { french, en });
            button.Click += action;
            return button;
        }

        private void Raise(EventHandler handler)
        {
            if (handler != null) { handler(this, EventArgs.Empty); }
        }

        public void SetLanguage(bool value)
        {
            english = value;
            AccessibleName = english ? "Jarvis PowerPoint presenter" : "Présentateur Jarvis PowerPoint";
            tabs.TabPages[0].Text = english ? "Presenter" : "Présentateur";
            tabs.TabPages[1].Text = english ? "Rehearsal" : "Répétition";
            tabs.TabPages[2].Text = english ? "Slide aliases" : "Alias de slides";
            foreach (var item in buttonTexts) { UiAccessibility.Caption(item.Key, item.Value[english ? 1 : 0]); }
            UiAccessibility.Caption(indicatorEnabled, english ? "Show compact listening &indicator on this screen:" : "Afficher l’&indicateur compact sur cet écran :");
            UiAccessibility.Caption(cautiousSearch, english ? "Con&firm ambiguous search results before moving" : "&Confirmer les résultats ambigus avant de changer de slide");
            UiAccessibility.Caption(hotkeysEnabled, english ? "Enable &keyboard fallback (Ctrl+Alt+Shift)" : "Ac&tiver les raccourcis de secours (Ctrl+Alt+Maj)");
            tooltips.SetToolTip(hotkeysEnabled, english
                ? "Right / Left: next / previous; Enter: resume; N: next match; B: black screen; S: restore."
                : "Droite / Gauche : suivant / précédent ; Entrée : reprise ; N : autre résultat ; B : écran noir ; S : afficher.");
            privacy.Text = english
                ? "This panel and indicator are visible on your screen. Select your presenter monitor. Share only the PowerPoint slide-show window, not your entire desktop."
                : "Ce panneau et l’indicateur sont visibles à l’écran. Choisissez votre écran présentateur. Partagez uniquement la fenêtre du diaporama, pas tout le bureau.";
            aliasHelp.Text = english
                ? "Saved presentation only. Name the CURRENT slide (e.g. the demo), then say “Jarvis, shortcut the demo”."
                : "Présentation enregistrée uniquement. Nommez le slide COURANT (ex. la démo), puis dites « Jarvis, raccourci la démo ».";
            budgetHelp.Text = english ? "Budget per slide (seconds):" : "Budget par slide (secondes) :";
            aliasList.Columns[2].Text = english ? "Title" : "Titre";
            rehearsalList.Columns[1].Text = english ? "Title" : "Titre";
            rehearsalList.Columns[2].Text = english ? "Time" : "Temps";
            rehearsalList.Columns[4].Text = english ? "Over budget" : "Dépassement";
            tooltips.SetToolTip(budget, english ? "Default for a new rehearsal; apply to a selected row to customize." : "Budget par défaut d’une nouvelle répétition ; personnalisable sur une ligne sélectionnée.");
            tabs.AccessibleName = english ? "Presenter tools" : "Outils du présentateur";
            audioLevel.AccessibleName = english ? "Microphone input level" : "Niveau d’entrée du microphone";
            displays.AccessibleName = english ? "Listening indicator screen" : "Écran de l’indicateur d’écoute";
            aliasInput.AccessibleName = english ? "Alias for the current slide" : "Alias du slide courant";
            aliasList.AccessibleName = english ? "Saved slide aliases" : "Alias de slides enregistrés";
            rehearsalList.AccessibleName = english ? "Rehearsal slide timings" : "Durées des slides de la répétition";
            budget.AccessibleName = budgetHelp.Text;
            hotkeysEnabled.AccessibleDescription = tooltips.GetToolTip(hotkeysEnabled);
            privacy.UseMnemonic = false;
            indicator.SetLanguage(english);
        }

        public void SetPresentation(string text)
        {
            text = text ?? string.Empty;
            if (presentation.Text != text) presentation.Text = text;
        }

        public void SetStatus(string state, string recognized, string result, int level)
        {
            state = state ?? string.Empty;
            string heardText = (english ? "Heard: " : "Entendu : ") + recognized;
            string outcomeText = (english ? "Result: " : "Résultat : ") + result;
            int clampedLevel = Math.Max(0, Math.Min(100, level));
            if (status.Text != state) status.Text = state;
            if (heard.Text != heardText) heard.Text = heardText;
            if (outcome.Text != outcomeText) outcome.Text = outcomeText;
            if (audioLevel.Value != clampedLevel) audioLevel.Value = clampedLevel;
            indicator.UpdateText(state, heard.Text, outcome.Text, audioLevel.Value);
        }

        public void SetSpeechControlsEnabled(bool enabled)
        {
            if (listenButton.Enabled != enabled) listenButton.Enabled = enabled;
            if (microphoneSetupButton.Enabled != enabled) microphoneSetupButton.Enabled = enabled;
        }

        public void SetAliases(IList<SlideAlias> aliases)
        {
            string selected = SelectedAliasName;
            aliasList.BeginUpdate();
            aliasList.Items.Clear();
            foreach (SlideAlias alias in aliases)
            {
                var item = new ListViewItem(new[] { alias.Name, alias.SlideNumber == 0 ? "?" : alias.SlideNumber.ToString(), alias.Title ?? string.Empty })
                    { ToolTipText = alias.Name + " · " + alias.Title };
                aliasList.Items.Add(item);
                item.Selected = alias.Name == selected;
            }
            aliasList.EndUpdate();
        }

        public void SetRehearsal(RehearsalTracker tracker)
        {
            int selected = SelectedRehearsalSlideId;
            int topId = rehearsalList.TopItem == null ? 0 : (int)rehearsalList.TopItem.Tag;
            rehearsalList.BeginUpdate();
            rehearsalList.Items.Clear();
            foreach (RehearsalEntry entry in tracker.Entries)
            {
                var item = new ListViewItem(new[]
                {
                    entry.SlideNumber.ToString(), entry.Title ?? string.Empty,
                    FormatTime(entry.Seconds), FormatTime(entry.BudgetSeconds),
                    entry.OverBudget ? (english ? "Yes" : "Oui") : ""
                }) { Tag = entry.SlideId, ToolTipText = entry.Title ?? "" };
                rehearsalList.Items.Add(item);
                item.Selected = entry.SlideId == selected;
            }
            if (topId != 0)
            {
                ListViewItem top = rehearsalList.Items.Cast<ListViewItem>().FirstOrDefault(item => (int)item.Tag == topId);
                if (top != null) { rehearsalList.TopItem = top; }
            }
            rehearsalList.EndUpdate();
            rehearsalSummary.Text = (tracker.IsRunning ? (english ? "Running · " : "En cours · ") : (english ? "Stopped · " : "Arrêtée · "))
                + FormatTime(tracker.TotalSeconds) + (english ? " total; see saved runs in Routes and history" : " au total ; bilans dans Parcours et historique");
            startRehearsal.Enabled = !tracker.IsRunning;
            stopRehearsal.Enabled = tracker.IsRunning;
        }

        private static string FormatTime(double seconds)
        {
            return ((int)seconds / 60).ToString("00") + ":" + ((int)seconds % 60).ToString("00");
        }

        private void UpdateIndicatorVisibility()
        {
            if (!indicatorEnabled.Checked) { indicator.Hide(); return; }
            Screen screen = Screen.AllScreens.FirstOrDefault(item => item.DeviceName == (string)displays.SelectedItem);
            if (screen == null) { indicatorEnabled.Checked = false; return; }
            indicator.Location = new Point(screen.WorkingArea.Right - indicator.Width - 16, screen.WorkingArea.Top + 16);
            indicator.Show();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { indicator.Dispose(); tooltips.Dispose(); }
            base.Dispose(disposing);
        }
    }

    internal sealed class ListeningIndicator : Form
    {
        private readonly Label text = new Label { Dock = DockStyle.Fill, Padding = new Padding(8), AutoEllipsis = true, UseMnemonic = false };
        private readonly ProgressBar level = new ProgressBar { Dock = DockStyle.Bottom, Height = 8 };
        protected override bool ShowWithoutActivation { get { return true; } }

        public ListeningIndicator()
        {
            Text = "Jarvis";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ControlBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(355, 138);
            Font = new Font("Segoe UI", 9);
            Controls.Add(text);
            Controls.Add(level);
            UiAccessibility.Initialize(this);
        }

        public void UpdateText(string state, string heard, string outcome, int audioLevel)
        {
            string statusText = state + "\n" + heard + "\n" + outcome;
            if (text.Text != statusText) text.Text = statusText;
            if (text.AccessibleName != statusText) text.AccessibleName = statusText;
            if (level.Value != audioLevel) level.Value = audioLevel;
        }

        public void SetLanguage(bool english)
        {
            AccessibleName = english ? "Compact listening indicator" : "Indicateur compact d’écoute";
            level.AccessibleName = english ? "Microphone input level" : "Niveau d’entrée du microphone";
        }
    }
}
