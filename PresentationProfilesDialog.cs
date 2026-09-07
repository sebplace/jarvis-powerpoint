using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;

namespace JarvisPowerPoint
{
    internal sealed class PresentationProfilesDialog : Form
    {
        private readonly PresentationSession session;
        private readonly PresentationProfileStore store;
        private readonly PresentationSnapshot origin;
        private readonly bool english;
        private readonly int defaultBudgetSeconds;
        private readonly ListBox available = new ListBox { SelectionMode = SelectionMode.MultiExtended, Dock = DockStyle.Fill, HorizontalScrollbar = true, IntegralHeight = false };
        private readonly ListBox routeSlides = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true, IntegralHeight = false };
        private readonly ComboBox routes = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
        private readonly TextBox routeName = new TextBox { Width = 140, MaxLength = 80 };
        private readonly NumericUpDown minutes = new NumericUpDown { Minimum = 1, Maximum = 240, Value = 5, Width = 60 };
        private readonly ListView budgets = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false, ShowItemToolTips = true };
        private readonly NumericUpDown seconds = new NumericUpDown { Minimum = 1, Maximum = 3600, Value = 90, Width = 90 };
        private readonly ComboBox olderRun = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 240 };
        private readonly ComboBox newerRun = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 240 };
        private readonly ListView comparison = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false, ShowItemToolTips = true };
        private readonly Label historySummary = new Label { AutoSize = true };
        private readonly Label routeStatus = new Label { AutoSize = true };
        private readonly Label routeEstimate = new Label { AutoSize = true };
        private readonly TextBox result = new TextBox { Dock = DockStyle.Top, Height = 55, Multiline = true,
            ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = SystemColors.Window, ForeColor = SystemColors.WindowText };
        private List<SlideChoice> slides;
        private Dictionary<int, int> knownBudgets = new Dictionary<int, int>();
        private bool refreshing;

        public PresentationProfilesDialog(bool useEnglish, PresentationSession presentationSession, PresentationProfileStore profileStore, int defaultBudget = 90)
        {
            if (defaultBudget < 1 || defaultBudget > 3600) { throw new ArgumentOutOfRangeException("defaultBudget"); }
            defaultBudgetSeconds = defaultBudget;
            english = useEnglish;
            session = presentationSession;
            store = profileStore;
            origin = session.Snapshot();
            if (string.IsNullOrWhiteSpace(origin.SavedPath))
            {
                throw new InvalidOperationException(T("Enregistrez la présentation avant de configurer ses parcours et budgets.",
                    "Save the presentation before configuring routes and budgets."));
            }
            Text = T("Parcours, budgets et historiques", "Routes, budgets and history") + " · " + origin.PresentationName;
            AccessibleName = T("Parcours, budgets et historiques", "Routes, budgets and history");
            Size = new Size(900, 650);
            MinimumSize = new Size(790, 550);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 10);
            MinimizeBox = false;
            ShowInTaskbar = false;
            var tabs = new TabControl { Dock = DockStyle.Fill, TabIndex = 0,
                AccessibleName = T("Outils de la présentation", "Presentation tools") };
            Controls.Add(tabs);
            var footer = new TableLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true,
                ColumnCount = 1, RowCount = 0, Padding = new Padding(8), TabIndex = 1 };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            result.AccessibleName = T("Résultat de l’opération", "Operation result");
            UiAccessibility.AddRow(footer, result);
            var close = new Button { AutoSize = true, DialogResult = DialogResult.Cancel, Anchor = AnchorStyles.Right };
            UiAccessibility.Caption(close, T("&Fermer", "&Close"));
            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
            buttons.Controls.Add(close);
            UiAccessibility.AddRow(footer, buttons);
            Controls.Add(footer);
            CancelButton = close;
            AcceptButton = null;
            var routeTab = new TabPage(T("Parcours", "Routes"));
            var budgetTab = new TabPage(T("Budgets", "Budgets"));
            var historyTab = new TabPage(T("Historique", "History"));
            tabs.TabPages.AddRange(new[] { routeTab, budgetTab, historyTab });
            BuildRoutes(routeTab);
            BuildBudgets(budgetTab);
            BuildHistory(historyTab);
            available.AccessibleName = T("Slides disponibles", "Available slides");
            available.AccessibleDescription = T("Ctrl ou Maj avec les flèches pour sélectionner plusieurs slides.",
                "Use Ctrl or Shift with arrow keys to select multiple slides.");
            routeSlides.AccessibleName = T("Ordre du parcours choisi", "Selected route order");
            routes.AccessibleName = T("Parcours enregistrés", "Saved routes");
            routeName.AccessibleName = T("Nom du parcours", "Route name");
            minutes.AccessibleName = T("Durée cible en minutes", "Target duration in minutes");
            budgets.AccessibleName = T("Budgets par slide", "Slide budgets");
            seconds.AccessibleName = T("Budget du slide sélectionné en secondes", "Selected slide budget in seconds");
            olderRun.AccessibleName = T("Ancienne répétition", "Earlier rehearsal");
            newerRun.AccessibleName = T("Nouvelle répétition", "Later rehearsal");
            comparison.AccessibleName = T("Comparaison des répétitions", "Rehearsal comparison");
            routes.SelectedIndexChanged += delegate { if (!refreshing) { LoadRoute(); } };
            budgets.SelectedIndexChanged += delegate
            {
                if (budgets.SelectedItems.Count > 0)
                {
                    seconds.Value = int.Parse(budgets.SelectedItems[0].SubItems[2].Text, CultureInfo.InvariantCulture);
                }
            };
            olderRun.SelectedIndexChanged += delegate { CompareRuns(); };
            newerRun.SelectedIndexChanged += delegate { CompareRuns(); };
            minutes.ValueChanged += delegate { UpdateRouteEstimate(); };
            Run(RefreshData);
            UiAccessibility.Initialize(this);
        }

        private string T(string french, string en) { return english ? en : french; }
        private Button Button(string french, string en, Action action)
        {
            var button = new Button { AutoSize = true, Margin = new Padding(4), Padding = new Padding(3) };
            UiAccessibility.Caption(button, T(french, en));
            button.Click += delegate { Run(action); };
            return button;
        }
        private static FlowLayoutPanel Bar()
        {
            return new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8), TabIndex = 0 };
        }

        private void BuildRoutes(TabPage tab)
        {
            var bar = Bar();
            bar.Controls.Add(new Label { AutoSize = true, Text = T("Enregistrés :", "Saved:") });
            bar.Controls.Add(routes);
            bar.Controls.Add(Button("&Nouveau", "&New", delegate
            {
                routeName.Text = "5 minutes"; minutes.Value = 5; routeSlides.Items.Clear(); UpdateRouteEstimate();
            }));
            bar.Controls.Add(new Label { AutoSize = true, Text = T("Nom :", "Name:") });
            bar.Controls.Add(routeName);
            bar.Controls.Add(minutes);
            bar.Controls.Add(new Label { AutoSize = true, Text = "min" });
            bar.Controls.Add(Button("&Enregistrer", "&Save", SaveRoute));
            bar.Controls.Add(Button("Acti&ver", "Acti&vate", ActivateRoute));
            bar.Controls.Add(Button("Supprimer", "Delete", DeleteRoute));
            bar.Controls.Add(Button("Parcours complet", "Full presentation", delegate
            {
                GuardDeck(); session.DeactivateRoute(); UpdateRouteStatus();
                result.Text = T("Parcours complet rétabli, sans déplacement.", "Full presentation restored without moving.");
            }));
            bar.Controls.Add(Button("&Actualiser", "&Refresh", RefreshData));
            bar.SetFlowBreak(bar.Controls[bar.Controls.Count - 1], true);
            bar.Controls.Add(routeStatus);
            bar.SetFlowBreak(routeStatus, true);
            bar.Controls.Add(routeEstimate);
            UiAccessibility.WrapLabel(routeStatus, bar);
            UiAccessibility.WrapLabel(routeEstimate, bar);
            var lists = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 2, Padding = new Padding(8), TabIndex = 1 };
            lists.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            lists.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            lists.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            lists.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lists.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            lists.Controls.Add(new Label { Text = T("Slides disponibles", "Available slides"), AutoSize = true }, 0, 0);
            lists.Controls.Add(new Label { Text = T("Ordre du parcours choisi", "Selected route order"), AutoSize = true }, 2, 0);
            lists.Controls.Add(available, 0, 1);
            lists.Controls.Add(routeSlides, 2, 1);
            available.TabIndex = 0;
            routeSlides.TabIndex = 2;
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true,
                FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, TabIndex = 1 };
            actions.Controls.Add(Button("Ajouter >", "Add >", AddSlides));
            actions.Controls.Add(Button("Retirer", "Remove", delegate
            {
                if (routeSlides.SelectedIndex >= 0) { routeSlides.Items.RemoveAt(routeSlides.SelectedIndex); UpdateRouteEstimate(); }
                else { throw new InvalidOperationException(T("Sélectionnez un slide du parcours.", "Select a route slide.")); }
            }));
            actions.Controls.Add(Button("Monter", "Up", delegate { MoveSlide(-1); }));
            actions.Controls.Add(Button("Descendre", "Down", delegate { MoveSlide(1); }));
            lists.Controls.Add(actions, 1, 1);
            UiAccessibility.ScrollableSection(tab, bar, lists, 200);
        }

        private void BuildBudgets(TabPage tab)
        {
            var bar = Bar();
            budgets.TabIndex = 1;
            bar.Controls.Add(new Label { AutoSize = true, Text = T("Budget du slide sélectionné (secondes) :", "Selected slide budget (seconds):") });
            bar.Controls.Add(seconds);
            bar.Controls.Add(Button("&Enregistrer", "&Save", delegate
            {
                GuardDeck();
                if (budgets.SelectedItems.Count != 1) { throw new InvalidOperationException(T("Sélectionnez un slide.", "Select a slide.")); }
                int id = (int)budgets.SelectedItems[0].Tag;
                Dictionary<int, int> saved = store.LoadBudgets(origin.SavedPath);
                saved[id] = (int)seconds.Value;
                store.SaveBudgets(origin.SavedPath, saved);
                knownBudgets = saved;
                UpdateRouteEstimate();
                budgets.SelectedItems[0].SubItems[2].Text = ((int)seconds.Value).ToString(CultureInfo.InvariantCulture);
                result.Text = T("Budget mémorisé pour cette présentation.", "Budget saved for this presentation.");
            }));
            budgets.Columns.Add("Slide", 65);
            budgets.Columns.Add(T("Titre", "Title"), 500);
            budgets.Columns.Add(T("Budget (s)", "Budget (s)"), 115);
            UiAccessibility.ScrollableSection(tab, bar, budgets, 160);
        }

        private void BuildHistory(TabPage tab)
        {
            var bar = Bar();
            comparison.TabIndex = 1;
            bar.Controls.Add(new Label { AutoSize = true, Text = T("Ancienne répétition :", "Earlier rehearsal:") });
            bar.Controls.Add(olderRun);
            bar.Controls.Add(new Label { AutoSize = true, Text = T("Nouvelle :", "Later:") });
            bar.Controls.Add(newerRun);
            bar.Controls.Add(Button("&Actualiser", "&Refresh", RefreshData));
            bar.SetFlowBreak(bar.Controls[bar.Controls.Count - 1], true);
            bar.Controls.Add(historySummary);
            UiAccessibility.WrapLabel(historySummary, bar);
            comparison.Columns.Add("Slide", 60);
            comparison.Columns.Add(T("Titre actuel", "Current title"), 300);
            comparison.Columns.Add(T("Avant (s)", "Before (s)"), 100);
            comparison.Columns.Add(T("Après (s)", "After (s)"), 100);
            comparison.Columns.Add(T("Écart (s)", "Change (s)"), 110);
            UiAccessibility.ScrollableSection(tab, bar, comparison, 160);
        }

        private void GuardDeck()
        {
            PresentationSnapshot now = session.Snapshot();
            if (now.SessionKey != origin.SessionKey || now.PresentationKey != origin.PresentationKey)
            {
                throw new InvalidOperationException(T("Le diaporama a changé. Fermez ce panneau et rouvrez-le.",
                    "The slide show changed. Close and reopen this panel."));
            }
        }

        private void RefreshData()
        {
            GuardDeck();
            slides = session.GetSlides();
            refreshing = true;
            try
            {
                available.Items.Clear();
                foreach (SlideChoice slide in slides) { available.Items.Add(new SlideItem(slide)); }
                routes.Items.Clear();
                foreach (SlideRoute route in session.GetRoutes()) { routes.Items.Add(new RouteItem(route)); }
                budgets.Items.Clear();
                Dictionary<int, int> saved = store.LoadBudgets(origin.SavedPath);
                knownBudgets = saved;
                foreach (SlideChoice slide in slides)
                {
                    int value;
                    if (!saved.TryGetValue(slide.SlideId, out value)) { value = defaultBudgetSeconds; }
                    budgets.Items.Add(new ListViewItem(new[] { slide.SlideNumber.ToString(), slide.Title ?? "", value.ToString() })
                        { Tag = slide.SlideId, ToolTipText = slide.Title ?? "" });
                }
                olderRun.Items.Clear();
                newerRun.Items.Clear();
                foreach (RehearsalRecord record in store.LoadRuns(origin.SavedPath).OrderByDescending(item => item.StartedUtc))
                {
                    olderRun.Items.Add(new RunItem(record));
                    newerRun.Items.Add(new RunItem(record));
                }
                if (newerRun.Items.Count > 0) { newerRun.SelectedIndex = 0; }
                if (olderRun.Items.Count > 0) { olderRun.SelectedIndex = olderRun.Items.Count > 1 ? 1 : 0; }
            }
            finally { refreshing = false; }
            UpdateRouteStatus();
            UpdateRouteEstimate();
            CompareRuns();
        }

        private void LoadRoute()
        {
            var selected = routes.SelectedItem as RouteItem;
            if (selected == null) { return; }
            routeName.Text = selected.Value.Name;
            minutes.Value = selected.Value.TargetMinutes;
            routeSlides.Items.Clear();
            foreach (int id in selected.Value.SlideIds)
            {
                SlideChoice slide = slides.FirstOrDefault(item => item.SlideId == id);
                routeSlides.Items.Add(new SlideItem(slide ?? new SlideChoice { SlideId = id, Title = T("(supprimé)", "(deleted)") }));
            }
            UpdateRouteEstimate();
        }

        private void AddSlides()
        {
            if (available.SelectedItems.Count == 0) { throw new InvalidOperationException(T("Sélectionnez des slides à ajouter.", "Select slides to add.")); }
            foreach (SlideItem item in available.SelectedItems)
            {
                if (!routeSlides.Items.Cast<SlideItem>().Any(existing => existing.Value.SlideId == item.Value.SlideId))
                    routeSlides.Items.Add(item);
            }
            UpdateRouteEstimate();
        }

        private void MoveSlide(int delta)
        {
            int index = routeSlides.SelectedIndex;
            if (index < 0) { throw new InvalidOperationException(T("Sélectionnez un slide du parcours.", "Select a route slide.")); }
            int target = index + delta;
            if (target < 0 || target >= routeSlides.Items.Count) { return; }
            object item = routeSlides.Items[index];
            routeSlides.Items.RemoveAt(index);
            routeSlides.Items.Insert(target, item);
            routeSlides.SelectedIndex = target;
        }

        private void SaveRoute()
        {
            GuardDeck();
            var route = new SlideRoute
            {
                Name = routeName.Text.Trim(), TargetMinutes = (int)minutes.Value,
                SlideIds = routeSlides.Items.Cast<SlideItem>().Select(item => item.Value.SlideId).ToList()
            };
            session.SaveRoute(route);
            RefreshData();
            for (int index = 0; index < routes.Items.Count; index++)
                if (((RouteItem)routes.Items[index]).Value.Name == route.Name) { routes.SelectedIndex = index; break; }
            result.Text = T("Parcours mémorisé sans modifier le PowerPoint.", "Route saved without modifying PowerPoint.");
        }

        private void ActivateRoute()
        {
            GuardDeck();
            var route = routes.SelectedItem as RouteItem;
            if (route == null) { throw new InvalidOperationException(T("Enregistrez et sélectionnez un parcours.", "Save and select a route.")); }
            if (MessageBox.Show(this,
                T("Démarrer au premier slide du parcours « ", "Start at the first slide of route \"") + route.Value.Name
                    + T(" » ? Les repères de reprise et de recherche seront effacés.", "\"? Return points and search results will be cleared."),
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) { return; }
            GuardDeck();
            session.ActivateRoute(route.Value.Name);
            UpdateRouteStatus();
            result.Text = T("Parcours activé. Suivant/précédent suivent cet ordre.", "Route active. Next/previous follow this order.");
        }

        private void DeleteRoute()
        {
            GuardDeck();
            var route = routes.SelectedItem as RouteItem;
            if (route == null) { throw new InvalidOperationException(T("Sélectionnez un parcours.", "Select a route.")); }
            if (MessageBox.Show(this, T("Supprimer le parcours ", "Delete route ") + route.Value.Name + " ?",
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) { return; }
            GuardDeck();
            session.RemoveRoute(route.Value.Name);
            routeSlides.Items.Clear();
            RefreshData();
        }

        private void UpdateRouteStatus()
        {
            routeStatus.Text = T("Actif : ", "Active: ") + (session.ActiveRouteName ?? T("complet", "full"));
        }

        private void UpdateRouteEstimate()
        {
            int total = 0;
            foreach (SlideItem item in routeSlides.Items)
            {
                int value;
                total += knownBudgets.TryGetValue(item.Value.SlideId, out value) ? value : defaultBudgetSeconds;
            }
            routeEstimate.Text = routeSlides.Items.Count + T(" slides · budgets : ", " slides · budgets: ")
                + (total / 60.0).ToString("F1") + T(" min / cible : ", " min / target: ") + minutes.Value + " min";
            if (total > (int)minutes.Value * 60)
                routeEstimate.Text += T(" — au-dessus de la cible", " — over target");
            routeEstimate.ForeColor = SystemColors.ControlText;
        }

        private void CompareRuns()
        {
            if (refreshing) { return; }
            comparison.Items.Clear();
            var before = olderRun.SelectedItem as RunItem;
            var after = newerRun.SelectedItem as RunItem;
            if (before == null || after == null)
            {
                historySummary.Text = T("Aucune répétition enregistrée.", "No saved rehearsal.");
                return;
            }
            foreach (int id in before.Value.Slides.Select(item => item.SlideId).Union(after.Value.Slides.Select(item => item.SlideId)))
            {
                SlideTiming left = before.Value.Slides.FirstOrDefault(item => item.SlideId == id);
                SlideTiming right = after.Value.Slides.FirstOrDefault(item => item.SlideId == id);
                SlideChoice slide = slides.FirstOrDefault(item => item.SlideId == id);
                comparison.Items.Add(new ListViewItem(new[]
                {
                    slide == null ? "?" : slide.SlideNumber.ToString(),
                    slide == null ? T("(supprimé)", "(deleted)") : slide.Title ?? "",
                    left == null ? "-" : left.Seconds.ToString("F1"),
                    right == null ? "-" : right.Seconds.ToString("F1"),
                    left == null || right == null ? "-" : (right.Seconds - left.Seconds).ToString("+0.0;-0.0;0.0")
                }));
            }
            historySummary.Text = T("Total avant/après : ", "Total before/after: ")
                + before.Value.Slides.Sum(item => item.Seconds).ToString("F1") + " / "
                + after.Value.Slides.Sum(item => item.Seconds).ToString("F1") + " s";
            if (before.Value.Id == after.Value.Id)
                historySummary.Text = T("Même répétition sélectionnée : choisissez deux bilans distincts pour comparer.",
                    "The same rehearsal is selected: choose two different runs to compare.");
        }

        private void Run(Action action)
        {
            try { action(); }
            catch (InvalidOperationException error) { result.Text = error.Message; }
            catch (IOException error) { result.Text = error.Message; }
            catch (XmlException error) { result.Text = error.Message; }
            catch (COMException error) { result.Text = error.Message; }
            catch (UnauthorizedAccessException error) { result.Text = error.Message; }
            catch (System.Security.SecurityException error) { result.Text = error.Message; }
            catch (System.ComponentModel.Win32Exception error) { result.Text = error.Message; }
            catch (ArgumentException error) { result.Text = error.Message; }
        }

        private sealed class SlideItem
        {
            public SlideItem(SlideChoice value) { Value = value; }
            public SlideChoice Value { get; private set; }
            public override string ToString() { return (Value.SlideNumber == 0 ? "?" : Value.SlideNumber.ToString()) + " · " + Value.Title; }
        }
        private sealed class RouteItem
        {
            public RouteItem(SlideRoute value) { Value = value; }
            public SlideRoute Value { get; private set; }
            public override string ToString() { return Value.Name + " (" + Value.TargetMinutes + " min)"; }
        }
        private sealed class RunItem
        {
            public RunItem(RehearsalRecord value) { Value = value; }
            public RehearsalRecord Value { get; private set; }
            public override string ToString() { return Value.StartedUtc.ToLocalTime().ToString("g") + " · " + Value.Slides.Sum(item => item.Seconds).ToString("F0") + " s"; }
        }
    }
}
