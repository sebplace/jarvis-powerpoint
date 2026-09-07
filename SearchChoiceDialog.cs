using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace JarvisPowerPoint
{
    internal sealed class SearchChoiceDialog : Form
    {
        private readonly ListView candidates = new ListView
        {
            Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true,
            MultiSelect = false, HideSelection = false, ShowItemToolTips = true, TabIndex = 0
        };
        private readonly Button confirm = new Button { AutoSize = true, Enabled = false, TabIndex = 0 };

        public int SelectedSlideId { get; private set; }

        public SearchChoiceDialog(SearchProposal proposal, bool english)
        {
            Text = english ? "Choose a search result" : "Choisir un résultat";
            AccessibleName = Text;
            Size = new Size(700, 440);
            MinimumSize = new Size(560, 400);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 10);
            ShowInTaskbar = false;
            MinimizeBox = false;
            var help = new Label
            {
                Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(8), UseMnemonic = false,
                Text = english
                    ? "Several slides have similar scores. Nothing has moved. Select with the arrow keys, then Tab to Go and press Space, or cancel (Esc). Enter in the list does not move the slides.\nThis window may be visible in a full-screen share."
                    : "Plusieurs diapositives ont des scores proches. Rien n’a bougé. Choisissez avec les flèches, puis Tab vers Ouvrir et Espace, ou annulez (Échap). Entrée dans la liste ne déplace pas les slides.\nCette fenêtre peut apparaître dans un partage d’écran complet."
            };
            UiAccessibility.Name(candidates, english ? "Search results" : "Résultats de recherche", 1);
            candidates.AccessibleDescription = help.Text;
            candidates.Columns.Add(english ? "Slide" : "Diapo", 65);
            candidates.Columns.Add(english ? "Title" : "Titre", 450);
            candidates.Columns.Add("Score", 70);
            foreach (SlideChoice choice in proposal.Candidates.Take(5))
            {
                candidates.Items.Add(new ListViewItem(new[]
                {
                    choice.SlideNumber.ToString(), choice.Title ?? "", choice.Score.ToString()
                }) { Tag = choice.SlideId, ToolTipText = choice.Title ?? "" });
            }
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(8), TabIndex = 2
            };
            var cancel = new Button { AutoSize = true, DialogResult = DialogResult.Cancel, TabIndex = 1 };
            UiAccessibility.Caption(cancel, english ? "&Cancel" : "&Annuler");
            UiAccessibility.Caption(confirm, english ? "&Go to selected slide" : "&Ouvrir la sélection");
            candidates.SelectedIndexChanged += delegate { confirm.Enabled = candidates.SelectedItems.Count == 1; };
            confirm.Click += delegate
            {
                if (candidates.SelectedItems.Count != 1) { return; }
                SelectedSlideId = (int)candidates.SelectedItems[0].Tag;
                DialogResult = DialogResult.OK;
                Close();
            };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(confirm);
            CancelButton = cancel;
            // Selection alone must never turn Enter in the results list into navigation.
            AcceptButton = null;
            var body = new Panel { Dock = DockStyle.Fill, Padding = new Padding(4), TabIndex = 0 };
            UiAccessibility.ScrollableSection(body, help, candidates, 160);
            Controls.Add(body);
            buttons.Dock = DockStyle.Bottom;
            buttons.TabIndex = 1;
            Controls.Add(buttons);
            Shown += delegate { candidates.Focus(); };
            UiAccessibility.Initialize(this);
        }
    }
}
