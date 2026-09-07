using System;
using System.Drawing;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JarvisPowerPoint
{
    internal sealed class UpdateDialog : Form
    {
        private readonly bool english;
        private readonly Func<bool> canInstall;
        private readonly Action closeApplication;
        private readonly Button check;
        private readonly Button install;
        private readonly Button close;
        private readonly Label status;
        private readonly TextBox details;
        private readonly TextBox notes;
        private CancellationTokenSource operation;
        private UpdateRelease release;
        private bool closing;
        private bool busy;

        public UpdateDialog(bool english, Func<bool> canInstall, Action closeApplication)
        {
            if (canInstall == null) throw new ArgumentNullException("canInstall");
            if (closeApplication == null) throw new ArgumentNullException("closeApplication");
            this.english = english;
            this.canInstall = canInstall;
            this.closeApplication = closeApplication;
            Text = T("Manual GitHub update", "Mise \u00e0 jour manuelle GitHub");
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(620, 480);
            Size = new Size(760, 650);
            Font = SystemFonts.MessageBoxFont;
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                AutoSize = true, MaximumSize = new Size(1000, 0),
                Text = T("Nothing is checked automatically. Only the buttons below contact GitHub.\r\n" +
                    "Public releases: sebplace/jarvis-powerpoint. No telemetry, audio, slides, or local paths are uploaded.",
                    "Aucune recherche automatique. Seuls les boutons ci-dessous contactent GitHub.\r\n" +
                    "Versions publiques : sebplace/jarvis-powerpoint. Aucun audio, diapositive, chemin local ou t\u00e9l\u00e9m\u00e9trie transmis.")
            }, 0, 0);
            status = new Label
            {
                AutoSize = true, Margin = new Padding(3, 12, 3, 8),
                Text = T("Installed version: ", "Version install\u00e9e : ") + UpdateManager.CurrentVersion.ToString(3)
            };
            layout.Controls.Add(status, 0, 1);
            details = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
                WordWrap = false, Text = T("No network request has been made.", "Aucune requ\u00eate r\u00e9seau effectu\u00e9e.")
            };
            layout.Controls.Add(details, 0, 2);
            layout.Controls.Add(new Label
            {
                AutoSize = true, Margin = new Padding(3, 10, 3, 6),
                Text = T("Release notes (unaltered plain text from GitHub; not instructions for this app):",
                    "Notes de version (texte brut GitHub inchang\u00e9 ; aucune instruction ex\u00e9cut\u00e9e) :")
            }, 0, 3);
            notes = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
                WordWrap = false, MaxLength = UpdateManager.MaximumMetadataBytes
            };
            layout.Controls.Add(notes, 0, 4);
            var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            check = new Button { AutoSize = true, Text = T("Check on GitHub", "Rechercher sur GitHub") };
            install = new Button { AutoSize = true, Text = T("Download and install...", "T\u00e9l\u00e9charger et installer..."), Enabled = false };
            close = new Button { AutoSize = true, Text = T("Close", "Fermer") };
            check.Click += CheckClicked;
            install.Click += InstallClicked;
            close.Click += delegate { Close(); };
            buttons.Controls.Add(check);
            buttons.Controls.Add(install);
            buttons.Controls.Add(close);
            layout.Controls.Add(buttons, 0, 5);
            Controls.Add(layout);
            CancelButton = close;
            if (ManagedDeployment.IsManaged)
            {
                check.Enabled = false;
                install.Enabled = false;
                details.Text = ManagedDeployment.GetMessage(english);
            }
            FormClosing += delegate
            {
                closing = true;
                if (operation != null) operation.Cancel();
            };
        }

        private string T(string en, string fr) { return english ? en : fr; }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                closing = true;
                if (operation != null) operation.Cancel();
            }
            base.Dispose(disposing);
        }

        private void SetBusy(bool value)
        {
            busy = value;
            check.Enabled = !value && !ManagedDeployment.IsManaged;
            install.Enabled = !value && release != null && !ManagedDeployment.IsManaged;
            close.Text = value ? T("Cancel / close", "Annuler / fermer") : T("Close", "Fermer");
            UseWaitCursor = value;
        }

        private async void CheckClicked(object sender, EventArgs args)
        {
            if (busy || ManagedDeployment.IsManaged) return;
            release = null;
            notes.Text = "";
            details.Text = "";
            SetBusy(true);
            status.Text = T("Checking GitHub...", "Recherche sur GitHub...");
            operation = new CancellationTokenSource(TimeSpan.FromSeconds(45));
            try
            {
                CancellationToken token = operation.Token;
                UpdateRelease found = await Task.Run(() => UpdateManager.Check(token), token);
                if (closing) return;
                release = found;
                if (found == null)
                    status.Text = T("No newer stable release is available.", "Aucune version stable plus r\u00e9cente.");
                else
                {
                    status.Text = T("Available version: ", "Version disponible : ") + found.Tag;
                    details.Text = T("Source: ", "Source : ") + found.DownloadUri.AbsoluteUri + "\r\n" +
                        T("Size: ", "Taille : ") + found.Size.ToString("N0", CultureInfo.CurrentCulture) +
                        T(" bytes", " octets") + "\r\nSHA-256: " + found.Sha256 + "\r\n" +
                        T("The download will be verified against GitHub's digest. This updater does not verify Authenticode publisher trust.",
                            "Le condensat GitHub sera v\u00e9rifi\u00e9. Cet outil ne v\u00e9rifie pas la confiance Authenticode de l'\u00e9diteur.");
                    notes.Text = found.Notes;
                }
            }
            catch (Exception error)
            {
                if (!UpdateManager.IsExpected(error)) throw;
                ShowFailure(error);
            }
            finally
            {
                operation.Dispose();
                operation = null;
                if (!closing) SetBusy(false);
            }
        }

        private async void InstallClicked(object sender, EventArgs args)
        {
            if (busy || release == null || ManagedDeployment.IsManaged) return;
            if (!canInstall())
            {
                MessageBox.Show(this, T("Stop all slide shows and rehearsal, and save any pending reports before installing.",
                    "Arr\u00eatez les diaporamas et la r\u00e9p\u00e9tition, et enregistrez les rapports en attente avant l'installation."),
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (MessageBox.Show(this,
                T("Download and install ", "T\u00e9l\u00e9charger et installer ") + release.Tag + "?\r\n\r\n" +
                T("Jarvis will close and restart. Your settings and aliases are retained.\r\n" +
                    "Unsaved reports kept only in memory will be lost when Jarvis closes.\r\n" +
                    "The executable will be checked against GitHub's SHA-256 digest and version, not publisher trust. " +
                    "A same-folder backup will be kept. No administrator elevation is requested.",
                    "Jarvis va se fermer et red\u00e9marrer. Vos param\u00e8tres et alias seront conserv\u00e9s.\r\n" +
                    "Les rapports non enregistr\u00e9s conserv\u00e9s uniquement en m\u00e9moire seront perdus \u00e0 la fermeture.\r\n" +
                    "Le condensat SHA-256 GitHub et la version seront v\u00e9rifi\u00e9s, pas la confiance de l'\u00e9diteur. " +
                    "Une sauvegarde sera gard\u00e9e dans le m\u00eame dossier. Aucune \u00e9l\u00e9vation administrateur."),
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                return;
            SetBusy(true);
            status.Text = T("Downloading and verifying...", "T\u00e9l\u00e9chargement et v\u00e9rification...");
            operation = new CancellationTokenSource(TimeSpan.FromMinutes(4));
            PreparedUpdate prepared = null;
            try
            {
                CancellationToken token = operation.Token;
                UpdateRelease selected = release;
                prepared = await Task.Run(() => UpdateManager.Prepare(selected, token), token);
                if (closing) return;
                status.Text = T("Preparing the update helper...", "Pr\u00e9paration de l'assistant...");
                await Task.Run(() => UpdateManager.StartHelperAndWait(prepared, token), token);
                if (closing) return;
                token.ThrowIfCancellationRequested();
                UpdateManager.AuthorizeAndClose(prepared, canInstall, closeApplication, token);
            }
            catch (Exception error)
            {
                if (!UpdateManager.IsExpected(error)) throw;
                ShowFailure(error);
            }
            finally
            {
                try { UpdateManager.Discard(prepared); }
                catch (Exception error)
                {
                    if (!UpdateManager.IsExpected(error)) throw;
                    ShowFailure(error);
                }
                finally
                {
                    operation.Dispose();
                    operation = null;
                    if (!closing) SetBusy(false);
                }
            }
        }

        private void ShowFailure(Exception error)
        {
            if (closing) return;
            bool cancelled = operation != null && operation.IsCancellationRequested;
            status.Text = cancelled ? T("Cancelled or timed out. Nothing was installed.", "Annul\u00e9 ou d\u00e9lai d\u00e9pass\u00e9. Rien n'a \u00e9t\u00e9 install\u00e9.") :
                T("Update could not complete. Jarvis remains open.", "Mise \u00e0 jour impossible. Jarvis reste ouvert.");
            if (!cancelled)
                MessageBox.Show(this, error.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
