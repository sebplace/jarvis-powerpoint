using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Forms;

namespace JarvisPowerPoint
{
    internal sealed class SearchProgressDialog : Form
    {
        private readonly Func<bool> advance;
        private readonly Func<string> describeProgress;
        private readonly Timer timer = new Timer { Interval = 15 };
        private readonly Label status;
        private Exception failure;

        public SearchProgressDialog(bool english, Func<bool> advance, Func<string> describeProgress)
        {
            if (advance == null) { throw new ArgumentNullException("advance"); }
            if (describeProgress == null) { throw new ArgumentNullException("describeProgress"); }
            this.advance = advance;
            this.describeProgress = describeProgress;
            Text = english ? "Searching the presentation" : "Recherche dans la présentation";
            ClientSize = new Size(480, 190);
            MinimumSize = new Size(380, 210);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, Padding = new Padding(12),
                ColumnCount = 1, RowCount = 3, AutoScroll = true
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            status = new Label
            {
                Dock = DockStyle.Fill, AutoSize = true,
                Text = english ? "Preparing search..." : "Préparation de la recherche...",
                AccessibleName = english ? "Search progress" : "Avancement de la recherche",
                TabStop = false
            };
            var notice = new Label
            {
                AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 8),
                Text = english ? "Cancel leaves the slideshow unchanged. This window is visible in desktop sharing."
                    : "Annuler laisse le diaporama inchangé. Cette fenêtre reste visible en partage du bureau.",
                TabStop = false
            };
            var cancel = new Button
            {
                Text = english ? "&Cancel" : "&Annuler", AutoSize = true,
                MinimumSize = new Size(100, 32), Anchor = AnchorStyles.Right,
                DialogResult = DialogResult.Cancel, TabIndex = 0,
                AccessibleName = english ? "Cancel search" : "Annuler la recherche"
            };
            CancelButton = cancel;
            layout.Controls.Add(status, 0, 0);
            layout.Controls.Add(notice, 0, 1);
            layout.Controls.Add(cancel, 0, 2);
            Controls.Add(layout);
            UiAccessibility.WrapLabel(status, layout);
            UiAccessibility.WrapLabel(notice, layout);
            timer.Tick += OnTick;
            Shown += delegate { cancel.Focus(); timer.Start(); };
            FormClosing += delegate { timer.Stop(); };
            UiAccessibility.Initialize(this);
        }

        public void ThrowIfFailed()
        {
            if (failure != null) { ExceptionDispatchInfo.Capture(failure).Throw(); }
        }

        private void OnTick(object sender, EventArgs args)
        {
            bool complete = false;
            try
            {
                complete = advance();
                status.Text = describeProgress();
            }
            catch (COMException error) { failure = error; }
            catch (InvalidOperationException error) { failure = error; }
            catch (IOException error) { failure = error; }
            catch (UnauthorizedAccessException error) { failure = error; }
            catch (SecurityException error) { failure = error; }
            catch (ArgumentException error) { failure = error; }
            catch (Win32Exception error) { failure = error; }
            if (failure != null || complete)
            {
                timer.Stop();
                DialogResult = failure == null ? DialogResult.OK : DialogResult.Cancel;
                Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { timer.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
