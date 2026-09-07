using System;
using System.Drawing;
using System.Windows.Forms;

namespace JarvisPowerPoint
{
    internal static class UiAccessibility
    {
        // Call after constructing authored 96-DPI geometry, before showing the form.
        // Legacy WinForms supports system DPI, not reliable per-monitor transitions.
        public static void Initialize(Form form)
        {
            form.AutoScaleDimensions = new SizeF(96, 96);
            form.AutoScaleMode = AutoScaleMode.Dpi;
            form.BackColor = SystemColors.Control;
            form.ForeColor = SystemColors.ControlText;
            TrackScrollableFocus(form);
            form.Shown += delegate { ConstrainToWorkingArea(form); };
        }

        private static void TrackScrollableFocus(Control root)
        {
            foreach (Control control in root.Controls)
            {
                if (control.TabStop)
                    control.GotFocus += delegate { RevealFocus(control); };
                TrackScrollableFocus(control);
            }
        }

        private static void RevealFocus(Control control)
        {
            // Nested scrolling panels do not consistently reveal descendants when
            // legacy WinForms advances keyboard focus. Preserve every authored tab index.
            for (Control ancestor = control.Parent; ancestor != null; ancestor = ancestor.Parent)
            {
                var scroll = ancestor as ScrollableControl;
                if (scroll != null && scroll.AutoScroll) scroll.ScrollControlIntoView(control);
            }
        }

        public static void Name(Control control, string name, int tabIndex)
        {
            control.AccessibleName = name;
            control.TabIndex = tabIndex;
        }

        public static void Caption(ButtonBase control, string caption)
        {
            control.Text = caption;
            control.AccessibleName = caption.Replace("&", "");
            control.UseVisualStyleBackColor = true;
        }

        public static void AddRow(TableLayoutPanel layout, Control control)
        {
            int row = layout.RowCount++;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            control.TabIndex = row;
            control.Dock = DockStyle.Top;
            Label label = control as Label;
            if (label != null)
            {
                label.AutoSize = true;
                label.MaximumSize = Size.Empty;
                label.UseMnemonic = false;
            }
            layout.Controls.Add(control, 0, row);
        }

        // Only the explicitly supplied label is changed: no recursive ordering or
        // rewriting of button text, user-entered content, or accessible names.
        public static void WrapLabel(Label label, Control container)
        {
            Action resize = delegate
            {
                int width = Math.Max(1, container.ClientSize.Width
                    - container.Padding.Horizontal - label.Margin.Horizontal);
                if (label.MaximumSize.Width != width)
                    label.MaximumSize = new Size(width, 0);
            };
            label.AutoSize = true;
            label.UseMnemonic = false;
            container.ClientSizeChanged += delegate { resize(); };
            resize();
        }

        public static void ScrollableSection(Control parent, Control header, Control content, int minimumBodyHeight)
        {
            var layout = new TableLayoutPanel {
                Dock = DockStyle.Fill, AutoScroll = true, ColumnCount = 1, RowCount = 2
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            header.Dock = DockStyle.Top;
            header.TabIndex = 0;
            content.Dock = DockStyle.Fill;
            content.TabIndex = 1;
            content.MinimumSize = new Size(0, minimumBodyHeight);
            layout.Controls.Add(header, 0, 0);
            layout.Controls.Add(content, 0, 1);
            layout.Layout += delegate
            {
                int height = header.Height + header.Margin.Vertical
                    + content.MinimumSize.Height + content.Margin.Vertical + layout.Padding.Vertical;
                if (layout.AutoScrollMinSize.Height != height)
                    layout.AutoScrollMinSize = new Size(0, height);
            };
            parent.Controls.Add(layout);
        }

        public static void ConstrainToWorkingArea(Form form)
        {
            Rectangle area = Screen.FromControl(form).WorkingArea;
            form.MinimumSize = new Size(Math.Min(form.MinimumSize.Width, area.Width),
                Math.Min(form.MinimumSize.Height, area.Height));
            form.Size = new Size(Math.Min(form.Width, area.Width), Math.Min(form.Height, area.Height));
            form.Location = new Point(Math.Max(area.Left, Math.Min(form.Left, area.Right - form.Width)),
                Math.Max(area.Top, Math.Min(form.Top, area.Bottom - form.Height)));
        }
    }
}
