[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$project = Split-Path -Parent $PSScriptRoot
$compiler = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (!$compiler) { throw "The Windows .NET Framework C# compiler is required." }
$output = Join-Path $PSScriptRoot ("accessibility-" + [Guid]::NewGuid().ToString("N"))
$code = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using JarvisPowerPoint;

// Production forms, isolated from Office. Any accidental presentation action fails.
namespace JarvisPowerPoint
{
    internal enum PresentationUnavailableReason { NoSlideShow, MultipleSlideShows }
    internal sealed class PresentationUnavailableException : InvalidOperationException
    {
        public PresentationUnavailableReason Reason { get; set; }
    }
    internal sealed class PresentationSession
    {
        public string ActiveRouteName { get { return null; } }
        public PresentationSnapshot Snapshot()
        {
            return new PresentationSnapshot { SessionKey = "synthetic", PresentationKey = "synthetic",
                PresentationName = "Synthetic UI test", SavedPath = @"C:\Synthetic\Accessibility.pptx",
                SlideId = 101, SlideNumber = 1, Title = "Synthetic slide" };
        }
        public List<SlideChoice> GetSlides()
        {
            return new List<SlideChoice> { new SlideChoice { SlideId = 101, SlideNumber = 1,
                Title = new string('W', 240) + " & synthetic title" } };
        }
        public List<SlideRoute> GetRoutes() { return new List<SlideRoute>(); }
        public void DeactivateRoute() { throw new Exception("Unexpected route action."); }
        public void SaveRoute(SlideRoute route) { throw new Exception("Unexpected save."); }
        public void ActivateRoute(string name) { throw new Exception("Unexpected navigation."); }
        public void RemoveRoute(string name) { throw new Exception("Unexpected deletion."); }
    }
}

internal static class AccessibilityTests
{
    private static int assertions;
    private static readonly BindingFlags Fields = BindingFlags.Instance | BindingFlags.NonPublic;

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            foreach (bool english in new[] { false, true })
                foreach (float scale in new[] { 1f, 1.5f, 2f })
                    TestForms(english, scale, args[0]);
            Console.WriteLine("Passed " + assertions + " synthetic accessibility/layout assertions (FR/EN, 100/150/200%).");
            Console.WriteLine("No Office automation, microphone enumeration/capture, settings writes, or global keystrokes.");
            Console.WriteLine("Scale/font simulation is not an OS-DPI, Narrator, contrast-theme or multi-monitor certification.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void Assert(bool value, string message)
    {
        if (!value) throw new Exception(message);
        assertions++;
    }
    private static T Field<T>(object instance, string name)
    {
        return (T)instance.GetType().GetField(name, Fields).GetValue(instance);
    }
    private static IEnumerable<Control> Descendants(Control parent)
    {
        foreach (Control child in parent.Controls)
        {
            yield return child;
            foreach (Control descendant in Descendants(child)) yield return descendant;
        }
    }
    private static void Layout(Control parent)
    {
        parent.PerformLayout();
        foreach (Control child in parent.Controls) Layout(child);
    }
    private static void Prepare(Form form, float scale)
    {
        Assert(form.AutoScaleMode == AutoScaleMode.Dpi, form.GetType().Name + ": explicit DPI mode missing.");
        form.AutoScaleMode = AutoScaleMode.None;
        var fonts = new List<Control> { form };
        fonts.AddRange(Descendants(form).Where(c => c.Parent != null && !c.Font.Equals(c.Parent.Font)));
        foreach (Control control in fonts)
            control.Font = new Font(control.Font.FontFamily, control.Font.Size * scale, control.Font.Style);
        form.Scale(new SizeF(scale, scale));
        form.ShowInTaskbar = false;
        form.Opacity = 0;
        form.Show();
        Layout(form);
        Application.DoEvents();
        Layout(form);
        Assert(form.Width <= Screen.FromControl(form).WorkingArea.Width, "Form exceeds working area width.");
        Assert(form.Height <= Screen.FromControl(form).WorkingArea.Height, "Form exceeds working area height.");
    }
    private static void CheckControls(Form form)
    {
        foreach (Control control in Descendants(form))
        {
            if (!control.Visible) continue;
            if (control is ButtonBase || control is ComboBox || control is ListControl
                || control is ListView || control is TextBox || control is NumericUpDown || control is DataGridView)
            {
                // NumericUpDown owns implementation-detail edit/buttons; the parent names it.
                if (control.Parent is NumericUpDown) continue;
                Assert(!string.IsNullOrWhiteSpace(control.AccessibleName),
                    form.GetType().Name + ": unnamed " + control.GetType().Name + " " + control.Text);
                Assert(control.TabIndex >= 0, "Invalid tab index.");
                Assert(control.Width > 0 && control.Height > 0, "Control has no usable geometry.");
            }
            if (control is Button)
            {
                Size text = TextRenderer.MeasureText(control.Text, control.Font,
                    Size.Empty, TextFormatFlags.SingleLine);
                Assert(control.ClientSize.Height >= text.Height,
                    form.GetType().Name + ": clipped button height " + control.Text);
                Assert(control.ClientSize.Width >= text.Width,
                    form.GetType().Name + ": clipped button width " + control.Text);
                var flow = control.Parent as FlowLayoutPanel;
                if (flow != null && !flow.AutoScroll)
                {
                    Assert(control.Right <= flow.ClientSize.Width - flow.Padding.Right + 1
                        && control.Bottom <= flow.ClientSize.Height - flow.Padding.Bottom + 1,
                        form.GetType().Name + ": clipped action " + control.Text + " " + control.Bounds + " in " + flow.ClientSize);
                }
            }
            if (control is Label && ((Label)control).AutoSize)
            {
                Size needed = control.GetPreferredSize(new Size(control.Width, 0));
                Assert(control.Height + 1 >= needed.Height,
                    form.GetType().Name + ": clipped wrapped label " + control.Text);
            }
            if (control is CheckBox)
            {
                Size needed = control.GetPreferredSize(new Size(control.Width, 0));
                Assert(control.Height + 1 >= needed.Height, form.GetType().Name + ": clipped checkbox " + control.Text);
            }
            if (control is Label || control is TextBox || control is ListView)
                Assert(control.ForeColor.IsSystemColor, form.GetType().Name + ": fixed text color does not follow contrast theme.");
            if (control is ListView || control is ListBox || control is DataGridView)
                Assert(control.Height >= control.Font.Height * 2,
                    form.GetType().Name + ": list/grid has no usable rows: " + control.AccessibleName
                        + " " + control.Bounds + " form " + form.ClientSize + " font " + control.Font.Height);
        }
    }
    private static void CheckTabTraversal(Form form)
    {
        var expected = Descendants(form).Where(c => c.Visible && c.Enabled && c.TabStop
            && (c is ButtonBase || c is ComboBox || c is ListBox || c is ListView
                || c is TextBox || c is NumericUpDown || c is DataGridView)
            && !(c.Parent is NumericUpDown)).ToList();
        var visited = new HashSet<Control>();
        Control current = null;
        for (int i = 0; i < expected.Count * 3 + 10; i++)
        {
            if (!form.SelectNextControl(current, true, true, true, true)) break;
            current = form.ActiveControl;
            while (current is ContainerControl && ((ContainerControl)current).ActiveControl != null)
                current = ((ContainerControl)current).ActiveControl;
            if (current == null) break;
            Control focused = current.Parent is NumericUpDown ? current.Parent : current;
            if (visited.Add(focused))
            {
                // Hidden test windows may not own OS focus. Exercise the local
                // focus notification without activating other applications or sending keys.
                typeof(Control).GetMethod("OnGotFocus", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(focused, new object[] { EventArgs.Empty });
                Application.DoEvents();
                Rectangle visible = focused.RectangleToScreen(focused.ClientRectangle);
                for (Control ancestor = focused.Parent; ancestor != null; ancestor = ancestor.Parent)
                    visible = Rectangle.Intersect(visible, ancestor.RectangleToScreen(ancestor.ClientRectangle));
                Assert(visible.Width > 0 && visible.Height >= Math.Min(focused.Font.Height, focused.ClientSize.Height),
                    form.GetType().Name + ": keyboard focus is clipped outside the viewport: " + focused.AccessibleName
                        + " bounds=" + focused.RectangleToScreen(focused.ClientRectangle) + " visible=" + visible
                        + " form=" + form.ClientSize + " focused=" + focused.Focused + " active=" + form.ContainsFocus
                        + " parents=" + Parents(focused));
            }
        }
        foreach (Control control in expected)
            Assert(visited.Contains(control), form.GetType().Name + ": Tab cannot reach " + control.AccessibleName);
    }
    private static string Parents(Control control)
    {
        string text = "";
        for (Control ancestor = control.Parent; ancestor != null; ancestor = ancestor.Parent)
        {
            var scroll = ancestor as ScrollableControl;
            text += ancestor.GetType().Name + ":" + ancestor.RectangleToScreen(ancestor.ClientRectangle)
                + (scroll == null ? "" : " scroll=" + scroll.AutoScrollPosition) + "; ";
        }
        return text;
    }
    private static void Narrow(Form form)
    {
        form.Size = form.MinimumSize;
        Layout(form);
        Application.DoEvents();
        Layout(form);
    }
    private static void TestTabs(Form form)
    {
        foreach (TabControl tabs in Descendants(form).OfType<TabControl>().ToList())
            foreach (TabPage page in tabs.TabPages)
            {
                tabs.SelectedTab = page;
                Layout(form);
                Application.DoEvents();
                CheckControls(form);
                CheckTabTraversal(form);
            }
    }
    private static void DialogKey(Form form, Keys key)
    {
        typeof(Form).GetMethod("ProcessDialogKey", BindingFlags.Instance | BindingFlags.NonPublic)
            .Invoke(form, new object[] { key });
    }
    private static void CheckPresenterRefresh(PresenterPanel panel, bool english)
    {
        Label title = Field<Label>(panel, "presentation");
        Label status = Field<Label>(panel, "status");
        Label heard = Field<Label>(panel, "heard");
        Label outcome = Field<Label>(panel, "outcome");
        ProgressBar audio = Field<ProgressBar>(panel, "audioLevel");
        ListeningIndicator indicator = Field<ListeningIndicator>(panel, "indicator");
        Label indicatorText = Field<Label>(indicator, "text");
        ProgressBar indicatorLevel = Field<ProgressBar>(indicator, "level");
        panel.SetPresentation("Synthetic refresh");
        panel.SetStatus("Paused", "Hello", "Ready", 42);
        int changes = 0;
        EventHandler onChange = delegate { changes++; };
        Label[] labels = { title, status, heard, outcome, indicatorText };
        foreach (Label label in labels) label.TextChanged += onChange;
        for (int i = 0; i < 10; i++)
        {
            panel.SetPresentation("Synthetic refresh");
            panel.SetStatus("Paused", "Hello", "Ready", 42);
        }
        Assert(changes == 0, "Identical status refresh changed visible text.");
        Assert(audio.Value == 42 && indicatorLevel.Value == 42, "Audio-only refresh lost its level.");
        Assert(heard.Text == (english ? "Heard: Hello" : "Entendu : Hello"), "Heard text lost localization.");
        Assert(outcome.Text == (english ? "Result: Ready" : "Résultat : Ready"), "Result text lost localization.");
        Assert(indicatorText.Text == "Paused\n" + heard.Text + "\n" + outcome.Text
            && indicatorText.AccessibleName == indicatorText.Text, "Indicator text and accessible name diverged.");
        panel.SetStatus("Paused", "Hello", "Ready", 200);
        Assert(audio.Value == 100 && indicatorLevel.Value == 100 && changes == 0,
            "Level-only refresh changed text or failed to clamp.");
        panel.SetStatus(null, null, null, -1);
        panel.SetPresentation(null);
        Assert(status.Text == string.Empty && title.Text == string.Empty
            && audio.Value == 0 && indicatorLevel.Value == 0, "Null status or lower audio limit mishandled.");
        int afterNull = changes;
        panel.SetStatus(null, null, null, -1);
        panel.SetPresentation(null);
        Assert(changes == afterNull, "Repeated null refresh changed text.");
        foreach (Label label in labels) label.TextChanged -= onChange;
    }
    private static void TestForms(bool english, float scale, string directory)
    {
        using (var panel = new PresenterPanel(english))
        {
            CheckPresenterRefresh(panel, english);
            panel.SetPresentation(new string('W', 180) + " & synthetic presentation");
            panel.SetStatus("Paused", new string('W', 180), new string('W', 220), 0);
            Prepare(panel, scale);
            TestTabs(panel);
            Narrow(panel);
            TestTabs(panel);
            Assert(panel.CautiousSearch && !panel.HotkeysEnabled && !panel.IndicatorEnabled, "Safe presenter defaults changed.");
            Assert(panel.AcceptButton == null, "Presenter must not default to a navigation action.");
            Assert(Field<Label>(panel, "presentation").UseMnemonic == false, "Presentation title ampersand lost.");
        }
        using (var setup = new SetupDialog(english ? "en-US" : "fr-FR", "synthetic", false))
        {
            Field<Label>(setup, "availability").Text = string.Join(" ", Enumerable.Repeat("Synthetic device unavailable.", 50));
            Field<Label>(setup, "availability").Visible = true;
            Prepare(setup, scale);
            CheckControls(setup);
            CheckTabTraversal(setup);
            Narrow(setup);
            CheckControls(setup);
            Assert(Field<object>(setup, "session") == null, "Setup implicitly opened microphone.");
            Assert(!((Control)setup.AcceptButton).Enabled, "Unverified synthetic setup can save.");
            Assert(setup.CancelButton != null, "Setup Escape missing.");
            var cancel = (Control)setup.CancelButton;
            Assert(cancel.RectangleToScreen(cancel.ClientRectangle).Bottom <= setup.RectangleToScreen(setup.ClientRectangle).Bottom,
                "Setup footer escaped viewport.");
            DialogKey(setup, Keys.Escape);
            Assert(setup.DialogResult == DialogResult.Cancel, "Setup Escape did not cancel.");
        }
        using (var preflight = new PreflightDialog(english,
            () => new PreflightContext { CultureName = english ? "en-US" : "fr-FR", MicrophoneId = "synthetic" },
            () => new PresentationSnapshot(), null, new FakeEnvironment()))
        {
            Prepare(preflight, scale);
            CheckControls(preflight);
            CheckTabTraversal(preflight);
            Narrow(preflight);
            CheckControls(preflight);
            var grid = Field<DataGridView>(preflight, "results");
            Assert(grid.Rows.Count == 6, "Synthetic preflight results missing.");
            Assert(grid.StandardTab && grid.ReadOnly, "Read-only grid must permit Tab to leave.");
            foreach (DataGridViewRow row in grid.Rows)
                Assert(row.Cells[0].Style.ForeColor.IsSystemColor && !string.IsNullOrEmpty((string)row.Cells[0].Value),
                    "Preflight severity relies on fixed color.");
            DialogKey(preflight, Keys.Escape);
            Assert(preflight.DialogResult == DialogResult.Cancel, "Preflight Escape did not cancel.");
        }
        using (var profiles = new PresentationProfilesDialog(english, new PresentationSession(),
            new PresentationProfileStore(directory)))
        {
            Prepare(profiles, scale);
            TestTabs(profiles);
            Narrow(profiles);
            TestTabs(profiles);
            Assert(Field<ListBox>(profiles, "available").HorizontalScrollbar, "Long slide titles cannot scroll.");
            Assert(profiles.AcceptButton == null && profiles.CancelButton != null, "Unsafe profile dialog defaults.");
            DialogKey(profiles, Keys.Escape);
            Assert(profiles.DialogResult == DialogResult.Cancel, "Profiles Escape did not close.");
        }
        var proposal = new SearchProposal { Candidates = Enumerable.Range(1, 5).Select(i =>
            new SlideChoice { SlideId = 100 + i, SlideNumber = i, Score = 500, Title = new string('W', 240) }).ToList() };
        using (var search = new SearchChoiceDialog(proposal, english))
        {
            Prepare(search, scale);
            CheckControls(search);
            CheckTabTraversal(search);
            Narrow(search);
            CheckControls(search);
            var list = Field<ListView>(search, "candidates");
            var confirm = Field<Button>(search, "confirm");
            Assert(search.AcceptButton == null && !confirm.Enabled && search.SelectedSlideId == 0,
                "Search default navigation is unsafe.");
            list.Items[1].Selected = true;
            list.Focus();
            Application.DoEvents();
            Assert(confirm.Enabled, "Explicit confirmation unavailable.");
            DialogKey(search, Keys.Enter);
            Assert(search.SelectedSlideId == 0 && search.DialogResult == DialogResult.None,
                "Enter in results navigated implicitly.");
            confirm.PerformClick();
            Assert(search.SelectedSlideId == 102 && search.DialogResult == DialogResult.OK, "Explicit selection failed.");
        }
        using (var search = new SearchChoiceDialog(proposal, english))
        {
            Prepare(search, scale);
            DialogKey(search, Keys.Escape);
            Assert(search.DialogResult == DialogResult.Cancel && search.SelectedSlideId == 0,
                "Search Escape navigated.");
        }
    }
    private sealed class FakeEnvironment : IPreflightEnvironment
    {
        public IList<string> GetRecognizerCultures() { return new[] { "fr-FR", "en-US" }; }
        public IList<string> GetMicrophoneIds() { return new[] { "synthetic" }; }
        public PreflightDisplays GetDisplays(IntPtr window)
        {
            return new PreflightDisplays { ScreenIds = new[] { "synthetic" } };
        }
    }
}
'@

New-Item -ItemType Directory -Path $output | Out-Null
try {
    $source = Join-Path $output "AccessibilityTests.cs"
    Set-Content -LiteralPath $source -Value $code -Encoding UTF8
    $executable = Join-Path $output "AccessibilityTests.exe"
    $arguments = @("/nologo", "/target:exe", "/out:$executable",
        "/reference:System.dll", "/reference:System.Core.dll", "/reference:System.Drawing.dll",
        "/reference:System.Windows.Forms.dll", "/reference:System.Xml.dll",
        ("/reference:" + (Join-Path (Split-Path -Parent $compiler) "WPF\System.Speech.dll")),
        ("/win32manifest:" + (Join-Path $project "app.manifest")), $source)
    foreach ($file in @("UiAccessibility.cs", "PresenterPanel.cs", "SetupDialog.cs",
        "PreflightDialog.cs", "PresentationProfilesDialog.cs", "SearchChoiceDialog.cs",
        "PresentationContracts.cs", "RehearsalTracker.cs", "PresentationProfileStore.cs",
        "AliasStore.cs", "DiagnosticTools.cs", "SpeechInput.cs")) {
        $arguments += Join-Path $project $file
    }
    & $compiler @arguments
    if ($LASTEXITCODE -ne 0) { throw "Accessibility test compilation failed." }
    & $executable $output
    if ($LASTEXITCODE -ne 0) { throw "Accessibility tests failed." }
} finally {
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Recurse -Force }
}
