[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$temporary = Join-Path ([IO.Path]::GetTempPath()) ("JarvisLifecycle-" + [Guid]::NewGuid().ToString("N"))
$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (!(Test-Path -LiteralPath $compiler)) { $compiler = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
$speech = Join-Path (Split-Path -Parent $compiler) "WPF\System.Speech.dll"
$powerPoint = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint" -Recurse -Filter Microsoft.Office.Interop.PowerPoint.dll |
    Select-Object -First 1 -ExpandProperty FullName
$office = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\office" -Recurse -Filter OFFICE.DLL |
    Select-Object -First 1 -ExpandProperty FullName
$code = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using JarvisPowerPoint;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static class LifecycleTests
{
    private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
    private static int assertions;
    private static object Field(object value, string name) { return value.GetType().GetField(name, Private).GetValue(value); }
    private static object Call(object value, string name, params object[] args)
    {
        try { return value.GetType().GetMethod(name, Private).Invoke(value, args); }
        catch (TargetInvocationException error) { throw error.InnerException; }
    }
    private static void Assert(bool value, string message)
    {
        if (!value) throw new Exception(message);
        assertions++;
    }
    private static void Release(object value)
    {
        if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
    private static IEnumerable<Control> Descendants(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (Control nested in Descendants(child)) yield return nested;
        }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        PowerPoint.Application app = null;
        PowerPoint.Presentations presentations = null;
        PowerPoint.Presentation deck = null;
        PowerPoint.Slides slides = null;
        PowerPoint.SlideShowSettings settings = null;
        PowerPoint.SlideShowWindow window = null;
        PowerPoint.SlideShowView view = null;
        JarvisApplicationContext context = null;
        try
        {
            try { app = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application"); }
            catch (COMException error)
            {
                if (error.ErrorCode != unchecked((int)0x800401E3)) throw;
                app = new PowerPoint.Application();
            }
            PowerPoint.SlideShowWindows shows = app.SlideShowWindows;
            try
            {
                if (shows.Count != 0) throw new InvalidOperationException("Close existing slide shows before running the opt-in lifecycle test.");
            }
            finally { Release(shows); }
            presentations = app.Presentations;
            deck = presentations.Add(Office.MsoTriState.msoTrue);
            slides = deck.Slides;
            int firstId = 0;
            for (int i = 1; i <= 4; i++)
            {
                PowerPoint.Slide slide = slides.Add(i, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                try
                {
                    if (i == 1) firstId = slide.SlideID;
                    PowerPoint.Shapes shapes = slide.Shapes;
                    PowerPoint.Shape title = shapes.Title;
                    PowerPoint.TextFrame frame = title.TextFrame;
                    PowerPoint.TextRange range = frame.TextRange;
                    try { range.Text = "Synthetic slide " + i; }
                    finally { Release(range); Release(frame); Release(title); Release(shapes); }
                }
                finally { Release(slide); }
            }
            string path = Path.Combine(args[0], "Lifecycle.pptx");
            deck.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
            settings = deck.SlideShowSettings;
            settings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeWindow;
            deck.Save();
            window = settings.Run();
            view = window.View;
            Thread.Sleep(400);

            Application.EnableVisualStyles();
            context = new JarvisApplicationContext();
            // Prevent startup/real microphone/settings initialization in this test process.
            EventHandler idle = (EventHandler)Delegate.CreateDelegate(typeof(EventHandler), context,
                typeof(JarvisApplicationContext).GetMethod("OnFirstIdle", Private));
            Application.Idle -= idle;
            var session = new PresentationSession(false, args[0]);
            var store = new PresentationProfileStore(args[0]);
            typeof(JarvisApplicationContext).GetField("session", Private).SetValue(context, session);
            typeof(JarvisApplicationContext).GetField("profiles", Private).SetValue(context, store);
            var panel = (PresenterPanel)Field(context, "panel");
            typeof(JarvisApplicationContext).GetField("applyingPreferences", Private).SetValue(context, true);
            panel.CautiousSearch = false;
            int progressWindows = 0;
            EventHandler cancelSearch = delegate
            {
                var search = Application.OpenForms.OfType<SearchProgressDialog>().FirstOrDefault();
                if (search != null)
                {
                    progressWindows++;
                    ((Button)search.CancelButton).PerformClick();
                }
            };
            Application.Idle += cancelSearch;
            string cancelled;
            try { cancelled = (string)Call(context, "SearchWithConfirmation", "Synthetic slide 4"); }
            finally { Application.Idle -= cancelSearch; }
            Assert(progressWindows == 1 && cancelled.Contains("annul"), "Main progress dismissal did not cancel.");
            Assert(session.Snapshot().SlideId == firstId && !session.HasReturnPoint,
                "Main cancelled search changed slide or checkpoint.");
            Call(context, "SearchWithConfirmation", "Synthetic slide 4");
            Assert(session.Snapshot().SlideNumber == 4 && session.HasReturnPoint,
                "Main incremental prepare/accept did not navigate.");
            session.Resume();
            panel.CautiousSearch = true;
            typeof(JarvisApplicationContext).GetField("applyingPreferences", Private).SetValue(context, false);
            EventHandler cancelChoice = delegate
            {
                var choice = Application.OpenForms.OfType<SearchChoiceDialog>().FirstOrDefault();
                if (choice != null) ((Button)choice.CancelButton).PerformClick();
            };
            Application.Idle += cancelChoice;
            try { cancelled = (string)Call(context, "SearchWithConfirmation", "Synthetic slide"); }
            finally { Application.Idle -= cancelChoice; }
            Assert(cancelled.Contains("annul") && session.Snapshot().SlideId == firstId && !session.HasReturnPoint,
                "Main ambiguous choice cancellation changed the presentation.");
            ((ToolStripMenuItem)Field(context, "refreshSearchItem")).PerformClick();
            Assert(((string)Field(context, "lastOutcome")).Contains("prochaine"),
                "Explicit index refresh has no visible result.");
            Assert((int)Field(context, "modalDepth") == 0 && Field(context, "activeSearch") == null,
                "Search left stale modal state behind.");
            store.SaveBudgets(path, new Dictionary<int, int> { { firstId, 7 } });
            Call(context, "StartRehearsal");
            var tracker = (RehearsalTracker)Field(context, "rehearsal");
            Assert(tracker.Entries[0].BudgetSeconds == 7, "Main startup ignored the persisted budget.");
            Thread.Sleep(30);
            session.GoTo(4);
            Call(context, "PollPresentation");
            Thread.Sleep(30);
            Call(context, "StopRehearsal");
            Assert(store.LoadRuns(path).Count == 1, "Main stop did not save history.");
            Assert(store.LoadRuns(path)[0].Slides.Count == 2, "Main report omitted a visited slide.");
            Call(context, "PersistRehearsal");
            Assert(store.LoadRuns(path).Count == 1, "Retry duplicated a completed run.");

            Call(context, "StartRehearsal");
            Thread.Sleep(30);
            view.Exit();
            Call(context, "PollPresentation");
            Assert(!tracker.IsRunning, "Main poll did not stop rehearsal when the show ended.");
            Assert(store.LoadRuns(path).Count == 2, "Main poll lost the completed report.");
            Assert(tracker.SavedPath == path, "Report persistence changed the origin path.");
            Release(view); view = null;
            Release(window); window = settings.Run();
            view = window.View;
            Thread.Sleep(300);

            using (var dialog = new PresentationProfilesDialog(false, session, store, 90))
            {
                Assert(((Control)Field(dialog, "result")).Text.Length == 0, "Profile dialog failed to load.");
                ListView budgets = (ListView)Field(dialog, "budgets");
                Assert(budgets.Items.Count == 4, "Profile dialog did not enumerate the saved deck.");
                Assert(((ComboBox)Field(dialog, "newerRun")).Items.Count == 2, "History dropdown omitted saved runs.");
                Assert(((ListView)Field(dialog, "comparison")).Items.Count > 0, "History comparison was not populated.");
                dialog.Show();
                TabControl tabs = Descendants(dialog).OfType<TabControl>().Single();
                tabs.SelectedIndex = 1;
                Application.DoEvents();
                budgets.Items[0].Selected = true;
                ((NumericUpDown)Field(dialog, "seconds")).Value = 11;
                Descendants(tabs.TabPages[1]).OfType<Button>().Single().PerformClick();
                Assert(store.LoadBudgets(path)[firstId] == 11, "Budget button did not persist its value.");
                Assert(((Control)Field(dialog, "result")).Text.Length > 0, "Budget save has no visible result.");
            }
            object recorder = Field(context, "diagnostics");
            var snapshot = (DiagnosticSnapshot)Call(context, "DiagnosticSnapshot");
            string report = ((DiagnosticRecorder)recorder).BuildReport(snapshot);
            Assert(!report.Contains(path) && !report.Contains("Synthetic slide"), "Main diagnostics leaked deck data.");
            Assert(Field(context, "recognizer") == null, "Lifecycle test unexpectedly opened speech recognition.");
            for (int i = 5; i <= 120; i++)
            {
                PowerPoint.Slide slide = slides.Add(i, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                PowerPoint.Shapes shapes = slide.Shapes;
                PowerPoint.Shape title = shapes.Title;
                PowerPoint.TextFrame frame = title.TextFrame;
                PowerPoint.TextRange range = frame.TextRange;
                try { range.Text = "Benchmark unique" + i + " " + new string('x', 1200); }
                finally { Release(range); Release(frame); Release(title); Release(shapes); Release(slide); }
            }
            deck.Save();
            session.InvalidateSearchIndex();
            var watch = Stopwatch.StartNew();
            SearchProposal coldProposal;
            using (SearchOperation operation = session.BeginSearch("unique120"))
            {
                while (!operation.Step(15)) { }
                Assert(!operation.IsWarm && operation.Proposal.Candidates.Count == 1, "Live cold index returned incorrect results.");
                coldProposal = operation.Proposal;
            }
            double coldMs = watch.Elapsed.TotalMilliseconds;
            watch.Restart();
            SearchProposal warmProposal;
            using (SearchOperation operation = session.BeginSearch("unique120"))
            {
                while (!operation.Step(15)) { }
                Assert(operation.IsWarm && operation.Proposal.Candidates[0].SlideId == coldProposal.Candidates[0].SlideId,
                    "Live warm index was not reused or changed results.");
                warmProposal = operation.Proposal;
            }
            double warmMs = watch.Elapsed.TotalMilliseconds;
            using (SearchOperation operation = session.BeginAcceptSearch(warmProposal, warmProposal.Candidates[0].SlideId))
            {
                while (!operation.Step(15)) { }
            }
            Assert(session.Snapshot().SlideNumber == 120, "Live indexed acceptance did not reach its stable target.");
            session.Resume();
            using (SearchOperation operation = session.BeginSearch("unique120", true))
            {
                operation.Step(1);
                operation.Step(1);
                Assert(!operation.IsComplete, "Synthetic large search unexpectedly completed before cancellation.");
                operation.Cancel();
            }
            Assert(session.Snapshot().SlideId == firstId && !session.HasReturnPoint,
                "Live partial extraction cancellation changed the slide/checkpoint.");
            PowerPoint.Slide edited = slides[120];
            PowerPoint.Shapes editedShapes = edited.Shapes;
            PowerPoint.Shape editedTitle = editedShapes.Title;
            PowerPoint.TextFrame editedFrame = editedTitle.TextFrame;
            PowerPoint.TextRange editedRange = editedFrame.TextRange;
            int editedId = edited.SlideID;
            try { editedRange.Text = "Changedunique120"; }
            finally { Release(editedRange); Release(editedFrame); Release(editedTitle); Release(editedShapes); Release(edited); }
            bool dirtyRejected = false;
            try
            {
                using (SearchOperation operation = session.BeginSearch("Changedunique120"))
                    while (!operation.Step(15)) { }
            }
            catch (InvalidOperationException) { dirtyRejected = true; }
            Assert(dirtyRejected && session.Snapshot().SlideId == firstId, "Live dirty search failed to stop safely.");
            deck.Save();
            using (SearchOperation operation = session.BeginSearch("Changedunique120"))
            {
                while (!operation.Step(15)) { }
                Assert(!operation.IsWarm && operation.Proposal.Candidates[0].SlideId == editedId,
                    "Saved edit reused stale text.");
            }
            Console.WriteLine("Live synthetic PowerPoint, 120 slides: cold=" + coldMs.ToString("F1")
                + "ms; warm=" + warmMs.ToString("F1") + "ms. Engine elapsed time, excludes progress-dialog timer delay.");
            Console.WriteLine("Passed " + assertions + " main lifecycle and live profile-dialog assertions.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            if (context != null) Call(context, "ExitApplication", null, EventArgs.Empty);
            if (view != null) view.Exit();
            if (deck != null) deck.Close();
            Release(view); Release(window); Release(settings); Release(slides);
            Release(deck); Release(presentations); Release(app);
            // Never quit the user's PowerPoint application.
        }
    }
}
'@
try {
    New-Item -ItemType Directory -Path $temporary | Out-Null
    $source = Join-Path $temporary "LifecycleTests.cs"
    [IO.File]::WriteAllText($source, $code, [Text.UTF8Encoding]::new($false))
    $exe = Join-Path $temporary "LifecycleTests.exe"
    $sources = @(Get-ChildItem -LiteralPath $root -Filter "*.cs" | Select-Object -ExpandProperty FullName)
    & $compiler /nologo /target:exe /langversion:5 /main:LifecycleTests "/out:$exe" `
        /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll `
        /reference:System.Windows.Forms.dll /reference:System.Xml.dll /reference:System.Runtime.Serialization.dll `
        "/reference:$speech" "/reference:$powerPoint" "/reference:$office" $sources $source
    if ($LASTEXITCODE -ne 0) { throw "Lifecycle test compilation failed." }
    & $exe $temporary
    if ($LASTEXITCODE -ne 0) { throw "Lifecycle tests failed ($LASTEXITCODE)." }
} finally {
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
