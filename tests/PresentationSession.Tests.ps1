[CmdletBinding()]
param(
    [string]$Executable,
    [switch]$IncludeCom
)

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Executable)) {
    $Executable = Join-Path $projectDirectory "bin\JarvisPowerPoint.exe"
}
$Executable = [IO.Path]::GetFullPath($Executable)
if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) {
    throw "Build the application first, or pass -Executable pointing to a candidate binary: $Executable"
}
$testDirectory = Join-Path $PSScriptRoot (".pst-" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
$compiler = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
$speech = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\WPF\System.Speech.dll",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\WPF\System.Speech.dll"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
$powerPoint = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint" `
    -Recurse -Filter Microsoft.Office.Interop.PowerPoint.dll | Select-Object -First 1 -ExpandProperty FullName
$office = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\office" `
    -Recurse -Filter OFFICE.DLL | Select-Object -First 1 -ExpandProperty FullName
if (-not $compiler -or -not $speech -or -not $powerPoint -or -not $office) {
    throw "The .NET Framework compiler and installed PowerPoint interop assemblies are required."
}

$testSource = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace JarvisPowerPoint
{
    internal static class PresentationSessionTests
    {
        private static int assertions;
        private static Assembly candidateAssembly;

        [STAThread]
        public static int Main(string[] args)
        {
            try
            {
                StateTests(Path.Combine(args[0], "state"));
                AliasTests(Path.Combine(args[0], "aliases"));
                CandidateContractTests(args[1], Path.Combine(args[0], "candidate"));
                if (args.Length > 2 && args[2] == "com") ComTests(Path.Combine(args[0], "com"));
                Console.WriteLine("Passed " + assertions + " presentation session assertions.");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
            assertions++;
        }

        private static T Throws<T>(Action action) where T : Exception
        {
            try { action(); }
            catch (T exception) { assertions++; return exception; }
            throw new Exception("Expected " + typeof(T).Name);
        }

        private static PresentationSession Session(FakeShow show, string directory)
        {
            return new PresentationSession(true, directory, () =>
            {
                if (show.Unavailable) throw new PresentationUnavailableException("No show.");
                if (show.OpenFailure) throw new COMException("Connection failed.");
                return show;
            });
        }

        private static void StateTests(string directory)
        {
            FakeShow show = new FakeShow();
            PresentationSession session = Session(show, directory);
            PresentationSnapshot snapshot = session.Snapshot();
            Assert(snapshot.SlideId == 101 && snapshot.SlideNumber == 1 && snapshot.SlideCount == 3, "Snapshot fields.");
            Assert(snapshot.SessionKey == session.Snapshot().SessionKey, "Unchanged show identity must be stable.");
            Assert(!session.HasReturnPoint && !session.QuestionsMode, "Initial state.");
            Throws<InvalidOperationException>(() => session.Resume());
            Throws<InvalidOperationException>(() => session.NextResult());
            Throws<InvalidOperationException>(() => session.GoTo(0));
            Throws<InvalidOperationException>(() => session.GoTo(4));
            Throws<InvalidOperationException>(() => session.Search("   "));
            Throws<InvalidOperationException>(() => session.Search("???"));
            Throws<InvalidOperationException>(() => session.Search(new string('a', 501)));
            Throws<InvalidOperationException>(() => session.Search("none"));
            Assert(!session.HasReturnPoint && show.Current == 101, "Failed searches cannot move or capture.");

            show.FailNavigation = true;
            Throws<COMException>(() => session.GoTo(2));
            Assert(!session.HasReturnPoint, "Failed navigation cannot capture a checkpoint.");
            show.FailNavigation = false;
            session.GoTo(1);
            Assert(!session.HasReturnPoint, "Navigating to the current slide is not an excursion.");
            show.Click = 2;
            session.Search("matches");
            Assert(show.Current == 202 && session.HasReturnPoint, "Search should capture original slide.");
            session.BeginQuestions();
            session.BeginQuestions();
            session.GoTo(3);
            Assert(session.QuestionsMode, "Nested navigation preserves questions mode.");
            Throws<InvalidOperationException>(() => session.Search("none"));
            session.NextResult();
            Assert(show.Current == 303, "A failed search must preserve the previous ranked result cursor.");
            session.NextResult();
            Assert(show.Current == 202, "Results must wrap.");

            show.Order = new List<int> { 303, 101, 202 };
            session.NextResult();
            Assert(show.Current == 303, "Ranking remains stable after reordering.");
            session.Resume();
            Assert(show.Current == 101 && show.Click == 2, "Resume uses slide ID after reorder and restores animation.");
            Assert(!session.HasReturnPoint && !session.QuestionsMode, "Successful resume clears the return point and mode.");
            Throws<InvalidOperationException>(() => session.NextResult());
            show.Current = 202;
            session.Search("matches");
            Assert(!session.HasReturnPoint, "Top result on current slide does not create a checkpoint.");
            session.NextResult();
            Assert(session.HasReturnPoint, "Cycling away from same-slide search captures a return point.");
            show.Order.Remove(202);
            Throws<InvalidOperationException>(() => session.Resume());
            Assert(session.HasReturnPoint, "Deleted return point must not clear history on failure.");
            session.NextResult();
            Assert(show.Current == 303, "Deleted ranked results are skipped.");
            show.Order.Remove(303);
            Throws<InvalidOperationException>(() => session.NextResult());

            show.NewShow("restart");
            string newKey = session.Snapshot().SessionKey;
            Assert(newKey != snapshot.SessionKey && !session.HasReturnPoint, "Restart resets checkpoint.");
            Throws<InvalidOperationException>(() => session.NextResult());
            session.BeginQuestions();
            show.Key = "different-deck";
            session.Snapshot();
            Assert(!session.HasReturnPoint && !session.QuestionsMode, "Deck changes reset state.");
            show.Path = Path.Combine(directory, "before", "Same name.pptx");
            PresentationSnapshot beforeSaveAs = session.Snapshot();
            session.Search("matches");
            session.BeginQuestions();
            show.Path = Path.Combine(directory, "after", "Same name.pptx");
            PresentationSnapshot afterSaveAs = session.Snapshot();
            Assert(afterSaveAs.PresentationKey != beforeSaveAs.PresentationKey &&
                afterSaveAs.SessionKey != beforeSaveAs.SessionKey, "Same-name Save As updates deck and session keys.");
            Assert(!session.HasReturnPoint && !session.QuestionsMode, "Same-name Save As clears checkpoint and questions.");
            Throws<InvalidOperationException>(() => session.NextResult());
            show.Current = 101;
            session.BeginQuestions();
            show.Time = 10;
            session.Snapshot();
            show.Time = 0;
            session.Snapshot();
            Assert(!session.HasReturnPoint, "Elapsed-time restart fallback resets state.");
            session.BeginQuestions();
            show.Unavailable = true;
            Throws<PresentationUnavailableException>(() => session.Snapshot());
            Assert(!session.HasReturnPoint && !session.QuestionsMode, "Ending the show clears state.");
            show.Unavailable = false;
            Assert(session.Snapshot().SessionKey != newKey, "A resumed connection has a fresh session identity.");
            show.OpenFailure = true;
            Throws<COMException>(() => session.Snapshot());
            show.OpenFailure = false;

            show.State = PowerPoint.PpSlideShowState.ppSlideShowPaused;
            session.SetBlackScreen(true);
            session.SetBlackScreen(true);
            Assert(session.Snapshot().IsBlack, "Black state snapshot.");
            session.SetBlackScreen(false);
            Assert(show.State == PowerPoint.PpSlideShowState.ppSlideShowPaused, "Unblank must restore paused state.");
            show.State = PowerPoint.PpSlideShowState.ppSlideShowRunning;
            session.SetBlackScreen(true);
            session.SetBlackScreen(false);
            Assert(show.State == PowerPoint.PpSlideShowState.ppSlideShowRunning, "Unblank restores running state.");
            show.State = PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen;
            session.SetBlackScreen(false);
            Assert(show.State == PowerPoint.PpSlideShowState.ppSlideShowRunning, "Unblank also clears a white screen.");
            show.State = PowerPoint.PpSlideShowState.ppSlideShowPaused;
            session.SetBlackScreen(true);
            show.State = PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen;
            session.SetBlackScreen(false);
            Assert(show.State == PowerPoint.PpSlideShowState.ppSlideShowPaused, "White screen restores remembered paused state.");
            show.State = PowerPoint.PpSlideShowState.ppSlideShowPaused;
            session.BeginQuestions();
            session.GoTo(3);
            show.FailNavigation = true;
            Throws<COMException>(() => session.Resume());
            Assert(session.HasReturnPoint && session.QuestionsMode, "Failed resume preserves all history.");
            show.FailNavigation = false;
            show.FailRestore = true;
            Throws<COMException>(() => session.Resume());
            Assert(session.HasReturnPoint, "Failed animation COM call preserves checkpoint.");
            show.FailRestore = false;
            session.Resume();
            Assert(show.State == PowerPoint.PpSlideShowState.ppSlideShowPaused, "Resume restores paused state.");
            session.BeginQuestions();
            show.AnimationSupported = false;
            Assert(session.Resume().Contains("Animation"), "Unsupported animation should report best-effort restoration.");
            Assert(!session.HasReturnPoint, "Unsupported animation does not prevent restoring the slide.");
            show.AnimationSupported = true;
            session.English = false;
            Assert(session.GoTo(2).StartsWith("Diapositive"), "French success.");
            Assert(Throws<InvalidOperationException>(() => session.Search("???")).Message.Contains("Saisissez"), "French validation.");
            session.Resume();
            session.Next();
            session.Previous();
            Assert(!session.HasReturnPoint, "Normal next/previous do not create return points.");
            int disposed = show.Disposals;
            Throws<InvalidOperationException>(() => session.GoTo(999));
            Assert(show.Disposals == disposed + 1, "Connection is disposed on validation failure.");
        }

        private static void AliasTests(string directory)
        {
            FakeShow show = new FakeShow();
            string deckOne = Path.Combine(directory, "one", "Deck.pptx");
            string deckTwo = Path.Combine(directory, "two", "Deck.pptx");
            show.Path = deckOne;
            PresentationSession session = Session(show, directory);
            Assert(session.GetAliases().Count == 0, "Missing store starts empty.");
            Throws<InvalidOperationException>(() => session.SaveAlias(""));
            Throws<InvalidOperationException>(() => session.SaveAlias("?! \u2014"));
            Throws<InvalidOperationException>(() => session.SaveAlias(new string('a', 81)));
            Throws<InvalidOperationException>(() => session.SaveAlias("bad\nname"));
            Throws<InvalidOperationException>(() => session.SaveAlias("bad\uD800"));
            Assert(!Directory.Exists(Path.Combine(directory, "Aliases")), "Invalid alias must not write files.");
            session.SaveAlias("  R\u00e9sum\u00e9 & <Q&A>  ");
            List<SlideAlias> aliases = session.GetAliases();
            Assert(aliases.Count == 1 && aliases[0].Name == "R\u00e9sum\u00e9 & <Q&A>" && aliases[0].SlideId == 101, "Trimmed XML-safe alias.");
            show.Current = 202;
            session.SaveAlias("r\u00e9sum\u00e9 & <q&a>");
            aliases = session.GetAliases();
            Assert(aliases.Count == 1 && aliases[0].SlideId == 202, "Case-insensitive name updates target.");
            Assert(!session.HasReturnPoint, "Saving an alias is not navigation.");
            session = Session(show, directory);
            Assert(session.GetAliases()[0].SlideId == 202, "Alias survives app restart.");
            show.Current = 101;
            show.Order = new List<int> { 202, 303, 101 };
            session.GoToAlias("R\u00c9SUM\u00c9 & <Q&A>");
            Assert(show.Current == 202 && session.HasReturnPoint, "Alias uses stable slide ID after reorder.");
            Assert(session.GetAliases()[0].SlideNumber == 1, "Aliases resolve current slide number.");
            session.Resume();
            Assert(show.Current == 101, "Alias resume returns original.");
            show.Path = deckTwo;
            Assert(session.GetAliases().Count == 0, "Same filename in another folder is isolated.");
            session.SaveAlias("Other deck");
            show.Path = deckOne;
            Assert(session.GetAliases().Count == 1 && session.GetAliases()[0].Name.Contains("r\u00e9sum\u00e9"), "Exact path binding.");
            show.Order.Remove(202);
            aliases = session.GetAliases();
            Assert(aliases.Count == 1 && aliases[0].SlideNumber == 0, "Deleted targets remain visible for removal.");
            Throws<InvalidOperationException>(() => session.GoToAlias("r\u00e9sum\u00e9 & <q&a>"));
            Assert(!session.HasReturnPoint, "Missing alias target does not capture.");
            session.RemoveAlias("r\u00e9sum\u00e9 & <q&a>");
            Assert(session.GetAliases().Count == 0, "Deleted-target aliases can be removed.");
            Throws<InvalidOperationException>(() => session.RemoveAlias("missing"));
            Throws<InvalidOperationException>(() => session.GoToAlias("missing"));
            show.Path = null;
            Assert(Throws<InvalidOperationException>(() => session.SaveAlias("unsaved")).Message.Contains("Save"), "Unsaved decks need explanation.");
            Throws<InvalidOperationException>(() => session.GetAliases());
            Throws<InvalidOperationException>(() => session.GoToAlias("missing"));
            Throws<InvalidOperationException>(() => session.RemoveAlias("missing"));
            show.Path = deckOne;
            session.SaveAlias("test");
            string file = Path.Combine(directory, "Aliases", AliasStore.PathKey(deckOne) + ".xml");
            Assert(Path.GetFileNameWithoutExtension(file).Length == 64, "Per-path SHA256 filename.");
            Assert(AliasStore.PathKey(deckOne) == AliasStore.PathKey(deckOne.ToUpperInvariant()), "Windows path casing is normalized.");
            Assert(AliasStore.PathKey("https://example.com/sites/Deck.pptx").Length == 64, "Cloud-saved deck paths are supported.");
            Assert(!Directory.Exists(Path.GetDirectoryName(deckOne)), "The deck directory is never created or modified.");
            string xml = File.ReadAllText(file);
            File.WriteAllText(file, xml.Replace("presentation=\"", "presentation=\"wrong"));
            Throws<InvalidDataException>(() => session.GetAliases());
            File.WriteAllText(file, "<!DOCTYPE slideAliases [<!ENTITY ext SYSTEM \"file:///not-read\">]><slideAliases>&ext;</slideAliases>");
            Throws<XmlException>(() => session.GetAliases());
            File.WriteAllText(file, xml);
            Assert(session.GetAliases().Count == 1, "Store remains readable after atomic replacements.");
            Assert(Directory.GetFiles(Path.Combine(directory, "Aliases"), "*.new").Length == 0, "No pending write artifacts.");

            string blocked = Path.Combine(directory, "blocked");
            Directory.CreateDirectory(blocked);
            File.WriteAllText(Path.Combine(blocked, "Aliases"), "blocked");
            PresentationSession denied = Session(show, blocked);
            Throws<IOException>(() => denied.SaveAlias("cannot write"));
            Assert(File.ReadAllText(Path.Combine(blocked, "Aliases")) == "blocked", "IO failure does not overwrite unrelated files.");
        }

        private static void CandidateContractTests(string executable, string directory)
        {
            candidateAssembly = Assembly.LoadFrom(executable);
            Type sessionType = candidateAssembly.GetType("JarvisPowerPoint.PresentationSession", true);
            Assert(sessionType.GetConstructor(new[] { typeof(bool), typeof(string) }) != null, "Candidate constructor.");
            foreach (string method in new[] { "Snapshot", "Next", "Previous", "NextResult", "BeginQuestions", "Resume", "GetAliases" })
                Assert(sessionType.GetMethod(method, Type.EmptyTypes) != null, "Candidate API: " + method);
            foreach (string method in new[] { "Search", "GoToAlias", "SaveAlias", "RemoveAlias" })
                Assert(sessionType.GetMethod(method, new[] { typeof(string) }) != null, "Candidate API: " + method);
            Assert(sessionType.GetMethod("GoTo", new[] { typeof(int) }) != null, "Candidate GoTo API.");
            Assert(sessionType.GetMethod("SetBlackScreen", new[] { typeof(bool) }) != null, "Candidate black screen API.");
            Assert(sessionType.GetProperty("English").CanWrite, "Candidate language is mutable.");
            CandidateSession session = new CandidateSession(directory);
            Assert(!session.HasReturnPoint && !session.QuestionsMode, "Candidate initial state.");
            Throws<InvalidOperationException>(() => session.Search("???"));
            Throws<InvalidOperationException>(() => session.SaveAlias("?! \u2014"));

            Type storeType = candidateAssembly.GetType("JarvisPowerPoint.AliasStore", true);
            Type aliasType = candidateAssembly.GetType("JarvisPowerPoint.SlideAlias", true);
            object store = Activator.CreateInstance(storeType, new object[] { directory });
            var aliases = (System.Collections.IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(aliasType));
            object alias = Activator.CreateInstance(aliasType, true);
            aliasType.GetProperty("Name").SetValue(alias, "Candidate & alias", null);
            aliasType.GetProperty("SlideId").SetValue(alias, 987, null);
            aliasType.GetProperty("Title").SetValue(alias, "Synthetic", null);
            aliases.Add(alias);
            string path = Path.Combine(directory, "Synthetic.pptx");
            storeType.GetMethod("Save").Invoke(store, new[] { (object)path, aliases });
            store = Activator.CreateInstance(storeType, new object[] { directory });
            var restored = (System.Collections.IList)storeType.GetMethod("Load").Invoke(store, new object[] { path });
            Assert(restored.Count == 1, "Candidate aliases persist.");
            Assert((int)aliasType.GetProperty("SlideId").GetValue(restored[0], null) == 987, "Candidate alias stable ID.");
            Assert((string)aliasType.GetProperty("Name").GetValue(restored[0], null) == "Candidate & alias", "Candidate XML escaping.");
        }

        private static void ComTests(string directory)
        {
            Directory.CreateDirectory(directory);
            PowerPoint.Application application = null;
            PowerPoint.Presentations presentations = null;
            PowerPoint.Presentation deck = null;
            PowerPoint.SlideShowWindows windows = null;
            PowerPoint.Slides slides = null;
            PowerPoint.SlideShowSettings settings = null;
            PowerPoint.SlideShowWindow window = null;
            PowerPoint.SlideShowView view = null;
            try
            {
                try { application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application"); }
                catch (COMException exception)
                {
                    if (exception.ErrorCode != unchecked((int)0x800401E3)) throw;
                    application = new PowerPoint.Application();
                }
                windows = application.SlideShowWindows;
                if (windows.Count != 0)
                    throw new InvalidOperationException("COM tests refuse to run while another slide show is active.");
                presentations = application.Presentations;
                deck = presentations.Add(Office.MsoTriState.msoTrue);
                slides = deck.Slides;
                AddSlide(slides, "Opening");
                AddSlide(slides, "Budget review");
                AddSlide(slides, "Budget details");
                string deckName = "Synthetic " + Guid.NewGuid().ToString("N") + ".pptx";
                string deckPath = Path.Combine(directory, deckName);
                deck.SaveAs(deckPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
                settings = deck.SlideShowSettings;
                settings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeWindow;
                window = settings.Run();
                view = window.View;
                Release(view); view = null;
                Release(window); window = null;
                Release(settings); settings = null;
                Release(slides); slides = null;
                Release(deck); deck = null;
                Release(presentations); presentations = null;
                Release(windows); windows = null;
                // Keep only the automation-created application alive; no show/deck RCWs remain.
                CandidateSession session = new CandidateSession(directory);
                PresentationSnapshot first = null;
                try
                {
                    first = session.Snapshot();
                    for (int index = 0; index < 5; index++)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        PresentationSnapshot current = session.Snapshot();
                        if (current.PresentationKey == first.PresentationKey)
                            Assert(current.SessionKey == first.SessionKey, "COM identity stable across RCW releases.");
                        else
                        {
                            Assert(current.SessionKey != first.SessionKey, "OneDrive path changes must invalidate session identity.");
                            first = current;
                        }
                    }
                }
                finally
                {
                    presentations = application.Presentations;
                    for (int index = 1; index <= presentations.Count; index++)
                    {
                        PowerPoint.Presentation candidate = presentations[index];
                        if (string.Equals(candidate.Name, deckName, StringComparison.Ordinal))
                        {
                            deck = candidate;
                            break;
                        }
                        Release(candidate);
                    }
                    Assert(deck != null, "Locate only the synthetic deck after releasing all COM references.");
                    slides = deck.Slides;
                    settings = deck.SlideShowSettings;
                    windows = application.SlideShowWindows;
                    if (windows.Count == 1)
                    {
                        window = windows[1];
                        view = window.View;
                    }
                }
                Assert(session.Snapshot().SessionKey == first.SessionKey, "Identity survives independent reacquisition.");
                session.Search("Budget");
                Assert(session.Snapshot().SlideNumber == 2 && session.HasReturnPoint, "Real COM search.");
                session.NextResult();
                Assert(session.Snapshot().SlideNumber == 3, "Real COM next result.");
                session.SaveAlias("Budget alias");
                session.Resume();
                Assert(session.Snapshot().SlideId == first.SlideId, "Real COM resume.");
                Throws<InvalidOperationException>(() => session.NextResult());
                session.GoToAlias("Budget alias");
                Assert(session.Snapshot().SlideNumber == 3, "Real COM alias navigation.");
                session.SetBlackScreen(true);
                Assert(session.Snapshot().IsBlack, "Real COM black screen.");
                session.SetBlackScreen(false);
                Assert(!session.Snapshot().IsBlack, "Real COM unblank.");
                view.State = PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen;
                session.SetBlackScreen(false);
                Assert(view.State == PowerPoint.PpSlideShowState.ppSlideShowRunning, "Real COM unblank clears white screen.");
                session.Resume();
                PowerPoint.Slide reorder = slides[1];
                try { reorder.MoveTo(3); }
                finally { Release(reorder); }
                session.GoTo(1);
                session.Resume();
                Assert(session.Snapshot().SlideId == first.SlideId, "Real COM reorder checkpoint.");
                session.BeginQuestions();
                view.Exit();
                Release(view); view = null;
                Release(window); window = null;
                Throws<PresentationUnavailableException>(() => session.Snapshot());
                Assert(!session.HasReturnPoint, "Real COM show end resets.");
                window = settings.Run();
                view = window.View;
                Assert(session.Snapshot().SessionKey != first.SessionKey, "Real COM restart identity.");
                session.BeginQuestions();
                string beforeRestart = session.Snapshot().SessionKey;
                view.Exit();
                Release(view); view = null;
                Release(window); window = null;
                window = settings.Run();
                view = window.View;
                Assert(session.Snapshot().SessionKey != beforeRestart && !session.HasReturnPoint,
                    "Real COM unobserved restart resets history.");
                view.Exit();
                Release(view); view = null;
                Release(window); window = null;
                Release(settings); settings = null;
                Release(slides); slides = null;
                deck.Saved = Office.MsoTriState.msoTrue;
                deck.Close();
                Release(deck); deck = null;
                deck = presentations.Add(Office.MsoTriState.msoTrue);
                slides = deck.Slides;
                AddSlide(slides, "Unsaved opening");
                AddSlide(slides, "Unsaved details");
                settings = deck.SlideShowSettings;
                settings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeWindow;
                window = settings.Run();
                view = window.View;
                PresentationSnapshot unsaved = session.Snapshot();
                Assert(unsaved.PresentationKey.StartsWith("unsaved:"), "Unsaved decks have runtime identities.");
                Assert(session.Snapshot().SessionKey == unsaved.SessionKey, "Unsaved show identity is stable.");
                session.BeginQuestions();
                Throws<InvalidOperationException>(() => session.SaveAlias("save first"));
                AmbiguousShowTest(presentations, session);
                Assert(session.Snapshot().SlideId == unsaved.SlideId, "Ambiguous commands leave the first show untouched.");
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("COM integration test: " + exception);
                throw;
            }
            finally
            {
                // Only our synthetic deck/show is touched; never quit a shared PowerPoint application.
                if (view != null) { try { view.Exit(); } finally { Release(view); } }
                Release(window);
                Release(settings);
                Release(slides);
                if (deck != null)
                {
                    try { deck.Saved = Office.MsoTriState.msoTrue; deck.Close(); }
                    catch (COMException exception)
                    {
                        if (exception.ErrorCode != unchecked((int)0x80048010)) throw;
                    }
                    finally { Release(deck); }
                }
                Release(presentations);
                Release(windows);
                Release(application);
            }
        }

        private static void AmbiguousShowTest(PowerPoint.Presentations presentations, CandidateSession session)
        {
            PowerPoint.Presentation other = null;
            PowerPoint.Slides slides = null;
            PowerPoint.SlideShowSettings settings = null;
            PowerPoint.SlideShowWindow window = null;
            PowerPoint.SlideShowView view = null;
            try
            {
                other = presentations.Add(Office.MsoTriState.msoTrue);
                slides = other.Slides;
                AddSlide(slides, "Second synthetic show");
                settings = other.SlideShowSettings;
                settings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeWindow;
                window = settings.Run();
                view = window.View;
                Assert(Throws<PresentationUnavailableException>(() => session.Snapshot()).Message.Contains("Several"),
                    "Two live shows must fail safely with an explanation.");
                Assert(!session.HasReturnPoint && !session.QuestionsMode, "Ambiguity clears potentially stale history.");
                Throws<PresentationUnavailableException>(() => session.GoTo(2));
            }
            finally
            {
                if (view != null) { try { view.Exit(); } finally { Release(view); } }
                Release(window);
                Release(settings);
                Release(slides);
                if (other != null)
                {
                    try { other.Saved = Office.MsoTriState.msoTrue; other.Close(); }
                    finally { Release(other); }
                }
            }
        }

        private static void AddSlide(PowerPoint.Slides slides, string title)
        {
            PowerPoint.Slide slide = null;
            PowerPoint.Shapes shapes = null;
            PowerPoint.Shape shape = null;
            PowerPoint.TextFrame frame = null;
            PowerPoint.TextRange text = null;
            try
            {
                slide = slides.Add(slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                shapes = slide.Shapes;
                shape = shapes.Title;
                frame = shape.TextFrame;
                text = frame.TextRange;
                text.Text = title;
            }
            finally
            {
                Release(text); Release(frame); Release(shape); Release(shapes); Release(slide);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        // Fake-connection invariants run against the source build; binary API/persistence
        // checks and optional live COM tests run against the explicitly selected executable.
        private sealed class CandidateSession
        {
            private readonly Type type;
            private readonly object instance;

            public CandidateSession(string directory)
            {
                type = candidateAssembly.GetType("JarvisPowerPoint.PresentationSession", true);
                instance = Activator.CreateInstance(type, new object[] { true, directory });
            }
            public bool HasReturnPoint { get { return (bool)type.GetProperty("HasReturnPoint").GetValue(instance, null); } }
            public bool QuestionsMode { get { return (bool)type.GetProperty("QuestionsMode").GetValue(instance, null); } }
            public void Search(string query) { Call("Search", query); }
            public void GoTo(int number) { Call("GoTo", number); }
            public void GoToAlias(string name) { Call("GoToAlias", name); }
            public void SaveAlias(string name) { Call("SaveAlias", name); }
            public void NextResult() { Call("NextResult"); }
            public void BeginQuestions() { Call("BeginQuestions"); }
            public void Resume() { Call("Resume"); }
            public void SetBlackScreen(bool black) { Call("SetBlackScreen", black); }
            public PresentationSnapshot Snapshot()
            {
                object result = Call("Snapshot");
                var snapshot = new PresentationSnapshot();
                foreach (PropertyInfo property in typeof(PresentationSnapshot).GetProperties())
                    property.SetValue(snapshot, result.GetType().GetProperty(property.Name).GetValue(result, null), null);
                return snapshot;
            }
            private object Call(string method, params object[] arguments)
            {
                try { return type.GetMethod(method).Invoke(instance, arguments); }
                catch (TargetInvocationException exception)
                {
                    if (exception.InnerException.GetType().FullName == "JarvisPowerPoint.PresentationUnavailableException")
                        throw new PresentationUnavailableException(exception.InnerException.Message);
                    throw exception.InnerException;
                }
            }
        }

        private sealed class FakeShow : PresentationConnection
        {
            public List<int> Order = new List<int> { 101, 202, 303 };
            public string RuntimeId = "show-one";
            public string Key = "deck-one";
            public string Path;
            public int Current = 101;
            public int Click;
            public float Time;
            public bool Unavailable, OpenFailure, FailNavigation, FailRestore;
            public bool AnimationSupported = true;
            public int Disposals;
            public override string Identity { get { return RuntimeId + Key; } }
            public override string PresentationKey { get { return Key + Path; } }
            public override string PresentationName { get { return "Synthetic"; } }
            public override string SavedPath { get { return Path; } }
            public override int SlideCount { get { return Order.Count; } }
            public override float Elapsed { get { return Time; } }
            public override int? ClickIndex { get { return AnimationSupported ? Click : (int?)null; } }
            public override PowerPoint.PpSlideShowState State { get; set; }
            public override PresentationSlide CurrentSlide() { return FindSlide(Current); }
            public override PresentationSlide SlideAt(int number) { return FindSlide(Order[number - 1]); }
            public override PresentationSlide FindSlide(int id)
            {
                int index = Order.IndexOf(id);
                return index < 0 ? null : new PresentationSlide { Id = id, Number = index + 1, Title = "Title " + id };
            }
            public override List<PresentationSlide> Search(string query)
            {
                var found = new List<PresentationSlide>();
                if (query == "matches")
                {
                    if (FindSlide(202) != null) found.Add(FindSlide(202));
                    if (FindSlide(303) != null) found.Add(FindSlide(303));
                }
                return found;
            }
            public override void GoTo(int id)
            {
                if (FailNavigation) throw new COMException("Navigation failed.");
                if (FindSlide(id) == null) throw new InvalidOperationException("Deleted.");
                Current = id;
                Click = 0;
            }
            public override void Move(bool next)
            {
                int index = Order.IndexOf(Current) + (next ? 1 : -1);
                if (index >= 0 && index < Order.Count) Current = Order[index];
            }
            public override bool RestoreClick(int? index)
            {
                if (FailRestore) throw new COMException("Animation failed.");
                if (!AnimationSupported || !index.HasValue) return false;
                Click = index.Value;
                return true;
            }
            public override void Dispose() { Disposals++; }
            public void NewShow(string id)
            {
                RuntimeId = id;
                Order = new List<int> { 101, 202, 303 };
                Current = 101;
                Time = 0;
            }
        }
    }
}
'@

try {
    New-Item -ItemType Directory -Path $testDirectory | Out-Null
    $testSourcePath = Join-Path $testDirectory "PresentationSessionTests.cs"
    $testExecutable = Join-Path $testDirectory "PresentationSessionTests.exe"
    [IO.File]::WriteAllText($testSourcePath, $testSource, [Text.UTF8Encoding]::new($false))
    $sources = @(Get-ChildItem $projectDirectory -Filter *.cs | Select-Object -ExpandProperty FullName)
    & $compiler /nologo /target:exe /langversion:5 /main:JarvisPowerPoint.PresentationSessionTests `
        "/out:$testExecutable" /reference:System.dll /reference:System.Core.dll `
        /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Xml.dll `
        "/reference:$speech" "/reference:$powerPoint" "/reference:$office" $sources $testSourcePath
    if ($LASTEXITCODE -ne 0) { throw "Session test compilation failed ($LASTEXITCODE)." }
    $arguments = @($testDirectory, $Executable)
    if ($IncludeCom) { $arguments += "com" }
    & $testExecutable @arguments
    if ($LASTEXITCODE -ne 0) { throw "Session tests failed ($LASTEXITCODE)." }
} finally {
    if (Test-Path -LiteralPath $testDirectory) {
        Remove-Item -LiteralPath $testDirectory -Recurse -Force
    }
}
