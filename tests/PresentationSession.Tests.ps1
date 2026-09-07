[CmdletBinding()]
param(
    [string]$Executable,
    [switch]$IncludeCom,
    [switch]$SourceOnly,
    [switch]$SearchBenchmark
)

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Executable)) {
    $Executable = Join-Path $projectDirectory "bin\JarvisPowerPoint.exe"
}
$Executable = [IO.Path]::GetFullPath($Executable)
if (-not $SourceOnly -and -not (Test-Path -LiteralPath $Executable -PathType Leaf)) {
    throw "Build the application first, or pass -Executable pointing to a candidate binary: $Executable"
}
$testDirectory = Join-Path $projectDirectory ("bin\search15\tests-" + [Guid]::NewGuid().ToString("N"))
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
using System.Diagnostics;
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
                ProposalTests(Path.Combine(args[0], "proposals"));
                IndexedSearchTests(Path.Combine(args[0], "indexed"));
                SearchLifetimeTests(Path.Combine(args[0], "search-lifetime"));
                RouteTests(Path.Combine(args[0], "routes"));
                CandidateContractTests(args[1], Path.Combine(args[0], "candidate"));
                if (args.Length > 2 && args[2] == "com") ComTests(Path.Combine(args[0], "com"));
                if (args.Length > 2 && args[2] == "benchmark") SearchBenchmark(Path.Combine(args[0], "benchmark"));
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
            Assert(Throws<PresentationUnavailableException>(() => session.Snapshot()).Reason ==
                PresentationUnavailableReason.NoSlideShow, "Missing show has a stable preflight reason.");
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

        private static void ProposalTests(string directory)
        {
            var show = new FakeShow();
            PresentationSession session = Session(show, directory);
            PresentationSnapshot first = session.Snapshot();
            Assert(first.WindowHandle == IntPtr.Zero && first.SavedPath == null, "Default fake snapshot metadata.");
            SearchProposal proposal = session.PrepareSearch("matches");
            Assert(proposal.SessionKey == first.SessionKey && proposal.OriginSlideId == 101, "Proposal captures origin.");
            Assert(proposal.IsAmbiguous && proposal.Candidates[0].Score == 100, "Real ranked scores determine ambiguity.");
            Assert(!session.HasReturnPoint && show.Current == 101, "Preparing never moves or captures.");
            Throws<InvalidOperationException>(() => session.NextResult());
            show.SecondScore = 84;
            Assert(!session.PrepareSearch("matches").IsAmbiguous, "More than 15 percent behind is not ambiguous.");
            show.SecondScore = 85;
            Assert(session.PrepareSearch("matches").IsAmbiguous, "The 15 percent ambiguity threshold is inclusive.");
            Throws<InvalidOperationException>(() => session.PrepareSearch("none"));
            Throws<InvalidOperationException>(() => session.PrepareSearch("bad\uD800"));
            Assert(!session.HasReturnPoint && show.Current == 101, "No matches and cancelled proposals leave history unchanged.");
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 101));
            proposal.Candidates.Add(new SlideChoice { SlideId = 101 });
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 101));
            Assert(!session.HasReturnPoint, "A caller cannot insert a candidate that was not ranked.");
            show.Current = 303;
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 202));
            Assert(!session.HasReturnPoint && show.Current == 303, "Same-deck slide change rejects a proposal.");
            show.Current = 101;
            proposal = session.PrepareSearch("matches");
            show.Order.Remove(202);
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 202));
            Assert(!session.HasReturnPoint && show.Current == 101, "Deleted target rejection does not capture.");
            show.Order.Insert(1, 202);
            proposal = session.PrepareSearch("matches");
            show.Key = "another-deck";
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 202));
            Assert(!session.HasReturnPoint, "Deck change rejects pending proposals.");
            proposal = session.PrepareSearch("matches");
            show.NewShow("restarted");
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 202));
            Assert(!session.HasReturnPoint, "Restart rejects a proposal even at the same slide ID.");

            session.Search("matches");
            session.PrepareSearch("matches");
            session.NextResult();
            Assert(show.Current == 303, "Cancelling a new proposal preserves the prior result list and cursor.");
            session.Resume();
            Assert(show.Current == 101, "Cancelling a new proposal preserves the original checkpoint.");
            show.Order = new List<int> { 101, 202, 303, 404, 505, 606, 707 };
            proposal = session.PrepareSearch("many");
            Assert(proposal.Candidates.Count == 6, "Proposals retain more than the five UI choices.");
            session.AcceptSearch(proposal, 707);
            Assert(show.Current == 707 && session.HasReturnPoint, "Accept navigates to the selected ranked candidate.");
            session.NextResult();
            Assert(show.Current == 202, "The full result snapshot wraps after the selected final candidate.");
            session.Resume();
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 707));
            Assert(!session.HasReturnPoint, "A successfully accepted proposal cannot be reused.");
        }

        private static SearchProposal Finish(SearchOperation operation)
        {
            int slices = 0;
            while (!operation.Step(1))
                Assert(++slices < 100000, "An incremental operation must terminate.");
            return operation.Proposal;
        }

        private static void IndexedSearchTests(string directory)
        {
            var show = new IndexedShow(4, 0);
            show.Titles[0] = "R\u00e9sum\u00e9";
            show.Contents[0] = "overview";
            show.Titles[1] = "The resume details";
            show.Contents[1] = "long term growth";
            show.Titles[2] = "Appendix";
            show.Contents[2] = "R\u00e9sum\u00e9 table cell grouped label alternative text";
            show.Titles[3] = "Resume details";
            show.Contents[3] = "overview";
            var session = new PresentationSession(true, directory, () => show);

            SearchProposal cold;
            using (SearchOperation operation = session.BeginSearch("R\u00c9SUM\u00c9!"))
            {
                Assert(show.Reads == 0, "BeginSearch must not open COM or extract text before the first UI paint.");
                cold = Finish(operation);
                Assert(!operation.IsWarm && operation.IsComplete && operation.Phase == SearchPhase.Complete,
                    "Cold operation completion.");
            }
            Assert(cold.Candidates.Count == 4 && cold.Candidates[0].Score == 1000 &&
                cold.Candidates[1].Score == 800 && cold.Candidates[2].Score == 800 &&
                cold.Candidates[3].Score == 500, "Preserve accent, punctuation and exact/phrase/content scores.");
            Assert(cold.Candidates[1].SlideId == 102 && cold.Candidates[2].SlideId == 104,
                "Equal scores sort by original slide number.");
            Assert(show.Reads == 4 && show.ActiveReaders == 0 && !session.HasReturnPoint && show.Current == 101,
                "Cold extraction visits every slide once, releases readers and never navigates.");
            int reads = show.Reads;
            using (SearchOperation operation = session.BeginSearch("alternative"))
            {
                SearchProposal warm = Finish(operation);
                Assert(operation.IsWarm && warm.Candidates[0].SlideId == 103 && show.Reads == reads,
                    "A different warm query reuses normalized text without extracting any slide.");
            }
            Throws<InvalidOperationException>(() => session.AcceptSearch(cold, 101));
            Assert(!session.HasReturnPoint, "A superseded proposal cannot navigate.");
            SearchProposal navigated = session.PrepareSearch("resume");
            session.GoTo(2);
            session.Resume();
            Throws<InvalidOperationException>(() => session.AcceptSearch(navigated, 102));
            Assert(show.Current == 101 && !session.HasReturnPoint,
                "Navigating away and back invalidates a proposal even when its origin matches again.");

            SearchProposal proposal = session.PrepareSearch("resume");
            show.Order.Reverse();
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 102));
            Assert(!session.HasReturnPoint, "Reordered IDs reject stale candidates even if metadata fails to change.");
            proposal = session.PrepareSearch("resume");
            Assert(proposal.Candidates[1].SlideId == 104 && proposal.Candidates[1].SlideNumber == 1,
                "Rebuild preserves IDs and updates tie order after reorder.");

            show.Dirty = true;
            show.Contents[2] = "first dirty edit needle";
            Assert(session.PrepareSearch("needle").Candidates[0].SlideId == 103, "First unsaved edit is searchable.");
            int dirtyReads = show.Reads;
            show.Contents[2] = "second dirty edit changed";
            proposal = session.PrepareSearch("changed");
            Assert(show.Reads >= dirtyReads + 8 && proposal.Candidates[0].SlideId == 103,
                "Saved=false cannot hide further edits; dirty searches reread and verify every slide.");
            show.Contents[2] = "third dirty edit replaced";
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 103));
            Assert(!session.HasReturnPoint && show.Current == 101 && show.ActiveReaders == 0,
                "Content edits during the choice dialog reject before navigation and release readers.");

            proposal = session.PrepareSearch("replaced");
            using (SearchOperation accept = session.BeginAcceptSearch(proposal, 103))
            {
                Finish(accept);
                Assert(accept.Result.StartsWith("Slide") && show.Current == 103 &&
                    session.HasReturnPoint && show.ActiveReaders == 0,
                    "Interactive acceptance supports dirty decks after full content validation.");
            }
            session.Resume();
            using (SearchOperation dirty = session.BeginSearch("replaced"))
            {
                proposal = Finish(dirty);
                Assert(!dirty.IsWarm && dirty.RequiresTextValidation && proposal.Candidates[0].SlideId == 103,
                    "Interactive dirty-deck searches reread content without requiring a save.");
            }
            Assert(session.AcceptSearch(proposal, 103).StartsWith("Slide") && show.Current == 103 && session.HasReturnPoint,
                "Synchronous compatibility retains dirty/unsaved searching and checkpoint behavior.");
            session.BeginQuestions();
            session.Resume();
            Assert(show.Current == 101 && !session.HasReturnPoint && !session.QuestionsMode,
                "Indexed search preserves Q&A/resume semantics.");

            show.Dirty = false;
            show.Revision++;
            proposal = session.PrepareSearch("replaced");
            reads = show.Reads;
            session.InvalidateSearchIndex();
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 103));
            session.PrepareSearch("replaced");
            Assert(show.Reads == reads + 4, "Explicit refresh invalidates proposals and cached text.");
            reads = show.Reads;
            using (SearchOperation refresh = session.BeginSearch("replaced", true)) Finish(refresh);
            Assert(show.Reads == reads + 4, "BeginSearch refresh forces a full rebuild.");

            proposal = session.PrepareSearch("replaced");
            show.Revision++;
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 103));
            Assert(!session.HasReturnPoint, "A saved content version change rejects pending proposals.");
            proposal = session.PrepareSearch("replaced");
            show.Path = "SaveAs.pptx";
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 103));
            proposal = session.PrepareSearch("replaced");
            show.RuntimeId = "new-show";
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 103));
            proposal = session.PrepareSearch("replaced");
            show.Order.Remove(103);
            Throws<InvalidOperationException>(() => session.AcceptSearch(proposal, 103));
            Assert(show.ActiveReaders == 0 && !session.HasReturnPoint,
                "Save As, show restart and deletion reject pending indexed proposals.");

            show = new IndexedShow(80, 60);
            session = new PresentationSession(true, directory, () => show);
            show.SlicesPerSlide = 80;
            using (SearchOperation cancel = session.BeginSearch("needle"))
            {
                Assert(!cancel.Step(1) && !cancel.Step(1) && show.ActiveReaders == 1, "Suspend inside a text reader.");
                cancel.Cancel();
                Assert(cancel.IsCancelled && show.ActiveReaders == 0 && show.Current == 101 &&
                    !session.HasReturnPoint, "Cancellation promptly disposes the in-flight reader without navigation.");
                Throws<OperationCanceledException>(() => cancel.Step());
            }
            show.SlicesPerSlide = 1;
            using (SearchOperation next = session.BeginSearch("needle"))
            {
                Finish(next);
                Assert(!next.IsWarm, "A cancelled partial index is never published.");
            }
            proposal = session.PrepareSearch("needle");
            using (SearchOperation validation = session.BeginSearch("needle"))
            {
                validation.Step(1);
                Assert(validation.IsWarm && validation.Phase == SearchPhase.Validating,
                    "A saved-deck warm search verifies stable slide order before scoring.");
                validation.Cancel();
                Assert(show.Current == 101 && !session.HasReturnPoint, "Warm validation cancellation is side-effect free.");
            }
            using (SearchOperation scoring = session.BeginSearch("needle"))
            {
                while (scoring.Phase != SearchPhase.Scoring && !scoring.IsComplete) scoring.Step(1);
                Assert(!scoring.IsComplete, "Warm scoring yields before completing a large result set.");
                scoring.Cancel();
                Assert(show.ActiveReaders == 0 && show.Current == 101 && !session.HasReturnPoint,
                    "Warm scoring cancellation cannot publish or navigate.");
            }
            show.Dirty = true;
            proposal = session.PrepareSearch("needle");
            using (SearchOperation dirty = session.BeginSearch("needle"))
            {
                reads = show.Reads;
                Assert(!dirty.Step() && show.Reads == reads, "Dirty search paints before extraction.");
                while (dirty.Phase != SearchPhase.Validating) dirty.Step(1);
                Assert(!dirty.IsComplete && show.Reads == reads + 80,
                    "Dirty extraction yields before the final uninterrupted check.");
                show.Contents[0] = "changed after extraction";
                Throws<InvalidOperationException>(() => dirty.Step(1));
                Assert(dirty.Proposal == null && show.ActiveReaders == 0 && !session.HasReturnPoint,
                    "An edit to an already extracted dirty slide rejects the entire proposal.");
            }
            show.Contents[0] = "needle restored";
            show.Path = null;
            using (SearchOperation unsaved = session.BeginSearch("needle"))
            {
                while (unsaved.Phase != SearchPhase.Validating) unsaved.Step(1);
                reads = show.Reads;
                Assert(unsaved.Step(1) && show.Reads == reads + 80 && unsaved.Proposal.Candidates.Count == 80,
                    "Never-saved decks validate all content and publish in one final step without saving.");
                proposal = unsaved.Proposal;
            }
            using (SearchOperation cancel = session.BeginAcceptSearch(proposal, 102))
            {
                Assert(!cancel.Step(1), "Unversioned acceptance yields before starting its final check.");
                cancel.Cancel();
                Assert(show.Current == 101 && !session.HasReturnPoint && show.ActiveReaders == 0,
                    "Cancelling before unversioned validation leaves navigation unchanged.");
            }
            using (SearchOperation edited = session.BeginAcceptSearch(proposal, 102))
            {
                edited.Step(1);
                show.Contents[0] = "edited while choosing";
                Throws<InvalidOperationException>(() => edited.Step(1));
                Assert(show.Current == 101 && !session.HasReturnPoint && show.ActiveReaders == 0,
                    "An unversioned edit to any slide during the choice rejects before navigation.");
            }
            proposal = session.PrepareSearch("needle");
            using (SearchOperation accept = session.BeginAcceptSearch(proposal, 102))
            {
                accept.Step(1);
                reads = show.Reads;
                Assert(accept.Step(1) && show.Reads == reads + 80 && show.Current == 102 && session.HasReturnPoint,
                    "Unversioned acceptance rereads every slide and navigates without an intervening timer yield.");
            }
            session.Resume();
            Assert(show.Path == null && show.Dirty, "Search and acceptance never save or mark a deck clean.");
            show.Contents[0] = "needle restored";
            show.Path = "Synthetic.pptx";
            show.Dirty = false;
            show.Revision++;
            proposal = session.PrepareSearch("needle");
            show.SlicesPerSlide = 80;
            using (SearchOperation cancel = session.BeginAcceptSearch(proposal, 102))
            {
                cancel.Step(1);
                cancel.Step(1);
                cancel.Cancel();
                Assert(show.ActiveReaders == 0 && show.Current == 101 && !session.HasReturnPoint,
                    "Cancellation while accepting cannot navigate or capture a checkpoint.");
            }
            show.SlicesPerSlide = 1;
            using (SearchOperation old = session.BeginSearch("needle"))
            using (SearchOperation newer = session.BeginSearch("needle"))
            {
                Throws<InvalidOperationException>(() => old.Step());
                Finish(newer);
            }
            using (SearchOperation stale = session.BeginSearch("needle", true))
            {
                stale.Step(1);
                show.Current = 102;
                Throws<InvalidOperationException>(() => stale.Step());
                Assert(show.ActiveReaders == 0 && !session.HasReturnPoint, "Origin changes invalidate in-progress work.");
            }
            show.Current = 101;
            show.FailRead = true;
            using (SearchOperation failing = session.BeginSearch("needle", true))
                Throws<COMException>(() => Finish(failing));
            Assert(show.ActiveReaders == 0 && !session.HasReturnPoint, "Read failures release all readers.");
            show.FailRead = false;

            using (SearchOperation editing = session.BeginSearch("needle", true))
            {
                editing.Step(1);
                editing.Step(1);
                show.Contents[79] = "modified after first pass";
                show.Dirty = true;
                Throws<InvalidOperationException>(() => Finish(editing));
                Assert(show.ActiveReaders == 0 && !session.HasReturnPoint,
                    "An edit during incremental saved-deck extraction cannot publish a mixed-content proposal.");
            }
            show.Contents[79] = "needle restored";
            session.Search("needle");
            session.GoTo(2);
            session.BeginQuestions();
            Throws<InvalidOperationException>(() => session.PrepareSearch("nomatchingword"));
            session.InvalidateSearchIndex();
            session.NextResult();
            Assert(show.Current == 102 && session.HasReturnPoint && session.QuestionsMode,
                "Refresh and failed indexed searches preserve ranked-result cursor and Q&A checkpoint.");
            session.Resume();
            show.Dirty = false;
            show.Revision++;
            proposal = session.PrepareSearch("needle");
            using (SearchOperation accept = session.BeginAcceptSearch(proposal, 102))
            {
                Finish(accept);
                Assert(accept.Result.StartsWith("Slide") && show.Current == 102 && session.HasReturnPoint,
                    "Saved-deck incremental acceptance validates then navigates in its final slice.");
            }
            session.Resume();

            using (SearchOperation operation = session.BeginSearch("needle"))
            {
                Exception wrongThread = null;
                var worker = new System.Threading.Thread(delegate()
                {
                    try { operation.Step(); }
                    catch (Exception exception) { wrongThread = exception; }
                });
                worker.Start();
                worker.Join();
                Assert(wrongThread is InvalidOperationException && show.ActiveReaders == 0,
                    "Cross-thread advancement is rejected before reaching the connection.");
                operation.Cancel();
            }

            Assert(PowerPointController.NormalizeForSearch("caf\u00e9--au lait 42") == "cafe au lait 42",
                "Normalization retains existing accent, boundary and number behavior.");
            string[] words = PowerPointController.GetSearchWords("the growth and growth");
            Assert(words.Length == 1 && words[0] == "growth", "Stop words and duplicate content words retain existing semantics.");
            words = PowerPointController.GetSearchWords("the the");
            Assert(words.Length == 2, "All-stopword fallback retains duplicate words as before.");
            Assert(PowerPointController.ScoreNormalizedSlide("long growth", "growth value", "growth vision",
                new[] { "growth", "vision" }) == 120, "Nonphrase title/content scoring weights are unchanged.");
            Assert(PowerPointController.ScoreNormalizedSlide("alphabet", "alphabetical", "alpha",
                new[] { "alpha" }) == 0, "Partial-word matches remain excluded.");

            show = new IndexedShow(180, 100000);
            session = new PresentationSession(true, directory, () => show);
            proposal = session.PrepareSearch("needle");
            Assert(proposal.Candidates.Count == 180, "Oversized indexes still search every slide without truncation.");
            reads = show.Reads;
            using (SearchOperation oversized = session.BeginSearch("needle"))
            {
                Finish(oversized);
                Assert(!oversized.IsWarm && show.Reads == reads + 180,
                    "Indexes exceeding the 16M-character memory cap are streamed, never retained as a warm cache.");
            }
            Assert(show.ActiveReaders == 0, "Over-budget streaming releases all extraction readers.");
        }

        private static void SearchLifetimeTests(string directory)
        {
            var show = new IndexedShow(80, 60);
            var session = new PresentationSession(true, directory, () => show);
            session.Search("needle");
            session.NextResult();
            session.BeginQuestions();
            show.SlicesPerSlide = 80;
            using (SearchOperation cancelled = session.BeginSearch("needle", true))
            {
                cancelled.Step(1);
                cancelled.Step(1);
                Assert(show.ActiveReaders == 1, "Lifetime test suspends inside an owned reader.");
                cancelled.Cancel();
                cancelled.Dispose();
                Assert(show.ActiveReaders == 0 && show.Current == 102 &&
                    session.HasReturnPoint && session.QuestionsMode,
                    "Cancel and repeated disposal preserve an existing checkpoint and Q&A mode.");
                Assert(typeof(SearchOperation).GetField("advance", BindingFlags.NonPublic | BindingFlags.Instance)
                    .GetValue(cancelled) == null, "A cancelled operation releases its work and text references.");
            }
            show.SlicesPerSlide = 1;
            SearchProposal proposal = session.PrepareSearch("needle");
            using (SearchOperation acceptance = session.BeginAcceptSearch(proposal, 104))
            {
                acceptance.Step(1);
                acceptance.Dispose();
                Assert(show.Current == 102 && session.HasReturnPoint && session.QuestionsMode,
                    "Disposing pending acceptance preserves the prior current slide and checkpoint.");
                Throws<ObjectDisposedException>(() => acceptance.Step());
            }
            session.NextResult();
            Assert(show.Current == 103, "Cancelled search and acceptance preserve the previous result cursor.");
            session.Resume();
            Assert(show.Current == 101 && !session.HasReturnPoint && !session.QuestionsMode,
                "Resume after cancellation returns to the original pre-search checkpoint.");

            show.SlicesPerSlide = 80;
            using (SearchOperation superseded = session.BeginSearch("needle", true))
            {
                superseded.Step(1);
                superseded.Step(1);
                Assert(show.ActiveReaders == 1, "Supersession test has a suspended reader.");
                using (SearchOperation replacement = session.BeginSearch("needle"))
                {
                    Assert(show.ActiveReaders == 0, "A newer search promptly disposes superseded reader references.");
                    Throws<ObjectDisposedException>(() => superseded.Step());
                }
            }
            using (SearchOperation pending = session.BeginSearch("needle"))
            {
                pending.Step(1);
                pending.Step(1);
                Assert(show.ActiveReaders == 1, "Session disposal test has an active reader.");
                session.Dispose();
                session.Dispose();
                Assert(show.ActiveReaders == 0 && pending.IsCancelled && show.Current == 101,
                    "Idempotent session disposal releases its active reader without navigation.");
                Throws<OperationCanceledException>(() => pending.Step());
            }
            Throws<ObjectDisposedException>(() => session.BeginSearch("needle"));
            Throws<ObjectDisposedException>(() => session.PrepareSearch("needle"));
            Throws<ObjectDisposedException>(() => session.BeginAcceptSearch(proposal, 102));
            Throws<ObjectDisposedException>(() => session.Next());

            show.SlicesPerSlide = 1;
            session = new PresentationSession(true, directory, () => show);
            using (SearchOperation completed = session.BeginSearch("needle"))
            {
                Finish(completed);
                Assert(typeof(SearchOperation).GetField("advance", BindingFlags.NonPublic | BindingFlags.Instance)
                    .GetValue(completed) == null, "Completed operations release their extraction work references.");
            }
            FieldInfo cache = typeof(PresentationSession).GetField("searchIndex", BindingFlags.NonPublic | BindingFlags.Instance);
            FieldInfo proposals = typeof(PresentationSession).GetField("proposals", BindingFlags.NonPublic | BindingFlags.Instance);
            object oldProposals = proposals.GetValue(session);
            Assert(cache.GetValue(session) != null, "A saved session has a reusable index before disposal.");
            session.Dispose();
            Assert(cache.GetValue(session) == null && !ReferenceEquals(oldProposals, proposals.GetValue(session)) &&
                !session.HasReturnPoint && session.ActiveRouteName == null,
                "Session disposal drops cached text, prepared proposals and presentation history.");
        }

        private static void SearchBenchmark(string directory)
        {
            const int count = 2500;
            var show = new IndexedShow(count, 3600);
            var session = new PresentationSession(true, directory, () => show);
            var clock = Stopwatch.StartNew();
            SearchProposal cold;
            int coldSlices = 0;
            long worstColdSlice = 0;
            using (SearchOperation operation = session.BeginSearch("needle"))
            {
                while (true)
                {
                    long start = Stopwatch.GetTimestamp();
                    bool done = operation.Step(10);
                    worstColdSlice = Math.Max(worstColdSlice, Stopwatch.GetTimestamp() - start);
                    coldSlices++;
                    if (done) break;
                }
                cold = operation.Proposal;
            }
            long coldTicks = clock.ElapsedTicks;
            int coldReads = show.Reads;
            clock.Restart();
            SearchProposal warm;
            int slices = 0;
            long worstSlice = 0;
            using (SearchOperation operation = session.BeginSearch("needle"))
            {
                while (true)
                {
                    long start = Stopwatch.GetTimestamp();
                    bool done = operation.Step(10);
                    worstSlice = Math.Max(worstSlice, Stopwatch.GetTimestamp() - start);
                    slices++;
                    if (done) break;
                }
                warm = operation.Proposal;
                Assert(operation.IsWarm, "Benchmark uses the populated saved-deck index.");
            }
            long warmTicks = clock.ElapsedTicks;
            Assert(coldReads == count && show.Reads == coldReads, "Warm benchmark must perform zero text reads.");
            Assert(cold.Candidates.Count == count && warm.Candidates.Count == count, "Benchmark results retain every candidate.");
            for (int i = 0; i < count; i++)
                Assert(cold.Candidates[i].SlideId == warm.Candidates[i].SlideId &&
                    cold.Candidates[i].Score == warm.Candidates[i].Score, "Cold/warm results must be identical.");
            Console.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "Search benchmark: {0} slides, ~{1} characters/slide; cold={2:F1}ms ({3} text reads), " +
                "warm={4:F1}ms (0 text reads), speedup={5:F2}x, warm slices={6}, max warm slice={7:F2}ms. " +
                "Cold slices={8}, max cold slice={9:F2}ms. " +
                "Synthetic managed data only; no PowerPoint was opened.",
                count, 3600, coldTicks * 1000.0 / Stopwatch.Frequency, coldReads,
                warmTicks * 1000.0 / Stopwatch.Frequency, (double)coldTicks / Math.Max(1, warmTicks), slices,
                worstSlice * 1000.0 / Stopwatch.Frequency, coldSlices,
                worstColdSlice * 1000.0 / Stopwatch.Frequency));
        }

        private static SlideRoute Route(string name, params int[] ids)
        {
            return new SlideRoute { Name = name, TargetMinutes = 15, SlideIds = new List<int>(ids) };
        }

        private static void RouteTests(string directory)
        {
            var show = new FakeShow();
            PresentationSession session = Session(show, directory);
            Assert(session.ActiveRouteName == null, "No route is active initially.");
            Throws<InvalidOperationException>(() => session.GetRoutes());
            Throws<InvalidOperationException>(() => session.SaveRoute(Route("unsaved", 101)));
            show.Path = Path.Combine(directory, "original", "Deck.pptx");
            Assert(session.Snapshot().SavedPath == show.Path, "Snapshot exposes the current saved path.");
            List<SlideChoice> choices = session.GetSlides();
            Assert(choices.Count == 3 && choices[1].SlideId == 202 && choices[1].SlideNumber == 2, "GetSlides preserves IDs and order.");
            Throws<InvalidOperationException>(() => session.SaveRoute(Route("", 101)));
            Throws<InvalidOperationException>(() => session.SaveRoute(Route("empty")));
            Throws<InvalidOperationException>(() => session.SaveRoute(Route("duplicates", 101, 101)));
            Throws<InvalidOperationException>(() => session.SaveRoute(Route("missing", 999)));
            SlideRoute invalid = Route("time", 101);
            invalid.TargetMinutes = 0;
            Throws<InvalidOperationException>(() => session.SaveRoute(invalid));
            invalid.TargetMinutes = 241;
            Throws<InvalidOperationException>(() => session.SaveRoute(invalid));
            session.SaveRoute(Route("Short", 101, 303));
            Assert(session.GetRoutes().Count == 1, "Route persisted.");
            session.SaveRoute(Route("SHORT", 101, 303));
            Assert(session.GetRoutes().Count == 1, "Route name updates are case-insensitive.");
            session = Session(show, directory);
            Assert(session.GetRoutes()[0].SlideIds[1] == 303, "Routes persist between sessions.");
            session.Search("matches");
            session.BeginQuestions();
            session.ActivateRoute("short");
            Assert(show.Current == 101 && session.ActiveRouteName == "SHORT", "Activating starts at the first route slide.");
            Assert(!session.HasReturnPoint && !session.QuestionsMode, "Route activation clears excursion state.");
            Throws<InvalidOperationException>(() => session.NextResult());
            int standardMoves = show.Moves;
            session.Next();
            Assert(show.Current == 303 && show.Moves == standardMoves, "Route advances whole slides, skipping ordinary order.");
            Assert(session.Next().Contains("End of route") && show.Current == 303, "Route end does not wrap or end the show.");
            session.Previous();
            Assert(show.Current == 101 && session.Previous().Contains("Start of route"), "Route start does not wrap.");
            show.Order = new List<int> { 202, 303, 101 };
            session.Next();
            Assert(show.Current == 303, "Route order uses stable IDs after slide reordering.");
            session.Previous();
            show.Order = new List<int> { 101, 202, 303 };
            session.BeginQuestions();
            session.Next();
            Assert(show.Current == 202 && show.Moves > standardMoves, "Questions use ordinary animation/slide navigation.");
            session.Resume();
            Assert(show.Current == 101 && session.ActiveRouteName == "SHORT", "Resume returns to the route point without deactivating it.");
            session.Search("matches");
            standardMoves = show.Moves;
            session.Next();
            Assert(show.Current == 303 && show.Moves == standardMoves + 1, "Browsing outside the route uses ordinary navigation.");
            session.Resume();
            Assert(show.Current == 101, "Search excursions resume at the route origin.");
            session.SaveRoute(Route("All", 101, 202, 303));
            session.ActivateRoute("All");
            show.Order.Remove(202);
            session.Next();
            Assert(show.Current == 303, "Route navigation skips deleted IDs.");
            show.Order = new List<int> { 404 };
            show.Current = 404;
            Throws<InvalidOperationException>(() => session.Next());
            Assert(session.ActiveRouteName == "All" && show.Current == 404, "All-deleted route fails without navigation.");
            session.DeactivateRoute();
            Assert(session.ActiveRouteName == null && show.Current == 404, "Deactivating never moves the slide.");
            Throws<InvalidOperationException>(() => session.ActivateRoute("All"));
            Assert(session.ActiveRouteName == null, "Activation with a deleted first slide fails safely.");
            show.NewShow("fresh-route-show");
            session.ActivateRoute("All");
            show.Unavailable = true;
            Throws<PresentationUnavailableException>(() => session.Snapshot());
            Assert(session.ActiveRouteName == null, "Ending a show clears the active route.");
            show.Unavailable = false;
            session.ActivateRoute("All");
            show.Path = Path.Combine(directory, "save-as", "Deck.pptx");
            session.Snapshot();
            Assert(session.ActiveRouteName == null && session.GetRoutes().Count == 0, "Same-name Save As resets route and path-bound profiles.");
            show.Path = Path.Combine(directory, "original", "Deck.pptx");
            session.ActivateRoute("All");
            session.RemoveRoute("all");
            Assert(session.ActiveRouteName == null && session.GetRoutes().Count == 1, "Removing the active route deactivates without touching other routes.");
            Throws<InvalidOperationException>(() => session.RemoveRoute("missing"));
            session.SaveRoute(Route("SHORT", 101, 303));
            show.FailNavigation = true;
            Throws<COMException>(() => session.ActivateRoute("SHORT"));
            Assert(session.ActiveRouteName == null, "A failed route activation cannot commit state.");
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
            Assert(sessionType.GetMethod("PrepareSearch", new[] { typeof(string) }) != null, "Candidate search proposal API.");
            Assert(sessionType.GetMethod("AcceptSearch") != null, "Candidate search acceptance API.");
            foreach (string name in new[] { "GetSlides", "GetRoutes", "SaveRoute", "RemoveRoute", "ActivateRoute", "DeactivateRoute" })
                Assert(sessionType.GetMethod(name) != null, "Candidate route API: " + name);
            Assert(sessionType.GetProperty("ActiveRouteName") != null, "Candidate active route property.");
            Assert(candidateAssembly.GetType("JarvisPowerPoint.PresentationUnavailableException", true).GetProperty("Reason") != null,
                "Candidate availability exception exposes its reason.");
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
                SearchTextComFixture(slides);
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
                Assert(!string.IsNullOrEmpty(first.SavedPath) && first.WindowHandle != IntPtr.Zero, "Live snapshot exposes saved path and slideshow window.");
                object proposed = session.PrepareSearch("Budget");
                var proposedSlides = (System.Collections.IList)proposed.GetType().GetProperty("Candidates").GetValue(proposed, null);
                Assert(proposedSlides.Count == 2 &&
                    (int)proposedSlides[0].GetType().GetProperty("Score").GetValue(proposedSlides[0], null) > 0,
                    "Live proposals carry the shared search engine scores.");
                Assert(session.Snapshot().SlideId == first.SlideId && !session.HasReturnPoint, "Live prepare is non-navigating.");
                int lastChoiceId = (int)proposedSlides[1].GetType().GetProperty("SlideId").GetValue(proposedSlides[1], null);
                session.AcceptSearch(proposed, lastChoiceId);
                Assert(session.Snapshot().SlideId == lastChoiceId, "Live acceptance selects the requested result.");
                session.Resume();
                session.SaveRoute(Route("Live route", first.SlideId, lastChoiceId));
                session.ActivateRoute("Live route");
                session.Next();
                Assert(session.Snapshot().SlideId == lastChoiceId, "Live route skips the middle slide.");
                session.Next();
                Assert(session.Snapshot().SlideId == lastChoiceId, "Live route boundary keeps the show open.");
                session.Previous();
                session.DeactivateRoute();
                Assert(session.Snapshot().SlideId == first.SlideId, "Live route deactivation does not move.");
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
                PresentationUnavailableException unavailable = Throws<PresentationUnavailableException>(() => session.Snapshot());
                Assert(unavailable.Message.Contains("Several"), "Two live shows must fail safely with an explanation.");
                Assert(unavailable.Reason == PresentationUnavailableReason.MultipleSlideShows,
                    "Ambiguity has a distinct preflight reason.");
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

        private static void SearchTextComFixture(PowerPoint.Slides slides)
        {
            PowerPoint.Slide slide = null;
            PowerPoint.Shapes shapes = null;
            PowerPoint.Shape left = null, right = null, group = null, tableShape = null, alternative = null;
            PowerPoint.ShapeRange range = null;
            PowerPoint.Table table = null;
            try
            {
                slide = slides[3];
                shapes = slide.Shapes;
                left = shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 10, 100, 180, 30);
                right = shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 200, 100, 180, 30);
                SetShapeText(left, "nestedone");
                SetShapeText(right, "nestedtwo");
                range = shapes.Range(new object[] { left.Name, right.Name });
                group = range.Group();
                group.AlternativeText = "groupalt";
                tableShape = shapes.AddTable(1, 2, 10, 160, 380, 60);
                table = tableShape.Table;
                for (int column = 1; column <= 2; column++)
                {
                    PowerPoint.Cell cell = null;
                    PowerPoint.Shape cellShape = null;
                    try
                    {
                        cell = table.Cell(1, column);
                        cellShape = cell.Shape;
                        SetShapeText(cellShape, column == 1 ? "tableone" : "tabletwo");
                    }
                    finally { Release(cellShape); Release(cell); }
                }
                alternative = shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 10, 240, 200, 30);
                alternative.AlternativeText = "picturealt";
                string text = PowerPointSearchText.ReadAll(slide);
                Assert(text.Contains(" nestedone nestedtwo groupalt") && text.Contains(" tableone tabletwo") &&
                    text.Contains(" picturealt"), "Grouped children, table cells and alternative text preserve walk order.");
                foreach (string query in new[] { "nestedone nestedtwo", "tableone tabletwo", "groupalt", "picturealt" })
                {
                    List<PowerPointController.SlideSearchResult> found = PowerPointController.FindSlides(slides, query);
                    Assert(found.Count == 1 && found[0].SlideNumber == 3 && found[0].Score == 500,
                        "Shared full-deck extraction searches synthetic grouped/table/alternative text.");
                }
                using (var reader = new PowerPointSearchSlideReader(slides[3]))
                {
                    reader.ReadNext();
                    reader.ReadNext();
                    // Simulates a status refresh while the incremental reader retains shapes.
                    Assert(PowerPointController.GetSlideTitle(slide) == "Budget details",
                        "Title reads cannot invalidate retained incremental COM references.");
                }
                Assert(PowerPointSearchText.ReadAll(slide) == text,
                    "Cancelling a COM iterator releases only its own references and leaves content unchanged.");
            }
            finally
            {
                Release(alternative); Release(table); Release(tableShape); Release(group); Release(range);
                Release(right); Release(left); Release(shapes); Release(slide);
            }
        }

        private static void SetShapeText(PowerPoint.Shape shape, string value)
        {
            PowerPoint.TextFrame frame = null;
            PowerPoint.TextRange range = null;
            try { frame = shape.TextFrame; range = frame.TextRange; range.Text = value; }
            finally { Release(range); Release(frame); }
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
            public object PrepareSearch(string query) { return Call("PrepareSearch", query); }
            public void AcceptSearch(object proposal, int slideId) { Call("AcceptSearch", proposal, slideId); }
            public void Next() { Call("Next"); }
            public void Previous() { Call("Previous"); }
            public void GoTo(int number) { Call("GoTo", number); }
            public void GoToAlias(string name) { Call("GoToAlias", name); }
            public void SaveAlias(string name) { Call("SaveAlias", name); }
            public void NextResult() { Call("NextResult"); }
            public void BeginQuestions() { Call("BeginQuestions"); }
            public void Resume() { Call("Resume"); }
            public void SetBlackScreen(bool black) { Call("SetBlackScreen", black); }
            public void SaveRoute(SlideRoute route)
            {
                Type routeType = candidateAssembly.GetType("JarvisPowerPoint.SlideRoute", true);
                object value = Activator.CreateInstance(routeType, true);
                routeType.GetProperty("Name").SetValue(value, route.Name, null);
                routeType.GetProperty("TargetMinutes").SetValue(value, route.TargetMinutes, null);
                routeType.GetProperty("SlideIds").SetValue(value, route.SlideIds, null);
                Call("SaveRoute", value);
            }
            public void ActivateRoute(string name) { Call("ActivateRoute", name); }
            public void DeactivateRoute() { Call("DeactivateRoute"); }
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
                    {
                        object reason = exception.InnerException.GetType().GetProperty("Reason").GetValue(exception.InnerException, null);
                        throw new PresentationUnavailableException(exception.InnerException.Message,
                            (PresentationUnavailableReason)Enum.Parse(typeof(PresentationUnavailableReason), reason.ToString()));
                    }
                    throw exception.InnerException;
                }
            }
        }

        private sealed class IndexedShow : PresentationConnection
        {
            public readonly List<int> Order = new List<int>();
            public readonly List<string> Titles = new List<string>();
            public readonly List<string> Contents = new List<string>();
            public int Current = 101;
            public int Reads;
            public int ActiveReaders;
            public int Revision;
            public int SlicesPerSlide = 1;
            public bool Dirty;
            public bool FailRead;
            public string Path = "Synthetic.pptx";
            public string RuntimeId = "indexed-show";
            public IndexedShow(int count, int contentSize)
            {
                for (int index = 0; index < count; index++)
                {
                    Order.Add(index + 101);
                    Titles.Add("Title " + index);
                    Contents.Add("needle " + new string('x', contentSize) + " synthetic grouped table alternative");
                }
            }
            public override string Identity { get { return RuntimeId; } }
            public override string PresentationKey { get { return Path; } }
            public override string PresentationName { get { return "Synthetic"; } }
            public override string SavedPath { get { return Path; } }
            public override int SlideCount { get { return Order.Count; } }
            public override float Elapsed { get { return 0; } }
            public override int? ClickIndex { get { return 0; } }
            public override int CurrentSlideId { get { return Current; } }
            public override PowerPoint.PpSlideShowState State { get; set; }
            public override bool SupportsSearchIndex { get { return true; } }
            public override string SearchVersion { get { return Dirty ? null : Revision.ToString(); } }
            public override PresentationSlide CurrentSlide() { return FindSlide(Current); }
            public override PresentationSlide SlideAt(int number) { return FindSlide(Order[number - 1]); }
            public override int SlideIdAt(int number) { return Order[number - 1]; }
            public override PresentationSlide FindSlide(int id)
            {
                int position = Order.IndexOf(id);
                return position < 0 ? null : new PresentationSlide {
                    Id = id, Number = position + 1, Title = Titles[id - 101] };
            }
            public override SearchSlideReader OpenSearchSlide(int number)
            {
                Reads++;
                ActiveReaders++;
                return new Reader(this, Order[number - 1], number);
            }
            public override List<PresentationSlide> Search(string query)
            {
                throw new Exception("Indexed searches must not call the legacy full-deck search.");
            }
            public override void GoTo(int id)
            {
                if (!Order.Contains(id)) throw new InvalidOperationException("Deleted.");
                Current = id;
            }
            public override void Move(bool next)
            {
                int target = Order.IndexOf(Current) + (next ? 1 : -1);
                if (target >= 0 && target < Order.Count) Current = Order[target];
            }
            public override bool RestoreClick(int? index) { return true; }
            public override void Dispose() { }

            private sealed class Reader : SearchSlideReader
            {
                private readonly IndexedShow show;
                private readonly int id;
                private readonly int number;
                private int steps;
                private bool disposed;
                private SearchDocument document;
                public Reader(IndexedShow show, int id, int number)
                {
                    this.show = show;
                    this.id = id;
                    this.number = number;
                }
                public override SearchDocument Document { get { return document; } }
                public override bool ReadNext()
                {
                    if (show.FailRead) throw new COMException("Synthetic extraction failure.");
                    if (++steps < show.SlicesPerSlide) return false;
                    document = new SearchDocument(id, number, show.Titles[id - 101], show.Contents[id - 101]);
                    return true;
                }
                public override void Dispose()
                {
                    if (!disposed) show.ActiveReaders--;
                    disposed = true;
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
            public int Moves;
            public int SecondScore = 90;
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
                return index < 0 ? null : new PresentationSlide { Id = id, Number = index + 1, Title = "Title " + id,
                    Score = id == 202 ? 100 : SecondScore };
            }
            public override List<PresentationSlide> Search(string query)
            {
                var found = new List<PresentationSlide>();
                if (query == "matches")
                {
                    if (FindSlide(202) != null) found.Add(FindSlide(202));
                    if (FindSlide(303) != null) found.Add(FindSlide(303));
                }
                if (query == "many")
                    foreach (int id in Order) if (id != 101) found.Add(FindSlide(id));
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
                Moves++;
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
    # Compile only this module and the exact shared controller implementation, not the UI under development.
    $mainSource = [IO.File]::ReadAllText((Join-Path $projectDirectory "JarvisPowerPoint.cs"))
    $namespaceStart = $mainSource.IndexOf("namespace JarvisPowerPoint")
    $controllerStart = $mainSource.IndexOf("    internal static class PowerPointController")
    if ($namespaceStart -lt 0 -or $controllerStart -lt 0) { throw "Cannot locate the shared PowerPoint search helpers." }
    $controllerSource = $mainSource.Substring(0, $namespaceStart) + "namespace JarvisPowerPoint {" +
        [Environment]::NewLine + $mainSource.Substring($controllerStart)
    $controllerPath = Join-Path $testDirectory "PowerPointController.cs"
    [IO.File]::WriteAllText($controllerPath, $controllerSource, [Text.UTF8Encoding]::new($false))
    $sources = @("PresentationSession.cs", "SearchEngine.cs", "SearchText.cs",
        "PresentationProfileStore.cs", "PresentationContracts.cs", "AliasStore.cs") |
        ForEach-Object { Join-Path $projectDirectory $_ }
    $sources += $controllerPath
    & $compiler /nologo /target:exe /langversion:5 /main:JarvisPowerPoint.PresentationSessionTests `
        "/out:$testExecutable" /reference:System.dll /reference:System.Core.dll `
        /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Xml.dll `
        /reference:System.Runtime.Serialization.dll `
        "/reference:$speech" "/reference:$powerPoint" "/reference:$office" $sources $testSourcePath
    if ($LASTEXITCODE -ne 0) { throw "Session test compilation failed ($LASTEXITCODE)." }
    $selectedExecutable = $Executable
    if ($SourceOnly) { $selectedExecutable = $testExecutable }
    $arguments = @($testDirectory, $selectedExecutable)
    if ($IncludeCom) { $arguments += "com" }
    elseif ($SearchBenchmark) { $arguments += "benchmark" }
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $testExecutable
    $startInfo.Arguments = ($arguments | ForEach-Object { '"' + $_ + '"' }) -join ' '
    $startInfo.UseShellExecute = $false
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        [void]$process.Start()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "Session tests failed ($($process.ExitCode))." }
    } finally {
        $process.Dispose()
    }
} finally {
    if (Test-Path -LiteralPath $testDirectory) {
        Remove-Item -LiteralPath $testDirectory -Recurse -Force
    }
}
