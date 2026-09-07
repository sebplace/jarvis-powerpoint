using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Threading;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace JarvisPowerPoint
{
    internal enum PresentationUnavailableReason
    {
        NoSlideShow,
        MultipleSlideShows
    }

    internal sealed class PresentationUnavailableException : InvalidOperationException
    {
        public PresentationUnavailableException(string message)
            : this(message, PresentationUnavailableReason.NoSlideShow) { }

        public PresentationUnavailableException(string message, PresentationUnavailableReason reason)
            : base(message) { Reason = reason; }

        public PresentationUnavailableReason Reason { get; private set; }
    }

    internal sealed partial class PresentationSession : IDisposable
    {
        private readonly int threadId = Thread.CurrentThread.ManagedThreadId;
        private bool disposed;
        private readonly AliasStore aliasStore;
        private readonly PresentationProfileStore profileStore;
        private readonly Func<PresentationConnection> connect;
        // The second-ranked result is ambiguous when it is within 15% of the best score.
        internal const double AmbiguityScoreRatio = 0.85;
        private ConditionalWeakTable<SearchProposal, PreparedSearch> proposals =
            new ConditionalWeakTable<SearchProposal, PreparedSearch>();
        private string identity;
        private string sessionKey;
        private string presentationKey;
        private float elapsed;
        private ReturnPoint returnPoint;
        private List<PresentationSlide> results;
        private int resultIndex;
        private PowerPoint.PpSlideShowState? beforeBlack;
        private SlideRoute activeRoute;

        public PresentationSession(bool english, string dataDirectory)
            : this(english, dataDirectory, null) { }

        internal PresentationSession(bool english, string dataDirectory, Func<PresentationConnection> connect)
        {
            English = english;
            aliasStore = new AliasStore(dataDirectory);
            profileStore = new PresentationProfileStore(dataDirectory);
            this.connect = connect ?? (() => new PowerPointConnection(English));
        }

        public bool English { get; set; }
        public bool HasReturnPoint { get { return returnPoint != null; } }
        public bool QuestionsMode { get; private set; }
        public string ActiveRouteName { get { return activeRoute == null ? null : activeRoute.Name; } }

        public void Dispose()
        {
            if (Thread.CurrentThread.ManagedThreadId != threadId)
                throw new InvalidOperationException("Dispose the presentation session on its creating STA thread.");
            if (disposed) return;
            disposed = true;
            try
            {
                if (activeSearchOperation != null) activeSearchOperation.Cancel();
            }
            finally
            {
                activeSearchOperation = null;
                Reset();
            }
        }

        private void ThrowIfDisposed()
        {
            if (disposed) throw new ObjectDisposedException("PresentationSession");
        }

        public PresentationSnapshot Snapshot()
        {
            using (PresentationConnection show = Open())
            {
                PresentationSlide slide = show.CurrentSlide();
                return new PresentationSnapshot
                {
                    SessionKey = sessionKey,
                    PresentationKey = presentationKey,
                    PresentationName = show.PresentationName,
                    SavedPath = show.SavedPath,
                    WindowHandle = show.WindowHandle,
                    SlideId = slide.Id,
                    SlideNumber = slide.Number,
                    SlideCount = show.SlideCount,
                    Title = slide.Title,
                    IsBlack = show.State == PowerPoint.PpSlideShowState.ppSlideShowBlackScreen
                };
            }
        }

        public string Next()
        {
            return Move(true);
        }

        public string Previous()
        {
            return Move(false);
        }

        public string GoTo(int slideNumber)
        {
            using (PresentationConnection show = Open())
            {
                if (slideNumber < 1 || slideNumber > show.SlideCount)
                    throw Error("Choose a slide between 1 and {0}.",
                        "Choisissez une diapositive entre 1 et {0}.", show.SlideCount);
                PresentationSlide target = show.SlideAt(slideNumber);
                Navigate(show, target);
                return SlideMessage(target);
            }
        }

        public string Search(string query)
        {
            SearchProposal proposal = PrepareSearch(query);
            return AcceptSearch(proposal, proposal.Candidates[0].SlideId);
        }

        public SearchProposal PrepareSearch(string query)
        {
            using (SearchOperation operation = BeginSearch(query))
            {
                while (!operation.Step()) { }
                return operation.Proposal;
            }
        }

        public string AcceptSearch(SearchProposal proposal, int selectedSlideId)
        {
            using (SearchOperation operation = BeginAcceptSearch(proposal, selectedSlideId))
            {
                while (!operation.Step()) { }
                return operation.Result;
            }
        }

        private string CommitSearch(SearchProposal proposal, int selectedSlideId)
        {
            using (PresentationConnection show = Open())
            {
                PreparedSearch prepared;
                if (proposal == null || !proposals.TryGetValue(proposal, out prepared) ||
                    proposal.SessionKey != sessionKey || prepared.SessionKey != sessionKey ||
                    proposal.OriginSlideId != prepared.Origin || show.CurrentSlideId != prepared.Origin ||
                    prepared.Generation != searchGeneration || !SearchVersionMatches(show, prepared))
                    throw Error("The presentation or current slide changed. Search again.",
                        "La présentation ou la diapositive actuelle a changé. Relancez la recherche.");
                int selectedIndex = prepared.Matches.FindIndex(match => match.Id == selectedSlideId);
                if (selectedIndex < 0)
                    throw Error("Choose a slide from the proposed search results.",
                        "Choisissez une diapositive parmi les résultats proposés.");
                PresentationSlide target = RequireSlide(show, selectedSlideId);
                Navigate(show, target);
                results = prepared.Matches;
                resultIndex = selectedIndex;
                proposals.Remove(proposal);
                return SlideMessage(target) + Text(" ({0} matches)", " ({0} résultats)", results.Count);
            }
        }

        public List<SlideChoice> GetSlides()
        {
            using (PresentationConnection show = Open())
            {
                var choices = new List<SlideChoice>();
                for (int index = 1; index <= show.SlideCount; index++) choices.Add(ToChoice(show.SlideAt(index)));
                return choices;
            }
        }

        public List<SlideRoute> GetRoutes()
        {
            using (PresentationConnection show = Open()) return profileStore.LoadRoutes(RequireSavedPath(show));
        }

        public void SaveRoute(SlideRoute route)
        {
            SlideRoute copy;
            try { copy = PresentationProfileStore.CopyRoute(route); }
            catch (ArgumentException)
            {
                throw Error("Use a route name of 1 to 80 characters, a target of 1 to 240 minutes, and a nonempty list of unique slide IDs.",
                    "Utilisez un nom de parcours de 1 à 80 caractères, une durée de 1 à 240 minutes et une liste non vide d’identifiants de diapositives uniques.");
            }
            using (PresentationConnection show = Open())
            {
                string path = RequireSavedPath(show);
                foreach (int id in copy.SlideIds) RequireSlide(show, id);
                List<SlideRoute> routes = profileStore.LoadRoutes(path);
                int index = routes.FindIndex(item => string.Equals(item.Name, copy.Name, StringComparison.OrdinalIgnoreCase));
                if (index < 0) routes.Add(copy);
                else routes[index] = copy;
                profileStore.SaveRoutes(path, routes);
                if (activeRoute != null && string.Equals(activeRoute.Name, copy.Name, StringComparison.OrdinalIgnoreCase))
                    activeRoute = copy;
            }
        }

        public void RemoveRoute(string name)
        {
            name = ValidateRouteName(name);
            using (PresentationConnection show = Open())
            {
                string path = RequireSavedPath(show);
                List<SlideRoute> routes = profileStore.LoadRoutes(path);
                if (routes.RemoveAll(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)) == 0)
                    throw Error("No route named \"{0}\" exists for this presentation.",
                        "Aucun parcours « {0} » pour cette présentation.", name);
                profileStore.SaveRoutes(path, routes);
                if (activeRoute != null && string.Equals(activeRoute.Name, name, StringComparison.OrdinalIgnoreCase))
                    activeRoute = null;
            }
        }

        public string ActivateRoute(string name)
        {
            name = ValidateRouteName(name);
            using (PresentationConnection show = Open())
            {
                SlideRoute route = profileStore.LoadRoutes(RequireSavedPath(show)).Find(item =>
                    string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
                if (route == null)
                    throw Error("No route named \"{0}\" exists for this presentation.",
                        "Aucun parcours « {0} » pour cette présentation.", name);
                PresentationSlide target = RequireSlide(show, route.SlideIds[0]);
                show.GoTo(target.Id);
                searchGeneration++;
                activeRoute = route;
                ClearHistory();
                return Text("Route \"{0}\" active. ", "Parcours « {0} » actif. ", route.Name) + SlideMessage(target);
            }
        }

        public string DeactivateRoute()
        {
            using (PresentationConnection show = Open()) activeRoute = null;
            return Text("Standard slide order restored.", "Ordre normal des diapositives rétabli.");
        }

        public string NextResult()
        {
            using (PresentationConnection show = Open())
            {
                if (results == null || results.Count == 0)
                    throw Error("Search for a slide first.", "Recherchez d’abord une diapositive.");
                for (int offset = 1; offset <= results.Count; offset++)
                {
                    int index = (resultIndex + offset) % results.Count;
                    PresentationSlide target = show.FindSlide(results[index].Id);
                    if (target == null) continue;
                    Navigate(show, target);
                    resultIndex = index;
                    return SlideMessage(target);
                }
                throw Error("The search results were deleted. Search again.",
                    "Les diapositives trouvées ont été supprimées. Relancez la recherche.");
            }
        }

        public string BeginQuestions()
        {
            using (PresentationConnection show = Open())
            {
                if (QuestionsMode)
                    return Text("Questions mode is already active.", "Le mode questions est déjà actif.");
                if (returnPoint == null) returnPoint = Capture(show);
                QuestionsMode = true;
                return Text("Questions mode. Your return point is saved.",
                    "Mode questions. Votre point de retour est mémorisé.");
            }
        }

        public string Resume()
        {
            using (PresentationConnection show = Open())
            {
                if (returnPoint == null)
                    throw Error("There is no return point. Search, jump to a slide or start questions first.",
                        "Aucun point de retour. Recherchez une diapositive, allez à une diapositive ou lancez les questions.");
                PresentationSlide target = RequireSlide(show, returnPoint.SlideId);
                show.GoTo(target.Id);
                searchGeneration++;
                bool animationRestored = show.RestoreClick(returnPoint.ClickIndex);
                show.State = returnPoint.State;
                beforeBlack = returnPoint.BeforeBlack;
                ClearHistory();
                return Text("Resumed: ", "Reprise : ") + SlideMessage(target) +
                    (animationRestored ? string.Empty : Text(
                        " Animation position could not be restored.",
                        " La position de l’animation n’a pas pu être restaurée."));
            }
        }

        public string SetBlackScreen(bool black)
        {
            using (PresentationConnection show = Open())
            {
                PowerPoint.PpSlideShowState state = show.State;
                if (black && state != PowerPoint.PpSlideShowState.ppSlideShowBlackScreen)
                {
                    show.State = PowerPoint.PpSlideShowState.ppSlideShowBlackScreen;
                    beforeBlack = state == PowerPoint.PpSlideShowState.ppSlideShowPaused
                        ? PowerPoint.PpSlideShowState.ppSlideShowPaused
                        : PowerPoint.PpSlideShowState.ppSlideShowRunning;
                }
                else if (!black && (state == PowerPoint.PpSlideShowState.ppSlideShowBlackScreen ||
                    state == PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen))
                {
                    show.State = beforeBlack ?? PowerPoint.PpSlideShowState.ppSlideShowRunning;
                    beforeBlack = null;
                }
                return black ? Text("Black screen.", "Écran noir.")
                    : Text("Presentation visible.", "Présentation visible.");
            }
        }

        public string GoToAlias(string name)
        {
            name = ValidateAliasName(name);
            using (PresentationConnection show = Open())
            {
                string path = RequireSavedPath(show);
                SlideAlias alias = aliasStore.Load(path).Find(item =>
                    string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
                if (alias == null)
                    throw Error("No alias named \"{0}\" exists for this presentation.",
                        "Aucun alias « {0} » pour cette présentation.", name);
                PresentationSlide target = RequireSlide(show, alias.SlideId);
                Navigate(show, target);
                return SlideMessage(target);
            }
        }

        public List<SlideAlias> GetAliases()
        {
            using (PresentationConnection show = Open())
            {
                List<SlideAlias> aliases = aliasStore.Load(RequireSavedPath(show));
                foreach (SlideAlias alias in aliases)
                {
                    PresentationSlide target = show.FindSlide(alias.SlideId);
                    alias.SlideNumber = target == null ? 0 : target.Number;
                    if (target != null) alias.Title = target.Title;
                }
                aliases.Sort((left, right) => StringComparer.CurrentCultureIgnoreCase.Compare(left.Name, right.Name));
                return aliases;
            }
        }

        public void SaveAlias(string name)
        {
            name = ValidateAliasName(name);
            using (PresentationConnection show = Open())
            {
                string path = RequireSavedPath(show);
                PresentationSlide current = show.CurrentSlide();
                List<SlideAlias> aliases = aliasStore.Load(path);
                SlideAlias alias = aliases.Find(item =>
                    string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
                if (alias == null)
                {
                    alias = new SlideAlias();
                    aliases.Add(alias);
                }
                alias.Name = name;
                alias.SlideId = current.Id;
                alias.Title = current.Title;
                aliasStore.Save(path, aliases);
            }
        }

        public void RemoveAlias(string name)
        {
            name = ValidateAliasName(name);
            using (PresentationConnection show = Open())
            {
                string path = RequireSavedPath(show);
                List<SlideAlias> aliases = aliasStore.Load(path);
                if (aliases.RemoveAll(item =>
                    string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)) == 0)
                    throw Error("No alias named \"{0}\" exists for this presentation.",
                        "Aucun alias « {0} » pour cette présentation.", name);
                aliasStore.Save(path, aliases);
            }
        }

        private PresentationConnection Open()
        {
            ThrowIfDisposed();
            PresentationConnection show;
            try
            {
                show = connect();
            }
            catch (PresentationUnavailableException)
            {
                Reset();
                throw;
            }
            bool success = false;
            try
            {
                string currentIdentity = show.Identity;
                string currentPresentationKey = show.PresentationKey;
                float currentElapsed = show.Elapsed;
                if (identity != currentIdentity || presentationKey != currentPresentationKey ||
                    currentElapsed + 0.1f < elapsed)
                {
                    Reset();
                    identity = currentIdentity;
                    sessionKey = Guid.NewGuid().ToString("N");
                    presentationKey = currentPresentationKey;
                }
                elapsed = currentElapsed;
                if (searchIndex != null && (searchIndex.Version != show.SearchVersion ||
                    searchIndex.Documents.Count != show.SlideCount))
                    InvalidateSearchIndex();
                success = true;
                return show;
            }
            finally
            {
                if (!success) show.Dispose();
            }
        }

        private void Reset()
        {
            InvalidateSearchIndex();
            identity = null;
            sessionKey = null;
            presentationKey = null;
            elapsed = 0;
            ClearHistory();
            activeRoute = null;
            beforeBlack = null;
        }

        private void ClearHistory()
        {
            returnPoint = null;
            results = null;
            resultIndex = 0;
            QuestionsMode = false;
        }

        private string Move(bool next)
        {
            using (PresentationConnection show = Open())
            {
                if (activeRoute != null && !QuestionsMode)
                {
                    int currentIndex = activeRoute.SlideIds.IndexOf(show.CurrentSlide().Id);
                    if (currentIndex >= 0)
                    {
                        int direction = next ? 1 : -1;
                        for (int index = currentIndex + direction; index >= 0 && index < activeRoute.SlideIds.Count; index += direction)
                        {
                            PresentationSlide target = show.FindSlide(activeRoute.SlideIds[index]);
                            if (target == null) continue;
                            // Route navigation deliberately skips whole slides, not individual animations.
                            show.GoTo(target.Id);
                            searchGeneration++;
                            return SlideMessage(target);
                        }
                        return next
                            ? Text("End of route \"{0}\". The slide show remains open.", "Fin du parcours « {0} ». Le diaporama reste ouvert.", activeRoute.Name)
                            : Text("Start of route \"{0}\".", "Début du parcours « {0} ».", activeRoute.Name);
                    }
                    if (!activeRoute.SlideIds.Exists(id => show.FindSlide(id) != null))
                        throw Error("All slides in this route were deleted. Edit or deactivate the route.",
                            "Toutes les diapositives de ce parcours ont été supprimées. Modifiez ou désactivez le parcours.");
                }
                show.Move(next);
                searchGeneration++;
                return next ? Text("Next slide.", "Diapositive suivante.")
                    : Text("Previous slide.", "Diapositive précédente.");
            }
        }

        private string ValidateRouteName(string name)
        {
            try
            {
                return PresentationProfileStore.NormalizeRouteName(name);
            }
            catch (ArgumentException)
            {
                throw Error("Use a route name of 1 to 80 characters, without control characters.",
                    "Utilisez un nom de parcours de 1 à 80 caractères, sans caractères de contrôle.");
            }
        }

        private static SlideChoice ToChoice(PresentationSlide slide)
        {
            return new SlideChoice { SlideId = slide.Id, SlideNumber = slide.Number, Title = slide.Title, Score = slide.Score };
        }

        private void Navigate(PresentationConnection show, PresentationSlide target)
        {
            ReturnPoint candidate = returnPoint == null ? Capture(show) : null;
            show.GoTo(target.Id);
            searchGeneration++;
            if (candidate != null && candidate.SlideId != target.Id)
                returnPoint = candidate;
        }

        private ReturnPoint Capture(PresentationConnection show)
        {
            return new ReturnPoint
            {
                SlideId = show.CurrentSlide().Id,
                ClickIndex = show.ClickIndex,
                State = show.State,
                BeforeBlack = beforeBlack
            };
        }

        private PresentationSlide RequireSlide(PresentationConnection show, int id)
        {
            PresentationSlide slide = show.FindSlide(id);
            if (slide == null)
                throw Error("The target slide was deleted. Choose another slide or remove its alias.",
                    "La diapositive cible a été supprimée. Choisissez une autre diapositive ou supprimez son alias.");
            return slide;
        }

        private string RequireSavedPath(PresentationConnection show)
        {
            if (string.IsNullOrEmpty(show.SavedPath))
                throw Error("Save this presentation in PowerPoint before using aliases or presentation profiles.",
                    "Enregistrez cette présentation dans PowerPoint avant d’utiliser les alias ou les profils.");
            return show.SavedPath;
        }

        private string ValidateAliasName(string name)
        {
            name = name == null ? string.Empty : name.Trim();
            if (name.Length < 1 || name.Length > 80)
                throw Error("Use an alias name of 1 to 80 characters.",
                    "Utilisez un nom d’alias de 1 à 80 caractères.");
            foreach (char character in name)
                if (char.IsControl(character))
                    throw Error("An alias name cannot contain control characters.",
                        "Un nom d’alias ne peut pas contenir de caractères de contrôle.");
            try { System.Xml.XmlConvert.VerifyXmlChars(name); }
            catch (System.Xml.XmlException)
            {
                throw Error("The alias name contains an invalid character.",
                    "Le nom d’alias contient un caractère invalide.");
            }
            if (string.IsNullOrEmpty(PowerPointController.NormalizeForSearch(name)))
                throw Error("An alias name must contain a word or number.",
                    "Un nom d’alias doit contenir un mot ou un nombre.");
            return name;
        }

        private string SlideMessage(PresentationSlide slide)
        {
            return Text("Slide {0}", "Diapositive {0}", slide.Number) +
                (string.IsNullOrWhiteSpace(slide.Title) ? "." : " — " + slide.Title);
        }

        private string Text(string english, string french, params object[] values)
        {
            return string.Format(CultureInfo.CurrentCulture, English ? english : french, values);
        }

        private InvalidOperationException Error(string english, string french, params object[] values)
        {
            return new InvalidOperationException(Text(english, french, values));
        }

        private sealed class ReturnPoint
        {
            public int SlideId;
            public int? ClickIndex;
            public PowerPoint.PpSlideShowState State;
            public PowerPoint.PpSlideShowState? BeforeBlack;
        }

        private sealed class PreparedSearch
        {
            public string SessionKey;
            public int Origin;
            public List<PresentationSlide> Matches;
            public long Generation;
            public string Version;
            public List<SearchDocument> Documents;
        }
    }

    // This short-lived boundary also lets the session invariants be tested without touching PowerPoint.
    internal abstract class PresentationConnection : IDisposable
    {
        public abstract string Identity { get; }
        public abstract string PresentationKey { get; }
        public abstract string PresentationName { get; }
        public abstract string SavedPath { get; }
        public virtual IntPtr WindowHandle { get { return IntPtr.Zero; } }
        public abstract int SlideCount { get; }
        public abstract float Elapsed { get; }
        public abstract int? ClickIndex { get; }
        public abstract PowerPoint.PpSlideShowState State { get; set; }
        public abstract PresentationSlide CurrentSlide();
        public virtual int CurrentSlideId { get { return CurrentSlide().Id; } }
        public abstract PresentationSlide SlideAt(int number);
        public virtual int SlideIdAt(int number) { return SlideAt(number).Id; }
        public abstract PresentationSlide FindSlide(int id);
        public abstract List<PresentationSlide> Search(string query);
        public virtual bool SupportsSearchIndex { get { return false; } }
        public virtual string SearchVersion { get { return null; } }
        public virtual SearchSlideReader OpenSearchSlide(int number) { throw new NotSupportedException(); }
        public abstract void GoTo(int id);
        public abstract void Move(bool next);
        public abstract bool RestoreClick(int? index);
        public abstract void Dispose();
    }

    internal sealed class PresentationSlide
    {
        public int Id;
        public int Number;
        public string Title;
        public int Score;
    }

    internal sealed class PowerPointConnection : PresentationConnection
    {
        private static readonly string WindowProperty = "JarvisPowerPoint.Session." + Guid.NewGuid().ToString("N");
        private static int nextWindowToken;
        private PowerPoint.Application application;
        private PowerPoint.SlideShowWindows windows;
        private PowerPoint.SlideShowWindow window;
        private PowerPoint.SlideShowView view;
        private PowerPoint.Presentation presentation;
        private PowerPoint.Slides slides;
        private readonly string identity;
        private readonly string savedPath;
        private readonly string presentationKey;
        private readonly bool english;

        public PowerPointConnection(bool english)
        {
            this.english = english;
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException(english
                    ? "PowerPoint connections must be opened on the UI STA thread."
                    : "Les connexions PowerPoint doivent être ouvertes sur le thread STA de l’interface.");
            bool success = false;
            try
            {
                try
                {
                    application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                }
                catch (COMException exception)
                {
                    if (exception.ErrorCode != unchecked((int)0x800401E3)) throw;
                    throw Unavailable();
                }
                windows = application.SlideShowWindows;
                if (windows.Count == 0) throw Unavailable();
                if (windows.Count != 1)
                    throw new PresentationUnavailableException(english
                        ? "Several slide shows are running. Keep only one running before using presentation controls."
                        : "Plusieurs diaporamas sont en cours. Gardez un seul diaporama actif pour utiliser les commandes.",
                        PresentationUnavailableReason.MultipleSlideShows);
                window = windows[1];
                view = window.View;
                if (view.State == PowerPoint.PpSlideShowState.ppSlideShowDone) throw Unavailable();
                presentation = window.Presentation;
                slides = presentation.Slides;
                savedPath = string.IsNullOrEmpty(presentation.Path) ? null : presentation.FullName;
                string windowIdentity = GetWindowIdentity(window.HWND);
                identity = windowIdentity + ":" + presentation.Name;
                presentationKey = savedPath == null ? "unsaved:" + windowIdentity : AliasStore.PathKey(savedPath);
                success = true;
            }
            finally
            {
                if (!success) Dispose();
            }
        }

        public override string Identity { get { return identity; } }
        public override string PresentationKey { get { return presentationKey; } }
        public override string PresentationName { get { return presentation.Name; } }
        public override string SavedPath { get { return savedPath; } }
        public override IntPtr WindowHandle { get { return new IntPtr(window.HWND); } }
        public override int SlideCount { get { return slides.Count; } }
        public override float Elapsed { get { return view.PresentationElapsedTime; } }
        public override bool SupportsSearchIndex { get { return true; } }
        public override string SearchVersion
        {
            get
            {
                // PowerPoint exposes no reliable edit revision. Never reuse dirty/unsaved
                // text: Saved=false remains false through any number of subsequent edits.
                if (presentation.Saved != Office.MsoTriState.msoTrue || savedPath == null) return null;
                try
                {
                    var file = new System.IO.FileInfo(savedPath);
                    if (!file.Exists) return null;
                    return file.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture) + ":" +
                        file.Length.ToString(CultureInfo.InvariantCulture);
                }
                catch (System.IO.IOException) { return null; }
                catch (UnauthorizedAccessException) { return null; }
            }
        }
        public override PowerPoint.PpSlideShowState State
        {
            get { return view.State; }
            set { view.State = value; }
        }

        public override int? ClickIndex
        {
            get
            {
                try
                {
                    int index = view.GetClickIndex();
                    return index < 0 ? (int?)null : index;
                }
                catch (COMException exception)
                {
                    if (!UnsupportedAnimationApi(exception)) throw;
                    return null;
                }
            }
        }

        public override PresentationSlide CurrentSlide()
        {
            PowerPoint.Slide slide = null;
            try
            {
                slide = view.Slide;
                return Describe(slide);
            }
            finally { Release(slide); }
        }

        public override int CurrentSlideId
        {
            get
            {
                PowerPoint.Slide slide = null;
                try { slide = view.Slide; return slide.SlideID; }
                finally { Release(slide); }
            }
        }

        public override int SlideIdAt(int number)
        {
            PowerPoint.Slide slide = null;
            try { slide = slides[number]; return slide.SlideID; }
            finally { Release(slide); }
        }

        public override SearchSlideReader OpenSearchSlide(int number)
        {
            PowerPoint.Slide slide = slides[number];
            try { return new PowerPointSearchSlideReader(slide); }
            catch { Release(slide); throw; }
        }

        public override PresentationSlide SlideAt(int number)
        {
            PowerPoint.Slide slide = null;
            try
            {
                slide = slides[number];
                return Describe(slide);
            }
            finally { Release(slide); }
        }

        public override PresentationSlide FindSlide(int id)
        {
            PowerPoint.Slide found = null;
            try
            {
                try { found = slides.FindBySlideID(id); }
                catch (COMException exception)
                {
                    // PowerPoint uses different missing-ID errors across versions. Preserve
                    // deleted-slide behavior, but never hide busy/disconnected COM.
                    if (exception.ErrorCode == unchecked((int)0x80010001) ||
                        exception.ErrorCode == unchecked((int)0x8001010A) ||
                        exception.ErrorCode == unchecked((int)0x80010108)) throw;
                }
                if (found != null) return Describe(found);
            }
            finally { Release(found); }
            for (int number = 1; number <= slides.Count; number++)
            {
                PowerPoint.Slide slide = null;
                try
                {
                    slide = slides[number];
                    if (slide.SlideID == id) return Describe(slide);
                }
                finally { Release(slide); }
            }
            return null;
        }

        public override List<PresentationSlide> Search(string query)
        {
            List<PowerPointController.SlideSearchResult> matches = PowerPointController.FindSlides(slides, query);
            var found = new List<PresentationSlide>();
            foreach (PowerPointController.SlideSearchResult match in matches)
            {
                PresentationSlide slide = SlideAt(match.SlideNumber);
                slide.Score = match.Score;
                found.Add(slide);
            }
            return found;
        }

        public override void GoTo(int id)
        {
            PresentationSlide target = FindSlide(id);
            if (target == null)
                throw new InvalidOperationException(english
                    ? "The target slide was deleted."
                    : "La diapositive cible a été supprimée.");
            view.GotoSlide(target.Number, Office.MsoTriState.msoFalse);
        }

        public override void Move(bool next)
        {
            if (next) view.Next();
            else view.Previous();
        }

        public override bool RestoreClick(int? index)
        {
            if (!index.HasValue) return false;
            try
            {
                if (index.Value > view.GetClickCount()) return false;
                view.GotoClick(index.Value);
                return true;
            }
            catch (COMException exception)
            {
                if (!UnsupportedAnimationApi(exception)) throw;
                return false;
            }
        }

        public override void Dispose()
        {
            Release(slides); slides = null;
            Release(presentation); presentation = null;
            Release(view); view = null;
            Release(window); window = null;
            Release(windows); windows = null;
            Release(application); application = null;
        }

        private PresentationSlide Describe(PowerPoint.Slide slide)
        {
            return new PresentationSlide
            {
                Id = slide.SlideID,
                Number = slide.SlideIndex,
                Title = PowerPointController.GetSlideTitle(slide)
            };
        }

        private PresentationUnavailableException Unavailable()
        {
            return new PresentationUnavailableException(english
                ? "Open PowerPoint and start a slide show."
                : "Ouvrez PowerPoint et démarrez un diaporama.");
        }

        private static string GetWindowIdentity(int handle)
        {
            IntPtr windowHandle = new IntPtr(handle);
            IntPtr token = GetProp(windowHandle, WindowProperty);
            if (token == IntPtr.Zero)
            {
                token = new IntPtr(Interlocked.Increment(ref nextWindowToken));
                if (!SetProp(windowHandle, WindowProperty, token))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            // COM proxy addresses change after release and HWNDs can be reused. A native
            // window property survives RCW release but is removed automatically on HWND destruction.
            return handle.ToString(CultureInfo.InvariantCulture) + ":" +
                token.ToInt64().ToString(CultureInfo.InvariantCulture);
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetProp(IntPtr window, string name);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetProp(IntPtr window, string name, IntPtr value);

        private static bool UnsupportedAnimationApi(COMException exception)
        {
            return exception.ErrorCode == unchecked((int)0x80020003) ||
                exception.ErrorCode == unchecked((int)0x80020006) ||
                exception.ErrorCode == unchecked((int)0x80004001);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
