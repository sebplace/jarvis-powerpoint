using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace JarvisPowerPoint
{
    internal sealed class PresentationUnavailableException : InvalidOperationException
    {
        public PresentationUnavailableException(string message) : base(message) { }
    }

    internal sealed class PresentationSession
    {
        private readonly AliasStore aliasStore;
        private readonly Func<PresentationConnection> connect;
        private string identity;
        private string sessionKey;
        private string presentationKey;
        private float elapsed;
        private ReturnPoint returnPoint;
        private List<PresentationSlide> results;
        private int resultIndex;
        private PowerPoint.PpSlideShowState? beforeBlack;

        public PresentationSession(bool english, string dataDirectory)
            : this(english, dataDirectory, null) { }

        internal PresentationSession(bool english, string dataDirectory, Func<PresentationConnection> connect)
        {
            English = english;
            aliasStore = new AliasStore(dataDirectory);
            this.connect = connect ?? (() => new PowerPointConnection(English));
        }

        public bool English { get; set; }
        public bool HasReturnPoint { get { return returnPoint != null; } }
        public bool QuestionsMode { get; private set; }

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
            using (PresentationConnection show = Open()) show.Move(true);
            return Text("Next slide.", "Diapositive suivante.");
        }

        public string Previous()
        {
            using (PresentationConnection show = Open()) show.Move(false);
            return Text("Previous slide.", "Diapositive précédente.");
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
            if (string.IsNullOrWhiteSpace(query) || query.Length > 500 ||
                string.IsNullOrEmpty(PowerPointController.NormalizeForSearch(query)))
                throw Error("Enter a search of 1 to 500 characters containing words or numbers.",
                    "Saisissez une recherche de 1 à 500 caractères contenant des mots ou des nombres.");
            using (PresentationConnection show = Open())
            {
                List<PresentationSlide> matches = show.Search(query);
                if (matches.Count == 0)
                    throw Error("No slide matches \"{0}\".", "Aucune diapositive ne correspond à « {0} ».", query);
                PresentationSlide target = RequireSlide(show, matches[0].Id);
                Navigate(show, target);
                results = matches;
                resultIndex = 0;
                return SlideMessage(target) + Text(" ({0} matches)", " ({0} résultats)", matches.Count);
            }
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
                bool animationRestored = show.RestoreClick(returnPoint.ClickIndex);
                show.State = returnPoint.State;
                beforeBlack = returnPoint.BeforeBlack;
                returnPoint = null;
                QuestionsMode = false;
                results = null;
                resultIndex = 0;
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
            identity = null;
            sessionKey = null;
            presentationKey = null;
            elapsed = 0;
            returnPoint = null;
            results = null;
            resultIndex = 0;
            QuestionsMode = false;
            beforeBlack = null;
        }

        private void Navigate(PresentationConnection show, PresentationSlide target)
        {
            ReturnPoint candidate = returnPoint == null ? Capture(show) : null;
            show.GoTo(target.Id);
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
                throw Error("Save this presentation in PowerPoint before using aliases.",
                    "Enregistrez cette présentation dans PowerPoint avant d’utiliser les alias.");
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
    }

    // This short-lived boundary also lets the session invariants be tested without touching PowerPoint.
    internal abstract class PresentationConnection : IDisposable
    {
        public abstract string Identity { get; }
        public abstract string PresentationKey { get; }
        public abstract string PresentationName { get; }
        public abstract string SavedPath { get; }
        public abstract int SlideCount { get; }
        public abstract float Elapsed { get; }
        public abstract int? ClickIndex { get; }
        public abstract PowerPoint.PpSlideShowState State { get; set; }
        public abstract PresentationSlide CurrentSlide();
        public abstract PresentationSlide SlideAt(int number);
        public abstract PresentationSlide FindSlide(int id);
        public abstract List<PresentationSlide> Search(string query);
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
                        : "Plusieurs diaporamas sont en cours. Gardez un seul diaporama actif pour utiliser les commandes.");
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
        public override int SlideCount { get { return slides.Count; } }
        public override float Elapsed { get { return view.PresentationElapsedTime; } }
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
                found.Add(SlideAt(match.SlideNumber));
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
