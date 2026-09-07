using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace JarvisPowerPoint
{
    internal enum SearchPhase { Initializing, Reading, Validating, Scoring, Complete }

    // Driven by the owning STA's WinForms timer, never a worker thread or DoEvents.
    // Extraction and saved-deck validation yield after at most 32 work units or their
    // time budget. Unversioned decks need a final uninterrupted content check.
    internal sealed class SearchOperation : IDisposable
    {
        private readonly int threadId = Thread.CurrentThread.ManagedThreadId;
        private Func<SearchOperation, int, bool> advance;
        private Action cleanup;
        private bool disposed;

        internal SearchOperation(Func<SearchOperation, int, bool> advance, Action cleanup)
        {
            this.advance = advance;
            this.cleanup = cleanup;
        }

        public int CompletedSlides { get; internal set; }
        public int TotalSlides { get; internal set; }
        public bool IsWarm { get; internal set; }
        public bool RequiresTextValidation { get; internal set; }
        public bool IsComplete { get; private set; }
        public bool IsCancelled { get; private set; }
        public SearchPhase Phase { get; internal set; }
        public SearchProposal Proposal { get; internal set; }
        public string Result { get; internal set; }

        public bool Step(int budgetMilliseconds = 15)
        {
            VerifyThread();
            if (IsCancelled) throw new OperationCanceledException();
            if (IsComplete) return true;
            if (disposed) throw new ObjectDisposedException("SearchOperation");
            try
            {
                IsComplete = advance(this, Math.Max(1, Math.Min(50, budgetMilliseconds)));
                if (IsComplete)
                {
                    Phase = SearchPhase.Complete;
                    Release();
                }
                return IsComplete;
            }
            catch
            {
                disposed = true;
                Release();
                throw;
            }
        }

        public void Cancel()
        {
            VerifyThread();
            if (IsComplete) return;
            IsCancelled = true;
            Dispose();
        }

        public void Dispose()
        {
            VerifyThread();
            disposed = true;
            Release();
        }

        private void Release()
        {
            Action action = cleanup;
            cleanup = null;
            advance = null;
            if (action != null) action();
        }

        private void VerifyThread()
        {
            if (Thread.CurrentThread.ManagedThreadId != threadId)
                throw new InvalidOperationException("Search operations must run on their creating STA thread.");
        }
    }

    internal sealed partial class PresentationSession
    {
        private long searchGeneration;
        private SearchIndex searchIndex;
        private SearchOperation activeSearchOperation;
        // About 32 MiB of UTF-16 text, including an allowance for per-slide objects. Larger decks
        // still score every slide; only the warm cache is omitted, never search data.
        internal const int MaximumSearchIndexCharacters = 16 * 1024 * 1024;

        public void InvalidateSearchIndex()
        {
            searchIndex = null;
            proposals = new System.Runtime.CompilerServices.ConditionalWeakTable<SearchProposal, PreparedSearch>();
            searchGeneration++;
        }

        public SearchOperation BeginSearch(string query, bool refresh = false)
        {
            ThrowIfDisposed();
            string normalized = ValidateSearchQuery(query);
            if (refresh) InvalidateSearchIndex();
            proposals = new System.Runtime.CompilerServices.ConditionalWeakTable<SearchProposal, PreparedSearch>();
            searchGeneration++;
            var work = new SearchWork(this, query, normalized, searchGeneration);
            return TrackSearch(work);
        }

        public SearchOperation BeginAcceptSearch(SearchProposal proposal, int selectedSlideId)
        {
            ThrowIfDisposed();
            PreparedSearch prepared;
            if (proposal == null || !proposals.TryGetValue(proposal, out prepared) ||
                prepared.Generation != searchGeneration ||
                prepared.Matches.FindIndex(match => match.Id == selectedSlideId) < 0)
                throw StaleSearch();
            var work = new SearchWork(this, proposal, prepared, selectedSlideId);
            return TrackSearch(work);
        }

        private SearchOperation TrackSearch(SearchWork work)
        {
            if (activeSearchOperation != null) activeSearchOperation.Dispose();
            SearchOperation operation = null;
            operation = new SearchOperation(work.Advance, delegate
            {
                try { work.Dispose(); }
                finally
                {
                    if (ReferenceEquals(activeSearchOperation, operation)) activeSearchOperation = null;
                }
            });
            activeSearchOperation = operation;
            return operation;
        }

        private string ValidateSearchQuery(string query)
        {
            string normalized = string.Empty;
            if (!string.IsNullOrWhiteSpace(query) && query.Length <= 500)
            {
                try { normalized = PowerPointController.NormalizeForSearch(query); }
                catch (ArgumentException)
                {
                    throw Error("The search contains an invalid character.",
                        "La recherche contient un caractère invalide.");
                }
            }
            if (normalized.Length == 0)
                throw Error("Enter a search of 1 to 500 characters containing words or numbers.",
                    "Saisissez une recherche de 1 à 500 caractères contenant des mots ou des nombres.");
            return normalized;
        }

        private InvalidOperationException StaleSearch()
        {
            return Error("The presentation or its content changed. Search again.",
                "La présentation ou son contenu a changé. Relancez la recherche.");
        }

        private bool SearchVersionMatches(PresentationConnection show, PreparedSearch prepared)
        {
            return prepared.Documents == null ||
                (show.SlideCount == prepared.Documents.Count && show.SearchVersion == prepared.Version);
        }

        private sealed class SearchIndex
        {
            public string SessionKey;
            public string Version;
            public List<SearchDocument> Documents;
        }

        private sealed class SearchWork : IDisposable
        {
            private readonly PresentationSession owner;
            private readonly string query;
            private readonly string normalizedQuery;
            private readonly string[] queryWords;
            private readonly SearchProposal accepting;
            private readonly int selectedSlideId;
            private long generation;
            private string key;
            private string version;
            private int origin;
            private int count;
            private int position;
            private bool initialized;
            private bool indexed;
            private bool verifyingText;
            private SearchSlideReader reader;
            private List<SearchDocument> documents;
            private List<SearchDocument> cacheDocuments;
            private long cacheCharacters;
            private readonly List<PresentationSlide> matches = new List<PresentationSlide>();

            public SearchWork(PresentationSession owner, string query, string normalizedQuery, long generation)
            {
                this.owner = owner;
                this.query = query;
                this.normalizedQuery = normalizedQuery;
                this.generation = generation;
                queryWords = PowerPointController.GetSearchWords(normalizedQuery);
            }

            public SearchWork(PresentationSession owner, SearchProposal proposal, PreparedSearch prepared, int selectedSlideId)
            {
                this.owner = owner;
                accepting = proposal;
                this.selectedSlideId = selectedSlideId;
                generation = prepared.Generation;
                key = prepared.SessionKey;
                version = prepared.Version;
                origin = prepared.Origin;
                documents = prepared.Documents;
            }

            public bool Advance(SearchOperation operation, int budget)
            {
                if (generation != owner.searchGeneration) throw owner.StaleSearch();
                var clock = Stopwatch.StartNew();
                using (PresentationConnection show = owner.Open())
                {
                    if (!initialized)
                    {
                        if (accepting == null)
                        {
                            // Opening the very first show establishes its session identity.
                            generation = owner.searchGeneration;
                            key = owner.sessionKey;
                            origin = show.CurrentSlideId;
                            version = show.SearchVersion;
                            count = show.SlideCount;
                            indexed = show.SupportsSearchIndex;
                            SearchIndex cached = owner.searchIndex;
                            if (cached != null && (cached.SessionKey != key || cached.Version != version ||
                                cached.Documents.Count != count || version == null))
                                owner.searchIndex = cached = null;
                            operation.IsWarm = indexed && cached != null;
                            documents = new List<SearchDocument>(count);
                            cacheDocuments = operation.IsWarm ? cached.Documents :
                                (indexed && version != null ? new List<SearchDocument>(count) : null);
                            operation.Phase = operation.IsWarm ? SearchPhase.Validating : SearchPhase.Reading;
                        }
                        else
                        {
                            indexed = documents != null;
                            count = indexed ? documents.Count : show.SlideCount;
                            operation.Phase = SearchPhase.Validating;
                            verifyingText = indexed && version == null;
                            operation.IsWarm = indexed && version != null;
                        }
                        initialized = true;
                        operation.TotalSlides = count;
                        operation.RequiresTextValidation = indexed && version == null;
                        Verify(show);
                        // Let the progress form paint before the first text extraction.
                        return false;
                    }

                    Verify(show);
                    if (count == 0) return Publish(operation, matches);
                    if (!indexed)
                    {
                        if (accepting != null) return Commit(operation);
                        return Publish(operation, show.Search(query));
                    }

                    int units = 0;
                    do
                    {
                        switch (operation.Phase)
                        {
                            case SearchPhase.Reading:
                                if (ReadSlide(show, false)) position++;
                                if (position == count)
                                {
                                    position = 0;
                                    verifyingText = version == null;
                                    operation.Phase = SearchPhase.Validating;
                                    if (verifyingText)
                                    {
                                        operation.CompletedSlides = 0;
                                        Verify(show);
                                        return false;
                                    }
                                }
                                break;
                            case SearchPhase.Validating:
                                if (verifyingText)
                                {
                                    // No timer yield between rereading unversioned content and using it.
                                    // External concurrent edits are still unsupported by PowerPoint.
                                    while (position < count)
                                        if (ReadSlide(show, true)) position++;
                                    Verify(show);
                                    return accepting != null ? Commit(operation) : CompleteSearch(operation);
                                }
                                if (position < count)
                                {
                                    SearchDocument expected = operation.IsWarm && accepting == null
                                        ? cacheDocuments[position] : documents[position];
                                    if (show.SlideIdAt(position + 1) != expected.Id) Changed();
                                    if (operation.IsWarm && accepting == null) documents.Add(expected.WithoutText());
                                    position++;
                                }
                                if (position == count)
                                {
                                    Verify(show);
                                    if (accepting != null) return Commit(operation);
                                    position = 0;
                                    operation.Phase = SearchPhase.Scoring;
                                }
                                break;
                            case SearchPhase.Scoring:
                                if (position < count)
                                {
                                    if (operation.IsWarm) Score(cacheDocuments[position]);
                                    position++;
                                }
                                if (position == count)
                                {
                                    Verify(show);
                                    return CompleteSearch(operation);
                                }
                                break;
                        }
                        operation.CompletedSlides = position;
                        units++;
                    }
                    while (units < 32 && clock.ElapsedMilliseconds < budget);
                    Verify(show);
                    return false;
                }
            }

            private bool ReadSlide(PresentationConnection show, bool verify)
            {
                if (reader == null)
                {
                    reader = show.OpenSearchSlide(position + 1);
                    return false;
                }
                if (!reader.ReadNext()) return false;
                SearchDocument document = reader.Document;
                reader.Dispose();
                reader = null;
                if (document == null || document.Number != position + 1) Changed();
                if (verify)
                {
                    if (!documents[position].SameContent(document)) Changed();
                }
                else
                {
                    Score(document);
                    documents.Add(document.WithoutText());
                    if (cacheDocuments != null)
                    {
                        cacheCharacters += 256L + document.Title.Length + document.NormalizedTitle.Length +
                            document.NormalizedContent.Length;
                        if (cacheCharacters > MaximumSearchIndexCharacters)
                            cacheDocuments = null;
                        else cacheDocuments.Add(document);
                    }
                }
                return true;
            }

            private void Score(SearchDocument document)
            {
                int score = PowerPointController.ScoreNormalizedSlide(document.NormalizedTitle,
                    document.NormalizedContent, normalizedQuery, queryWords);
                if (score > 0)
                    matches.Add(new PresentationSlide { Id = document.Id, Number = document.Number,
                        Title = document.Title, Score = score });
            }

            private void Verify(PresentationConnection show)
            {
                if (generation != owner.searchGeneration || key != owner.sessionKey ||
                    origin != show.CurrentSlideId || count != show.SlideCount ||
                    (indexed && version != show.SearchVersion) ||
                    (accepting != null && (accepting.SessionKey != key || accepting.OriginSlideId != origin)))
                    Changed();
            }

            private void Changed()
            {
                owner.InvalidateSearchIndex();
                throw owner.StaleSearch();
            }

            private bool CompleteSearch(SearchOperation operation)
            {
                matches.Sort(delegate(PresentationSlide left, PresentationSlide right)
                {
                    int score = right.Score.CompareTo(left.Score);
                    return score == 0 ? left.Number.CompareTo(right.Number) : score;
                });
                if (cacheDocuments != null)
                    owner.searchIndex = new SearchIndex { SessionKey = key, Version = version,
                        Documents = cacheDocuments };
                return Publish(operation, matches);
            }

            private bool Publish(SearchOperation operation, List<PresentationSlide> found)
            {
                if (found.Count == 0)
                    throw owner.Error("No slide matches \"{0}\".",
                        "Aucune diapositive ne correspond à « {0} ».", query);
                var proposal = new SearchProposal { SessionKey = key, OriginSlideId = origin,
                    Candidates = found.ConvertAll(ToChoice),
                    IsAmbiguous = found.Count >= 2 && found[1].Score >= found[0].Score * AmbiguityScoreRatio };
                owner.proposals.Add(proposal, new PreparedSearch { SessionKey = key, Origin = origin,
                    Matches = found, Generation = generation, Version = version, Documents = indexed ? documents : null });
                operation.Proposal = proposal;
                operation.CompletedSlides = count;
                return true;
            }

            private bool Commit(SearchOperation operation)
            {
                // No message-pump yield between the last validation and navigation.
                operation.Result = owner.CommitSearch(accepting, selectedSlideId);
                operation.Proposal = accepting;
                operation.CompletedSlides = count;
                return true;
            }

            public void Dispose()
            {
                try
                {
                    if (reader != null) reader.Dispose();
                }
                finally
                {
                    reader = null;
                    documents = null;
                    cacheDocuments = null;
                }
            }
        }
    }
}
