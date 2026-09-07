using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace JarvisPowerPoint
{
    internal sealed class RehearsalEntry
    {
        public int SlideId { get; set; }
        public int SlideNumber { get; set; }
        public string Title { get; set; }
        public double Seconds { get; set; }
        public int BudgetSeconds { get; set; }
        public bool OverBudget { get { return Seconds > BudgetSeconds; } }
    }

    internal sealed class RehearsalTracker
    {
        private readonly Dictionary<int, RehearsalEntry> entries = new Dictionary<int, RehearsalEntry>();
        private string sessionKey;
        private int activeSlideId;
        private bool counting;
        private TimeSpan lastObservation;
        private Dictionary<int, int> savedBudgets = new Dictionary<int, int>();

        public bool IsRunning { get; private set; }
        public int DefaultBudgetSeconds { get; set; }
        public string PresentationName { get; private set; }
        public string SavedPath { get; private set; }
        public string PresentationKey { get; private set; }
        public string RunId { get; private set; }
        public DateTime StartedUtc { get; private set; }
        public double TotalSeconds { get { return entries.Values.Sum(entry => entry.Seconds); } }
        public IList<RehearsalEntry> Entries
        {
            get { return entries.Values.OrderBy(entry => entry.SlideNumber).ToList().AsReadOnly(); }
        }

        public RehearsalTracker()
        {
            DefaultBudgetSeconds = 90;
        }

        public void Start(PresentationSnapshot snapshot, TimeSpan now)
        {
            Start(snapshot, now, null);
        }

        public void Start(PresentationSnapshot snapshot, TimeSpan now, IDictionary<int, int> budgets)
        {
            if (snapshot == null) { throw new ArgumentNullException("snapshot"); }
            if (DefaultBudgetSeconds < 1) { throw new InvalidOperationException("The slide budget must be positive."); }
            var validated = new Dictionary<int, int>();
            if (budgets != null)
            {
                foreach (var entry in budgets)
                {
                    if (entry.Key <= 0 || entry.Value < 1 || entry.Value > 3600) { throw new ArgumentException("Invalid saved slide budget.", "budgets"); }
                    validated.Add(entry.Key, entry.Value);
                }
            }
            savedBudgets = validated;
            entries.Clear();
            sessionKey = snapshot.SessionKey;
            PresentationName = snapshot.PresentationName;
            SavedPath = snapshot.SavedPath;
            PresentationKey = snapshot.PresentationKey;
            RunId = Guid.NewGuid().ToString("N");
            StartedUtc = DateTime.UtcNow;
            lastObservation = now;
            activeSlideId = 0;
            counting = false;
            IsRunning = true;
            Observe(snapshot, now);
        }

        public void Observe(PresentationSnapshot snapshot, TimeSpan now)
        {
            if (!IsRunning) { return; }
            Accumulate(now);
            if (snapshot == null || snapshot.SessionKey != sessionKey)
            {
                IsRunning = false;
                counting = false;
                return;
            }

            RehearsalEntry entry;
            if (!entries.TryGetValue(snapshot.SlideId, out entry))
            {
                int budget;
                if (!savedBudgets.TryGetValue(snapshot.SlideId, out budget)) { budget = DefaultBudgetSeconds; }
                entry = new RehearsalEntry { SlideId = snapshot.SlideId, BudgetSeconds = budget };
                entries.Add(entry.SlideId, entry);
            }
            entry.SlideNumber = snapshot.SlideNumber;
            entry.Title = snapshot.Title;
            activeSlideId = entry.SlideId;
            counting = !snapshot.IsBlack;
        }

        public void Stop(TimeSpan now)
        {
            if (IsRunning) { Accumulate(now); }
            IsRunning = false;
            counting = false;
        }

        public void SetBudget(int slideId, int seconds)
        {
            if (seconds < 1 || seconds > 3600) { throw new ArgumentOutOfRangeException("seconds"); }
            RehearsalEntry entry;
            if (!entries.TryGetValue(slideId, out entry)) { throw new InvalidOperationException("Select a recorded slide first."); }
            entry.BudgetSeconds = seconds;
            savedBudgets[slideId] = seconds;
        }

        private void Accumulate(TimeSpan now)
        {
            if (now < lastObservation) { throw new ArgumentOutOfRangeException("now", "Rehearsal time must be monotonic."); }
            if (counting) { entries[activeSlideId].Seconds += (now - lastObservation).TotalSeconds; }
            lastObservation = now;
        }

        public void Export(string path)
        {
            using (var writer = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("Presentation,Slide,Title,Seconds,BudgetSeconds,OverBudget");
                foreach (RehearsalEntry entry in Entries)
                {
                    writer.WriteLine(string.Join(",", new[]
                    {
                        CsvText(PresentationName),
                        entry.SlideNumber.ToString(CultureInfo.InvariantCulture),
                        CsvText(entry.Title),
                        entry.Seconds.ToString("F1", CultureInfo.InvariantCulture),
                        entry.BudgetSeconds.ToString(CultureInfo.InvariantCulture),
                        entry.OverBudget ? "true" : "false"
                    }));
                }
            }
        }

        private static string CsvText(string value)
        {
            value = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
            string trimmed = value.TrimStart();
            if (trimmed.Length > 0 && "=+-@\t".IndexOf(trimmed[0]) >= 0) { value = "'" + value; }
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
