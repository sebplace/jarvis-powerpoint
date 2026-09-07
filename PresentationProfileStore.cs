using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace JarvisPowerPoint
{
    internal sealed class SlideChoice
    {
        public int SlideId { get; set; }
        public int SlideNumber { get; set; }
        public int Score { get; set; }
        public string Title { get; set; }
    }

    internal sealed class SearchProposal
    {
        public string SessionKey { get; set; }
        public int OriginSlideId { get; set; }
        public List<SlideChoice> Candidates { get; set; }
        public bool IsAmbiguous { get; set; }
    }

    internal sealed class SlideRoute
    {
        public SlideRoute() { SlideIds = new List<int>(); }
        public string Name { get; set; }
        public int TargetMinutes { get; set; }
        public List<int> SlideIds { get; set; }
    }

    internal sealed class SlideTiming
    {
        public int SlideId { get; set; }
        public double Seconds { get; set; }
        public int BudgetSeconds { get; set; }
    }

    internal sealed class RehearsalRecord
    {
        public RehearsalRecord() { Slides = new List<SlideTiming>(); }
        public string Id { get; set; }
        public DateTime StartedUtc { get; set; }
        public List<SlideTiming> Slides { get; set; }
    }

    internal sealed class PresentationProfileStore
    {
        private const int MaximumRuns = 20;
        private const long MaximumCharacters = 4 * 1024 * 1024;
        private static readonly object SyncRoot = new object();
        private readonly string directory;

        public PresentationProfileStore(string dataDirectory)
        {
            if (string.IsNullOrWhiteSpace(dataDirectory))
                throw new ArgumentException("A local data directory is required.", "dataDirectory");
            directory = Path.Combine(Path.GetFullPath(dataDirectory), "Profiles");
        }

        public Dictionary<int, int> LoadBudgets(string savedPath)
        {
            lock (SyncRoot) return Read(savedPath).Budgets;
        }

        public void SaveBudgets(string savedPath, IDictionary<int, int> budgets)
        {
            if (budgets == null) throw new ArgumentNullException("budgets");
            var copy = new Dictionary<int, int>();
            foreach (KeyValuePair<int, int> budget in budgets)
            {
                ValidateBudget(budget.Key, budget.Value);
                copy.Add(budget.Key, budget.Value);
            }
            lock (SyncRoot)
            {
                Profile profile = Read(savedPath);
                profile.Budgets = copy;
                Write(savedPath, profile);
            }
        }

        public List<SlideRoute> LoadRoutes(string savedPath)
        {
            lock (SyncRoot) return Read(savedPath).Routes;
        }

        public void SaveRoutes(string savedPath, IList<SlideRoute> routes)
        {
            if (routes == null) throw new ArgumentNullException("routes");
            var copies = new List<SlideRoute>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SlideRoute route in routes)
            {
                SlideRoute copy = CopyRoute(route);
                if (!names.Add(copy.Name)) throw new ArgumentException("Route names must be unique.", "routes");
                copies.Add(copy);
            }
            lock (SyncRoot)
            {
                Profile profile = Read(savedPath);
                profile.Routes = copies;
                Write(savedPath, profile);
            }
        }

        public void SaveRun(string savedPath, RehearsalRecord record)
        {
            RehearsalRecord copy = CopyRun(record);
            lock (SyncRoot)
            {
                Profile profile = Read(savedPath);
                profile.Runs.RemoveAll(run => string.Equals(run.Id, copy.Id, StringComparison.Ordinal));
                profile.Runs.Add(copy);
                SortRuns(profile.Runs);
                if (profile.Runs.Count > MaximumRuns)
                    profile.Runs.RemoveRange(MaximumRuns, profile.Runs.Count - MaximumRuns);
                Write(savedPath, profile);
            }
        }

        public List<RehearsalRecord> LoadRuns(string savedPath)
        {
            lock (SyncRoot) return Read(savedPath).Runs;
        }

        internal static SlideRoute CopyRoute(SlideRoute route)
        {
            if (route == null) throw new ArgumentNullException("route");
            var copy = new SlideRoute { Name = ValidateName(route.Name, "route.Name"), TargetMinutes = route.TargetMinutes };
            if (copy.TargetMinutes < 1 || copy.TargetMinutes > 240)
                throw new ArgumentException("A route target must be between 1 and 240 minutes.", "route");
            if (route.SlideIds == null || route.SlideIds.Count == 0)
                throw new ArgumentException("A route must contain at least one slide.", "route");
            var ids = new HashSet<int>();
            foreach (int id in route.SlideIds)
            {
                if (id <= 0 || !ids.Add(id))
                    throw new ArgumentException("Route slide IDs must be positive and unique.", "route");
                copy.SlideIds.Add(id);
            }
            return copy;
        }

        internal static string NormalizeRouteName(string name)
        {
            return ValidateName(name, "route.Name");
        }

        private static RehearsalRecord CopyRun(RehearsalRecord record)
        {
            if (record == null) throw new ArgumentNullException("record");
            var copy = new RehearsalRecord { Id = ValidateName(record.Id, "record.Id"), StartedUtc = record.StartedUtc };
            if (record.StartedUtc.Kind != DateTimeKind.Utc)
                throw new ArgumentException("The rehearsal start must be a UTC date.", "record");
            if (record.Slides == null)
                throw new ArgumentException("Rehearsal slide timings are required.", "record");
            var ids = new HashSet<int>();
            foreach (SlideTiming slide in record.Slides)
            {
                if (slide == null) throw new ArgumentException("A rehearsal timing cannot be null.", "record");
                ValidateBudget(slide.SlideId, slide.BudgetSeconds);
                if (!ids.Add(slide.SlideId) || double.IsNaN(slide.Seconds) ||
                    double.IsInfinity(slide.Seconds) || slide.Seconds < 0)
                    throw new ArgumentException("Rehearsal durations must be finite and nonnegative, with unique slide IDs.", "record");
                copy.Slides.Add(new SlideTiming
                {
                    SlideId = slide.SlideId,
                    Seconds = slide.Seconds,
                    BudgetSeconds = slide.BudgetSeconds
                });
            }
            return copy;
        }

        private static string ValidateName(string value, string parameter)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Trim().Length > 80)
                throw new ArgumentException("Use a name or ID of 1 to 80 characters.", parameter);
            value = value.Trim();
            foreach (char character in value)
                if (char.IsControl(character))
                    throw new ArgumentException("Names and IDs cannot contain control characters.", parameter);
            try { XmlConvert.VerifyXmlChars(value); }
            catch (XmlException exception)
            {
                throw new ArgumentException("The name or ID contains an invalid XML character.", parameter, exception);
            }
            return value;
        }

        private static void ValidateBudget(int id, int seconds)
        {
            if (id <= 0 || seconds < 1 || seconds > 3600)
                throw new ArgumentException("Slide IDs must be positive and budgets must be between 1 and 3600 seconds.");
        }

        private Profile Read(string savedPath)
        {
            string path = FilePath(savedPath);
            var profile = new Profile();
            // File.GetAttributes, unlike Exists, does not hide denied access.
            FileAttributes attributes;
            try { attributes = File.GetAttributes(directory); }
            catch (FileNotFoundException) { return profile; }
            catch (DirectoryNotFoundException) { return profile; }
            if ((attributes & FileAttributes.Directory) == 0)
                throw new IOException("The profile directory is a file.");
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaximumCharacters
            };
            var document = new XmlDocument { XmlResolver = null };
            try
            {
                using (XmlReader reader = XmlReader.Create(path, settings)) document.Load(reader);
            }
            catch (FileNotFoundException) { return profile; }
            XmlElement root = document.DocumentElement;
            if (root == null || root.Name != "presentationProfile" || root.GetAttribute("version") != "1" ||
                !string.Equals(root.GetAttribute("presentation"), AliasStore.CanonicalPath(savedPath), StringComparison.Ordinal))
                throw new InvalidDataException("The profile does not belong to this presentation or has an unsupported version.");
            CheckShape(root, 2, false);
            var sections = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                foreach (XmlElement section in Children(root))
                {
                    CheckShape(section, 0, false);
                    if (!sections.Add(section.Name)) throw InvalidProfile();
                    if (section.Name == "budgets")
                    {
                        foreach (XmlElement element in Children(section))
                        {
                            if (element.Name != "slide") throw InvalidProfile();
                            CheckShape(element, 2, true);
                            int id = IntAttribute(element, "id");
                            int seconds = IntAttribute(element, "seconds");
                            ValidateBudget(id, seconds);
                            if (profile.Budgets.ContainsKey(id)) throw InvalidProfile();
                            profile.Budgets.Add(id, seconds);
                        }
                    }
                    else if (section.Name == "routes")
                    {
                        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        foreach (XmlElement element in Children(section))
                        {
                            if (element.Name != "route") throw InvalidProfile();
                            CheckShape(element, 2, false);
                            var route = new SlideRoute
                            {
                                Name = element.GetAttribute("name"),
                                TargetMinutes = IntAttribute(element, "minutes")
                            };
                            foreach (XmlElement slide in Children(element))
                            {
                                if (slide.Name != "slide") throw InvalidProfile();
                                CheckShape(slide, 1, true);
                                route.SlideIds.Add(IntAttribute(slide, "id"));
                            }
                            route = CopyRoute(route);
                            if (!names.Add(route.Name)) throw InvalidProfile();
                            profile.Routes.Add(route);
                        }
                    }
                    else if (section.Name == "runs")
                    {
                        var ids = new HashSet<string>(StringComparer.Ordinal);
                        foreach (XmlElement element in Children(section))
                        {
                            if (element.Name != "run" || profile.Runs.Count >= MaximumRuns) throw InvalidProfile();
                            CheckShape(element, 2, false);
                            DateTime started;
                            if (!DateTime.TryParseExact(element.GetAttribute("startedUtc"), "o", CultureInfo.InvariantCulture,
                                DateTimeStyles.RoundtripKind, out started)) throw InvalidProfile();
                            var run = new RehearsalRecord { Id = element.GetAttribute("id"), StartedUtc = started };
                            foreach (XmlElement slide in Children(element))
                            {
                                if (slide.Name != "slide") throw InvalidProfile();
                                CheckShape(slide, 3, true);
                                double seconds;
                                if (!double.TryParse(slide.GetAttribute("seconds"), NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out seconds)) throw InvalidProfile();
                                run.Slides.Add(new SlideTiming
                                {
                                    SlideId = IntAttribute(slide, "id"),
                                    Seconds = seconds,
                                    BudgetSeconds = IntAttribute(slide, "budgetSeconds")
                                });
                            }
                            run = CopyRun(run);
                            if (!ids.Add(run.Id)) throw InvalidProfile();
                            profile.Runs.Add(run);
                        }
                    }
                    else throw InvalidProfile();
                }
            }
            catch (ArgumentException exception)
            {
                throw new InvalidDataException("The profile contains invalid values.", exception);
            }
            if (sections.Count != 3) throw InvalidProfile();
            SortRuns(profile.Runs);
            return profile;
        }

        private void Write(string savedPath, Profile profile)
        {
            string path = FilePath(savedPath);
            Directory.CreateDirectory(directory);
            string pending = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".new");
            try
            {
                using (var stream = new FileStream(pending, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings
                    {
                        Encoding = new UTF8Encoding(false), Indent = true, CloseOutput = false
                    }))
                    {
                        writer.WriteStartElement("presentationProfile");
                        writer.WriteAttributeString("version", "1");
                        writer.WriteAttributeString("presentation", AliasStore.CanonicalPath(savedPath));
                        writer.WriteStartElement("budgets");
                        var slideIds = new List<int>(profile.Budgets.Keys);
                        slideIds.Sort();
                        foreach (int id in slideIds)
                        {
                            writer.WriteStartElement("slide");
                            Attribute(writer, "id", id);
                            Attribute(writer, "seconds", profile.Budgets[id]);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                        writer.WriteStartElement("routes");
                        foreach (SlideRoute route in profile.Routes)
                        {
                            writer.WriteStartElement("route");
                            writer.WriteAttributeString("name", route.Name);
                            Attribute(writer, "minutes", route.TargetMinutes);
                            foreach (int id in route.SlideIds)
                            {
                                writer.WriteStartElement("slide");
                                Attribute(writer, "id", id);
                                writer.WriteEndElement();
                            }
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                        writer.WriteStartElement("runs");
                        foreach (RehearsalRecord run in profile.Runs)
                        {
                            writer.WriteStartElement("run");
                            writer.WriteAttributeString("id", run.Id);
                            writer.WriteAttributeString("startedUtc", run.StartedUtc.ToString("o", CultureInfo.InvariantCulture));
                            foreach (SlideTiming slide in run.Slides)
                            {
                                writer.WriteStartElement("slide");
                                Attribute(writer, "id", slide.SlideId);
                                writer.WriteAttributeString("seconds", slide.Seconds.ToString("R", CultureInfo.InvariantCulture));
                                Attribute(writer, "budgetSeconds", slide.BudgetSeconds);
                                writer.WriteEndElement();
                            }
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    if (stream.Length > MaximumCharacters)
                        throw new InvalidDataException("The presentation profile exceeds the size limit.");
                    stream.Flush(true);
                }
                if (File.Exists(path)) File.Replace(pending, path, null);
                else File.Move(pending, path);
            }
            finally
            {
                if (File.Exists(pending)) File.Delete(pending);
            }
        }

        private string FilePath(string savedPath)
        {
            if (string.IsNullOrWhiteSpace(savedPath))
                throw new InvalidOperationException("Save the presentation before using its profile.");
            return Path.Combine(directory, AliasStore.PathKey(savedPath) + ".xml");
        }

        private static IEnumerable<XmlElement> Children(XmlElement parent)
        {
            foreach (XmlNode node in parent.ChildNodes)
            {
                XmlElement element = node as XmlElement;
                if (element != null) yield return element;
                else if (node.NodeType != XmlNodeType.Comment && !string.IsNullOrWhiteSpace(node.Value))
                    throw InvalidProfile();
            }
        }

        private static int IntAttribute(XmlElement element, string name)
        {
            int value;
            if (!int.TryParse(element.GetAttribute(name), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out value)) throw InvalidProfile();
            return value;
        }

        private static void CheckShape(XmlElement element, int attributes, bool leaf)
        {
            if (element.Attributes.Count != attributes) throw InvalidProfile();
            if (leaf)
                foreach (XmlElement child in Children(element)) throw InvalidProfile();
        }

        private static InvalidDataException InvalidProfile()
        {
            return new InvalidDataException("The presentation profile has an invalid structure or value.");
        }

        private static void Attribute(XmlWriter writer, string name, int value)
        {
            writer.WriteAttributeString(name, value.ToString(CultureInfo.InvariantCulture));
        }

        private static void SortRuns(List<RehearsalRecord> runs)
        {
            runs.Sort((left, right) =>
            {
                int date = right.StartedUtc.CompareTo(left.StartedUtc);
                return date != 0 ? date : StringComparer.Ordinal.Compare(left.Id, right.Id);
            });
        }

        private sealed class Profile
        {
            public Dictionary<int, int> Budgets = new Dictionary<int, int>();
            public List<SlideRoute> Routes = new List<SlideRoute>();
            public List<RehearsalRecord> Runs = new List<RehearsalRecord>();
        }
    }
}
