[CmdletBinding()]
param([string]$Executable)

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $PSScriptRoot
if (-not [string]::IsNullOrWhiteSpace($Executable)) {
    $Executable = [IO.Path]::GetFullPath($Executable)
    if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) { throw "Candidate binary not found: $Executable" }
}
# Profile data lives on a local drive in production; do not race OneDrive sync in fixtures.
$testDirectory = Join-Path ([IO.Path]::GetTempPath()) ("JarvisProfiles-" + [Guid]::NewGuid().ToString("N"))
$compiler = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $compiler) { throw "The Windows .NET Framework compiler is required." }

$source = @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Xml;

namespace JarvisPowerPoint
{
    internal static class PresentationProfilesTests
    {
        private static int assertions;
        private static Assembly assembly;

        public static int Main(string[] args)
        {
            try
            {
                assembly = args.Length > 1 ? Assembly.LoadFrom(args[1]) : typeof(PresentationProfileStore).Assembly;
                Run(args[0]);
                Console.WriteLine("Passed " + assertions + " presentation profile assertions.");
                return 0;
            }
            catch (Exception exception) { Console.Error.WriteLine(exception); return 1; }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
            assertions++;
        }

        private static void Throws<T>(Action action) where T : Exception
        {
            try { action(); }
            catch (T) { assertions++; return; }
            throw new Exception("Expected " + typeof(T).Name);
        }

        private static SlideRoute Route(string name, params int[] ids)
        {
            return new SlideRoute { Name = name, TargetMinutes = 15, SlideIds = new List<int>(ids) };
        }

        private static RehearsalRecord Record(string id, DateTime started)
        {
            return new RehearsalRecord
            {
                Id = id, StartedUtc = started,
                Slides = new List<SlideTiming> { new SlideTiming { SlideId = 101, Seconds = 12.25, BudgetSeconds = 60 } }
            };
        }

        private static void Run(string directory)
        {
            string deck = Path.Combine(directory, "one", "Same deck.pptx");
            string other = Path.Combine(directory, "two", "Same deck.pptx");
            var store = new Store(directory);
            Assert(store.Budgets(deck).Count == 0 && store.Routes(deck).Count == 0 && store.Runs(deck).Count == 0, "New profiles are empty.");
            Throws<InvalidOperationException>(() => store.Budgets(null));
            Throws<InvalidOperationException>(() => store.SaveBudgets("", new Dictionary<int, int>()));
            Throws<ArgumentException>(() => store.SaveBudgets(deck, new Dictionary<int, int> { { 0, 60 } }));
            Throws<ArgumentException>(() => store.SaveBudgets(deck, new Dictionary<int, int> { { 101, 0 } }));
            Throws<ArgumentException>(() => store.SaveBudgets(deck, new Dictionary<int, int> { { 101, 3601 } }));
            Assert(!Directory.Exists(Path.Combine(directory, "Profiles")), "Invalid inputs do not create profile files.");
            store.SaveBudgets(deck, new Dictionary<int, int> { { 101, 1 }, { 202, 3600 } });
            store.SaveRoutes(deck, new List<SlideRoute> { Route("R\u00e9sum\u00e9 & <route>", 202, 101) });
            DateTime start = new DateTime(2026, 9, 7, 0, 0, 0, DateTimeKind.Utc);
            store.SaveRun(deck, Record("first", start));
            store = new Store(directory);
            Assert(store.Budgets(deck)[202] == 3600, "Budgets persist through route and history writes.");
            Assert(store.Routes(deck)[0].Name == "R\u00e9sum\u00e9 & <route>" && store.Routes(deck)[0].SlideIds[0] == 202, "Routes persist in stable-ID order with XML-safe names.");
            Assert(store.Runs(deck)[0].StartedUtc.Kind == DateTimeKind.Utc && store.Runs(deck)[0].Slides[0].Seconds == 12.25, "UTC date and fractional duration round-trip.");
            Assert(store.Budgets(other).Count == 0 && store.Routes(other).Count == 0 && store.Runs(other).Count == 0, "Same-filename decks in different folders are isolated.");
            Assert(store.Budgets(deck.ToUpperInvariant()).Count == 2, "Windows path casing is normalized.");
            store.SaveBudgets(deck, new Dictionary<int, int> { { 303, 90 } });
            Assert(store.Routes(deck).Count == 1 && store.Runs(deck).Count == 1, "Updating budgets preserves unrelated sections.");
            store.SaveRoutes(deck, new List<SlideRoute>());
            Assert(store.Budgets(deck)[303] == 90 && store.Runs(deck).Count == 1, "Removing routes preserves budgets and history.");
            Assert(!Directory.Exists(Path.GetDirectoryName(deck)), "Persistence never creates or edits a presentation file.");

            var duplicateRoutes = new List<SlideRoute> { Route("same", 101), Route("SAME", 202) };
            Throws<ArgumentException>(() => store.SaveRoutes(deck, duplicateRoutes));
            Throws<ArgumentException>(() => store.SaveRoutes(deck, new List<SlideRoute> { Route("empty") }));
            Throws<ArgumentException>(() => store.SaveRoutes(deck, new List<SlideRoute> { Route("duplicate IDs", 101, 101) }));
            Throws<ArgumentException>(() => store.SaveRoutes(deck, new List<SlideRoute> { Route("invalid ID", -1) }));
            Throws<ArgumentException>(() => store.SaveRoutes(deck, new List<SlideRoute> { Route("bad\nname", 101) }));
            Throws<ArgumentException>(() => store.SaveRoutes(deck, new List<SlideRoute> { Route(new string('n', 81), 101) }));
            SlideRoute badMinutes = Route("time", 101);
            badMinutes.TargetMinutes = 241;
            Throws<ArgumentException>(() => store.SaveRoutes(deck, new List<SlideRoute> { badMinutes }));
            badMinutes.TargetMinutes = 0;
            Throws<ArgumentException>(() => store.SaveRoutes(deck, new List<SlideRoute> { badMinutes }));
            store.SaveRoutes(deck, new List<SlideRoute> { Route("Valid", 101, 202) });
            foreach (double value in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity, -0.001 })
            {
                RehearsalRecord invalid = Record("invalid", start);
                invalid.Slides[0].Seconds = value;
                Throws<ArgumentException>(() => store.SaveRun(deck, invalid));
            }
            RehearsalRecord invalidRecord = Record("badbudget", start);
            invalidRecord.Slides[0].BudgetSeconds = 0;
            Throws<ArgumentException>(() => store.SaveRun(deck, invalidRecord));
            invalidRecord = Record("duplicates", start);
            invalidRecord.Slides.Add(invalidRecord.Slides[0]);
            Throws<ArgumentException>(() => store.SaveRun(deck, invalidRecord));
            Throws<ArgumentException>(() => store.SaveRun(deck, Record("", start)));
            Throws<ArgumentException>(() => store.SaveRun(deck, Record("local", DateTime.SpecifyKind(start, DateTimeKind.Local))));
            Assert(store.Runs(deck).Count == 1 && store.Routes(deck).Count == 1, "Rejected writes leave prior sections intact.");

            for (int index = 0; index < 25; index++) store.SaveRun(deck, Record("run-" + index, start.AddMinutes(index + 1)));
            List<RehearsalRecord> runs = store.Runs(deck);
            Assert(runs.Count == 20 && runs[0].Id == "run-24" && runs[19].Id == "run-5", "History retains the newest twenty runs.");
            RehearsalRecord replacement = Record("run-24", start.AddMinutes(25));
            replacement.Slides[0].Seconds = 0;
            store.SaveRun(deck, replacement);
            Assert(store.Runs(deck).Count == 20 && store.Runs(deck)[0].Slides[0].Seconds == 0, "Saving the same run ID updates rather than duplicates.");
            runs[0].Slides[0].Seconds = 999;
            List<SlideRoute> routes = store.Routes(deck);
            routes[0].SlideIds.Clear();
            Dictionary<int, int> budgets = store.Budgets(deck);
            budgets.Clear();
            Assert(store.Runs(deck)[0].Slides[0].Seconds == 0 && store.Routes(deck)[0].SlideIds.Count == 2 &&
                store.Budgets(deck).Count == 1, "Loaded values do not retain mutable references to persisted state.");
            store.SaveRun(other, Record("other", start));
            Assert(store.Runs(other).Count == 1 && store.Runs(deck).Count == 20, "Run history is bound to the explicitly passed origin path.");
            string cloud = "https://example.com/sites/team/Deck.pptx";
            store.SaveBudgets(cloud, new Dictionary<int, int> { { 909, 45 } });
            Assert(store.Budgets(cloud)[909] == 45, "Cloud-saved paths retain local profile storage.");

            string file = Path.Combine(directory, "Profiles", AliasStore.PathKey(deck) + ".xml");
            string xml = File.ReadAllText(file);
            Assert(Path.GetFileNameWithoutExtension(file).Length == 64, "Profile filename uses the shared SHA256 path key.");
            Assert(xml.IndexOf("title", StringComparison.OrdinalIgnoreCase) < 0 &&
                xml.IndexOf("content", StringComparison.OrdinalIgnoreCase) < 0, "No slide titles or content are persisted.");
            Assert(Directory.GetFiles(Path.Combine(directory, "Profiles"), "*.new").Length == 0, "Atomic writes leave no staging files.");
            File.WriteAllText(file, xml.Replace("presentation=\"", "presentation=\"wrong"));
            Throws<InvalidDataException>(() => store.Budgets(deck));
            Throws<InvalidDataException>(() => store.SaveBudgets(deck, new Dictionary<int, int>()));
            Assert(File.ReadAllText(file).Contains("presentation=\"wrong"), "A corrupt profile is not overwritten by updates.");
            File.WriteAllText(file, "<!DOCTYPE p [<!ENTITY ext SYSTEM \"file:///never-read\">]><p>&ext;</p>");
            Throws<XmlException>(() => store.Runs(deck));
            File.WriteAllText(file, xml.Replace("seconds=\"0\"", "seconds=\"NaN\""));
            Throws<InvalidDataException>(() => store.Runs(deck));
            File.WriteAllText(file, xml.Replace("version=\"1\"", "version=\"9\""));
            Throws<InvalidDataException>(() => store.Routes(deck));
            File.WriteAllText(file, xml.Replace("<budgets>", "<budgets unexpected=\"true\">"));
            Throws<InvalidDataException>(() => store.Budgets(deck));
            File.WriteAllText(file, xml.Replace("seconds=\"0\"", "seconds=\"0\" title=\"must not be retained\""));
            Throws<InvalidDataException>(() => store.Runs(deck));
            File.WriteAllText(file, "<presentationProfile>" + new string('x', 4 * 1024 * 1024) + "</presentationProfile>");
            Throws<XmlException>(() => store.Budgets(deck));
            File.WriteAllText(file, xml);
            Assert(store.Budgets(deck)[303] == 90, "A restored valid profile remains readable.");

            Exception backgroundFailure = null;
            Thread budgetsThread = new Thread(() =>
            {
                try
                {
                    for (int index = 0; index < 6; index++)
                        new Store(directory).SaveBudgets(deck, new Dictionary<int, int> { { 303, 100 + index } });
                }
                catch (Exception exception) { backgroundFailure = exception; }
            });
            Thread routesThread = new Thread(() =>
            {
                try
                {
                    for (int index = 0; index < 6; index++)
                        new Store(directory).SaveRoutes(deck, new List<SlideRoute> { Route("Concurrent " + index, 101, 202) });
                }
                catch (Exception exception) { backgroundFailure = exception; }
            });
            budgetsThread.Start(); routesThread.Start();
            budgetsThread.Join(); routesThread.Join();
            if (backgroundFailure != null) throw backgroundFailure;
            Assert(store.Budgets(deck)[303] == 105 && store.Routes(deck)[0].Name == "Concurrent 5" &&
                store.Runs(deck).Count == 20, "Independent store instances preserve concurrent section updates within the process.");
            string blocked = Path.Combine(directory, "blocked");
            Directory.CreateDirectory(blocked);
            File.WriteAllText(Path.Combine(blocked, "Profiles"), "unchanged");
            Throws<IOException>(() => new Store(blocked).SaveBudgets(deck, new Dictionary<int, int>()));
            Assert(File.ReadAllText(Path.Combine(blocked, "Profiles")) == "unchanged", "IO failures do not overwrite unrelated files.");
        }

        // The same assertions can exercise either compiled source or a selected candidate binary.
        private sealed class Store
        {
            private readonly Type type;
            private readonly object instance;
            public Store(string directory)
            {
                type = assembly.GetType("JarvisPowerPoint.PresentationProfileStore", true);
                instance = Activator.CreateInstance(type, new object[] { directory });
            }
            public Dictionary<int, int> Budgets(string path) { return (Dictionary<int, int>)Call("LoadBudgets", path); }
            public List<SlideRoute> Routes(string path) { return (List<SlideRoute>)ConvertValue(typeof(List<SlideRoute>), Call("LoadRoutes", path)); }
            public List<RehearsalRecord> Runs(string path) { return (List<RehearsalRecord>)ConvertValue(typeof(List<RehearsalRecord>), Call("LoadRuns", path)); }
            public void SaveBudgets(string path, IDictionary<int, int> budgets) { Call("SaveBudgets", path, budgets); }
            public void SaveRoutes(string path, List<SlideRoute> routes) { Call("SaveRoutes", path, routes); }
            public void SaveRun(string path, RehearsalRecord record) { Call("SaveRun", path, record); }
            private object Call(string name, params object[] values)
            {
                MethodInfo method = type.GetMethod(name);
                ParameterInfo[] parameters = method.GetParameters();
                for (int index = 0; index < values.Length; index++)
                    values[index] = ConvertValue(parameters[index].ParameterType, values[index]);
                try { return method.Invoke(instance, values); }
                catch (TargetInvocationException exception) { throw exception.InnerException; }
            }
        }

        private static object ConvertValue(Type target, object value)
        {
            if (value == null || target.IsInstanceOfType(value)) return value;
            if (target.IsGenericType && typeof(IEnumerable).IsAssignableFrom(target))
            {
                Type itemType = target.GetGenericArguments()[0];
                IList list = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(itemType));
                foreach (object item in (IEnumerable)value) list.Add(ConvertValue(itemType, item));
                return list;
            }
            object result = Activator.CreateInstance(target, true);
            foreach (PropertyInfo property in target.GetProperties())
                property.SetValue(result, ConvertValue(property.PropertyType,
                    value.GetType().GetProperty(property.Name).GetValue(value, null)), null);
            return result;
        }
    }
}
'@

try {
    New-Item -ItemType Directory -Path $testDirectory | Out-Null
    $sourcePath = Join-Path $testDirectory "ProfilesTests.cs"
    $testExecutable = Join-Path $testDirectory "ProfilesTests.exe"
    [IO.File]::WriteAllText($sourcePath, $source, [Text.UTF8Encoding]::new($false))
    & $compiler /nologo /target:exe /langversion:5 /main:JarvisPowerPoint.PresentationProfilesTests `
        "/out:$testExecutable" /reference:System.dll /reference:System.Core.dll /reference:System.Xml.dll `
        /reference:System.Runtime.Serialization.dll `
        (Join-Path $projectDirectory "PresentationProfileStore.cs") `
        (Join-Path $projectDirectory "AliasStore.cs") (Join-Path $projectDirectory "PresentationContracts.cs") $sourcePath
    if ($LASTEXITCODE -ne 0) { throw "Profile test compilation failed ($LASTEXITCODE)." }
    $arguments = @($testDirectory)
    if (-not [string]::IsNullOrWhiteSpace($Executable)) { $arguments += $Executable }
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $testExecutable
    $startInfo.Arguments = ($arguments | ForEach-Object { '"' + $_ + '"' }) -join ' '
    $startInfo.UseShellExecute = $false
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        [void]$process.Start()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "Profile tests failed ($($process.ExitCode))." }
    } finally { $process.Dispose() }
} finally {
    if (Test-Path -LiteralPath $testDirectory) { Remove-Item -LiteralPath $testDirectory -Recurse -Force }
}
