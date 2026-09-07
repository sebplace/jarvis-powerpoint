[CmdletBinding()]
param([string]$Executable, [string]$ReadOnlyPath)

$ErrorActionPreference = "Stop"
$updaterAssertionCount = 0
$projectDirectory = Split-Path -Parent $PSScriptRoot
$fixtureRoot = Join-Path $PSScriptRoot (".updater-fixtures-" + [Guid]::NewGuid().ToString("N"))
$compiler = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path -LiteralPath $compiler)) {
    $compiler = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path -LiteralPath $compiler)) { throw "Windows .NET Framework C# compiler not found." }

function Invoke-Child([string]$path, [string]$arguments, [int]$expected, [int]$timeout = 30000,
    [int]$outputTimeout = $timeout) {
    # Diagnostics.Process avoids the null ExitCode seen with Start-Process in Windows PowerShell 5.
    $process = New-Object Diagnostics.Process
    $process.StartInfo = New-Object Diagnostics.ProcessStartInfo
    $process.StartInfo.FileName = $path
    $process.StartInfo.Arguments = $arguments
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.RedirectStandardError = $true
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try {
        Write-Host "Running updater child: $arguments"
        if (-not $process.Start()) { throw "Could not start isolated test process." }
        $outTask = $process.StandardOutput.ReadToEndAsync()
        $errTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($timeout)) {
            # Only the exact test process is terminated; the installed app is never targeted.
            $process.Kill()
            $null = $process.WaitForExit(5000)
            $outText = if ($outTask.Wait(5000)) { $outTask.GetAwaiter().GetResult() } else { "[stdout still open]" }
            $errText = if ($errTask.Wait(5000)) { $errTask.GetAwaiter().GetResult() } else { "[stderr still open]" }
            throw "Updater child timed out after $timeout ms: $arguments`n$outText`n$errText"
        }
        # Descendants can inherit the pipes even after this process exits; drain within the same deadline.
        $remaining = [Math]::Max(1, [Math]::Min($outputTimeout, $timeout - [int]$watch.ElapsedMilliseconds))
        if (-not [Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($outTask, $errTask), $remaining)) {
            throw "Updater child exited but inherited output pipes did not close within $remaining ms: $arguments"
        }
        $outText = $outTask.GetAwaiter().GetResult()
        $errText = $errTask.GetAwaiter().GetResult()
        if ($process.ExitCode -ne $expected) {
            throw "Expected exit $expected, got $($process.ExitCode).`n$outText`n$errText"
        }
        if ($outText) { Write-Host $outText.TrimEnd() }
        foreach ($match in [regex]::Matches($outText, 'Passed ([0-9]+) offline updater assertions')) {
            $script:updaterAssertionCount += [int]$match.Groups[1].Value
        }
        Write-Host ("Updater child finished in {0:N1}s." -f $watch.Elapsed.TotalSeconds)
    } finally { $process.Dispose() }
}

$candidateSource = @'
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
[assembly: AssemblyVersion("2.0.0.0")]
[assembly: AssemblyFileVersion("2.0.0.0")]
internal static class Candidate
{
    private static void Main(string[] args)
    {
        string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        if (!File.Exists(Path.Combine(folder, "fixture.marker"))) { Environment.ExitCode = 87; return; }
        using (Process self = Process.GetCurrentProcess())
            File.WriteAllText(Path.Combine(folder, "new-app-started.txt"),
                self.Id + "|" + self.StartTime.ToUniversalTime().Ticks + "|" + args.Length);
        if (File.Exists(Path.Combine(folder, "fail-new-startup"))) { Environment.ExitCode = 23; return; }
        Thread.Sleep(3500);
    }
}
'@

$harnessSource = @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using JarvisPowerPoint;
[assembly: AssemblyVersion("1.4.0.0")]
[assembly: AssemblyFileVersion("1.4.0.0")]

internal sealed class FakeTransport : IUpdateTransport
{
    internal readonly Queue<UpdateResponse> Responses = new Queue<UpdateResponse>();
    internal readonly List<Uri> Requests = new List<Uri>();
    public UpdateResponse Get(Uri uri, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        Requests.Add(uri);
        if (Responses.Count == 0) throw new Exception("Unexpected network request.");
        return Responses.Dequeue();
    }
    internal void Add(byte[] bytes) { Responses.Enqueue(new UpdateResponse {
        StatusCode = 200, Body = new MemoryStream(bytes), ContentLength = bytes.Length }); }
    internal void Redirect(string location) { Responses.Enqueue(new UpdateResponse {
        StatusCode = 302, Location = location, Body = new MemoryStream(), ContentLength = 0 }); }
}

internal sealed class CancellingStream : MemoryStream
{
    private readonly CancellationTokenSource source;
    internal CancellingStream(byte[] bytes, CancellationTokenSource source) : base(bytes) { this.source = source; }
    public override int Read(byte[] bytes, int offset, int count)
    {
        int result = base.Read(bytes, offset, count);
        source.Cancel();
        return result;
    }
}

internal sealed class AbortedStream : MemoryStream
{
    private readonly CancellationTokenSource source;
    internal AbortedStream(CancellationTokenSource source) { this.source = source; }
    public override int Read(byte[] bytes, int offset, int count)
    {
        if (source != null) source.Cancel();
        throw new IOException("The response stream was aborted.");
    }
}

internal sealed class FaultFiles : IUpdateFiles
{
    private readonly WindowsUpdateFiles real = new WindowsUpdateFiles();
    internal string FailMode;
    internal int Replacements;
    public bool Exists(string path) { return real.Exists(path); }
    public string Hash(string path) { return real.Hash(path); }
    public Version FileVersion(string path) { return real.FileVersion(path); }
    public long Length(string path) { return real.Length(path); }
    public void Copy(string from, string to)
    {
        if (FailMode == "copy" && to.Contains(".update-")) {
            File.WriteAllText(to, "partial");
            throw new IOException("Injected partial copy failure.");
        }
        real.Copy(from, to);
        if (FailMode == "restore-corrupt" && to.Contains(".restore-"))
            File.AppendAllText(to, "corrupt rollback copy");
        if (FailMode == "changed-before-rollback" && to.Contains(".restore-"))
            File.AppendAllText(Path.Combine(Path.GetDirectoryName(to), "JarvisPowerPoint.exe"), "external change");
        if (FailMode == "managed-before-replace" && to.Contains(".update-"))
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(to), ManagedDeployment.MarkerFileName), "managed");
    }
    public void Replace(string from, string target, string backup)
    {
        Replacements++;
        if (FailMode == "replace-before" && Replacements == 1)
            throw new IOException("Injected pre-replacement failure.");
        if (FailMode == "rollback" && Replacements == 2)
            throw new IOException("Injected rollback failure.");
        real.Replace(from, target, backup);
        if (FailMode == "changed-after-replace" && Replacements == 1) {
            File.AppendAllText(target, "external change");
            throw new IOException("Target changed externally after replacement.");
        }
        if (FailMode == "managed-after-replace")
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(target), ManagedDeployment.MarkerFileName), "managed");
        if (FailMode == "replace-after" && Replacements == 1)
            throw new IOException("Injected post-replacement failure.");
    }
    public void Move(string from, string to) { real.Move(from, to); }
    public void Delete(string path) { real.Delete(path); }
}

internal static class Harness
{
    private static int assertions;
    private static readonly WindowsUpdateFiles Files = new WindowsUpdateFiles();
    private static readonly Version Current = new Version(1, 4, 0, 0);
    private const string HelperRoot = "__TEST_STAGE_ROOT__";
    private static string root, candidate;
    private static UpdateRelease offered;

    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length == 3 && args[0] == "--release-managed-helper") {
            try { ReleaseManagedHelper(args[1], args[2]); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.ExitCode = 1; }
            return;
        }
        if (args.Length == 2 && args[0] == "--hold-output") {
            using (Process self = Process.GetCurrentProcess())
                File.WriteAllText(args[1], self.Id + "|" + self.StartTime.ToUniversalTime().Ticks);
            Thread.Sleep(8000);
            return;
        }
        if (args.Length == 2 && args[0] == "--pipe-owner") {
            using (Process child = StartOwned(Assembly.GetExecutingAssembly().Location,
                "--hold-output " + UpdateManager.QuoteArgument(args[1])))
                WaitForFile(args[1], child);
            return;
        }
        if (UpdateManager.TryHandleCommandLine(args, HelperRoot, Console.Error.WriteLine)) return;
        if (args.Length == 2 && args[0] == "--fixture-parent") {
            File.WriteAllText(args[1] + ".ready", "ready");
            var watch = Stopwatch.StartNew();
            while (!File.Exists(args[1]) && watch.Elapsed < TimeSpan.FromSeconds(40)) Thread.Sleep(50);
            return;
        }
        string ownFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        if (args.Length == 0 && File.Exists(Path.Combine(ownFolder, "fixture.marker"))) {
            using (Process self = Process.GetCurrentProcess())
                File.WriteAllText(Path.Combine(ownFolder, "old-app-restarted.txt"),
                    self.Id + "|" + self.StartTime.ToUniversalTime().Ticks + "|0");
            Thread.Sleep(3500);
            return;
        }
        if ((args.Length != 3 && args.Length != 4) || args[0] != "--run-tests") {
            Console.WriteLine("NORMAL-STARTUP"); Environment.ExitCode = 77; return;
        }
        root = args[1];
        candidate = Path.Combine(root, "candidate-fixture.exe");
        try {
            offered = UpdateManager.ValidateMetadata(Json(MetadataModel()), Current);
            switch (args[2]) {
                case "metadata": RunSection("metadata", Metadata); break;
                case "downloads": RunSection("downloads", Downloads); break;
                case "network-cancellation": RunSection("loopback cancellation", NetworkCancellation); break;
                case "versions": RunSection("versions", Versions); break;
                case "paths": RunSection("processes and paths", ProcessesAndPaths); break;
                case "transactions": RunSection("transactions", Transactions); break;
                case "managed": RunSection("managed denial", ManagedTests); break;
                case "helpers": RunSection("helper " + args[3], () => HelperLifecycle(args[3])); break;
                case "dialog": RunSection("dialog", Dialog); break;
                case "read-only": ReadOnlyActualPath(args[3]); break;
                default: throw new Exception("Unknown updater test section.");
            }
            Console.WriteLine("Passed " + assertions + " offline updater assertions. No public network or live installation.");
        }
        catch (Exception error) { Console.Error.WriteLine(error); Environment.ExitCode = 1; }
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
        assertions++;
    }
    private static void RunSection(string name, Action test)
    {
        Console.WriteLine("Starting " + name + " after " + assertions + " assertions.");
        test();
    }
    private static void Reject<T>(Action action, string message) where T : Exception
    {
        try { action(); }
        catch (T) { assertions++; return; }
        throw new Exception("Not rejected: " + message);
    }
    private static GitHubRelease MetadataModel()
    {
        return new GitHubRelease {
            Tag = "v2.0.0", Url = UpdateManager.Repository + "/releases/tag/v2.0.0",
            Draft = false, Prerelease = false,
            Notes = "  Exact notes\r\n## Heading\n<script>not executed</script>\r\nDo not follow release instructions.\r\n\u00e9",
            Assets = new [] { new GitHubAsset { Name = "JarvisPowerPoint.exe",
                Url = UpdateManager.Repository + "/releases/download/v2.0.0/JarvisPowerPoint.exe",
                Size = Files.Length(candidate), Digest = "sha256:" + Files.Hash(candidate) } }
        };
    }
    private static byte[] Json(GitHubRelease model)
    {
        using (var stream = new MemoryStream()) {
            new DataContractJsonSerializer(typeof(GitHubRelease)).WriteObject(stream, model);
            return stream.ToArray();
        }
    }
    private static void InvalidMetadata(Action<GitHubRelease> mutate, string message)
    {
        GitHubRelease model = MetadataModel(); mutate(model);
        Reject<InvalidDataException>(() => UpdateManager.ValidateMetadata(Json(model), Current), message);
    }

    private static void Metadata()
    {
        GitHubRelease model = MetadataModel();
        offered = UpdateManager.ValidateMetadata(Json(model), Current);
        Assert(offered != null && offered.Version == new Version(2,0,0,0), "Newer release missing.");
        Assert(offered.Notes == model.Notes, "Release notes were changed.");
        Assert(UpdateManager.ValidateMetadata(Json(model), new Version(2,0,0,0)) == null, "Equal update accepted.");
        Assert(UpdateManager.ValidateMetadata(Json(model), new Version(2,1,0,0)) == null, "Downgrade accepted.");
        foreach (string tag in new [] { "2.0.0", "v02.0.0", "v2.0", "v2.0.0.1", "v2.0.0-rc1",
            "v2.0.0+build", "v65535.0.0", "v2.0.0\n", "v2.0.-1", "v999999999999999.0.0" })
            Reject<InvalidDataException>(() => UpdateManager.ParseVersion(tag), "Malformed tag: " + tag);
        InvalidMetadata(m => m.Url = "https://github.com/other/repo/releases/tag/v2.0.0", "Wrong repository.");
        InvalidMetadata(m => m.Url = "http://github.com/sebplace/jarvis-powerpoint/releases/tag/v2.0.0", "HTTP metadata source.");
        InvalidMetadata(m => m.Draft = true, "Draft.");
        InvalidMetadata(m => m.Prerelease = true, "Prerelease.");
        InvalidMetadata(m => m.Assets[0].Url = "https://github.com/other/repo/releases/download/v2.0.0/JarvisPowerPoint.exe", "Wrong asset repository.");
        InvalidMetadata(m => m.Assets[0].Url += "?extra=1", "Noncanonical asset query.");
        InvalidMetadata(m => m.Assets[0].Name = "JarvisPowerPoint.exe.zip", "Wrong asset name.");
        InvalidMetadata(m => m.Assets = new [] {m.Assets[0], m.Assets[0]}, "Duplicate executable.");
        InvalidMetadata(m => m.Assets[0].Digest = null, "Missing digest.");
        InvalidMetadata(m => m.Assets[0].Digest = "sha1:" + new string('a', 64), "Wrong digest algorithm.");
        InvalidMetadata(m => m.Assets[0].Digest = "sha256:" + new string('a', 63), "Short digest.");
        InvalidMetadata(m => m.Assets[0].Digest = "sha256:" + new string('a', 64) + "\n", "Digest with trailing data.");
        InvalidMetadata(m => m.Assets[0].Size = 0, "Zero size.");
        InvalidMetadata(m => m.Assets[0].Size = -1, "Negative size.");
        InvalidMetadata(m => m.Assets[0].Size = UpdateManager.MaximumAssetBytes + 1, "Oversized asset.");
        Reject<InvalidDataException>(() => UpdateManager.ValidateMetadata(new byte[UpdateManager.MaximumMetadataBytes + 1], Current), "Oversized metadata.");
        Reject<SerializationException>(() => UpdateManager.ValidateMetadata(Encoding.UTF8.GetBytes("{}"), Current), "Missing required fields.");
        var transport = new FakeTransport(); transport.Add(Json(model));
        Assert(UpdateManager.Check(transport, Current, CancellationToken.None) != null, "Check failed.");
        Assert(transport.Requests.Count == 1 && transport.Requests[0].AbsoluteUri == UpdateManager.LatestApi, "Wrong public API request.");
        transport = new FakeTransport(); transport.Redirect("https://example.com/release");
        Reject<InvalidDataException>(() => UpdateManager.Check(transport, Current, CancellationToken.None), "Metadata redirect.");
        Assert(transport.Requests.Count == 1, "Metadata followed redirect.");
        transport = new FakeTransport();
        transport.Responses.Enqueue(new UpdateResponse { StatusCode=200, ContentLength=-1,
            Body=new MemoryStream(new byte[UpdateManager.MaximumMetadataBytes + 1]) });
        Reject<InvalidDataException>(() => UpdateManager.Check(transport, Current, CancellationToken.None), "Metadata streaming limit.");
    }

    private static void Downloads()
    {
        byte[] bytes = File.ReadAllBytes(candidate);
        string output = Path.Combine(root, "download.exe");
        var good = new FakeTransport();
        good.Redirect("https://release-assets.githubusercontent.com/github-production-release-asset/test?token=public");
        good.Add(bytes);
        UpdateManager.Download(offered, output, good, Files, CancellationToken.None);
        Assert(Files.Hash(output) == offered.Sha256 && good.Requests.Count == 2, "Valid asset download failed.");
        File.Delete(output);
        foreach (string uri in new [] { "https://evil.example/file", "http://objects.githubusercontent.com/file",
            "https://objects.githubusercontent.com.evil.example/file", "https://user@objects.githubusercontent.com/file",
            "https://objects.githubusercontent.com:444/file", "https://objects.githubusercontent.com/file#fragment",
            "https://github.com/evil/repo/releases/download/v2.0.0/JarvisPowerPoint.exe" }) {
            var bad = new FakeTransport(); bad.Redirect(uri);
            Reject<InvalidDataException>(() => UpdateManager.Download(offered, output, bad, Files, CancellationToken.None), uri);
            Assert(bad.Requests.Count == 1 && !File.Exists(output), "Untrusted redirect caused a request or file.");
        }
        UpdateManager.ValidateAssetUri(new Uri("https://objects.githubusercontent.com/asset?token=public"), offered.Tag, true);
        assertions++;
        var loop = new FakeTransport();
        for (int i=0; i<5; i++) loop.Redirect("https://objects.githubusercontent.com/loop");
        Reject<InvalidDataException>(() => UpdateManager.Download(offered, output, loop, Files, CancellationToken.None), "Redirect loop.");
        Assert(loop.Requests.Count == 5, "Redirect bound is incorrect.");
        var partial = new FakeTransport();
        partial.Responses.Enqueue(new UpdateResponse { StatusCode=200, ContentLength=-1,
            Body=new MemoryStream(new byte[bytes.Length - 1]) });
        Reject<InvalidDataException>(() => UpdateManager.Download(offered, output, partial, Files, CancellationToken.None), "Partial download.");
        Assert(!File.Exists(output), "Partial download retained.");
        var overflow = new FakeTransport();
        overflow.Responses.Enqueue(new UpdateResponse { StatusCode=200, ContentLength=-1,
            Body=new MemoryStream(new byte[bytes.Length + 1]) });
        Reject<InvalidDataException>(() => UpdateManager.Download(offered, output, overflow, Files, CancellationToken.None), "Oversized response.");
        Assert(!File.Exists(output), "Oversized partial retained.");
        var mismatch = new FakeTransport(); mismatch.Add(new byte[1]);
        Reject<InvalidDataException>(() => UpdateManager.Download(offered, output, mismatch, Files, CancellationToken.None), "Content-Length mismatch.");
        using (var cancellation = new CancellationTokenSource()) {
            var none = new FakeTransport(); cancellation.Cancel();
            Reject<OperationCanceledException>(() => UpdateManager.Check(none, Current, cancellation.Token), "Cancelled metadata request.");
            Reject<OperationCanceledException>(() => UpdateManager.Download(offered, output, none, Files, cancellation.Token), "Cancelled asset request.");
            Assert(none.Requests.Count == 0, "Cancellation still requested network.");
        }
        using (var cancellation = new CancellationTokenSource()) {
            var during = new FakeTransport();
            during.Responses.Enqueue(new UpdateResponse { StatusCode=200, ContentLength=bytes.Length,
                Body=new CancellingStream(bytes, cancellation) });
            Reject<OperationCanceledException>(() => UpdateManager.Download(offered, output, during, Files, cancellation.Token), "Cancellation during body read.");
            Assert(!File.Exists(output), "Cancelled partial retained.");
        }
        using (var cancellation = new CancellationTokenSource()) {
            var during = new FakeTransport();
            during.Responses.Enqueue(new UpdateResponse { StatusCode=200, ContentLength=-1,
                Body=new AbortedStream(cancellation) });
            Reject<OperationCanceledException>(() => UpdateManager.Download(offered, output, during, Files, cancellation.Token),
                "An aborted response body must preserve cancellation semantics.");
            Assert(!File.Exists(output), "Aborted partial retained.");
        }
        using (var input = new AbortedStream(null))
        using (var outputStream = new MemoryStream())
            Reject<IOException>(() => UpdateManager.CopyBounded(input, outputStream, 1, null, CancellationToken.None),
                "Non-cancelled stream failures must remain IO errors.");
        var wrongHash = new FakeTransport(); bytes[bytes.Length - 1] ^= 1; wrongHash.Add(bytes);
        Reject<InvalidDataException>(() => UpdateManager.Download(offered, output, wrongHash, Files, CancellationToken.None), "Downloaded digest mismatch.");
        Assert(!File.Exists(output), "Hash-rejected candidate retained.");
    }

    private static void Versions()
    {
        UpdateManager.VerifyCandidate(candidate, offered, Files); assertions++;
        var badVersion = new UpdateRelease("v2.0.1", new Version(2,0,1,0), "", UpdateManager.CanonicalAsset("v2.0.1"),
            offered.Size, offered.Sha256);
        Reject<InvalidDataException>(() => UpdateManager.VerifyCandidate(candidate, badVersion, Files), "Embedded PE version mismatch.");
        var badHash = new UpdateRelease(offered.Tag, offered.Version, "", offered.DownloadUri, offered.Size, new string('0',64));
        Reject<InvalidDataException>(() => UpdateManager.VerifyCandidate(candidate, badHash, Files), "Candidate hash mismatch.");
        var badSize = new UpdateRelease(offered.Tag, offered.Version, "", offered.DownloadUri, offered.Size+1, offered.Sha256);
        Reject<InvalidDataException>(() => UpdateManager.VerifyCandidate(candidate, badSize, Files), "Candidate size mismatch.");
        string notPe = Path.Combine(root, "not-pe.exe"); File.WriteAllText(notPe, "not a PE executable");
        var invalidPe = new UpdateRelease(offered.Tag, offered.Version, "", offered.DownloadUri, Files.Length(notPe), Files.Hash(notPe));
        Reject<BadImageFormatException>(() => UpdateManager.VerifyCandidate(notPe, invalidPe, Files), "Non-PE candidate.");
    }

    private static void NetworkCancellation()
    {
        foreach (bool body in new[] { false, true }) {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            IWebProxy previousProxy = WebRequest.DefaultWebProxy;
            WebRequest.DefaultWebProxy = null;
            using (var cancellation = new CancellationTokenSource())
            using (var requestReceived = new ManualResetEventSlim())
            using (var bodyReading = new ManualResetEventSlim())
            using (var releaseServer = new ManualResetEventSlim()) {
                Exception serverFailure = null, clientFailure = null;
                var server = new Thread(() => {
                    try {
                        using (TcpClient client = listener.AcceptTcpClient())
                        using (NetworkStream stream = client.GetStream()) {
                            stream.ReadTimeout = 5000;
                            string headers = "";
                            while (!headers.EndsWith("\r\n\r\n", StringComparison.Ordinal)) {
                                int next = stream.ReadByte();
                                if (next < 0 || headers.Length > 16384) throw new IOException("Invalid loopback request.");
                                headers += (char)next;
                            }
                            if (body) {
                                byte[] response = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Length: 100\r\nConnection: close\r\n\r\nx");
                                stream.Write(response, 0, response.Length);
                                stream.Flush();
                            }
                            requestReceived.Set();
                            if (!releaseServer.Wait(10000)) throw new Exception("Loopback server was not released.");
                        }
                    }
                    catch (Exception error) { serverFailure = error; requestReceived.Set(); }
                }) { IsBackground = true };
                server.Start();
                Task clientTask = Task.Run(() => {
                    try {
                        var uri = new Uri("http://127.0.0.1:" + ((IPEndPoint)listener.LocalEndpoint).Port + "/");
                        using (UpdateResponse response = new GitHubUpdateTransport().Get(uri, cancellation.Token))
                        using (var output = new MemoryStream()) {
                            if (response.Body.ReadByte() != 'x') throw new IOException("Loopback response prefix missing.");
                            bodyReading.Set();
                            UpdateManager.CopyBounded(response.Body, output, 100, null, cancellation.Token);
                        }
                    }
                    catch (Exception error) { clientFailure = error; }
                });
                try {
                    Assert(requestReceived.Wait(5000) && serverFailure==null, "Loopback request did not reach fixture server.");
                    if (body) Assert(bodyReading.Wait(5000), "Response-body read was not reached.");
                    Thread.Sleep(100);
                    cancellation.Cancel();
                    Assert(clientTask.Wait(5000), "Cancellation did not interrupt a blocked HTTP " + (body ? "body" : "header") + " read.");
                    Assert(clientFailure is OperationCanceledException, "Aborted HTTP request lost cancellation semantics.");
                }
                finally {
                    releaseServer.Set();
                    listener.Stop();
                    bool serverStopped = server.Join(5000);
                    bool clientStopped = clientTask.Wait(5000);
                    WebRequest.DefaultWebProxy = previousProxy;
                    if (!serverStopped || !clientStopped) throw new Exception("Loopback cancellation fixture did not stop.");
                }
                Assert(serverFailure==null, "Loopback server failed: " + serverFailure);
            }
        }
    }

    private static void ProcessesAndPaths()
    {
        for (uint variant = 0; variant < 16; variant++)
            Assert(UpdateManager.IsCloudReparseTag(0x9000001AU | (variant << 12)), "Known OneDrive cloud tag rejected.");
        foreach (uint tag in new uint[] {0, 0xA0000003U, 0xA000000CU, 0x80000022U, 0xB000001AU, 0x9001001AU, 0x9000001BU})
            Assert(!UpdateManager.IsCloudReparseTag(tag), "Unknown/name-surrogate reparse tag accepted.");
        Reject<IOException>(() => UpdateManager.EnsureSafeLocalPath(@"\\server\share\JarvisPowerPoint.exe"), "UNC update path.");
        var manifest = new UpdateManifest { ParentId=42, ParentStartUtcTicks=123456, Target=@"C:\Jarvis App\JarvisPowerPoint.exe" };
        Assert(UpdateManager.MatchesProcess(new UpdateProcessIdentity {Id=42, StartUtcTicks=123456, Executable=manifest.Target}, manifest), "Original identity rejected.");
        Assert(!UpdateManager.MatchesProcess(null, manifest), "Missing process accepted.");
        Assert(!UpdateManager.MatchesProcess(new UpdateProcessIdentity {Id=42, StartUtcTicks=789, Executable=manifest.Target}, manifest), "Reused PID accepted.");
        Assert(!UpdateManager.MatchesProcess(new UpdateProcessIdentity {Id=43, StartUtcTicks=123456, Executable=manifest.Target}, manifest), "Wrong PID accepted.");
        Assert(!UpdateManager.MatchesProcess(new UpdateProcessIdentity {Id=42, StartUtcTicks=123456, Executable=@"C:\elsewhere.exe"}, manifest), "Wrong process path accepted.");
        Assert(UpdateManager.QuoteArgument(@"C:\Folder with spaces\file.xml") == "\"C:\\Folder with spaces\\file.xml\"", "Spaces not quoted.");
        Assert(UpdateManager.QuoteArgument("C:\\ends\\") == "\"C:\\ends\\\\\"", "Trailing backslash not quoted.");
        Assert(UpdateManager.QuoteArgument("a\"b") == "\"a\\\"b\"", "Quote not escaped.");
        Reject<ArgumentException>(() => UpdateManager.QuoteArgument("bad\0argument"), "NUL argument.");
        string stageRoot = Path.Combine(root, "staging");
        string good = Path.Combine(stageRoot, Guid.NewGuid().ToString("N"));
        Assert(UpdateManager.IsStagePath(good, stageRoot), "Generated staging path rejected.");
        Assert(!UpdateManager.IsStagePath(Path.Combine(good, "nested"), stageRoot), "Nested stage accepted.");
        Assert(!UpdateManager.IsStagePath(Path.Combine(stageRoot, "not-guid"), stageRoot), "Arbitrary stage accepted.");
        Assert(!UpdateManager.IsStagePath(Path.Combine(root, Guid.NewGuid().ToString("N")), stageRoot), "External staging accepted.");
        Assert(!UpdateManager.TryHandleCommandLine(new string[0]), "Normal startup is handled as update.");
        Assert(!UpdateManager.TryHandleCommandLine(new [] {"--next"}), "Speech/navigation CLI intercepted.");
        string manifestPath = Path.Combine(root, "manifest.xml");
        manifest.Tag = "v2.0.0"; manifest.OldHash = offered.Sha256; manifest.Sha256 = offered.Sha256; manifest.Size = offered.Size;
        UpdateManager.WriteManifest(manifestPath, manifest);
        Assert(UpdateManager.ReadManifest(manifestPath).Target == manifest.Target, "Manifest round trip changed target.");
        File.WriteAllText(manifestPath, "<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///C:/Windows/win.ini'>]><x>&e;</x>");
        Reject<SerializationException>(() => UpdateManager.ReadManifest(manifestPath), "XML external entity.");
        var prepared = new PreparedUpdate(good, manifest);
        bool closed = false;
        Reject<InvalidOperationException>(() => UpdateManager.AuthorizeAndClose(prepared, () => false, () => closed=true), "Final presentation/rehearsal guard.");
        Assert(!closed && !File.Exists(Path.Combine(good, "proceed")), "Blocked install still authorized or closed.");
        Directory.CreateDirectory(good);
        using (var cancellation = new CancellationTokenSource()) {
            Reject<OperationCanceledException>(() => UpdateManager.AuthorizeAndClose(prepared,
                () => { cancellation.Cancel(); return true; }, () => closed=true, cancellation.Token),
                "Cancellation during final guard must refuse authorization.");
            Assert(!closed && !File.Exists(Path.Combine(good, "proceed")) && !prepared.InstallAuthorized,
                "Cancelled authorization closed the app or left permission to install.");
        }
    }

    private static void Transactions()
    {
        foreach (string mode in new [] {"success", "copy", "replace-before", "replace-after", "launch", "both-launches", "rollback",
            "changed-target", "changed-candidate", "managed-before-replace", "managed-after-replace",
            "changed-after-replace", "changed-before-rollback", "restore-corrupt"})
        {
            string folder = Path.Combine(root, "Transaction " + mode + " \u00e9");
            Directory.CreateDirectory(folder);
            string stage = Path.Combine(folder, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(stage);
            string target = Path.Combine(folder, "JarvisPowerPoint.exe");
            File.Copy(Assembly.GetExecutingAssembly().Location, target);
            File.Copy(candidate, Path.Combine(stage, "candidate.exe"));
            string oldHash = Files.Hash(target);
            var manifest = new UpdateManifest { Target=target, OldHash=oldHash, Tag=offered.Tag, Size=offered.Size, Sha256=offered.Sha256 };
            var files = new FaultFiles { FailMode=mode };
            int launches = 0;
            Action<string> launch = path => {
                Assert(path == target, "Launch path changed.");
                launches++;
                if ((mode=="launch" || mode=="rollback" || mode=="changed-before-rollback" || mode=="restore-corrupt") &&
                    launches==1 || mode=="both-launches")
                    throw new Win32Exception("Injected launch failure.");
            };
            if (mode=="changed-target") File.AppendAllText(target, "changed");
            if (mode=="changed-candidate") File.AppendAllText(Path.Combine(stage,"candidate.exe"), "changed");
            if (mode=="success") {
                UpdateManager.ApplyTransaction(manifest, stage, files, launch);
                Assert(Files.Hash(target) == offered.Sha256 && launches==1, "Successful replacement incorrect.");
            } else if (mode.StartsWith("managed-")) {
                Reject<InvalidOperationException>(() => UpdateManager.ApplyTransaction(manifest, stage, files, launch), mode);
                Assert(launches==0, "Newly managed target was launched.");
                Assert(files.Replacements==(mode=="managed-before-replace" ? 0 : 1), "Newly managed target was replaced or rolled back.");
                Assert(Files.Hash(target)==(mode=="managed-before-replace" ? oldHash : offered.Sha256),
                    "Newly managed target was mutated after managed detection.");
            } else {
                Reject<IOException>(() => UpdateManager.ApplyTransaction(manifest, stage, files, launch), mode);
                if (mode=="changed-target") Assert(Files.Hash(target) != oldHash && launches==0 && files.Replacements==0, "Changed target overwritten.");
                else if (mode=="changed-after-replace" || mode=="changed-before-rollback")
                    Assert(Files.Hash(target)!=oldHash && Files.Hash(target)!=offered.Sha256 && files.Replacements==1 &&
                        launches==(mode=="changed-after-replace" ? 0 : 1), "Rollback overwrote or launched an externally changed target.");
                else if (mode=="restore-corrupt")
                    Assert(Files.Hash(target)==offered.Sha256 && files.Replacements==1 && launches==1,
                        "Unverified rollback copy replaced the candidate.");
                else if (mode=="rollback") Assert(Files.Hash(target)==offered.Sha256, "Injected rollback failure not exercised.");
                else Assert(Files.Hash(target)==oldHash, "Original was not preserved: " + mode);
                if (mode=="launch" || mode=="both-launches") Assert(launches==2, "Original restart was not attempted.");
                if (mode=="copy" || mode=="changed-candidate") Assert(files.Replacements==0 && launches==1, "Preflight failure did not restart the verified original.");
            }
            string backup = Path.Combine(folder, ".JarvisPowerPoint.backup-" + Path.GetFileName(stage) + ".exe");
            if (mode=="success" || mode=="replace-after" || mode=="launch" || mode=="both-launches" || mode=="rollback" ||
                mode=="managed-after-replace" || mode=="changed-after-replace" || mode=="changed-before-rollback" || mode=="restore-corrupt")
                Assert(File.Exists(backup) && Files.Hash(backup)==oldHash, "Runnable old backup missing: " + mode);
            Assert(Directory.GetFiles(folder, ".JarvisPowerPoint.update-*.exe").Length==0, "Disposable incoming executable retained: " + mode);
        }
    }

    private static void ReadOnlyActualPath(string path)
    {
        DateTime modified = File.GetLastWriteTimeUtc(path);
        UpdateManager.EnsureSafeLocalPath(path);
        Assert(Files.Hash(path).Length == 64, "Installed executable could not be read/hashed.");
        Assert(Files.FileVersion(path) != null, "Installed executable PE version unreadable.");
        Assert(File.GetLastWriteTimeUtc(path) == modified, "Read-only actual-path check changed file data.");
        int cloudPaths = 0;
        string cursor = Path.GetFullPath(path);
        while (!string.IsNullOrEmpty(cursor)) {
            uint tag = UpdateManager.ReadReparseTag(cursor);
            if (tag != 0) {
                Assert(UpdateManager.IsCloudReparseTag(tag), "Actual-path ancestor has an unsafe tag.");
                cloudPaths++;
            }
            cursor = Path.GetDirectoryName(cursor);
        }
        Console.WriteLine("Read-only installed-path validation passed; " + cloudPaths +
            " cloud-tagged path components. Executable was only read, never launched or replaced.");
    }

    private static void WaitForFile(string path, Process process)
    {
        var watch = Stopwatch.StartNew();
        while (!File.Exists(path)) {
            if (process.HasExited) throw new Exception("Fixture process exited before writing " + Path.GetFileName(path));
            if (watch.Elapsed > TimeSpan.FromSeconds(15)) throw new Exception("Fixture file wait timed out: " + path);
            Thread.Sleep(50);
        }
    }

    private static Process StartOwned(string path, string arguments)
    {
        Process process = Process.Start(new ProcessStartInfo(path, arguments) {
            UseShellExecute=false, WorkingDirectory=Path.GetDirectoryName(path) });
        if (process == null) throw new Exception("Fixture process failed to start.");
        return process;
    }

    private static void WaitForOwnedRestart(string marker, string expectedTarget)
    {
        string[] identity = File.ReadAllText(marker).Split('|');
        Assert(identity.Length==3 && identity[2]=="0", "Restart unexpectedly passed CLI arguments.");
        Process process;
        try { process = Process.GetProcessById(int.Parse(identity[0])); }
        catch (ArgumentException) { return; }
        using (process) {
            IntPtr handle = process.Handle;
            if (process.StartTime.ToUniversalTime().Ticks != long.Parse(identity[1]) ||
                !string.Equals(process.MainModule.FileName, expectedTarget, StringComparison.OrdinalIgnoreCase))
                return;
            if (!process.WaitForExit(10000)) {
                process.Kill();
                process.WaitForExit();
                throw new Exception("Owned fixture restart did not exit.");
            }
        }
    }

    private static void HelperLifecycle(string selectedMode)
    {
        Directory.CreateDirectory(HelperRoot);
        foreach (string mode in new[] {"success", "launch-failure", "candidate-changed", "identity-mismatch", "managed", "became-managed",
            "cancel-before-proceed", "cancel-after-proceed", "authorization-failure"}) {
            if (mode != selectedMode) continue;
            string folder = Path.Combine(root, "Live helper " + mode + " \u00e9");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "fixture.marker"), "synthetic offline fixture");
            string stage = Path.Combine(HelperRoot, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stage);
            string target = Path.Combine(folder, "JarvisPowerPoint.exe");
            string stop = Path.Combine(folder, "parent.exit");
            string helperPath = Path.Combine(stage, "updater.exe");
            File.Copy(Assembly.GetExecutingAssembly().Location, target);
            File.Copy(target, helperPath);
            File.Copy(candidate, Path.Combine(stage, "candidate.exe"));
            string oldHash = Files.Hash(target);
            Process parent = null, helper = null;
            try {
                parent = StartOwned(target, "--fixture-parent " + UpdateManager.QuoteArgument(stop));
                WaitForFile(stop + ".ready", parent);
                var manifest = new UpdateManifest { Target=target, OldHash=oldHash, ParentId=parent.Id,
                    ParentStartUtcTicks=parent.StartTime.ToUniversalTime().Ticks,
                    Tag=offered.Tag, Size=offered.Size, Sha256=offered.Sha256 };
                if (mode=="identity-mismatch") manifest.ParentStartUtcTicks++;
                UpdateManager.WriteManifest(Path.Combine(stage, "manifest.xml"), manifest);
                if (mode=="launch-failure") File.WriteAllText(Path.Combine(folder, "fail-new-startup"), "fail");
                if (mode=="managed") File.WriteAllText(Path.Combine(folder, ManagedDeployment.MarkerFileName), "managed");
                helper = StartOwned(helperPath, "--apply-update " + UpdateManager.QuoteArgument(Path.Combine(stage, "manifest.xml")));
                if (mode=="identity-mismatch" || mode=="managed") {
                    Assert(helper.WaitForExit(15000) && helper.ExitCode==2, "Invalid/managed target helper did not fail.");
                    Assert(!File.Exists(Path.Combine(stage, "ready")) && !parent.HasExited &&
                        Files.Hash(target)==oldHash, "Identity failure disturbed the original process.");
                    File.WriteAllText(stop, "exit");
                    Assert(parent.WaitForExit(10000), "Fixture original did not exit.");
                } else {
                    WaitForFile(Path.Combine(stage, "ready"), helper);
                    Assert(!parent.HasExited && Files.Hash(target)==oldHash, "Helper replaced an active original.");
                    if (mode.StartsWith("cancel-") || mode=="authorization-failure") {
                        var prepared = new PreparedUpdate(stage, manifest) { HelperStarted=true };
                        if (mode=="authorization-failure")
                            Reject<IOException>(() => UpdateManager.AuthorizeAndClose(prepared, () => true,
                                () => { throw new IOException("Injected application close failure."); }),
                                "Failed application shutdown must revoke installation.");
                        if (mode=="cancel-after-proceed") {
                            File.WriteAllText(Path.Combine(stage, "proceed"), "authorized fixture update");
                            Thread.Sleep(250);
                        }
                        UpdateManager.Discard(prepared);
                        Assert(helper.WaitForExit(5000) && helper.ExitCode==2, "Cancelled helper did not exit promptly.");
                        Assert(!parent.HasExited && Files.Hash(target)==oldHash && !prepared.InstallAuthorized,
                            "Cancellation changed or stopped the original.");
                        Assert(File.ReadAllText(Path.Combine(stage, "result.txt")).Contains("cancelled"),
                            "Helper failed for a different reason instead of observing cancellation.");
                        Assert(!File.Exists(Path.Combine(folder, "new-app-started.txt")) &&
                            !File.Exists(Path.Combine(folder, "old-app-restarted.txt")), "Cancelled helper launched an application.");
                        Console.WriteLine("Synthetic helper lifecycle passed: " + mode + ".");
                        continue;
                    }
                    if (mode=="candidate-changed") File.AppendAllText(Path.Combine(stage, "candidate.exe"), "changed");
                    if (mode=="became-managed") File.WriteAllText(Path.Combine(folder, ManagedDeployment.MarkerFileName), "managed");
                    if (mode=="success") {
                        var prepared = new PreparedUpdate(stage, manifest) { HelperStarted=true };
                        bool closed = false;
                        UpdateManager.AuthorizeAndClose(prepared, () => true, () => closed=true);
                        UpdateManager.Discard(prepared);
                        Assert(closed && prepared.InstallAuthorized && !File.Exists(Path.Combine(stage, "cancel")),
                            "Successful authorization was cancelled by cleanup.");
                    } else File.WriteAllText(Path.Combine(stage, "proceed"), "authorized fixture update");
                    Thread.Sleep(200);
                    Assert(!helper.HasExited && Files.Hash(target)==oldHash, "Helper did not wait for the original process.");
                    File.WriteAllText(stop, "exit");
                    Assert(parent.WaitForExit(10000), "Fixture original did not exit after authorization.");
                    bool helperExited = helper.WaitForExit(20000);
                    if (!helperExited) {
                        string resultPath = Path.Combine(stage, "result.txt");
                        Console.Error.WriteLine("Timed-out helper stage: " + string.Join(", ", Directory.GetFiles(stage)));
                        Console.Error.WriteLine("Timed-out helper target folder: " + string.Join(", ", Directory.GetFiles(folder)));
                        if (File.Exists(resultPath)) Console.Error.WriteLine(File.ReadAllText(resultPath));
                    }
                    Assert(helperExited, "Fixture helper did not finish: " + mode);
                    Assert(helper.ExitCode == (mode=="success" ? 0 : 2), "Incorrect helper result exit code.");
                    string result = File.ReadAllText(Path.Combine(stage, "result.txt"));
                    Assert(result.StartsWith(mode=="success" ? "SUCCESS:" : "Jarvis update failed:"), "Durable helper result incorrect.");
                    if (mode=="success") {
                        Assert(Files.Hash(target)==offered.Sha256, "Live helper did not install candidate.");
                        Assert(!File.Exists(Path.Combine(stage, "candidate.exe")), "Successful helper retained disposable candidate.");
                        WaitForOwnedRestart(Path.Combine(folder, "new-app-started.txt"), target);
                    } else if (mode=="became-managed") {
                        Assert(Files.Hash(target)==oldHash, "Target that became managed was changed.");
                        Assert(!File.Exists(Path.Combine(folder, "old-app-restarted.txt")), "Managed application must not be restarted by helper.");
                    } else {
                        Assert(Files.Hash(target)==oldHash, "Live helper did not preserve/restore the original.");
                        Assert(File.Exists(Path.Combine(folder, "old-app-restarted.txt")), "Live helper did not restart original after failure.");
                        WaitForOwnedRestart(Path.Combine(folder, "old-app-restarted.txt"), target);
                    }
                    if (mode!="candidate-changed" && mode!="became-managed") {
                        string backup = Path.Combine(folder, ".JarvisPowerPoint.backup-" + Path.GetFileName(stage) + ".exe");
                        Assert(Files.Hash(backup)==oldHash, "Live helper original backup missing.");
                    }
                }
                Console.WriteLine("Synthetic helper lifecycle passed: " + mode + ".");
            }
            finally {
                if (parent != null) {
                    File.WriteAllText(stop, "exit");
                    if (!parent.HasExited && !parent.WaitForExit(5000)) { parent.Kill(); parent.WaitForExit(); }
                    parent.Dispose();
                }
                if (helper != null) {
                    if (!helper.HasExited && !helper.WaitForExit(5000)) { helper.Kill(); helper.WaitForExit(); }
                    helper.Dispose();
                }
            }
        }
    }

    private static void ManagedTests()
    {
        string folder = Path.Combine(root, "Managed target");
        Directory.CreateDirectory(folder);
        string target = Path.Combine(folder, "JarvisPowerPoint.exe");
        File.Copy(Assembly.GetExecutingAssembly().Location, target);
        string before = Files.Hash(target);
        Assert(!ManagedDeployment.IsManagedExecutable(target), "Portable fixture detected as managed.");
        Assert(ManagedDeployment.MatchesInstallPath(folder, folder.ToUpperInvariant() + "\\"), "Install path normalization.");
        Assert(!ManagedDeployment.MatchesInstallPath(folder + " portable", folder), "Install path prefix must not match.");
        Assert(!ManagedDeployment.MatchesInstallPath(folder, "relative"), "Relative registry paths must not match.");
        Assert(ManagedDeployment.IsEnabled(1) && !ManagedDeployment.IsEnabled("1") && !ManagedDeployment.IsEnabled(0),
            "Policy must be an enabled DWORD.");
        File.WriteAllText(Path.Combine(folder, ManagedDeployment.MarkerFileName), "managed");
        Assert(ManagedDeployment.IsManagedExecutable(target), "Managed marker not recognized.");
        Assert(!ManagedDeployment.IsManagedExecutable(candidate), "Marker leaked into unrelated portable directory.");
        var manifest = new UpdateManifest { Target=target };
        int launches = 0;
        Reject<InvalidOperationException>(() => UpdateManager.ApplyTransaction(manifest, folder, Files, p => launches++),
            "Managed transaction must refuse before reading candidate.");
        Assert(Files.Hash(target)==before && launches==0, "Managed transaction changed/launched target.");
        var prepared = new PreparedUpdate(folder, manifest);
        Reject<InvalidOperationException>(() => UpdateManager.StartHelperAndWait(prepared, CancellationToken.None),
            "Managed helper launch refused.");
        Reject<InvalidOperationException>(() => UpdateManager.AuthorizeAndClose(prepared, () => true, () => launches++),
            "Managed authorization refused.");
        Assert(!prepared.HelperStarted && launches==0 && !File.Exists(Path.Combine(folder, "proceed")),
            "Managed authorization had side effects.");
        File.Delete(Path.Combine(folder, ManagedDeployment.MarkerFileName));
        Reject<InvalidOperationException>(() => UpdateManager.AuthorizeAndClose(prepared, () => {
            File.WriteAllText(Path.Combine(folder, ManagedDeployment.MarkerFileName), "new managed deployment");
            return true;
        }, () => launches++), "Managed state must be rechecked after the final application guard.");
        Assert(!prepared.InstallAuthorized && launches==0 && !File.Exists(Path.Combine(folder, "proceed")),
            "A newly managed target was authorized.");
        string ownMarker = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ManagedDeployment.MarkerFileName);
        File.WriteAllText(ownMarker, "managed test");
        try {
            Reject<InvalidOperationException>(() => UpdateManager.Check(CancellationToken.None), "Managed check must not contact GitHub.");
            Reject<InvalidOperationException>(() => UpdateManager.Prepare(offered, CancellationToken.None), "Managed prepare refused.");
            using (var dialog = new UpdateDialog(true, () => true, () => launches++)) {
                var check = (System.Windows.Forms.Button)typeof(UpdateDialog).GetField("check", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(dialog);
                Assert(!check.Enabled, "Managed dialog exposes public update check.");
            }
        } finally { File.Delete(ownMarker); }
    }

    private static void ReleaseManagedHelper(string executable, string fixtureRoot)
    {
        string folder = Path.Combine(fixtureRoot, "Managed release target");
        string stage = Path.Combine(HelperRoot, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        Directory.CreateDirectory(stage);
        string target = Path.Combine(folder, "JarvisPowerPoint.exe");
        string helper = Path.Combine(stage, "updater.exe");
        string manifestPath = Path.Combine(stage, "manifest.xml");
        File.Copy(executable, target);
        File.Copy(executable, helper);
        string stagedCandidate = Path.Combine(stage, "candidate.exe");
        File.Copy(Path.Combine(fixtureRoot, "candidate-fixture.exe"), stagedCandidate);
        string sourceHash = Files.Hash(executable);
        File.WriteAllText(Path.Combine(folder, ManagedDeployment.MarkerFileName), "managed release fixture");
        var manifest = new UpdateManifest { Target=target, OldHash=sourceHash, ParentId=int.MaxValue,
            ParentStartUtcTicks=DateTime.UtcNow.Ticks, Tag="v2.0.0",
            Size=Files.Length(stagedCandidate), Sha256=Files.Hash(stagedCandidate) };
        UpdateManager.WriteManifest(manifestPath, manifest);

        // Exercise the actual release assembly's helper dispatcher; only staging/reporting use test seams.
        Assembly release = Assembly.LoadFrom(helper);
        Type managed = release.GetType("JarvisPowerPoint.ManagedDeployment", true);
        Assert(!(bool)managed.GetMethod("IsManagedExecutable").Invoke(null, new object[] { helper }),
            "The helper must be an unmarked portable copy, not the managed target.");
        MethodInfo dispatch = release.GetType("JarvisPowerPoint.UpdateManager", true).GetMethod(
            "TryHandleCommandLine", BindingFlags.Static | BindingFlags.NonPublic);
        int reports = 0;
        bool handled = (bool)dispatch.Invoke(null, new object[] {
            new[] { "--apply-update", manifestPath }, HelperRoot, new Action<string>(message => reports++) });
        Assert(handled && Environment.ExitCode==2, "Release helper did not reject managed target.");
        Assert(File.ReadAllText(Path.Combine(stage, "result.txt")).Contains("managed by your organization"),
            "Release helper failed for a different reason instead of its managed-target guard.");
        Assert(!File.Exists(Path.Combine(stage, "ready")) && reports==0,
            "Release helper authorized an update or reported original application shutdown.");
        Assert(Files.Hash(target)==sourceHash && Files.Hash(executable)==sourceHash,
            "Release helper changed the target or source executable.");
        Assert(Directory.GetFiles(folder).Length==2, "Release helper created replacement/backup/launch files.");
        Environment.ExitCode = 0;
        Console.WriteLine("Passed " + assertions + " actual-release portable-helper managed-denial assertions.");
    }

    private static void Dialog()
    {
        using (var form = new UpdateDialog(true, () => true, () => { throw new Exception("Unexpected app close."); })) {
            Assert(form.Text == "Manual GitHub update", "English dialog missing.");
            // Constructing and disposing the dialog must not perform IO/network or invoke app shutdown.
            Assert(form.Controls.Count == 1, "Dialog not constructed.");
        }
        using (var form = new UpdateDialog(false, () => false, () => {}))
            Assert(form.Text.Contains("GitHub"), "French dialog missing.");
        using (var cancellation = new CancellationTokenSource()) {
            var form = new UpdateDialog(true, () => true, () => {});
            typeof(UpdateDialog).GetField("operation", BindingFlags.NonPublic | BindingFlags.Instance)
                .SetValue(form, cancellation);
            form.Dispose();
            Assert(cancellation.IsCancellationRequested, "Disposing the dialog did not cancel pending work.");
            typeof(UpdateDialog).GetField("operation", BindingFlags.NonPublic | BindingFlags.Instance)
                .SetValue(form, null);
        }
    }
}
'@

try {
    New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
    $security = New-Object Security.AccessControl.DirectorySecurity
    $security.SetAccessRuleProtection($true, $false)
    $owner = [Security.Principal.WindowsIdentity]::GetCurrent().User
    $security.SetOwner($owner)
    $rule = New-Object Security.AccessControl.FileSystemAccessRule($owner, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
    $security.AddAccessRule($rule)
    Set-Acl -LiteralPath $fixtureRoot -AclObject $security
    $candidateCs = Join-Path $fixtureRoot "Candidate.cs"
    $harnessCs = Join-Path $fixtureRoot "Harness.cs"
    $candidateExe = Join-Path $fixtureRoot "candidate-fixture.exe"
    $harnessExe = Join-Path $fixtureRoot "UpdaterTests.exe"
    Set-Content -LiteralPath $candidateCs -Value $candidateSource -Encoding ASCII
    $testStageRoot = (Join-Path $fixtureRoot "private-helper-staging").Replace('\', '\\').Replace('"', '\"')
    Set-Content -LiteralPath $harnessCs -Value $harnessSource.Replace("__TEST_STAGE_ROOT__", $testStageRoot) -Encoding ASCII
    & $compiler /nologo /langversion:5 /target:exe "/out:$candidateExe" $candidateCs
    if ($LASTEXITCODE -ne 0) { throw "Candidate fixture compilation failed." }
    & $compiler /nologo /langversion:5 /target:exe /main:Harness "/out:$harnessExe" `
        /reference:System.dll /reference:System.Core.dll /reference:System.Runtime.Serialization.dll `
        /reference:System.Xml.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll `
        (Join-Path $projectDirectory "UpdateModels.cs") (Join-Path $projectDirectory "UpdateManager.cs") `
        (Join-Path $projectDirectory "ManagedDeployment.cs") `
        (Join-Path $projectDirectory "UpdateDialog.cs") $harnessCs
    if ($LASTEXITCODE -ne 0) { throw "Updater test compilation failed." }
    # Bound each section independently so a timeout identifies the failing section, not just the suite.
    foreach ($section in @("metadata", "downloads", "network-cancellation", "versions", "paths", "transactions", "managed", "dialog")) {
        Invoke-Child $harnessExe "--run-tests `"$fixtureRoot`" $section" 0 90000
    }
    foreach ($mode in @("success", "launch-failure", "candidate-changed", "identity-mismatch", "managed", "became-managed",
        "cancel-before-proceed", "cancel-after-proceed", "authorization-failure")) {
        Invoke-Child $harnessExe "--run-tests `"$fixtureRoot`" helpers $mode" 0 60000
    }
    if ($ReadOnlyPath) {
        Invoke-Child $harnessExe "--run-tests `"$fixtureRoot`" read-only `"$([IO.Path]::GetFullPath($ReadOnlyPath))`"" 0
    }
    Write-Host "Passed $updaterAssertionCount total offline updater assertions."
    $pipeMarker = Join-Path $fixtureRoot "pipe-owner.txt"
    try {
        try {
            Invoke-Child $harnessExe "--pipe-owner `"$pipeMarker`"" 0 30000 1000
            throw "Inherited output pipe timeout was not enforced."
        } catch {
            if ($_.Exception.Message -notlike "Updater child exited but inherited output pipes did not close*") { throw }
            Write-Host "Passed inherited-output-pipe timeout regression."
        }
    } finally {
        if (Test-Path -LiteralPath $pipeMarker) {
            $identity = (Get-Content -LiteralPath $pipeMarker -Raw).Split('|')
            $owned = Get-Process -Id ([int]$identity[0]) -ErrorAction SilentlyContinue
            if ($owned) {
                try {
                    $null = $owned.Handle
                    if ($owned.StartTime.ToUniversalTime().Ticks -eq [long]$identity[1] -and
                        $owned.MainModule.FileName -eq $harnessExe -and -not $owned.WaitForExit(10000)) {
                        $owned.Kill()
                        throw "Owned output-pipe fixture did not exit."
                    }
                } finally { $owned.Dispose() }
            }
        }
    }
    foreach ($arguments in @("--apply-update", "--apply-update no-such-manifest.xml",
        "--apply-update=bad", "--APPLY-UPDATE bad", "--next --apply-update bad")) {
        Invoke-Child $harnessExe $arguments 2
    }
    if ($Executable) {
        # Optional integration check: these malformed flags must exit before speech/UI initialization.
        $fullExecutable = [IO.Path]::GetFullPath($Executable)
        Invoke-Child $fullExecutable "--apply-update" 2
        Invoke-Child $fullExecutable "--apply-update no-such-manifest.xml" 2
        Invoke-Child $harnessExe "--release-managed-helper `"$fullExecutable`" `"$fixtureRoot`"" 0
    }
    Write-Host "Passed updater CLI failure-mode checks (no normal GUI startup)."
} finally {
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
