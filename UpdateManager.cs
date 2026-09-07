using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Win32.SafeHandles;

namespace JarvisPowerPoint
{
    internal static class UpdateManager
    {
        internal const long MaximumAssetBytes = 64L * 1024 * 1024;
        internal const int MaximumMetadataBytes = 512 * 1024;
        internal const string Repository = "https://github.com/sebplace/jarvis-powerpoint";
        internal const string LatestApi = "https://api.github.com/repos/sebplace/jarvis-powerpoint/releases/latest";
        private static readonly Regex HashPattern = new Regex(@"\A[a-fA-F0-9]{64}\z", RegexOptions.CultureInvariant);
        private static readonly string[] StageFiles =
            { "candidate.exe", "updater.exe", "manifest.xml", "ready", "proceed", "cancel", "result.txt" };

        public static Version CurrentVersion { get { return Assembly.GetExecutingAssembly().GetName().Version; } }
        public static string StagingRoot
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "JarvisPowerPoint", "Updates"); }
        }

        public static Version ParseVersion(string tag)
        {
            if (tag == null || !Regex.IsMatch(tag, @"\Av(0|[1-9][0-9]{0,4})\.(0|[1-9][0-9]{0,4})\.(0|[1-9][0-9]{0,4})\z",
                RegexOptions.CultureInvariant))
                throw new InvalidDataException("Only stable tags in vMAJOR.MINOR.PATCH format are accepted.");
            string[] parts = tag.Substring(1).Split('.');
            int major = int.Parse(parts[0], CultureInfo.InvariantCulture);
            int minor = int.Parse(parts[1], CultureInfo.InvariantCulture);
            int patch = int.Parse(parts[2], CultureInfo.InvariantCulture);
            if (major > 65534 || minor > 65534 || patch > 65534)
                throw new InvalidDataException("Release version exceeds the Windows file-version range.");
            return new Version(major, minor, patch, 0);
        }

        internal static Uri CanonicalAsset(string tag)
        {
            ParseVersion(tag);
            return new Uri(Repository + "/releases/download/" + tag + "/JarvisPowerPoint.exe");
        }

        internal static void ValidateAssetUri(Uri uri, string tag, bool redirect)
        {
            if (uri == null || !uri.IsAbsoluteUri || uri.Scheme != Uri.UriSchemeHttps ||
                !uri.IsDefaultPort || uri.UserInfo.Length != 0 || uri.Fragment.Length != 0)
                throw new InvalidDataException("Update URLs must be HTTPS without credentials or custom ports.");
            if (!redirect || uri.Host.Equals("github.com", StringComparison.OrdinalIgnoreCase))
            {
                if (!uri.AbsoluteUri.Equals(CanonicalAsset(tag).AbsoluteUri, StringComparison.Ordinal))
                    throw new InvalidDataException("The asset URL is not the exact sebplace/jarvis-powerpoint release.");
            }
            else if (!uri.Host.Equals("release-assets.githubusercontent.com", StringComparison.OrdinalIgnoreCase) &&
                !uri.Host.Equals("objects.githubusercontent.com", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("GitHub redirected the asset to an untrusted host.");
        }

        internal static UpdateRelease ValidateMetadata(byte[] bytes, Version current)
        {
            if (bytes == null || bytes.Length == 0 || bytes.Length > MaximumMetadataBytes)
                throw new InvalidDataException("Release metadata is empty or too large.");
            GitHubRelease release;
            using (var stream = new MemoryStream(bytes, false))
            {
                var serializer = new DataContractJsonSerializer(typeof(GitHubRelease));
                release = (GitHubRelease)serializer.ReadObject(stream);
            }
            if (release == null || release.Draft || release.Prerelease)
                throw new InvalidDataException("Draft and prerelease updates are not accepted.");
            Version version = ParseVersion(release.Tag);
            if (!string.Equals(release.Url, Repository + "/releases/tag/" + release.Tag, StringComparison.Ordinal))
                throw new InvalidDataException("Release metadata is for a different repository.");
            if (release.Assets == null || release.Assets.Length > 100)
                throw new InvalidDataException("Release assets are missing or excessive.");
            GitHubAsset chosen = null;
            foreach (GitHubAsset asset in release.Assets)
            {
                if (asset != null && string.Equals(asset.Name, "JarvisPowerPoint.exe", StringComparison.Ordinal))
                {
                    if (chosen != null) throw new InvalidDataException("The release has duplicate executable assets.");
                    chosen = asset;
                }
            }
            if (chosen == null) throw new InvalidDataException("The release has no JarvisPowerPoint.exe asset.");
            if (chosen.Size <= 0 || chosen.Size > MaximumAssetBytes)
                throw new InvalidDataException("The advertised executable size is invalid (maximum 64 MiB).");
            if (chosen.Digest == null || !chosen.Digest.StartsWith("sha256:", StringComparison.Ordinal) ||
                !HashPattern.IsMatch(chosen.Digest.Substring(7)))
                throw new InvalidDataException("GitHub did not provide a valid SHA-256 digest. Update refused.");
            Uri uri;
            if (!Uri.TryCreate(chosen.Url, UriKind.Absolute, out uri))
                throw new InvalidDataException("The executable URL is invalid.");
            ValidateAssetUri(uri, release.Tag, false);
            if (version <= current) return null;
            return new UpdateRelease(release.Tag, version, release.Notes ?? "", uri, chosen.Size,
                chosen.Digest.Substring(7).ToLowerInvariant());
        }

        public static UpdateRelease Check(CancellationToken token)
        {
            ManagedDeployment.EnsurePortableUpdateAllowed(Assembly.GetExecutingAssembly().Location);
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(45));
                return Check(new GitHubUpdateTransport(), CurrentVersion, timeout.Token);
            }
        }

        internal static UpdateRelease Check(IUpdateTransport transport, Version current, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            using (UpdateResponse response = transport.Get(new Uri(LatestApi), token))
            {
                if (response.StatusCode != 200)
                    throw new InvalidDataException("GitHub release check returned HTTP " + response.StatusCode +
                        ". No update was downloaded.");
                if (response.ContentLength > MaximumMetadataBytes)
                    throw new InvalidDataException("Release metadata is too large.");
                using (var bytes = new MemoryStream())
                {
                    CopyBounded(response.Body, bytes, MaximumMetadataBytes,
                        response.ContentLength >= 0 ? (long?)response.ContentLength : null, token);
                    return ValidateMetadata(bytes.ToArray(), current);
                }
            }
        }

        internal static long CopyBounded(Stream input, Stream output, long limit, long? expected,
            CancellationToken token)
        {
            var buffer = new byte[32768];
            long total = 0;
            while (true)
            {
                token.ThrowIfCancellationRequested();
                int read;
                try { read = input.Read(buffer, 0, buffer.Length); }
                catch (Exception error)
                {
                    if (!IsExpected(error)) throw;
                    token.ThrowIfCancellationRequested();
                    throw;
                }
                token.ThrowIfCancellationRequested();
                if (read == 0) break;
                total += read;
                if (total > limit) throw new InvalidDataException("The download exceeded its allowed size.");
                output.Write(buffer, 0, read);
            }
            if (expected.HasValue && total != expected.Value)
                throw new InvalidDataException("The download is incomplete or its size differs from GitHub.");
            return total;
        }

        internal static void Download(UpdateRelease release, string destination, IUpdateTransport transport,
            IUpdateFiles files, CancellationToken token)
        {
            ValidateAssetUri(release.DownloadUri, release.Tag, false);
            Uri uri = release.DownloadUri;
            bool created = false;
            try
            {
                for (int redirects = 0; redirects <= 4; redirects++)
                {
                    token.ThrowIfCancellationRequested();
                    using (UpdateResponse response = transport.Get(uri, token))
                    {
                        if (response.StatusCode == 301 || response.StatusCode == 302 || response.StatusCode == 303 ||
                            response.StatusCode == 307 || response.StatusCode == 308)
                        {
                            Uri next;
                            if (redirects == 4 || string.IsNullOrEmpty(response.Location) ||
                                !Uri.TryCreate(uri, response.Location, out next))
                                throw new InvalidDataException("Too many or invalid download redirects.");
                            ValidateAssetUri(next, release.Tag, true);
                            uri = next;
                            continue;
                        }
                        if (response.StatusCode != 200)
                            throw new InvalidDataException("GitHub download returned HTTP " + response.StatusCode + ".");
                        if (response.ContentLength >= 0 && response.ContentLength != release.Size)
                            throw new InvalidDataException("The download Content-Length differs from GitHub metadata.");
                        using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                        {
                            created = true;
                            CopyBounded(response.Body, output, release.Size, release.Size, token);
                            output.Flush(true);
                        }
                        token.ThrowIfCancellationRequested();
                        VerifyCandidate(destination, release, files);
                        token.ThrowIfCancellationRequested();
                        return;
                    }
                }
            }
            catch
            {
                if (created) DeleteDisposable(destination);
                throw;
            }
        }

        internal static void VerifyCandidate(string path, UpdateRelease release, IUpdateFiles files)
        {
            if (release.Size <= 0 || release.Size > MaximumAssetBytes || release.Sha256 == null ||
                !HashPattern.IsMatch(release.Sha256) || ParseVersion(release.Tag) != release.Version)
                throw new InvalidDataException("Invalid candidate verification parameters.");
            if (files.Length(path) != release.Size)
                throw new InvalidDataException("The executable size does not match the release.");
            if (!string.Equals(files.Hash(path), release.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("SHA-256 verification failed. Update refused.");
            if (files.FileVersion(path) != release.Version)
                throw new InvalidDataException("The embedded executable version does not match the release tag.");
        }

        public static PreparedUpdate Prepare(UpdateRelease release, CancellationToken token)
        {
            ManagedDeployment.EnsurePortableUpdateAllowed(Assembly.GetExecutingAssembly().Location);
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                timeout.CancelAfter(TimeSpan.FromMinutes(3));
                return PrepareCore(release, timeout.Token);
            }
        }

        private static PreparedUpdate PrepareCore(UpdateRelease release, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (release.Version <= CurrentVersion) throw new InvalidDataException("Downgrades are not allowed.");
            string target = Path.GetFullPath(Assembly.GetExecutingAssembly().Location);
            ManagedDeployment.EnsurePortableUpdateAllowed(target);
            if (!string.Equals(Path.GetFileName(target), "JarvisPowerPoint.exe", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Run the installed JarvisPowerPoint.exe before requesting an update.");
            EnsureSafeLocalPath(target);
            if ((File.GetAttributes(target) & FileAttributes.ReadOnly) != 0)
                throw new IOException("The application is read-only. Move it to a writable local folder and try again.");
            string probe = Path.Combine(Path.GetDirectoryName(target), ".JarvisPowerPoint.write-" +
                Guid.NewGuid().ToString("N") + ".tmp");
            using (var stream = new FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { }
            File.Delete(probe);
            CreatePrivateRoot();
            CleanupExpiredStages();
            if (Directory.GetDirectories(StagingRoot).Length >= 32)
                throw new IOException("Update staging is full. Remove old update folders from " + StagingRoot + ".");
            string directory = Path.Combine(StagingRoot, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var files = new WindowsUpdateFiles();
                string oldHash = files.Hash(target);
                Download(release, Path.Combine(directory, "candidate.exe"), new GitHubUpdateTransport(), files, token);
                File.Copy(target, Path.Combine(directory, "updater.exe"), false);
                if (!string.Equals(oldHash, files.Hash(Path.Combine(directory, "updater.exe")), StringComparison.Ordinal) ||
                    !string.Equals(oldHash, files.Hash(target), StringComparison.Ordinal))
                    throw new IOException("The running application changed during preparation. Retry after restarting.");
                UpdateManifest manifest;
                using (Process parent = Process.GetCurrentProcess())
                {
                    manifest = new UpdateManifest
                    {
                        Target = target, OldHash = oldHash, ParentId = parent.Id,
                        ParentStartUtcTicks = parent.StartTime.ToUniversalTime().Ticks,
                        Tag = release.Tag, Size = release.Size, Sha256 = release.Sha256
                    };
                }
                WriteManifest(Path.Combine(directory, "manifest.xml"), manifest);
                token.ThrowIfCancellationRequested();
                return new PreparedUpdate(directory, manifest);
            }
            catch
            {
                CleanupStage(directory);
                throw;
            }
        }

        public static void StartHelperAndWait(PreparedUpdate prepared, CancellationToken token)
        {
            ManagedDeployment.EnsurePortableUpdateAllowed(prepared.Manifest.Target);
            token.ThrowIfCancellationRequested();
            ValidateStage(prepared.DirectoryPath, StagingRoot);
            using (var helper = new Process())
            {
                helper.StartInfo = new ProcessStartInfo(Path.Combine(prepared.DirectoryPath, "updater.exe"),
                    "--apply-update " + QuoteArgument(Path.Combine(prepared.DirectoryPath, "manifest.xml")))
                {
                    UseShellExecute = false, WorkingDirectory = prepared.DirectoryPath
                };
                if (!helper.Start()) throw new IOException("Windows could not start the update helper.");
                prepared.HelperStarted = true;
                var watch = Stopwatch.StartNew();
                while (!File.Exists(Path.Combine(prepared.DirectoryPath, "ready")))
                {
                    token.ThrowIfCancellationRequested();
                    if (helper.HasExited)
                        throw new IOException("The update helper could not prepare the installation. " +
                            ReadResult(prepared.DirectoryPath));
                    if (watch.Elapsed > TimeSpan.FromSeconds(20))
                        throw new IOException("The update helper did not become ready. Jarvis has not been closed.");
                    token.WaitHandle.WaitOne(100);
                }
                token.ThrowIfCancellationRequested();
            }
        }

        public static void AuthorizeAndClose(PreparedUpdate prepared, Func<bool> canInstall, Action closeApplication,
            CancellationToken token = default(CancellationToken))
        {
            ManagedDeployment.EnsurePortableUpdateAllowed(prepared.Manifest.Target);
            // The guard runs on the UI thread after all asynchronous work, immediately before shutdown.
            if (!canInstall())
                throw new InvalidOperationException("Stop all slide shows and rehearsal, and save any pending reports before installing an update.");
            token.ThrowIfCancellationRequested();
            ManagedDeployment.EnsurePortableUpdateAllowed(prepared.Manifest.Target);
            EnsureNotCancelled(prepared.DirectoryPath);
            using (var stream = new FileStream(Path.Combine(prepared.DirectoryPath, "proceed"),
                FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            {
                stream.WriteByte(1);
                stream.Flush(true);
            }
            try
            {
                token.ThrowIfCancellationRequested();
                closeApplication();
                prepared.InstallAuthorized = true;
            }
            catch
            {
                CancelPrepared(prepared);
                throw;
            }
        }

        public static void Discard(PreparedUpdate prepared)
        {
            if (prepared == null || prepared.InstallAuthorized) return;
            if (prepared.HelperStarted) CancelPrepared(prepared);
            else CleanupStage(prepared.DirectoryPath);
        }

        private static void CancelPrepared(PreparedUpdate prepared)
        {
            // Leave a durable veto: removing proceed alone races a helper that has already read it.
            File.WriteAllText(Path.Combine(prepared.DirectoryPath, "cancel"), "cancel", Encoding.ASCII);
        }

        private static void EnsureNotCancelled(string directory)
        {
            if (File.Exists(Path.Combine(directory, "cancel")))
                throw new OperationCanceledException("Installation was cancelled. Nothing was replaced.");
        }

        internal static string QuoteArgument(string value)
        {
            if (value == null || value.IndexOf('\0') >= 0) throw new ArgumentException("Invalid process argument.");
            var output = new StringBuilder("\"");
            int slashes = 0;
            foreach (char c in value)
            {
                if (c == '\\') { slashes++; continue; }
                if (c == '"') output.Append('\\', slashes * 2 + 1).Append('"');
                else output.Append('\\', slashes).Append(c);
                slashes = 0;
            }
            return output.Append('\\', slashes * 2).Append('"').ToString();
        }

        internal static bool MatchesProcess(UpdateProcessIdentity identity, UpdateManifest manifest)
        {
            return identity != null && identity.Id == manifest.ParentId &&
                identity.StartUtcTicks == manifest.ParentStartUtcTicks &&
                string.Equals(identity.Executable, manifest.Target, StringComparison.OrdinalIgnoreCase);
        }

        internal static UpdateRelease ManifestRelease(UpdateManifest manifest)
        {
            Version version = ParseVersion(manifest.Tag);
            return new UpdateRelease(manifest.Tag, version, "", CanonicalAsset(manifest.Tag),
                manifest.Size, manifest.Sha256);
        }

        internal static void ApplyTransaction(UpdateManifest manifest, string directory, IUpdateFiles files,
            Action<string> launch)
        {
            ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
            string id = Path.GetFileName(directory);
            string folder = Path.GetDirectoryName(manifest.Target);
            string incoming = Path.Combine(folder, ".JarvisPowerPoint.update-" + id + ".exe");
            string backup = Path.Combine(folder, ".JarvisPowerPoint.backup-" + id + ".exe");
            string restore = Path.Combine(folder, ".JarvisPowerPoint.restore-" + id + ".exe");
            UpdateRelease release = ManifestRelease(manifest);
            bool replacementAttempted = false;
            bool incomingCreated = false;
            try
            {
                VerifyCandidate(Path.Combine(directory, "candidate.exe"), release, files);
                if (!string.Equals(files.Hash(manifest.Target), manifest.OldHash, StringComparison.OrdinalIgnoreCase))
                    throw new IOException("The installed executable changed. It was not replaced.");
                if (files.Exists(incoming) || files.Exists(backup) || files.Exists(restore))
                    throw new IOException("An update transaction path already exists. It was not overwritten.");
                ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
                incomingCreated = true;
                files.Copy(Path.Combine(directory, "candidate.exe"), incoming);
                VerifyCandidate(incoming, release, files);
                if (!string.Equals(files.Hash(manifest.Target), manifest.OldHash, StringComparison.OrdinalIgnoreCase))
                    throw new IOException("The installed executable changed just before replacement.");
                ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
                replacementAttempted = true;
                files.Replace(incoming, manifest.Target, backup);
                VerifyCandidate(manifest.Target, release, files);
                ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
                launch(manifest.Target);
            }
            catch (Exception error)
            {
                if (!IsExpected(error)) throw;
                // Do not restart or roll back files that became managed while the helper was waiting.
                ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
                string recovery;
                try
                {
                    bool unchanged = files.Exists(manifest.Target) &&
                        string.Equals(files.Hash(manifest.Target), manifest.OldHash, StringComparison.OrdinalIgnoreCase);
                    if (!unchanged)
                    {
                        if (!replacementAttempted)
                            throw new IOException("The installed executable changed or disappeared; it is not safe to restart.");
                        if (files.Exists(manifest.Target) &&
                            !string.Equals(files.Hash(manifest.Target), release.Sha256, StringComparison.OrdinalIgnoreCase))
                            throw new IOException("The installed executable changed after replacement; automatic rollback was refused.");
                        if (!files.Exists(backup) ||
                            !string.Equals(files.Hash(backup), manifest.OldHash, StringComparison.OrdinalIgnoreCase))
                            throw new IOException("The original backup could not be verified.");
                        ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
                        files.Copy(backup, restore);
                        if (!string.Equals(files.Hash(restore), manifest.OldHash, StringComparison.OrdinalIgnoreCase))
                            throw new IOException("The rollback copy failed verification.");
                        ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
                        if (files.Exists(manifest.Target))
                        {
                            if (!string.Equals(files.Hash(manifest.Target), release.Sha256, StringComparison.OrdinalIgnoreCase))
                                throw new IOException("The installed executable changed just before rollback.");
                            files.Replace(restore, manifest.Target, null);
                        }
                        else files.Move(restore, manifest.Target);
                        if (!string.Equals(files.Hash(manifest.Target), manifest.OldHash, StringComparison.OrdinalIgnoreCase))
                            throw new IOException("Restored executable failed verification.");
                    }
                    recovery = replacementAttempted ? "The original application is restored. Backup: " + backup + "." :
                        "The original application was not replaced.";
                    try
                    {
                        ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
                        launch(manifest.Target);
                        recovery += " The original application was restarted.";
                    }
                    catch (Exception restartError)
                    {
                        if (!IsExpected(restartError)) throw;
                        recovery += " Restart it manually: " + manifest.Target + ". " + restartError.Message;
                    }
                }
                catch (Exception rollbackError)
                {
                    if (!IsExpected(rollbackError)) throw;
                    recovery = replacementAttempted ?
                        "Automatic rollback could not complete: " + rollbackError.Message +
                            " Do not delete the runnable original backup: " + backup + "." :
                        "Nothing was replaced, but automatic restart was refused: " + rollbackError.Message +
                            " Restore a verified original executable and launch it manually: " + manifest.Target + ".";
                }
                throw new IOException("Update failed: " + error.Message + " " + recovery +
                    " Check folder permissions, free disk space, and antivirus; no elevation was attempted.", error);
            }
            finally
            {
                if (incomingCreated)
                {
                    try { if (files.Exists(incoming)) files.Delete(incoming); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
        }

        public static bool TryHandleCommandLine(string[] args)
        {
            return TryHandleCommandLine(args, StagingRoot, delegate(string message)
            {
                MessageBox.Show(message, "Jarvis PowerPoint update", MessageBoxButtons.OK, MessageBoxIcon.Error);
            });
        }

        internal static bool TryHandleCommandLine(string[] args, string stagingRoot, Action<string> reportFailure)
        {
            bool requested = false;
            foreach (string arg in args ?? new string[0])
                if (arg != null && arg.StartsWith("--apply-update", StringComparison.OrdinalIgnoreCase)) requested = true;
            if (!requested) return false;
            Environment.ExitCode = 2;
            string directory = null;
            bool originalExited = false;
            try
            {
                if (args.Length != 2 || args[0] != "--apply-update")
                    throw new InvalidDataException("Usage: --apply-update <private manifest.xml>");
                string path = Path.GetFullPath(args[1]);
                directory = Path.GetDirectoryName(path);
                ValidateStage(directory, stagingRoot);
                if (!string.Equals(path, Path.Combine(directory, "manifest.xml"), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Path.GetFullPath(Assembly.GetExecutingAssembly().Location),
                        Path.Combine(directory, "updater.exe"), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The helper must run from its own private update staging folder.");
                EnsureSafeLocalPath(path);
                UpdateManifest manifest = ReadManifest(path);
                ValidateManifest(manifest, directory, stagingRoot);
                var files = new WindowsUpdateFiles();
                if (!string.Equals(files.Hash(Path.Combine(directory, "updater.exe")), manifest.OldHash,
                    StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The helper does not match the original executable.");
                if (!string.Equals(files.Hash(manifest.Target), manifest.OldHash, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The original executable changed before the helper started.");
                VerifyCandidate(Path.Combine(directory, "candidate.exe"), ManifestRelease(manifest), files);
                EnsureNotCancelled(directory);
                using (Process parent = Process.GetProcessById(manifest.ParentId))
                {
                    // Pin the process before reading identity: later PID reuse cannot change this handle.
                    IntPtr handle = parent.Handle;
                    if (!MatchesProcess(new UpdateProcessIdentity { Id = parent.Id,
                        StartUtcTicks = parent.StartTime.ToUniversalTime().Ticks,
                        Executable = parent.MainModule.FileName }, manifest))
                        throw new InvalidDataException("The original process identity does not match. Nothing was installed.");
                    File.WriteAllText(Path.Combine(directory, "ready"), "ready", Encoding.ASCII);
                    var watch = Stopwatch.StartNew();
                    while (!File.Exists(Path.Combine(directory, "proceed")))
                    {
                        EnsureNotCancelled(directory);
                        if (parent.HasExited || watch.Elapsed > TimeSpan.FromSeconds(30))
                            throw new IOException("Installation was not authorized before the timeout.");
                        Thread.Sleep(100);
                    }
                    watch.Restart();
                    while (!parent.WaitForExit(100))
                    {
                        EnsureNotCancelled(directory);
                        if (watch.Elapsed > TimeSpan.FromSeconds(60))
                            throw new IOException("Jarvis did not close within 60 seconds. No process was killed; nothing was installed.");
                    }
                    EnsureNotCancelled(directory);
                    originalExited = true;
                }
                EnsureSafeLocalPath(directory);
                EnsureSafeLocalPath(manifest.Target);
                EnsureNotCancelled(directory);
                ApplyTransaction(manifest, directory, files, LaunchApplication);
                File.WriteAllText(Path.Combine(directory, "result.txt"),
                    "SUCCESS: Installed " + manifest.Tag + ". The previous executable remains in the same-directory backup.",
                    Encoding.UTF8);
                Environment.ExitCode = 0;
                DeleteDisposable(Path.Combine(directory, "candidate.exe"));
            }
            catch (Exception error)
            {
                if (!IsExpected(error)) throw;
                string message = "Jarvis update failed: " + error.Message;
                Console.Error.WriteLine(message);
                if (directory != null && IsStagePath(directory, stagingRoot))
                {
                    try
                    {
                        ValidateStage(directory, stagingRoot);
                        File.WriteAllText(Path.Combine(directory, "result.txt"), message, Encoding.UTF8);
                    }
                    catch (Exception writeError)
                    {
                        if (!IsExpected(writeError)) throw;
                        message += "\r\nUnable to save the result: " + writeError.Message;
                    }
                }
                if (originalExited) reportFailure(message);
            }
            return true;
        }

        internal static void ValidateManifest(UpdateManifest manifest, string directory, string stagingRoot)
        {
            if (manifest == null || manifest.Target == null || !Path.IsPathRooted(manifest.Target) ||
                !string.Equals(Path.GetFullPath(manifest.Target), manifest.Target, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Path.GetFileName(manifest.Target), "JarvisPowerPoint.exe", StringComparison.OrdinalIgnoreCase) ||
                manifest.Target.StartsWith(stagingRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                manifest.ParentId <= 0 || manifest.ParentId == Process.GetCurrentProcess().Id ||
                manifest.ParentStartUtcTicks <= 0 || manifest.OldHash == null || !HashPattern.IsMatch(manifest.OldHash))
                throw new InvalidDataException("Invalid update manifest or target path.");
            ManagedDeployment.EnsurePortableUpdateAllowed(manifest.Target);
            if (ParseVersion(manifest.Tag) <= CurrentVersion)
                throw new InvalidDataException("The helper refuses equal or older versions.");
            EnsureSafeLocalPath(manifest.Target);
            EnsureSafeLocalPath(Path.Combine(directory, "candidate.exe"));
        }

        private static void LaunchApplication(string target)
        {
            using (Process process = Process.Start(new ProcessStartInfo(target)
                { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(target), Arguments = "" }))
            {
                if (process == null) throw new IOException("Windows did not start Jarvis.");
                if (process.WaitForExit(2000))
                    throw new IOException("Jarvis exited during startup (exit code " + process.ExitCode + ").");
            }
        }

        internal static bool IsExpected(Exception error)
        {
            return error is IOException || error is InvalidDataException || error is UnauthorizedAccessException || error is SecurityException ||
                error is Win32Exception || error is WebException || error is SerializationException ||
                error is XmlException || error is ArgumentException || error is InvalidOperationException ||
                error is OperationCanceledException || error is BadImageFormatException || error is CryptographicException;
        }

        internal static bool IsCloudReparseTag(uint tag)
        {
            // Only CLOUD and CLOUD_1..CLOUD_F; no name-surrogate, symlink, mount-point or unknown tags.
            return (tag & 0x20000000U) == 0 && (tag & 0xFFFF0FFFU) == 0x9000001AU;
        }

        internal static uint ReadReparseTag(string path)
        {
            using (SafeFileHandle handle = CreateFile(path, 0x80, 7, IntPtr.Zero, 3,
                0x02000000 | 0x00200000, IntPtr.Zero))
            {
                if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error(),
                    "Cannot inspect the update path reparse tag: " + path);
                FileAttributeTagInfo information;
                if (!GetFileInformationByHandleEx(handle, 9, out information,
                    (uint)Marshal.SizeOf(typeof(FileAttributeTagInfo))))
                    throw new Win32Exception(Marshal.GetLastWin32Error(),
                        "Cannot verify the update path reparse tag: " + path);
                var buffer = new byte[16384];
                uint returned;
                if (DeviceIoControl(handle, 0x000900A8, IntPtr.Zero, 0, buffer, (uint)buffer.Length,
                    out returned, IntPtr.Zero))
                {
                    if (returned < 8) throw new IOException("Windows returned an incomplete reparse tag: " + path);
                    uint tag = BitConverter.ToUInt32(buffer, 0);
                    if (tag == 0 || (information.ReparseTag != 0 && information.ReparseTag != tag))
                        throw new IOException("Windows returned inconsistent reparse tags: " + path);
                    return tag;
                }
                int error = Marshal.GetLastWin32Error();
                if (error != 4390)
                    throw new Win32Exception(error, "Cannot query the update path reparse data: " + path);
                if (information.ReparseTag != 0 || (information.FileAttributes & (uint)FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Windows reported a reparse point without verifiable data: " + path);
                return 0;
            }
        }

        internal static void EnsureSafeLocalPath(string path)
        {
            string cursor = Path.GetFullPath(path);
            if (cursor.StartsWith(@"\\", StringComparison.Ordinal))
                throw new IOException("Updates require a local drive. Move Jarvis to a writable local folder.");
            while (!string.IsNullOrEmpty(cursor))
            {
                if (File.Exists(cursor) || Directory.Exists(cursor))
                {
                    // .NET Framework can mask cloud reparse attributes; inspect each existing component natively.
                    uint tag = ReadReparseTag(cursor);
                    if (tag != 0 && !IsCloudReparseTag(tag))
                        throw new IOException("Update paths cannot contain symlinks, junctions, or unrecognized reparse points (tag 0x" +
                            tag.ToString("X8", CultureInfo.InvariantCulture) + "). Move Jarvis to a writable local folder.");
                }
                cursor = Path.GetDirectoryName(cursor);
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FileAttributeTagInfo
        {
            public uint FileAttributes;
            public uint ReparseTag;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string fileName, uint access, uint share,
            IntPtr securityAttributes, uint creationDisposition, uint flags, IntPtr template);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandleEx(SafeFileHandle handle, int informationClass,
            out FileAttributeTagInfo information, uint size);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(SafeFileHandle handle, uint controlCode, IntPtr input,
            uint inputSize, [Out] byte[] output, uint outputSize, out uint returned, IntPtr overlapped);

        private static void CreatePrivateRoot()
        {
            EnsureSafeLocalPath(StagingRoot);
            var security = new DirectorySecurity();
            security.SetAccessRuleProtection(true, false);
            SecurityIdentifier owner = WindowsIdentity.GetCurrent().User;
            security.SetOwner(owner);
            security.AddAccessRule(new FileSystemAccessRule(owner, FileSystemRights.FullControl,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
            security.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null),
                FileSystemRights.FullControl, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                PropagationFlags.None, AccessControlType.Allow));
            Directory.CreateDirectory(StagingRoot, security);
            Directory.SetAccessControl(StagingRoot, security);
        }

        internal static bool IsStagePath(string directory, string root)
        {
            if (string.IsNullOrEmpty(directory)) return false;
            string normalized = Path.GetFullPath(directory);
            Guid id;
            return string.Equals(Path.GetDirectoryName(normalized), Path.GetFullPath(root), StringComparison.OrdinalIgnoreCase) &&
                Guid.TryParseExact(Path.GetFileName(normalized), "N", out id);
        }

        internal static void ValidateStage(string directory, string root)
        {
            if (!IsStagePath(directory, root)) throw new InvalidDataException("Invalid private update staging path.");
            EnsureSafeLocalPath(directory);
        }

        internal static void WriteManifest(string path, UpdateManifest manifest)
        {
            using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            {
                new DataContractSerializer(typeof(UpdateManifest)).WriteObject(stream, manifest);
                stream.Flush(true);
            }
        }

        internal static UpdateManifest ReadManifest(string path)
        {
            if (new FileInfo(path).Length > 16384) throw new InvalidDataException("Update manifest is too large.");
            using (XmlReader reader = XmlReader.Create(path, new XmlReaderSettings
                { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16384 }))
                return (UpdateManifest)new DataContractSerializer(typeof(UpdateManifest)).ReadObject(reader);
        }

        private static string ReadResult(string directory)
        {
            string path = Path.Combine(directory, "result.txt");
            if (!File.Exists(path)) return "See " + directory + " for details.";
            if (new FileInfo(path).Length > 16384) return "The update result is too large to display.";
            return File.ReadAllText(path);
        }

        internal static void CleanupExpiredStages()
        {
            foreach (string directory in Directory.GetDirectories(StagingRoot))
                if (IsStagePath(directory, StagingRoot) && Directory.GetLastWriteTimeUtc(directory) < DateTime.UtcNow.AddDays(-7))
                    CleanupStage(directory);
        }

        internal static void CleanupStage(string directory)
        {
            try
            {
                ValidateStage(directory, StagingRoot);
                if (!Directory.Exists(directory)) return;
                foreach (string name in StageFiles)
                {
                    string path = Path.Combine(directory, name);
                    EnsureSafeLocalPath(path);
                    if (File.Exists(path)) File.Delete(path);
                }
                // Never recurse: unknown files, junctions, or a running helper are retained, not followed.
                if (Directory.GetFileSystemEntries(directory).Length == 0) Directory.Delete(directory, false);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            catch (SecurityException) { }
        }

        private static void DeleteDisposable(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    internal sealed class WindowsUpdateFiles : IUpdateFiles
    {
        public bool Exists(string path) { return File.Exists(path); }
        public long Length(string path) { return new FileInfo(path).Length; }
        public void Copy(string from, string to) { File.Copy(from, to, false); }
        public void Replace(string from, string target, string backup) { File.Replace(from, target, backup, true); }
        public void Move(string from, string to) { File.Move(from, to); }
        public void Delete(string path) { File.Delete(path); }
        public string Hash(string path)
        {
            UpdateManager.EnsureSafeLocalPath(path);
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (SHA256 algorithm = SHA256.Create())
                return BitConverter.ToString(algorithm.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
        }
        public Version FileVersion(string path)
        {
            FileVersionInfo info = FileVersionInfo.GetVersionInfo(path);
            var version = new Version(info.FileMajorPart, info.FileMinorPart, info.FileBuildPart, info.FilePrivatePart);
            if (AssemblyName.GetAssemblyName(path).Version != version)
                throw new InvalidDataException("The PE file version and assembly version differ.");
            return version;
        }
    }

    internal sealed class GitHubUpdateTransport : IUpdateTransport
    {
        public UpdateResponse Get(Uri uri, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = "GET";
            request.AllowAutoRedirect = false;
            request.UserAgent = "JarvisPowerPoint-ManualUpdater";
            request.Accept = uri.Host == "api.github.com" ? "application/vnd.github+json" : "application/octet-stream";
            request.Headers["X-GitHub-Api-Version"] = "2022-11-28";
            request.UseDefaultCredentials = false;
            request.Credentials = null;
            request.PreAuthenticate = false;
            if (request.Proxy != null) request.Proxy = new AnonymousUpdateProxy(request.Proxy);
            request.Timeout = 20000;
            request.ReadWriteTimeout = 20000;
            request.AutomaticDecompression = DecompressionMethods.None;
            CancellationTokenRegistration registration = token.Register(request.Abort);
            try
            {
                var response = (HttpWebResponse)request.GetResponse();
                return new UpdateResponse
                {
                    StatusCode = (int)response.StatusCode, Location = response.Headers["Location"],
                    ContentLength = response.ContentLength, Body = response.GetResponseStream(),
                    Close = delegate { try { response.Close(); } finally { registration.Dispose(); } }
                };
            }
            catch (WebException error)
            {
                if (error.Response != null) error.Response.Close();
                registration.Dispose();
                token.ThrowIfCancellationRequested();
                throw;
            }
            catch
            {
                registration.Dispose();
                token.ThrowIfCancellationRequested();
                throw;
            }
        }

        internal sealed class AnonymousUpdateProxy : IWebProxy
        {
            private readonly IWebProxy systemProxy;
            public AnonymousUpdateProxy(IWebProxy systemProxy) { this.systemProxy = systemProxy; }
            public Uri GetProxy(Uri destination) { return systemProxy.GetProxy(destination); }
            public bool IsBypassed(Uri host) { return systemProxy.IsBypassed(host); }
            public ICredentials Credentials
            {
                get { return null; }
                set { if (value != null) throw new InvalidOperationException("The updater does not send proxy credentials."); }
            }
        }
    }
}
