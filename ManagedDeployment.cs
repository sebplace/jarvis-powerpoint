using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;

namespace JarvisPowerPoint
{
    internal static class ManagedDeployment
    {
        internal const string MarkerFileName = "JarvisPowerPoint.managed";
        internal const string DeploymentKey = @"Software\JarvisPowerPoint\Deployment";
        internal const string PolicyKey = @"Software\Policies\JarvisPowerPoint";

        public static bool IsManaged
        {
            get { return IsManagedExecutable(Assembly.GetExecutingAssembly().Location); }
        }

        public static string GetMessage(bool english)
        {
            return english
                ? "This installation is managed by your organization. Ask IT to update or repair Jarvis with its MSI package. The portable updater and per-user shortcut changes are disabled."
                : "Cette installation est g\u00e9r\u00e9e par votre organisation. Contactez le service informatique pour mettre \u00e0 jour ou r\u00e9parer Jarvis avec le package MSI. La mise \u00e0 jour portable et les modifications du raccourci personnel sont d\u00e9sactiv\u00e9es.";
        }

        public static void EnsurePortableUpdateAllowed(string executablePath)
        {
            if (IsManagedExecutable(executablePath))
                throw new InvalidOperationException(GetMessage(true));
        }

        public static bool IsManagedExecutable(string executablePath)
        {
            try
            {
                string directory = Path.GetDirectoryName(Path.GetFullPath(executablePath));
                if (string.IsNullOrEmpty(directory)) return true;
                directory = NormalizeDirectory(directory);
                string marker = Path.Combine(directory, MarkerFileName);
                // An adjacent marker only protects this directory, not unrelated portable copies.
                try { File.GetAttributes(marker); return true; }
                catch (FileNotFoundException) { }
                catch (DirectoryNotFoundException) { }

                if (IsInstallationDirectory(directory, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)) ||
                    IsInstallationDirectory(directory, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)) ||
                    IsInstallationDirectory(directory, Environment.GetEnvironmentVariable("ProgramW6432"))) return true;

                // Always inspect both registry views, including from an emulated process.
                foreach (RegistryView view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
                {
                    if (view == RegistryView.Registry64 && !Environment.Is64BitOperatingSystem) continue;
                    using (RegistryKey machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                    {
                        using (RegistryKey policy = machine.OpenSubKey(PolicyKey))
                            if (policy != null && IsEnabled(policy.GetValue("DisablePortableUpdates"))) return true;
                        using (RegistryKey deployment = machine.OpenSubKey(DeploymentKey))
                            if (deployment != null && MatchesInstallPath(directory, deployment.GetValue("InstallPath") as string))
                                return true;
                    }
                }
                return false;
            }
            catch (IOException) { return true; }
            catch (UnauthorizedAccessException) { return true; }
            catch (SecurityException) { return true; }
            catch (ArgumentException) { return true; }
            catch (NotSupportedException) { return true; }
            catch (Win32Exception) { return true; }
        }

        internal static bool IsEnabled(object value)
        {
            return value is int && (int)value == 1;
        }

        private static bool IsInstallationDirectory(string directory, string programFiles)
        {
            return !string.IsNullOrWhiteSpace(programFiles) &&
                MatchesInstallPath(directory, Path.Combine(programFiles, "Jarvis PowerPoint"));
        }

        internal static bool MatchesInstallPath(string directory, string installPath)
        {
            if (string.IsNullOrWhiteSpace(installPath) || !Path.IsPathRooted(installPath)) return false;
            return string.Equals(NormalizeDirectory(directory), NormalizeDirectory(installPath),
                StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeDirectory(string path)
        {
            string fullPath = TrimDirectory(Path.GetFullPath(path));
            // Resolve existing directories so short names and directory aliases cannot evade registration.
            using (SafeFileHandle handle = CreateFile(fullPath, 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero))
            {
                if (handle.IsInvalid)
                {
                    int error = Marshal.GetLastWin32Error();
                    if (error == 2 || error == 3) return fullPath;
                    throw new Win32Exception(error, "Cannot inspect the installation directory.");
                }
                var buffer = new StringBuilder(512);
                uint length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
                if (length >= buffer.Capacity)
                {
                    buffer.Capacity = checked((int)length + 1);
                    length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
                }
                if (length == 0 || length >= buffer.Capacity)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot resolve the installation directory.");
                string resolved = buffer.ToString();
                if (resolved.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
                    resolved = @"\\" + resolved.Substring(8);
                else if (resolved.StartsWith(@"\\?\", StringComparison.Ordinal))
                    resolved = resolved.Substring(4);
                return TrimDirectory(resolved);
            }
        }

        private static string TrimDirectory(string path)
        {
            return string.Equals(path, Path.GetPathRoot(path), StringComparison.OrdinalIgnoreCase)
                ? path : path.TrimEnd('\\', '/');
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string name, uint access, uint share,
            IntPtr security, uint disposition, uint flags, IntPtr template);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFinalPathNameByHandle(SafeFileHandle handle, StringBuilder path,
            uint length, uint flags);
    }
}
