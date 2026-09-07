using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Microsoft.Win32;

namespace JarvisPowerPoint
{
    internal enum ShortcutResult
    {
        NotNeeded,
        Declined,
        Created
    }

    internal sealed class StartMenuShortcut
    {
        private const string PromptValueName = "StartMenuShortcutPrompted";
        private readonly string executablePath;
        private readonly string settingsKey;

        public StartMenuShortcut(string programsFolder, string executablePath, string settingsKey)
        {
            if (string.IsNullOrWhiteSpace(programsFolder))
            {
                throw new DirectoryNotFoundException("Windows Start menu folder is unavailable.");
            }

            this.executablePath = Path.GetFullPath(executablePath);
            this.settingsKey = settingsKey;
            ShortcutPath = Path.Combine(programsFolder, "Jarvis PowerPoint.lnk");
        }

        public string ShortcutPath { get; private set; }

        public ShortcutResult Configure(Func<bool> confirm, bool fromMenu)
        {
            if (ManagedDeployment.IsManagedExecutable(executablePath)) return ShortcutResult.NotNeeded;
            if (!fromMenu)
            {
                // Leave existing shortcuts alone; the menu can repair one after a move.
                if (File.Exists(ShortcutPath))
                {
                    return ShortcutResult.NotNeeded;
                }

                object prompted = Registry.GetValue(settingsKey, PromptValueName, null);
                if (prompted is int && (int)prompted == 1)
                {
                    return ShortcutResult.NotNeeded;
                }
            }

            bool accepted = confirm();
            if (ManagedDeployment.IsManagedExecutable(executablePath)) return ShortcutResult.NotNeeded;
            if (!accepted)
            {
                RememberChoice();
                return ShortcutResult.Declined;
            }

            CreateShortcut();
            RememberChoice();
            return ShortcutResult.Created;
        }

        private void RememberChoice()
        {
            Registry.SetValue(settingsKey, PromptValueName, 1, RegistryValueKind.DWord);
        }

        private void CreateShortcut()
        {
            if (!File.Exists(executablePath))
            {
                throw new FileNotFoundException("Jarvis PowerPoint executable was not found.", executablePath);
            }

            Directory.CreateDirectory(Path.GetDirectoryName(ShortcutPath));
            object link = new ShellLink();
            try
            {
                var shortcut = (IShellLinkW)link;
                shortcut.SetPath(executablePath);
                shortcut.SetWorkingDirectory(Path.GetDirectoryName(executablePath));
                shortcut.SetDescription("Jarvis PowerPoint - Voice control for slide shows");
                shortcut.SetIconLocation(executablePath, 0);
                shortcut.SetShowCmd(1);
                ((IPersistFile)link).Save(ShortcutPath, true);
            }
            finally
            {
                Marshal.ReleaseComObject(link);
            }
        }

        [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
        private class ShellLink
        {
        }

        // IShellLinkW methods must retain the Windows COM vtable order.
        [ComImport, Guid("000214F9-0000-0000-C000-000000000046"),
            InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellLinkW
        {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder path, int length,
                IntPtr findData, uint flags);
            void GetIDList(out IntPtr itemList);
            void SetIDList(IntPtr itemList);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder description, int length);
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string description);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder directory, int length);
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string directory);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder arguments, int length);
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string arguments);
            void GetHotkey(out short hotkey);
            void SetHotkey(short hotkey);
            void GetShowCmd(out int showCommand);
            void SetShowCmd(int showCommand);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder iconPath,
                int length, out int index);
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string iconPath, int index);
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string path, uint reserved);
            void Resolve(IntPtr window, uint flags);
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string path);
        }
    }
}
