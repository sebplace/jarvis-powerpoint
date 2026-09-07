[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$temporary = Join-Path ([IO.Path]::GetTempPath()) ("JarvisSearchUI-" + [Guid]::NewGuid().ToString("N"))
$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (!(Test-Path -LiteralPath $compiler)) { $compiler = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
$code = @'
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using JarvisPowerPoint;

internal static class SearchProgressTests
{
    private static int assertions;
    private static void Assert(bool value, string message)
    {
        if (!value) throw new Exception(message);
        assertions++;
    }
    [STAThread]
    private static int Main()
    {
        Application.EnableVisualStyles();
        foreach (bool english in new[] { false, true })
        {
            int steps = 0;
            using (var dialog = new SearchProgressDialog(english, delegate { return ++steps == 3; }, delegate { return steps.ToString(); }))
            {
                Assert(steps == 0, "Constructor must not start extraction.");
                Assert(dialog.AcceptButton == null, "Enter must not confirm navigation.");
                Assert(dialog.CancelButton != null, "Escape must cancel.");
                Assert(dialog.ShowDialog() == DialogResult.OK && steps == 3, "Incremental search did not complete.");
                dialog.ThrowIfFailed();
            }
            steps = 0;
            using (var dialog = new SearchProgressDialog(english, delegate { steps++; return false; }, delegate { return ""; }))
            {
                dialog.Shown += delegate { ((Button)dialog.CancelButton).PerformClick(); };
                Assert(dialog.ShowDialog() == DialogResult.Cancel, "Cancel result missing.");
                Assert(steps == 0, "Cancelled search continued extraction.");
            }
            using (var dialog = new SearchProgressDialog(english, delegate { throw new IOException("Synthetic search failure"); }, delegate { return ""; }))
            {
                Assert(dialog.ShowDialog() == DialogResult.Cancel, "Failure appeared successful.");
                bool surfaced = false;
                try { dialog.ThrowIfFailed(); }
                catch (IOException error) { surfaced = error.Message == "Synthetic search failure"; }
                Assert(surfaced, "Original failure was swallowed.");
            }
            using (var dialog = new SearchProgressDialog(english, delegate { return true; }, delegate { throw new InvalidOperationException("Stale progress"); }))
            {
                Assert(dialog.ShowDialog() == DialogResult.Cancel, "Progress failure appeared successful.");
                bool surfaced = false;
                try { dialog.ThrowIfFailed(); }
                catch (InvalidOperationException) { surfaced = true; }
                Assert(surfaced, "Progress failure was swallowed.");
            }
        }
        Console.WriteLine("Passed " + assertions + " cancellable search UI assertions.");
        return 0;
    }
}
'@
try {
    New-Item -ItemType Directory -Path $temporary | Out-Null
    $source = Join-Path $temporary "SearchProgressTests.cs"
    [IO.File]::WriteAllText($source, $code, [Text.UTF8Encoding]::new($false))
    $exe = Join-Path $temporary "SearchProgressTests.exe"
    & $compiler /nologo /target:exe /langversion:5 "/out:$exe" `
        /reference:System.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll `
        (Join-Path $root "SearchProgressDialog.cs") (Join-Path $root "UiAccessibility.cs") $source
    if ($LASTEXITCODE -ne 0) { throw "Search progress test compilation failed." }
    & $exe
    if ($LASTEXITCODE -ne 0) { throw "Search progress tests failed ($LASTEXITCODE)." }
} finally {
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
