[CmdletBinding()]
param([string]$Executable = (Join-Path (Split-Path -Parent $PSScriptRoot) "bin\JarvisPowerPoint.exe"))

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Speech
$assembly = [Reflection.Assembly]::LoadFrom($Executable)
$script:assertions = 0

function Assert-True([bool]$condition, [string]$message) {
    if (!$condition) { throw $message }
    $script:assertions++
}
function New-Instance([string]$type, [object[]]$arguments = @()) {
    [Activator]::CreateInstance($assembly.GetType("JarvisPowerPoint.$type", $true), $arguments)
}
function New-Snapshot([string]$session, [int]$id, [int]$number, [bool]$black = $false) {
    $item = New-Instance "PresentationSnapshot"
    $item.SessionKey = $session
    $item.PresentationKey = "Synthetic"
    $item.PresentationName = "=Synthetic deck"
    $item.SlideId = $id
    $item.SlideNumber = $number
    $item.SlideCount = 3
    $item.Title = "@Title $id"
    $item.IsBlack = $black
    $item
}
function Time([double]$seconds) { [TimeSpan]::FromSeconds($seconds) }

$tracker = New-Instance "RehearsalTracker"
$tracker.DefaultBudgetSeconds = 15
$one = New-Snapshot "A" 101 1
$two = New-Snapshot "A" 202 2
$tracker.Start($one, (Time 0))
$tracker.Observe($two, (Time 10))
$tracker.Observe($one, (Time 30))
$tracker.Observe((New-Snapshot "A" 101 1 $true), (Time 35))
$tracker.Observe($one, (Time 60))
$tracker.Stop((Time 65))
Assert-True (!$tracker.IsRunning) "Stopping the rehearsal failed."
Assert-True ($tracker.TotalSeconds -eq 40) "Black-screen time was not excluded or revisit totals are wrong."
Assert-True ($tracker.Entries.Count -eq 2) "Revisited slides must aggregate by SlideID."
Assert-True ($tracker.Entries[0].Seconds -eq 20) "Incorrect accumulated time on slide 1."
Assert-True ($tracker.Entries[1].Seconds -eq 20) "Incorrect accumulated time on slide 2."
Assert-True ($tracker.Entries[0].OverBudget) "Exceeded budget was not flagged."
$tracker.SetBudget(101, 25)
Assert-True (!$tracker.Entries[0].OverBudget) "Per-slide budget override failed."
$tracker.Observe($two, (Time 100))
Assert-True ($tracker.TotalSeconds -eq 40) "A stopped rehearsal accumulated time."

$report = Join-Path ([IO.Path]::GetTempPath()) ("Jarvis-rehearsal-" + [Guid]::NewGuid().ToString("N") + ".csv")
try {
    $tracker.Export($report)
    $data = Import-Csv -LiteralPath $report
    Assert-True ($data.Count -eq 2) "CSV row count is wrong."
    Assert-True ($data[0].Presentation.StartsWith("'=")) "CSV presentation text can execute a formula."
    Assert-True ($data[0].Title.StartsWith("'@")) "CSV slide text can execute a formula."
    Assert-True ($data[0].Seconds -eq "20.0") "CSV numeric output is culture-dependent."
} finally {
    if (Test-Path -LiteralPath $report) { Remove-Item -LiteralPath $report }
}

$tracker.Start($one, (Time 100))
$tracker.Observe((New-Snapshot "B" 101 1), (Time 105))
Assert-True (!$tracker.IsRunning) "Switching the slideshow must stop rehearsal."
Assert-True ($tracker.TotalSeconds -eq 5) "Slideshow change attributed time to another deck."
$tracker.Start($one, (Time 200))
$tracker.Observe($null, (Time 204))
Assert-True (!$tracker.IsRunning) "Ending a slideshow must stop rehearsal."
Assert-True ($tracker.TotalSeconds -eq 4) "End-of-show totals incorrect."

$panel = New-Instance "PresenterPanel" ([object[]]@($false))
try {
    $panel.SetRehearsal($tracker)
    $panel.SetStatus("Listening", "Jarvis test", "No slide changed", 100)
    $panel.SetLanguage($true)
    $tabs = $panel.Controls[0]
    Assert-True ($tabs.TabPages.Count -eq 3) "Presenter panel tabs missing."
    Assert-True ($tabs.TabPages[0].Text -eq "Presenter") "Presenter panel language switch failed."
    Assert-True ($panel.BudgetSeconds -eq 90) "Default rehearsal budget incorrect."
    $panel.Show()
    [System.Windows.Forms.Application]::DoEvents()
    Assert-True ($panel.Visible) "Presenter panel does not open."
    $panel.Close()
    Assert-True (!$panel.Visible -and !$panel.IsDisposed) "Closing presenter panel must hide, not exit Jarvis."
} finally { $panel.Dispose() }

$type = $assembly.GetType("JarvisPowerPoint.JarvisApplicationContext", $true)
$context = [Runtime.Serialization.FormatterServices]::GetUninitializedObject($type)
$flags = [Reflection.BindingFlags]"Instance,NonPublic"
$cases = @(
    @{Culture="fr-FR"; Phrases=@(
        @("Jarvis reprends", "JarvisActions", "resume"),
        @("Jarvis mode questions", "JarvisActions", "questions"),
        @("Jarvis autre r\u00e9sultat", "JarvisActions", "results"),
        @("Jarvis \u00e9cran noir", "JarvisActions", "black"),
        @("Jarvis affiche", "JarvisActions", "display"),
        @("Jarvis d\u00e9marre la r\u00e9p\u00e9tition", "JarvisActions", "rehearse"),
        @("Jarvis arr\u00eate la r\u00e9p\u00e9tition", "JarvisActions", "stopRehearsal"),
        @("Jarvis raccourci la d\u00e9mo", "JarvisAlias", ""),
        @("Jarvis cherche budget", "JarvisSearch", ""),
        @("Jarvis va au slide vingt et un", "JarvisGoToSlide", "21")
    )},
    @{Culture="en-US"; Phrases=@(
        @("Jarvis resume", "JarvisActions", "resume"),
        @("Jarvis questions mode", "JarvisActions", "questions"),
        @("Jarvis another result", "JarvisActions", "results"),
        @("Jarvis black screen", "JarvisActions", "black"),
        @("Jarvis restore slides", "JarvisActions", "display"),
        @("Jarvis start rehearsal", "JarvisActions", "rehearse"),
        @("Jarvis stop rehearsal", "JarvisActions", "stopRehearsal"),
        @("Jarvis shortcut the demo", "JarvisAlias", ""),
        @("Jarvis search for budget", "JarvisSearch", ""),
        @("Jarvis go to slide twenty one", "JarvisGoToSlide", "21")
    )}
)
foreach ($case in $cases) {
    $info = [System.Speech.Recognition.SpeechRecognitionEngine]::InstalledRecognizers() |
        Where-Object { $_.Culture.Name -eq $case.Culture } | Select-Object -First 1
    if (!$info) { throw "Required test recognizer missing: $($case.Culture)" }
    $engine = [System.Speech.Recognition.SpeechRecognitionEngine]::new($info)
    try {
        foreach ($name in @("CreateActionGrammar", "CreateAliasGrammar", "CreateSearchGrammar", "CreateGoToGrammar")) {
            $grammar = $type.GetMethod($name, $flags).Invoke($context, @($info.Culture))
            $engine.LoadGrammar($grammar)
        }
        foreach ($phrase in $case.Phrases) {
            $text = [regex]::Unescape($phrase[0])
            $result = $engine.EmulateRecognize($text)
            Assert-True ($null -ne $result) "Phrase not recognized: $text"
            Assert-True ($result.Grammar.Name -eq $phrase[1]) "Wrong grammar won: $text"
            if ($phrase[1] -eq "JarvisActions") {
                Assert-True ($result.Semantics["action"].Value -eq $phrase[2]) "Wrong action: $text"
            } elseif ($phrase[1] -eq "JarvisGoToSlide") {
                Assert-True ($result.Semantics["slideNumber"].Value -eq 21) "Number navigation regressed."
            }
        }
    } finally { $engine.Dispose() }
}
Write-Host "Passed $assertions presenter, rehearsal and grammar assertions."
