[CmdletBinding()]
param([string]$Executable)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Executable)) {
    $Executable = Join-Path (Split-Path -Parent $PSScriptRoot) "bin\JarvisPowerPoint.exe"
}
Add-Type -AssemblyName System.Windows.Forms
$assembly = [Reflection.Assembly]::LoadFrom([IO.Path]::GetFullPath($Executable))
$flags = [Reflection.BindingFlags]"Instance,NonPublic"
$script:assertions = 0
function Assert-True([bool]$condition, [string]$message) {
    if (!$condition) { throw $message }
    $script:assertions++
}
function New-AppObject([string]$name, [object[]]$arguments = @()) {
    [Activator]::CreateInstance($assembly.GetType("JarvisPowerPoint.$name", $true), $arguments)
}
function Field($object, [string]$name) {
    $object.GetType().GetField($name, $flags).GetValue($object)
}
function New-AppList([string]$name) {
    $definition = [type]'System.Collections.Generic.List`1'
    $closed = $definition.MakeGenericType([type[]]@($assembly.GetType("JarvisPowerPoint.$name", $true)))
    return ,([Activator]::CreateInstance($closed))
}

$snapshot = New-AppObject "PresentationSnapshot"
$snapshot.SessionKey = "origin-session"
$snapshot.PresentationKey = "origin-deck"
$snapshot.SavedPath = "C:\Synthetic\Deck.pptx"
$snapshot.PresentationName = "Synthetic"
$snapshot.SlideId = 101
$snapshot.SlideNumber = 1
$snapshot.Title = "Slide 1"
$savedBudgets = New-Object 'Collections.Generic.Dictionary[int,int]'
$savedBudgets.Add(101, 12)
$savedBudgets.Add(202, 45)
$tracker = New-AppObject "RehearsalTracker"
$tracker.Start($snapshot, [TimeSpan]::Zero, $savedBudgets)
Assert-True ($tracker.Entries[0].BudgetSeconds -eq 12) "Saved budget was not applied on the first slide."
Assert-True ($tracker.SavedPath -eq $snapshot.SavedPath) "The report lost its origin path."
Assert-True ($tracker.PresentationKey -eq $snapshot.PresentationKey) "The report lost its origin identity."
Assert-True ($tracker.StartedUtc.Kind -eq [DateTimeKind]::Utc) "Run timestamps must be UTC."
$runId = $tracker.RunId
Assert-True ($runId -match '^[a-f0-9]{32}$') "Run ID must be stable and suitable for retry/upsert."
$snapshot.SlideId = 202
$snapshot.SlideNumber = 2
$tracker.Observe($snapshot, [TimeSpan]::FromSeconds(15))
Assert-True ($tracker.Entries[1].BudgetSeconds -eq 45) "Saved budget did not follow the next slide."
$tracker.SetBudget(202, 60)
$tracker.Observe($snapshot, [TimeSpan]::FromSeconds(20))
Assert-True ($tracker.Entries[1].BudgetSeconds -eq 60) "Manual override was overwritten by observation."
$snapshot.SessionKey = "different-session"
$snapshot.SavedPath = "C:\Synthetic\Other.pptx"
$tracker.Observe($snapshot, [TimeSpan]::FromSeconds(25))
Assert-True (!$tracker.IsRunning) "Changing deck must stop rehearsal."
Assert-True ($tracker.SavedPath -eq "C:\Synthetic\Deck.pptx") "A deck switch redirected history to the new deck."
Assert-True ($tracker.RunId -eq $runId) "Stopping must not change the idempotence key."

$panel = New-AppObject "PresenterPanel" ([object[]]@($true))
try {
    Assert-True ($panel.CautiousSearch) "Cautious search should be the safe default."
    Assert-True (!$panel.HotkeysEnabled) "Global hotkeys must be opt-in."
    Assert-True (!$panel.IndicatorEnabled) "Indicator must not appear automatically."
    $panel.CautiousSearch = $false
    Assert-True (!$panel.CautiousSearch) "Direct search opt-out is missing."
    $panel.HotkeysEnabled = $true
    Assert-True ($panel.HotkeysEnabled) "Hotkey toggle is missing."
    $panel.SetLanguage($false)
    Assert-True ((Field $panel "cautiousSearch").Text -match "Confirmer") "French cautious-search label missing."
    $panel.SetLanguage($true)
    Assert-True ((Field $panel "hotkeysEnabled").Text -match "Ctrl") "Keyboard shortcut description missing."
} finally { $panel.Dispose() }

$proposal = New-AppObject "SearchProposal"
$choices = New-AppList "SlideChoice"
foreach ($number in 1..7) {
    $choice = New-AppObject "SlideChoice"
    $choice.SlideId = 100 + $number
    $choice.SlideNumber = $number
    $choice.Score = 500
    $choice.Title = "Synthetic $number"
    $choices.Add($choice)
}
$proposal.Candidates = $choices
$dialog = New-AppObject "SearchChoiceDialog" ([object[]]@($proposal, $true))
try {
    $list = Field $dialog "candidates"
    Assert-True ($list.Items.Count -eq 5) "Ambiguous picker must bound the visible shortlist."
    $confirm = Field $dialog "confirm"
    Assert-True ($null -eq $dialog.AcceptButton) "Ambiguous search must not have a default acceptance action."
    Assert-True (!$confirm.Enabled) "No candidate must be confirmed by default."
    Assert-True ($dialog.SelectedSlideId -eq 0) "Opening the dialog selected a slide implicitly."
    $dialog.Show()
    [Windows.Forms.Application]::DoEvents()
    $list.Items[1].Selected = $true
    [Windows.Forms.Application]::DoEvents()
    Assert-True ($confirm.Enabled) "Selected result cannot be confirmed."
    Assert-True ($null -eq $dialog.AcceptButton) "Selecting a candidate must not make Enter navigate."
    $confirm.PerformClick()
    Assert-True ($dialog.DialogResult -eq [Windows.Forms.DialogResult]::OK) "Confirmation did not return OK."
    Assert-True ($dialog.SelectedSlideId -eq 102) "Picker returned an index instead of the stable slide ID."
} finally { $dialog.Dispose() }
$dialog = New-AppObject "SearchChoiceDialog" ([object[]]@($proposal, $false))
try {
    $dialog.Show()
    [Windows.Forms.Application]::DoEvents()
    $dialog.CancelButton.PerformClick()
    Assert-True ($dialog.SelectedSlideId -eq 0) "Cancel must not select a slide."
    Assert-True ($dialog.DialogResult -eq [Windows.Forms.DialogResult]::Cancel) "Cancel was not returned."
} finally { $dialog.Dispose() }

$script:closedForUpdate = $false
$update = New-AppObject "UpdateDialog" ([object[]]@($true, [Func[bool]]{return $true}, [Action]{$script:closedForUpdate=$true}))
try {
    Assert-True (!(Field $update "install").Enabled) "Installation must be unavailable before a successful check."
    Assert-True ((Field $update "check").Enabled) "Manual update check is missing."
    Assert-True ((Field $update "details").Text -match "No network request") "Opening updater should not start networking."
    Assert-True (!$script:closedForUpdate) "Opening updater closed the app without consent."
    $update.Close()
} finally { $update.Dispose() }

Write-Host "Passed $assertions rehearsal-origin, privacy-default and confirmation UI assertions."
