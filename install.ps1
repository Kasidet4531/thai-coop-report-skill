# Install the coop-report-docx skill for Claude Code, Codex and/or Antigravity (Windows).
#
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Target codex,antigravity
#   & ([scriptblock]::Create((irm https://raw.githubusercontent.com/Kasidet4531/thai-coop-report-skill/main/install.ps1))) -Target codex,antigravity
#
# Targets: claude, codex, antigravity, antigravity-cli, project, all (default).
# "project" installs into .\.agents\skills of the current folder (read by Codex and Antigravity).
param([string[]]$Target = @('all'))
$ErrorActionPreference = 'Stop'

$Repo = 'Kasidet4531/thai-coop-report-skill'
$Skill = 'coop-report-docx'
$Home_ = $env:USERPROFILE
$dest = @{
    'claude'          = Join-Path $Home_ '.claude\skills'
    'codex'           = Join-Path $Home_ '.agents\skills'
    'antigravity'     = Join-Path $Home_ '.gemini\config\skills'
    'antigravity-cli' = Join-Path $Home_ '.gemini\antigravity-cli\skills'
    'project'         = Join-Path (Get-Location) '.agents\skills'
}

$targets = @()
foreach ($t in ($Target -split ',')) {
    $t = $t.Trim().ToLower()
    if ($t -eq 'all') { $targets += 'claude', 'codex', 'antigravity' }
    elseif ($dest.ContainsKey($t)) { $targets += $t }
    else { throw "unknown target: $t (use claude, codex, antigravity, antigravity-cli, project, all)" }
}

# Locate the skill source: local clone, or download the repo zip into a temp dir.
$tmp = $null
$local = if ($PSScriptRoot) { Join-Path $PSScriptRoot "skills\$Skill" } else { $null }
if ($local -and (Test-Path (Join-Path $local 'SKILL.md'))) {
    $src = $local
} else {
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("coop-skill-" + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    Write-Output "Downloading $Repo ..."
    $zip = Join-Path $tmp 'repo.zip'
    Invoke-WebRequest -UseBasicParsing "https://github.com/$Repo/archive/refs/heads/main.zip" -OutFile $zip
    Expand-Archive -Path $zip -DestinationPath $tmp
    $src = (Get-ChildItem -Path $tmp -Directory -Recurse | Where-Object { $_.FullName -like "*\skills\$Skill" } | Select-Object -First 1).FullName
}

try {
    foreach ($t in $targets) {
        $d = $dest[$t]
        New-Item -ItemType Directory -Force -Path $d | Out-Null
        $out = Join-Path $d $Skill
        if (Test-Path $out) { Remove-Item -Recurse -Force $out }
        Copy-Item -Recurse $src $out
        $evals = Join-Path $out 'evals'   # test cases are for developing the skill, not using it
        if (Test-Path $evals) { Remove-Item -Recurse -Force $evals }
        Write-Output "installed for $t -> $out"
    }
} finally {
    if ($tmp) { Remove-Item -Recurse -Force $tmp }
}

# Python packages the scripts need
$py = Get-Command python -ErrorAction SilentlyContinue
if ($py) {
    & $py.Source -m pip install --quiet --user python-docx pythainlp
    if ($LASTEXITCODE -eq 0) { Write-Output 'python packages ready: python-docx, pythainlp' }
    else { Write-Output "warning: run 'pip install python-docx pythainlp' yourself" }
} else {
    Write-Output 'warning: Python not found - install Python 3, then: pip install python-docx pythainlp'
}
Write-Output 'Done. Restart the app (or start a new session) so it picks up the skill.'
