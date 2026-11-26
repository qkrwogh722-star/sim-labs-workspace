# PreToolUse hook: block edits/deletes of the source datasets in sim-data/data/.
# Reads the hook payload (JSON) from stdin. Denies by printing a PreToolUse decision.
# Reading the files (Read tool, ZipFile.OpenRead, Get-Content ...) stays allowed.

$ErrorActionPreference = 'Stop'
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Write-Deny([string]$why) {
    $out = @{
        hookSpecificOutput = @{
            hookEventName            = 'PreToolUse'
            permissionDecision       = 'deny'
            permissionDecisionReason = "sim-data/data/ is protected (read-only source data). $why Write outputs to another folder (e.g. sim-data/output/ or the scratchpad) instead."
        }
    }
    $out | ConvertTo-Json -Depth 5 -Compress
    exit 0
}

$raw = [Console]::In.ReadToEnd()
if (-not $raw) { exit 0 }
$payload = $raw | ConvertFrom-Json
$tool  = [string]$payload.tool_name
$input_ = $payload.tool_input

# Matches the data folder by path, or one of its dataset files by name.
$dataPath = '(?i)(sim-data[\\/]+data([\\/]|\b))|((^|[\s''"\\/])data[\\/])|((apple_A|wafer_test_W)\d+\.xlsx)'

switch -Regex ($tool) {
    '^(Edit|Write|MultiEdit|NotebookEdit)$' {
        $p = [string]$input_.file_path
        if (-not $p) { $p = [string]$input_.notebook_path }
        if ($p -match '(?i)sim-data[\\/]+data[\\/]') {
            Write-Deny "Tool '$tool' may not modify '$p'."
        }
        exit 0
    }
    '^(Bash|PowerShell)$' {
        $cmd = [string]$input_.command
        if ($cmd -notmatch $dataPath) { exit 0 }
        # Commands/APIs that write, move, rename or delete files.
        $mutating = '(?i)(\b(Remove-Item|Clear-Content|Set-Content|Add-Content|Out-File|Move-Item|Rename-Item|Copy-Item|New-Item|Set-ItemProperty|Clear-Item|Expand-Archive|Compress-Archive)\b)' +
                    '|((^|[\s;|&(])(rm|del|erase|rd|rmdir|ri|mv|move|ren|rename|cp|copy|cpi|mi|rni|sc|ac|ni|truncate|shred|unlink|tee)(\s|$))' +
                    '|(\bsed\s+(-\w*i|--in-place))' +
                    '|(\[(System\.)?IO\.(File|Directory)\]::(Delete|Move|Copy|Replace|Write\w*|Append\w*|Create\w*|Open\b|SetAttributes))' +
                    '|(ZipFile\]::Open\s*\([^)]*(Update|Create)|ZipArchiveMode|::CreateFromDirectory|\.Save(As)?\s*\(|\.Delete\s*\(|\.MoveTo\s*\(|\.CopyTo\s*\(|\.Create\w*\s*\(|OpenWrite|\bWrite-ZipEntry)' +
                    '|(\battrib\b|\bicacls\b|\btakeown\b)' +
                    '|((^|[^<>=!-])>{1,2}(?!\s*\$null)(?!\s*&))'
        if ($cmd -match $mutating) {
            Write-Deny "This $tool command references the data folder and looks like it writes, moves or deletes files (matched '$($Matches[0].Trim())')."
        }
        exit 0
    }
    default { exit 0 }
}
