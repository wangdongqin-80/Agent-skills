param(
    [string]$RepoUrl = "https://github.com/wangdongqin-80/AI-HOT",
    [string]$RepoPath = "",
    [ValidateSet("once", "timer", "watch")]
    [string]$Mode = "once",
    [int]$IntervalSeconds = 300,
    [switch]$EnableWrites,
    [switch]$NoPush,
    [switch]$Toast
)

$ErrorActionPreference = "Stop"

function Write-Notice {
    param([string]$Level, [string]$Message)
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$stamp][$Level] $Message"

    if ($Toast) {
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show($Message, "AI-HOT Git Sync: $Level") | Out-Null
        }
        catch {
            Write-Host "[$stamp][WARN] Toast notification failed: $($_.Exception.Message)"
        }
    }
}

function Invoke-Git {
    param([string[]]$Args, [string]$Cwd)
    $result = & git -C $Cwd @Args 2>&1
    [pscustomobject]@{
        Code = $LASTEXITCODE
        Text = ($result -join "`n")
    }
}

function Stop-OnGitError {
    param($Result, [string]$Action)
    if ($Result.Code -ne 0) {
        throw "$Action failed: $($Result.Text)"
    }
}

function Resolve-RepoPath {
    if ($RepoPath) {
        $resolved = Resolve-Path -LiteralPath $RepoPath -ErrorAction SilentlyContinue
        if ($resolved) {
            return $resolved.Path
        }
        return $RepoPath
    }

    $inside = & git rev-parse --show-toplevel 2>$null
    if ($LASTEXITCODE -eq 0 -and $inside) {
        return $inside.Trim()
    }

    return (Join-Path (Get-Location) "AI-HOT")
}

function Ensure-Repository {
    param([string]$Path)

    if (Test-Path (Join-Path $Path ".git")) {
        return
    }

    if (-not $EnableWrites) {
        throw "Repository not found at '$Path'. Re-run with -EnableWrites to clone $RepoUrl."
    }

    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }

    Write-Notice "INFO" "Cloning $RepoUrl into $Path"
    & git clone $RepoUrl $Path
    if ($LASTEXITCODE -ne 0) {
        throw "git clone failed."
    }
}

function Get-CurrentBranch {
    param([string]$Path)
    $branch = Invoke-Git -Args @("branch", "--show-current") -Cwd $Path
    Stop-OnGitError $branch "git branch --show-current"
    if (-not $branch.Text.Trim()) {
        throw "Detached HEAD is not supported. Check out a branch before enabling sync."
    }
    return $branch.Text.Trim()
}

function Get-Upstream {
    param([string]$Path, [string]$Branch)
    $upstream = Invoke-Git -Args @("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}") -Cwd $Path
    if ($upstream.Code -eq 0 -and $upstream.Text.Trim()) {
        return $upstream.Text.Trim()
    }
    return "origin/$Branch"
}

function Get-AheadBehind {
    param([string]$Path, [string]$Upstream)
    $counts = Invoke-Git -Args @("rev-list", "--left-right", "--count", "$Upstream...HEAD") -Cwd $Path
    Stop-OnGitError $counts "git rev-list"
    $parts = $counts.Text.Trim() -split "\s+"
    [pscustomobject]@{
        Behind = [int]$parts[0]
        Ahead = [int]$parts[1]
    }
}

function Commit-LocalChanges {
    param([string]$Path)
    $status = Invoke-Git -Args @("status", "--short") -Cwd $Path
    Stop-OnGitError $status "git status"

    if (-not $status.Text.Trim()) {
        return
    }

    if (-not $EnableWrites) {
        Write-Notice "DRYRUN" "Local changes detected. Would stage and commit them."
        return
    }

    $add = Invoke-Git -Args @("add", "-A") -Cwd $Path
    Stop-OnGitError $add "git add -A"

    $changedPaths = ($status.Text -split "`n" | Select-Object -First 8) -join "; "
    $message = "chore: auto-sync AI-HOT changes"
    $body = "Created by git-ai-hot-auto-sync.`n`nChanged paths: $changedPaths"
    $commit = Invoke-Git -Args @("commit", "-m", $message, "-m", $body) -Cwd $Path
    if ($commit.Code -ne 0 -and $commit.Text -match "nothing to commit") {
        Write-Notice "INFO" "No local changes to commit."
        return
    }
    Stop-OnGitError $commit "git commit"
    Write-Notice "INFO" "Committed local changes."
}

function Pull-RemoteChanges {
    param([string]$Path)
    if (-not $EnableWrites) {
        Write-Notice "DRYRUN" "Remote updates detected. Would run git pull --no-rebase --ff-only."
        return
    }

    $pull = Invoke-Git -Args @("pull", "--no-rebase", "--ff-only") -Cwd $Path
    if ($pull.Code -ne 0) {
        throw "Pull stopped. Manual merge or conflict resolution may be required: $($pull.Text)"
    }
    Write-Notice "INFO" "Pulled remote updates."
}

function Push-LocalCommits {
    param([string]$Path)
    if ($NoPush) {
        Write-Notice "INFO" "Push skipped because -NoPush was set."
        return
    }
    if (-not $EnableWrites) {
        Write-Notice "DRYRUN" "Local commits are ahead. Would run git push."
        return
    }

    $push = Invoke-Git -Args @("push") -Cwd $Path
    Stop-OnGitError $push "git push"
    Write-Notice "INFO" "Pushed local commits."
}

function Invoke-SyncOnce {
    $path = Resolve-RepoPath
    Ensure-Repository $path

    $remote = Invoke-Git -Args @("remote", "get-url", "origin") -Cwd $path
    Stop-OnGitError $remote "git remote get-url origin"
    if ($remote.Text -notmatch "wangdongqin-80/AI-HOT") {
        throw "Refusing to sync non-AI-HOT remote: $($remote.Text.Trim())"
    }

    Commit-LocalChanges $path

    $fetch = Invoke-Git -Args @("fetch", "--prune") -Cwd $path
    Stop-OnGitError $fetch "git fetch --prune"

    $branch = Get-CurrentBranch $path
    $upstream = Get-Upstream $path $branch
    $sync = Get-AheadBehind $path $upstream

    if ($sync.Behind -gt 0 -and $sync.Ahead -gt 0) {
        throw "Local and remote branches diverged ($($sync.Ahead) ahead, $($sync.Behind) behind). Resolve manually."
    }
    if ($sync.Behind -gt 0) {
        Pull-RemoteChanges $path
    }

    $syncAfterPull = Get-AheadBehind $path $upstream
    if ($syncAfterPull.Ahead -gt 0) {
        Push-LocalCommits $path
    }

    Write-Notice "OK" "Sync complete for $path"
}

function Start-TimerMode {
    while ($true) {
        try {
            Invoke-SyncOnce
        }
        catch {
            Write-Notice "STOP" $_.Exception.Message
        }
        Start-Sleep -Seconds $IntervalSeconds
    }
}

function Start-WatchMode {
    $path = Resolve-RepoPath
    Ensure-Repository $path
    Invoke-SyncOnce

    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $path
    $watcher.IncludeSubdirectories = $true
    $watcher.EnableRaisingEvents = $true
    $watcher.Filter = "*"

    $script:lastRun = Get-Date "2000-01-01"
    $action = {
        $now = Get-Date
        if (($now - $script:lastRun).TotalSeconds -lt 10) {
            return
        }
        $script:lastRun = $now
        try {
            Invoke-SyncOnce
        }
        catch {
            Write-Notice "STOP" $_.Exception.Message
        }
    }

    Register-ObjectEvent $watcher Changed -Action $action | Out-Null
    Register-ObjectEvent $watcher Created -Action $action | Out-Null
    Register-ObjectEvent $watcher Deleted -Action $action | Out-Null
    Register-ObjectEvent $watcher Renamed -Action $action | Out-Null

    Write-Notice "INFO" "Watching $path for changes. Press Ctrl+C to stop."
    while ($true) {
        Wait-Event -Timeout 60 | Out-Null
    }
}

try {
    switch ($Mode) {
        "once" { Invoke-SyncOnce }
        "timer" { Start-TimerMode }
        "watch" { Start-WatchMode }
    }
}
catch {
    Write-Notice "STOP" $_.Exception.Message
    exit 1
}
