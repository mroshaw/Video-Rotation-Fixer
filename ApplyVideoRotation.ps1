Write-Host "Executing: $($MyInvocation.MyCommand.Path)"
# ============================================================
# Configuration
# ============================================================

# Full path of the source to process. File and subfolders will all be processed
[string]$sourceFolderRoot = "Z:\Home Movies\"

# Full path of the folder where backups will be written. Only files rotated will
# be backed up. Subfolders will be created, aligned to the structure of the source
[string]$backupFolder = "E:\Media\Rotation Backup"

# Name of the log file. Will be created in the same folder as this script
# with the execution timestamp appended to the name
[string]$logFileBaseName = "PlexRotationFixLog"

# Set to '$true' to generate output without making any changes.
# Set to '$false' to apply rotation to files
# RECOMMENDED on your first run
[bool]$reportOnly = $false

# Set to '$true' to backup files before making any changes
# RECOMMENDED!!!
[bool]$backupFiles = $true

# Set to '$true' to abort the entire process if any errors are encountered
# Set to '$false' to skip the file that caused the error and continue
[bool]$abortOnError = $false

# Set to the path of the 'ffmpeg' executable (ffmpeg.exe)
[string]$ffmpegFolderPath = "E:\Dev\ffmpeg\bin"

# Add additional file types, if these don't cover what you need
[string[]]$videoExtensions = @("*.mp4", "*.mov", "*.mts")

# ============================================================
# DO NOT MAKE CHANGES BEYOND THIS POINT!!!!!
# ============================================================

# ============================================================
# Derived paths
# ============================================================
[string]$ffmpegPath  = Join-Path $ffmpegFolderPath "ffmpeg.exe"
[string]$ffprobePath = Join-Path $ffmpegFolderPath "ffprobe.exe"

[string]$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
[string]$scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
[string]$logFileName = "$logFileBaseName`_$timestamp.txt"
[string]$logFile = Join-Path $scriptFolder $logFileName

# ============================================================
# Create log
# ============================================================
try {
    if (!(Test-Path $logFile)) {
        New-Item -ItemType File -Path $logFile | Out-Null
    }
} catch {
    Write-Host "Failed to create log file! ABORTING PROCESS!"
    Write-Host "Exception type : $($_.Exception.GetType().FullName)"
    Write-Host "Message        : $($_.Exception.Message)"
    Write-Host "Source         : $($_.Exception.Source)"
    Write-Host "StackTrace     :"
    Write-Host $_.Exception.StackTrace
    return
}

# ============================================================
# Main file processing
# ============================================================
function ProcessFile {
    param (
        [string]$filePath
    )

    [string]$fileLog = ""

    $fileInfo = Get-Item $filePath
    [string]$fileName = $fileInfo.Name
    [long]$fileSize = $fileInfo.Length

    # Get meta data via ffprobe
    [string]$widthStr = GetFfprobeValue $ffprobePath @("-v", "error","-select_streams", "v:0","-show_entries", "stream=width","-of", "default=nw=1:nk=1",$filePath)
    [string]$heightStr = GetFfprobeValue $ffprobePath @("-v", "error", "-select_streams", "v:0", "-show_entries", "stream=height", "-of", "default=nw=1:nk=1", $filePath)
    [string]$displayMatrix = GetFfprobeValue $ffprobePath @("-v", "error", "-select_streams", "v:0", "-show_entries", "stream_side_data=rotation", "-of", "default=nw=1:nk=1", $filePath)

    # Safely cast to int
    $width = 0
    if ([int]::TryParse($widthStr, [ref]$width) -eq $false) {
        $width = 0
    }

    $height = 0
    if ([int]::TryParse($heightStr, [ref]$height) -eq $false) {
        $height = 0
    }

    # Get rotation needed
    $rotation = 0
    if ([int]::TryParse($displayMatrix, [ref]$rotation) -eq $false) {
        $rotation = 0
    }

    [string]$filterParams = ""

    switch ($rotation) {
        0   { $filterParams = "" }
        90  { $filterParams = "transpose=2" }
        180 { $filterParams = "transpose=2,transpose=2" }
        -90 { $filterParams = "transpose=1" }
        -180 { $filterParams = "transpose=1,transpose=1" }
        default {
            LogMessage "$fileName : Unknown rotation: $rotation — skipped"
            return
        }
    }

    $fileLog = "$filePath (META WxH=${width}x${height} DisplayMatrix=$displayMatrix RotationToApply=$rotation)"

    # No rotation required, so we can skip
    if ($rotation -eq 0) {
        LogMessage "$fileLog. No rotation needed, SKIPPING"
        return
    }

    # Track impact
    $script:impactedFileCount++
    $script:impactedSourceBytes += $fileSize

    # Backup, if we configured in parameters
    if ($backupFiles) {
        $relativePath = GetRelativePath $filePath $sourceFolderRoot
        if (-not $reportOnly) {
            BackupFile $filePath $fileSize $backupFolder $sourceFolderRoot $relativePath
        }
        $fileLog += ". Backed up to: $backupFolder $relativePath"
    }

    # If rotation is required, rotate via ffmpeg
    if ($rotation -ne 0) {
        if (-not $reportOnly) {
            RotateVideo $filePath $fileName $ffmpegPath $filterParams                
        }
        $fileLog += ". Rotated using parameters: $filterParams"
    }
    LogMessage $fileLog
}

# ============================================================
# Helper Functions
# ============================================================

# ============================================================
# Validate the script parameters
# ============================================================
function ValidateParameters {
    param (
        [string]$ffmpegPath,
        [string]$ffprobePath,
        [string]$sourceFolderRoot,
        [string]$backupFolder
    )
    if (!(Test-Path $ffmpegPath)) {
        throw "Invalid configuration: ffmpeg.exe not found at path: $ffmpegPath"
    }

    if (!(Test-Path $ffprobePath)) {
        throw "Invalid configuration: ffprobe.exe not found at path: $ffprobePath"
    }

    if (!(Test-Path $sourceFolderRoot)) {
        throw "Invalid configuration: Source folder not found at path: $sourceFolderRoot"
    }

    # Ensure backup folder is not within source folder
    [string]$resolvedSourceRoot = (Resolve-Path $sourceFolderRoot).Path
    [string]$resolvedBackupRoot = [System.IO.Path]::GetFullPath($backupFolder)

    # Normalise trailing separators
    $resolvedSourceRoot = $resolvedSourceRoot.TrimEnd('\') + '\'
    $resolvedBackupRoot = $resolvedBackupRoot.TrimEnd('\') + '\'

    if ($resolvedBackupRoot.StartsWith($resolvedSourceRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Invalid configuration: BackupFolder must not be inside SourceFolderRoot. BackupFolder=$backupFolder"
    }
} 

# ============================================================
# Logging helper
# ============================================================
function LogMessage {
    param (
        [string]$message
    )
    $message = $message.Trim()
    [string]$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    [string]$line = "$timestamp - $message"

    Write-Host $line
    Add-Content -Path $logFile -Value $line
}



# ============================================================
# Returns the relative path of the given file
# ============================================================
function GetRelativePath {
    param (
        [string]$fullPath,
        [string]$rootPath
    )

    return $fullPath.Substring($rootPath.Length).TrimStart("\")
}

# ============================================================
# Ensures the given directory path exists. Create if it doesn't
# ============================================================
function EnsureDirectory {
    param (
        [string]$path
    )

    if (!(Test-Path $path)) {
        if (-not $script:loggedFolders.ContainsKey($path)) {
            $script:loggedFolders[$path] = $true

            if (-not $reportOnly) {
                New-Item -ItemType Directory -Path $path | Out-Null
            }                
            LogMessage "Created folder: $path"
        }
    }
}

# ========================================================
# Backup handling
# ========================================================
function BackupFile {
    param (
        [string]$filePath,
        [long]$fileSize,
        [string]$backupFolder,
        [string]$sourceFolderRoot,
        [string]$relativePath
    )
    [string]$backupTargetPath = Join-Path $backupFolder $relativePath
    [string]$backupTargetDir  = Split-Path $backupTargetPath -Parent

    EnsureDirectory $backupTargetDir
    $script:backupBytes += $fileSize

    Copy-Item -Path $filePath -Destination $backupTargetPath -Force
}

# ========================================================
# Safely retrieve meta properties using ffprobe
# ========================================================
function GetFfprobeValue {
    param (
        [string]$ffprobePath,
        [string[]]$arguments
    )

    $value = & $ffprobePath @arguments 2>$null

    if ($null -eq $value) {
        return ""
    }

    # Join multiple lines into one string (first non-empty line)
    if ($value -is [System.Array]) {
        $value = ($value | Where-Object { $_.Trim() -ne "" })[0]
    }

    return $value.ToString().Trim()
}

# ========================================================
# In-place rotation
# ========================================================
function RotateVideo {
    param (
        [string]$filePath,
        [string]$fileName,
        [string]$ffmpegPath,
        [string]$filterParams
    )
    
    [string]$tempPath = Join-Path (Split-Path $filePath) ("temp_" + $fileName)

    $ffmpegArgs = @(
        "-noautorotate",
        "-i", "`"$filePath`"",
        "-vf", $filterParams,
        "-map_metadata", "0",
        "-metadata:s:v:0", "rotate=0",
        "-c:v", "libx264",
        "-crf", "18",
        "-preset", "slow",
        "-c:a", "copy",
        "`"$tempPath`""
    )

    $process = Start-Process -FilePath $ffmpegPath -ArgumentList $ffmpegArgs -NoNewWindow -Wait -PassThru

    if ($process.ExitCode -eq 0 -and (Test-Path $tempPath)) {
        RestoreFile $filePath $tempPath $fileName
    } else {
        if (Test-Path $tempPath) {
            RemoveFile $filePath $tempPath
        }
        throw "$filePath : FAILED to process! Command was: $ffmpegPath $ffmpegArgs"
    }
}

# ========================================================
# Safely remove file, waiting for file handle to be released
# ========================================================
function RemoveFile {
    param (
        [string]$filePath,
        [string]$tempPath
    )

    $maxAttempts = 5
    $attempt = 0
    $success = $false

    while (-not $success -and $attempt -lt $maxAttempts) {
        try {
            Remove-Item -Path $filePath -Force -ErrorAction Stop
            $success = $true
            return
        }
        catch {
            $attempt++
            Start-Sleep -Milliseconds 300
        }
    }

    if (-not $success) {
        if (Test-Path $tempPath) {
            Remove-Item $tempPath -Force
        }
        throw "$filePath : FAILED to replace original after $maxAttempts attempts"
    }
}

# ========================================================
# Safely restore file, waiting for file handle to be released
# ========================================================
function RestoreFile {
    param (
        [string]$filePath,
        [string]$tempPath,
        [string]$fileName
    )

    $maxAttempts = 5
    $attempt = 0
    $success = $false

    while (-not $success -and $attempt -lt $maxAttempts) {
        try {
            Remove-Item -Path $filePath -Force -ErrorAction Stop
            Rename-Item -Path $tempPath -NewName $fileName -ErrorAction Stop
            $success = $true
            return
        }
        catch {
            $attempt++
            Start-Sleep -Milliseconds 300
        }
    }

    if (-not $success) {
        if (Test-Path $tempPath) {
            Remove-Item $tempPath -Force
        }
        throw "$filePath : FAILED to replace original after $maxAttempts attempts"
    }
}

# ========================================================
# Formats bytes into MB, GB, TB
# ========================================================
function FormatBytes {
    param (
        [long]$bytes
    )

    if ($bytes -ge 1TB) { "{0:N2} TB" -f ($bytes / 1TB) }
    elseif ($bytes -ge 1GB) { "{0:N2} GB" -f ($bytes / 1GB) }
    elseif ($bytes -ge 1MB) { "{0:N2} MB" -f ($bytes / 1MB) }
    else { "$bytes bytes" }
}

# ============================================================
# Validate parameters
# ============================================================
try {
    ValidateParameters $ffmpegPath $ffprobePath $sourceFolderRoot $backupFolder
    
} catch {
    LogMessage "Parameter validation failed!"
    LogMessage "Exception type : $($_.Exception.GetType().FullName)"
    LogMessage "Message        : $($_.Exception.Message)"
    LogMessage "Source         : $($_.Exception.Source)"
    return
}

# ============================================================
# State tracking
# ============================================================
[int]$script:impactedFileCount = 0
[long]$script:impactedSourceBytes = 0
[long]$script:backupBytes = 0

# Track folders already created/logged
$script:loggedFolders = @{}

# ============================================================
# Main execution
# ============================================================
LogMessage "--------------------------------------------------"
LogMessage "Scan started!"
LogMessage "ReportOnly: $reportOnly, BackupFiles: $backupFiles"
LogMessage "Logging to: $logFile"
LogMessage "Source Folder Root: $sourceFolderRoot"
LogMessage "--------------------------------------------------"
if ($backupFiles) {
    LogMessage "Backup Folder: $backupFolder"
}

try {

    foreach ($ext in $videoExtensions) {
        Get-ChildItem -Path $sourceFolderRoot -Filter $ext -Recurse | ForEach-Object {
            $currentFile = $_
            $filePath = $currentFile.FullName

            try {
                ProcessFile $filePath
            } catch {
            
                LogMessage "An error occurred while processing: $filePath"
                LogMessage "Exception type : $($_.Exception.GetType().FullName)"
                LogMessage "Message        : $($_.Exception.Message)"
                LogMessage "Source         : $($_.Exception.Source)"

                if($abortOnError) {
                    throw
                }
                LogMessage "Skipping $filePath"
            }
        }
    }
} catch {
    LogMessage "ABORTING PROCESS!"
}

LogMessage "--------------------------------------------------"
LogMessage "SUMMARY"
LogMessage ""
LogMessage "Mode                  : $(if ($reportOnly) { 'REPORT ONLY' } else { 'APPLIED' })"
LogMessage "Impacted files        : $script:impactedFileCount"
LogMessage "Source data impacted  : $(FormatBytes $script:impactedSourceBytes)"
LogMessage "Backup data size      : $(FormatBytes $script:backupBytes)"
LogMessage "--------------------------------------------------"