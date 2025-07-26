# ===========================================================================================================================
# = Rename-MediaFiles.ps1                                                                                                   =
# = Enhanced PowerShell script to rename media files based on EXIF or file date, with backup, logging, and dry-run support. =
# ===========================================================================================================================
#            .---.
#            |[X]|
#    _.==._.""""".___n__
#    d __ ___.-''-. _____b
#    |[__]  /."""".\ _   |
#    |     // /""\ \\_)  |
#    |     \\ \__/ //    |
#    |      \`.__.'/     |
#    \=======`-..-'======/
#     `-----------------' 
#
#  This script automatically renames your photo and video files in a folder, using the date taken (from EXIF for images, or file date for videos),
#  and appends the original filename. It creates a backup of each file before renaming, supports dry-run mode, and logs all actions.
#
#  - Supported file types: .jpg, .jpeg, .png, .tiff, .mp4, .mov, .avi, .heic (customizable)
#  - Output format: YYYYMMDD_HHmmss - originalname.extension
#  - No files are overwritten; name conflicts are handled automatically.
#  - Use -DryRun to preview changes without modifying files.
#

param (
    [string]$sourceFolderPath = (Get-Location).Path, # Default source folder: the directory where the script is run
    [string]$backupFolderName = "backup",            # Default backup subfolder name
    [string[]]$includeExtensions = @(".jpg", ".jpeg", ".png", ".tiff", ".mp4", ".mov", ".avi", ".heic"), # File types to process
    [switch]$DryRun,                                  # If set, simulate actions only
    [string]$LogFile = "RenameMediaFiles.log"         # Log file path
)

# --- Define Paths ---
# Set up the source and backup folder paths
$sourceFolder = $sourceFolderPath
$backupFolder = Join-Path -Path $sourceFolder -ChildPath $backupFolderName

# --- Logging Function ---
# Logs messages to both the console and a log file, with timestamp and level.
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogFile -Value $logEntry
}

# --- Load System.Drawing Assembly ---
# This assembly is required to read EXIF metadata from image files.
# For video files, this part will likely fail, and the script will fall back to file system dates.
Add-Type -AssemblyName System.Drawing

# --- Log the folders being used ---
Write-Log "Source Folder: $sourceFolder"
Write-Log "Backup Folder: $backupFolder"

# --- Create Backup Folder ---
# Check if the backup folder exists, and create it if it doesn't.
if (-not (Test-Path $backupFolder)) {
    Write-Log "Creating backup folder: $backupFolder"
    if (-not $DryRun) {
        New-Item -Path $backupFolder -ItemType Directory | Out-Null
    }
}

# --- Counters for Summary ---
# These variables track statistics for the summary report at the end
$totalFiles = 0
$processedFiles = 0
$renamedFiles = 0
$errorFiles = 0
$skippedFiles = 0 # (not used, but placeholder for future skip logic)
$logErrors = @()

# --- Process Files ---
# Get all files in the source folder matching the allowed extensions
$files = Get-ChildItem -Path $sourceFolder -File | Where-Object { $includeExtensions -contains $_.Extension.ToLower() }
$totalFiles = $files.Count

# Loop through each file and process
foreach ($file in $files) {
    $extension = $file.Extension
    $newDate = $null
    $originalFullPath = $file.FullName
    $backupFullPath = Join-Path -Path $backupFolder -ChildPath $file.Name
    $processedFiles++

    # --- Attempt to Read EXIF Date (Primarily for Images) ---
    # Try to extract the original date from EXIF metadata if the file is an image
    $isImage = $extension -in @(".jpg", ".jpeg", ".tiff", ".png", ".heic")
    if ($isImage) {
        try {
            $image = [System.Drawing.Image]::FromFile($originalFullPath)
            $propIdDateTimeOriginal = 36867
            if ($image.PropertyIdList -contains $propIdDateTimeOriginal) {
                $propItem = $image.GetPropertyItem($propIdDateTimeOriginal)
                $exifDate = [System.Text.Encoding]::ASCII.GetString($propItem.Value) -replace "`0", ""
                $newDate = [datetime]::ParseExact($exifDate, "yyyy:MM:dd HH:mm:ss", $null)
            }
            $image.Dispose()
        }
        catch {
            # If EXIF reading fails, fall back to file system date
            Write-Log "Could not read EXIF data for '$($file.Name)'. Using file system's last write time." "WARN"
        }
    }

    # --- Fallback to File System Date ---
    # If no EXIF date was found, use the file's last write time
    if ($newDate -eq $null) {
        $newDate = $file.LastWriteTime
    }

    # --- Format New File Name ---
    # Build the new file name: date + original name (no extension) + extension
    $originalNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $baseFileName = $newDate.ToString("yyyyMMdd_HHmmss") + " - " + $originalNameNoExt
    $newFileName = $baseFileName + $extension
    $newFullFileNameInSource = Join-Path -Path $file.DirectoryName -ChildPath $newFileName

    # --- Handle Name Conflicts for Renamed File ---
    # If a file with the new name already exists, append a numeric suffix
    try {
        if (Test-Path $newFullFileNameInSource) {
            Write-Log "Renamed file '$newFileName' already exists in the source folder. Appending numeric suffix." "WARN"
            $counter = 1
            do {
                $tempNewFileName = $baseFileName + "_" + $counter.ToString("00") + $extension
                $newFullFileNameInSource = Join-Path -Path $file.DirectoryName -ChildPath $tempNewFileName
                $counter++
            } while (Test-Path $newFullFileNameInSource)
            $newFileName = $tempNewFileName
        }

        # --- Backup and Rename Operations ---
        # 1. Copy the original file to the backup folder
        Write-Log "Backing up '$($file.Name)' to '$backupFolder'..."
        if ($DryRun) {
            Write-Log "[DryRun] Would copy '$originalFullPath' to '$backupFullPath'"
        } else {
            Copy-Item -Path $originalFullPath -Destination $backupFullPath -Force
        }

        # 2. Rename the original file in the source folder
        Write-Log "Renaming '$($file.Name)' to '$newFileName' in the source folder..."
        if ($DryRun) {
            Write-Log "[DryRun] Would rename '$originalFullPath' to '$newFileName'"
        } else {
            Rename-Item -Path $originalFullPath -NewName $newFileName
            $renamedFiles++
        }
    }
    catch {
        # Log and count any errors during backup/rename
        $errorFiles++
        $logErrors += "Error processing '$($file.Name)': $($_.Exception.Message)"
        Write-Log "Error processing '$($file.Name)': $($_.Exception.Message)" "ERROR"
    }
}

# --- Summary Report ---
# Print a summary of the script's actions
Write-Log "--- SUMMARY ---"
Write-Log "Total files found: $totalFiles"
Write-Log "Files processed: $processedFiles"
Write-Log "Files renamed: $renamedFiles"
Write-Log "Errors: $errorFiles"
if ($logErrors.Count -gt 0) {
    Write-Log "Error details:"
    foreach ($err in $logErrors) { Write-Log $err "ERROR" }
}
Write-Log "Script completed. Log saved to $LogFile."