# PowerShell script to merge Word documents in a selected subfolder
#Have not tested this yet

# Ensure Word is available
$word = New-Object -ComObject Word.Application
if (-not $word) {
    Write-Host "Microsoft Word is required but not found." -ForegroundColor Red
    Exit
}

# Set base directory
$BaseDir = "C:\Users\ar94\Documents\Combine_WordDocs"

# Function to list subdirectories
function Get-Subfolders($path) {
    Get-ChildItem -Path $path -Directory | Select-Object -ExpandProperty Name
}

# Function to get user-selected folder
function Get-UserSelection($folders) {
    if ($folders.Count -eq 0) {
        Write-Host "No subdirectories found in the base directory."
        return $null
    }

    Write-Host "`nAvailable Folders:"
    for ($i = 0; $i -lt $folders.Count; $i++) {
        Write-Host "$($i+1). $($folders[$i])"
    }

    while ($true) {
        $choice = Read-Host "`nEnter the number of the folder to use"
        if ($choice -match "^\d+$" -and [int]$choice -ge 1 -and [int]$choice -le $folders.Count) {
            return $folders[$choice - 1]
        } else {
            Write-Host "Invalid selection. Please enter a valid folder number."
        }
    }
}

# Function to get common filename prefix
function Get-CommonPrefix($files) {
    $prefixes = $files | ForEach-Object { ($_ -split '\.')[0] } | Sort-Object -Unique
    return $(if ($prefixes.Count -eq 1) { $prefixes[0] } else { "combined_document" })
}

# Function to combine Word documents
function Combine-WordDocuments($inputFolder, $outputFolder) {
    # Ensure output directory exists
    if (!(Test-Path $outputFolder)) { New-Item -ItemType Directory -Path $outputFolder | Out-Null }

    # Get all .docx files and ignore temp files (~$)
    $wordFiles = Get-ChildItem -Path $inputFolder -Filter "*.docx" | Where-Object { $_.Name -notmatch "^~\$" } | Sort-Object Name

    if ($wordFiles.Count -eq 0) {
        Write-Host "No valid Word documents found in the folder."
        return
    }

    # Get output filename
    $outputFilename = "$(Get-CommonPrefix $wordFiles.Name).docx"
    $outputFilePath = Join-Path -Path $outputFolder -ChildPath $outputFilename

    # Open first document
    $masterDoc = $word.Documents.Open($wordFiles[0].FullName)

    # Append remaining documents
    for ($i = 1; $i -lt $wordFiles.Count; $i++) {
        Write-Host "Adding: $($wordFiles[$i].Name)"
        $masterDoc.Application.Selection.EndKey(6)  # Move cursor to end of document
        $masterDoc.Application.Selection.InsertBreak(1)  # Insert page break
        $masterDoc.Application.Selection.InsertFile($wordFiles[$i].FullName)
    }

    # Save and close
    $masterDoc.SaveAs([ref] $outputFilePath)
    $masterDoc.Close($false)
    $word.Quit()

    Write-Host "Merged document saved as: $outputFilePath"
    Write-Host "Total number of Word documents combined: $($wordFiles.Count)"
}

# Main execution
$folders = Get-Subfolders -path $BaseDir
$selectedFolder = Get-UserSelection -folders $folders

if ($selectedFolder) {
    $inputFolder = Join-Path -Path $BaseDir -ChildPath $selectedFolder
    $outputFolder = Join-Path -Path $BaseDir -ChildPath "Convert"  # Keep output in 'Convert' folder
    Combine-WordDocuments -inputFolder $inputFolder -outputFolder $outputFolder
} else {
    Write-Host "No folder selected. Exiting script."
}