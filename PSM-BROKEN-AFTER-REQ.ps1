clear-host

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define the base script path
$global:preloadedMediaPlayer = $null
$global:textBoxBackupCount = $null

# Retrieve or Set Base Script Path
$global:BaseScriptPath = Get-ConfigScriptPath
if (-not $global:BaseScriptPath) {
    Set-ScriptsPath
}



# Adding Class Requirements for Requirements DataGrid, etc.
class Requirement {
    [string]$Requirement
    [bool]$Fulfilled
}

$requirementsList = New-Object System.ComponentModel.BindingList[Requirement]


$colorList = @()

[System.Enum]::GetValues([System.Drawing.KnownColor]) | ForEach-Object {
    $color = [System.Drawing.Color]::FromKnownColor($_)
    $colorList += [PSCustomObject]@{
        Name = $_
        Color = $color
    }
}

# Updates the title of the given form and the text of the version TextBox
function UpdateFormTitle {
    param (
        [System.Windows.Forms.Form]$form,
        [System.Windows.Forms.TextBox]$textBoxVersion,  # Pass the TextBox object as a parameter
        [string]$newVersion
    )

    $form.Text = "PowerShell Script Manager v$newVersion"
    $textBoxVersion.Text = $newVersion
}

# Retrieves the script path from the 'config.json' file
function Get-ConfigScriptPath {
    $configPath = Join-Path $PSScriptRoot "config.json"
    if (Test-Path -Path $configPath) {
        $config = Get-Content -Path $configPath | ConvertFrom-Json
        return $config.ScriptPath
    } else {
        return $null
    }
}

# Saves the configuration object to the 'config.json' file
function Save-Config {
    param ($path)
    $config = @{
        scriptPath = $path
    }
    $configFile = Join-Path $PSScriptRoot "config.json"
    $config | ConvertTo-Json | Set-Content -Path $configFile
}

# Sets the script path in the 'config.json' file
function Set-ConfigScriptPath {
    param (
        [string]$path
    )
    $config = @{
        ScriptPath = $path
    } | ConvertTo-Json

    $configPath = Join-Path $PSScriptRoot "config.json"
    $config | Set-Content -Path $configPath
}

# Initializes the environment by creating a backup folder and copying the sounds folder
function Initialize-Environment($scriptPath) {
    # Create backup folder if it doesn't exist
    $backupFolderPath = Join-Path $scriptPath "backup"
    if (-not (Test-Path -Path $backupFolderPath)) {
        New-Item -Path $backupFolderPath -ItemType Directory
    }

    # Copy sounds folder if it doesn't exist in the script path
    $sourceSoundsPath = Join-Path $PSScriptRoot "sounds"
    $destinationSoundsPath = Join-Path $scriptPath "sounds"
    if (-not (Test-Path -Path $destinationSoundsPath) -and (Test-Path -Path $sourceSoundsPath)) {
        Copy-Item -Path $sourceSoundsPath -Destination $destinationSoundsPath -Recurse
    }
}

# Opens a folder dialog for the user to select a new scripts directory and initializes the environment
function Set-ScriptSPath {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select a new scripts directory"
    $dialog.RootFolder = [Environment+SpecialFolder]::MyComputer
    $dialog.SelectedPath = $textBoxScriptsPath.Text # Use the current path as the default

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxScriptsPath.Text = $dialog.SelectedPath
        Set-ConfigScriptPath $textBoxScriptsPath.Text

        # Copy necessary files
        Copy-ScriptsAndConfig $textBoxScriptsPath.Text

        # Call the initialization function
        Initialize-Environment $textBoxScriptsPath.Text

         # Display path information
        $pathInfo = Get-ScriptPaths -selectedScript $global:selectedScript -baseScriptPath $global:BaseScriptPath

        Write-Host "Selected Script Paths:"
        Write-Host "Script Path: $($pathInfo.ScriptPath)" -ForegroundColor Yellow
        Write-Host "JSON Path: $($pathInfo.JsonPath)" -ForegroundColor Yellow
    }
}

# Copies the necessary scripts and configuration files to the destination path
function Copy-ScriptsAndConfig {
    param (
        [string]$destinationPath
    )
    # Define the source paths for the files you want to copy
    $psmPath = Join-Path $PSScriptRoot "PSM.ps1"
    $jsonPath = Join-Path $PSScriptRoot "PSM.json"
    $configPath = Join-Path $PSScriptRoot "config.json"

    # Function to copy the file if it does not exist at the destination
    function Copy-IfNotExists {
        param (
            [string]$source,
            [string]$destination
        )
        $destinationFile = Join-Path $destination (Split-Path $source -Leaf)
        if (-Not (Test-Path $destinationFile)) {
            Copy-Item -Path $source -Destination $destination
        } else {
            Write-Host "File $($source) already exists at the destination."
        }
    }

    # Copy each file to the destination path, if it doesn't exist there already
    Copy-IfNotExists -source $psmPath -destination $destinationPath
    Copy-IfNotExists -source $jsonPath -destination $destinationPath
    Copy-IfNotExists -source $configPath -destination $destinationPath
}

# Function to generate a new copy name with a dynamic increment
function Get-NewCopyName {
    param (
        [string]$baseName,
        [string]$folderPath
    )

    $counter = 1
    $copyName = $baseName
    while (Test-Path -Path (Join-Path -Path $folderPath -ChildPath "$copyName.ps1")) {
        $copyName = "${baseName}_Copy$counter"
        $counter++
    }

    return $copyName
}

function Refresh-ScriptsInListBox {
    # Clear the list box
    $listBoxScripts.Items.Clear()

    try {
        $scriptFiles = Get-ChildItem -Path $textBoxScriptsPath.Text -Filter "*.ps1" | Select-Object -ExpandProperty Name | Sort-Object
        foreach ($scriptFile in $scriptFiles) {
            $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($scriptFile).ToUpper()
            $listBoxScripts.Items.Add($scriptName)
        }
    } catch {
        Write-Error $_
        [System.Windows.Forms.MessageBox]::Show("An error occurred while loading the scripts: `n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to Get Scripts File Path by joining the user-defined scripts path with the provided script name and appending the .ps1 extension.
function Get-ScriptFilePath {
    param (
        [string]$scriptName
    )
    return Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$scriptName.ps1"
}

# Function to Get JSON File Path by joining the user-defined scripts path with the provided script name and appending the .json extension.
function Get-JsonFilePath {
    param (
        [string]$scriptName
    )
    return Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$scriptName.json"
}

# Declare global variables
$global:preloadedMediaPlayer = $null
$global:textBoxBackupCount = $null
$global:BaseScriptPath = Get-ConfigScriptPath
$global:selectedScript = $null
$global:DefaultPath = [System.Environment]::GetFolderPath('MyDocuments')
$global:BackupFolderPath = Join-Path -Path $global:BaseScriptPath -ChildPath "backup"
# Constructs the full path for the specified sound file
# Define a function to get paths and related variables
function Get-ScriptPaths {
    param (
        [string]$selectedScript,
        [string]$baseScriptPath
    )
   if ([string]::IsNullOrEmpty($baseScriptPath) -or [string]::IsNullOrEmpty($selectedScript)) {
        Write-Host "Error: baseScriptPath or selectedScript is null or empty."
        return # Return early to prevent further execution
    }
    # Debugging lines to check the input parameters
    Write-Host "Debug: Selected Script: $selectedScript" -ForegroundColor Green
    Write-Host "Debug: Base Script Path: $baseScriptPath" -ForegroundColor Green

    # Construct the full paths
    $scriptPath = Join-Path -Path $baseScriptPath -ChildPath "$selectedScript.ps1"
    $jsonPath = Join-Path -Path $baseScriptPath -ChildPath "$selectedScript.json"
    
    # Get the script name without extension
    $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($scriptPath)
    $jsonName = [System.IO.Path]::GetFileNameWithoutExtension($jsonPath)
    
    # Create an object to store the variables
    $pathInfo = @{
        BasePath = $baseScriptPath
        ScriptPath = $scriptPath
        JsonPath = $jsonPath
        ScriptName = $scriptName
        JsonName = $jsonName
    }

    # Colorful output for better visibility
    Write-Host ("`n" + "-" * 50) -ForegroundColor Cyan
    Write-Host "Path Information:" -ForegroundColor Cyan
    Write-Host ("-" * 50) -ForegroundColor Cyan
    Write-Host "Base Path: $($pathInfo.BasePath)" -ForegroundColor Yellow
    Write-Host "Script Path: $($pathInfo.ScriptPath)" -ForegroundColor Yellow
    Write-Host "JSON Path: $($pathInfo.JsonPath)" -ForegroundColor Yellow
    Write-Host "Script Name: $($pathInfo.ScriptName)" -ForegroundColor Yellow
    Write-Host "JSON Name: $($pathInfo.JsonName)" -ForegroundColor Yellow
    Write-Host ("-" * 50) -ForegroundColor Cyan

    return $pathInfo
}


function Get-SoundFilePath {
    param (
        [string]$soundFileName
    )

    $soundsFolderPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "sounds"
    return Join-Path -Path $soundsFolderPath -ChildPath $soundFileName
}

# Function to preload a cool sound
function PreloadCoolSound {
    param (
        [string]$SoundFileName
    )

    # Get the full path of the sound file using the path helper function
    $SoundFilePath = Get-SoundFilePath -soundFileName $SoundFileName

    if (Test-Path -Path $SoundFilePath) {
        try {
            # Create a new media player
            $global:preloadedMediaPlayer = New-Object System.Windows.Media.MediaPlayer
            # Register the event handler for media being fully loaded
            $global:preloadedMediaPlayer.Add_MediaOpened({
                # When the media is fully loaded, it will automatically start to play
                $global:preloadedMediaPlayer.Play()
            })
            # Open the media file (the event will be triggered when loading is done)
            $global:preloadedMediaPlayer.Open([System.Uri]$SoundFilePath)
        } catch {
            Write-Host "Error occurred while preloading the sound: $_"
        }
    } else {
        Write-Host "Cool sound file not found: $SoundFilePath"
    }
}

# Plays the preloaded cool sound file
function PlayPreloadedCoolSound {
    if ($global:preloadedMediaPlayer -ne $null) {
        try {
            # Play the preloaded sound file
            $global:preloadedMediaPlayer.Play()
        } catch {
            Write-Host "Error occurred while playing the preloaded sound: $_"
        }
    } else {
        Write-Host "Cool sound is not preloaded or was not loaded successfully."
    }
}

# Function to play cool sounds
function PlayCoolSound {
    param (
        [string]$SoundFileName
    )

    # Get the full path of the sound file using the path helper function
    $SoundFilePath = Get-SoundFilePath -soundFileName $SoundFileName

    if (Test-Path -Path $SoundFilePath) {
        try {
            # Create a new media player and play the sound file
            $mediaPlayer = New-Object System.Windows.Media.MediaPlayer
            $mediaPlayer.Open([System.Uri]$SoundFilePath)
            $mediaPlayer.Play()
        } catch {
            Write-Host "Error occurred while playing the sound: $_"
        }
    } else {
        Write-Host "Cool sound file not found: $SoundFilePath"
    }
}

<## Function to toggle between Tagging Mode and Search Mode
function ToggleMode($mode) {
    if ($mode -eq "Tagging") {
        # Switch to Tagging Mode
        $currentMode = 1
        $buttonTaggingMode.Visible = $false
        $buttonSearchMode.Visible = $true
        $labelModeIndicator.Text = "Tagging Mode"
        $labelModeIndicator.ForeColor = [System.Drawing.Color]::Green # Green color for Tagging Mode
        
        # Hide the search ListBox and show the original ListBox
        #$listBoxTags.Visible = $true
        #$listBoxSearchTags.Visible = $false
    } else {
        # Switch to Search Mode
        $currentMode = 2
        $buttonTaggingMode.Visible = $true
        $buttonSearchMode.Visible = $false
        $labelModeIndicator.Text = "Search Mode"
        $labelModeIndicator.ForeColor = [System.Drawing.Color]::Blue

        # Show the search ListBox and hide the original ListBox
        #$listBoxTagsFilter.Visible = $false
        #$listBoxSearchTags.Visible = $true

       # foreach ($tag in $tags) {
            #$listBoxTagsSearch.Items.Add($tag.Name)
        #}
    }
}
#>
# Function to display a Random Quote in the RichTextBox at the bottom of the form
function DisplayRandomQuote {
    $randomQuote = Get-Random -InputObject $inspirationalQuotes
    $richTextBoxOutput.Clear()
    $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Segoe Script", 20, [System.Drawing.FontStyle]::Regular)
    $richTextBoxOutput.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Center  # Center the text
    $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::White
    $richTextBoxOutput.ForeColor = [System.Drawing.Color]::CornSilk
    $richTextBoxOutput.AppendText($randomQuote)
    #foreach ($char in $randomQuote.ToCharArray()) {
        #$richTextBoxOutput.AppendText($char)
        #$richTextBoxOutput.Refresh() # Redraw the control to make sure the update is visible
        #Start-Sleep -Milliseconds 5 # Delay for 40 ms between characters
   # }
}

# Update the Backup Counter
function Update-BackupCounter {
    param (
        [string]$ScriptName,
        [string]$BackupFolderPath,
        [System.Windows.Forms.Label]$LabelBackupCountValue
    )
    #Clear-Host
    $backupFiles = Get-ChildItem -Path $BackupFolderPath -Filter "${ScriptName}_backup*.ps1" -File
    $backupCount = $backupFiles.Count

    if ($LabelBackupCountValue -ne $null) {
        $action = [Action[string]] {
            param($count)
            $LabelBackupCountValue.Text = $count
            if ($count -eq '0') {
                $LabelBackupCountValue.ForeColor = [System.Drawing.Color]::Red
            } else {
                $LabelBackupCountValue.ForeColor = [System.Drawing.Color]::Black
            }
        }
        $LabelBackupCountValue.Invoke($action, $backupCount.ToString())
    }
} 

# Update the Backup Counter
function Get-BackupScriptCounter {
    param (
        [string]$ScriptPath,
        [string]$BackupFolderPath,
        [int]$MaxBackupNumber
    )

    $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $backupFiles = Get-ChildItem -Path $BackupFolderPath -Filter "$scriptName*_backup*.*" -File
    $backupCount = $backupFiles.Count

    # Clean up any excess backup files if the count exceeds the maximum
    if ($backupCount -gt $MaxBackupNumber) {
        $excessBackups = $backupFiles | Sort-Object -Property LastWriteTime -Descending | Select-Object -Skip $MaxBackupNumber
        $excessBackups | Remove-Item -Force
    }

    [PSCustomObject]@{
        BackupFiles = $backupCount
    }
}

# Helper function for creating DefaultScript if it doesn't exist
function CreateDefaultScriptAndJson {
    param (
        [string]$scriptPath
    )
    
    Write-Host "Creating default script and JSON in path: $scriptPath" -ForegroundColor Green
    $defaultScriptPath = Join-Path -Path $scriptPath -ChildPath "DefaultScript.ps1"
    $defaultJsonPath = Join-Path -Path $scriptPath -ChildPath "DefaultScript.json"

    enum TaskStatus {
        NotStarted
        InProgress
        Completed
    }

    try {
        if (-not (Test-Path -Path $defaultScriptPath)) {
            Set-Content -Path $defaultScriptPath -Value "# Default script"
        }
        if (-not (Test-Path -Path $defaultJsonPath)) {
            $defaultJsonContent = @{
                "Name" = "DefaultScript"
                "ScriptPath" = $defaultScriptPath
                "ScriptJson" = $defaultJsonPath
                "Author" = $env:username.ToUpper()
                "Description" = "Default script"
                "Version" = "0.0"
                "Tags" = @("🆕 New Script", "📜 PowerShell", "👩‍💻 Development", "🗑️ Junk")
                "ModifiedBy" = $env:username.ToUpper()
                "Keywords" = @("POWERSHELL")
                "Modifications" = @(
                    @{
                        "Date" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
                        "ModifiedBy" = $env:username.ToUpper()
                        "Modification" = "Default script created."
                        "Version" = "0.1"
                    }
                )
                "ToDoList" = @(
                    @{
                        "Title" = "Review Script"
                        "Task" = "Review Script for errors and Update!"
                        "DateCompleted" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt") 
                        "Status" = ([TaskStatus]::NotStarted).ToString() # Convert the enum value to a string
                        "CompletedBy" = ""
                    }
                ) # You can add more to-do list items here
                "Requirements" = @()
            }

            $defaultJsonContent | ConvertTo-Json -Depth 2 | Out-File -FilePath $defaultJsonPath -Encoding UTF8
        }
    } catch {
        Write-Error $_
    }
}

# Helper function for updating JSON content
function UpdateJsonContent {
    param (
        $jsonContent,
        $selectedScript,
        $scriptPath,
        $jsonPath
    )

    $requiredProperties = @(
        "Name", "ScriptPath", "ScriptJson", "Author", "Description",
        "Version", "Tags", "ModifiedBy", "Keywords", "Modifications", "ToDoList", "Requirements"  # Added "Requirements"
    )

    # Default property values
    $defaultValues = @{
        'Name' = $selectedScript.ToUpper()
        'Author' = $env:USERNAME.ToUpper()
        'Description' = "New Script $selectedScript"
        'Version' = "0.1"
        'Tags' = @("📜 PowerShell")
        'ModifiedBy' = $env:USERNAME.ToUpper()
        'Keywords' = @("POWERSHELL")
        'Modifications' = @(
            @{
                "Date" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
                "Version" = "0.1"
                "ModifiedBy" = $env:USERNAME.ToUpper()
                "Modification" = "Created $($selectedScript) a JSON Buddy"
            }
        )
        'ToDoList' = @(
            @{
                "Title" = "Review Script"
                "Task" = "Review Script for errors and Update!"
                "DateCompleted" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
                "Status" = "NotStarted"  # No enum required here
                "CompletedBy" = ""
            }
        )
        'Requirements' = @()  # Default empty array for "Requirements"
    }

    # Validate and populate missing properties
    foreach ($property in $requiredProperties) {
        if ($property -notin $jsonContent.PSObject.Properties.Name) {
            $defaultValue = $defaultValues[$property]
            Add-Member -InputObject $jsonContent -NotePropertyName $property -NotePropertyValue $defaultValue
        }
    }

    # Update existing properties
    $jsonContent.Name = $selectedScript.ToUpper()
    $jsonContent.ScriptPath = $scriptPath
    $jsonContent.ScriptJson = $jsonPath
    $jsonContent.Author = if ($jsonContent.Author -ne $null) { $jsonContent.Author } else { $env:USERNAME.ToUpper() }
    $jsonContent.Description = if ($jsonContent.Description -ne $null) { $jsonContent.Description } else { "New Script $selectedScript" }
    
    # Update the Version property based on modifications
    $highestVersion, $newJsonContent = Get-HighestVersionFromModifications -modifications $jsonContent.Modifications -jsonPath $jsonPath
    
    if ($null -eq $highestVersion) {
        $highestVersion = "0.1"
    }
    $jsonContent.Version = $highestVersion.ToString()

    # Update Tags if null or empty
    if ($null -eq $jsonContent.Tags) {
        $jsonContent.Tags = @("📜 PowerShell")
    } elseif ($jsonContent.Tags.Count -eq 0) {
        $jsonContent.Tags = @("📜 PowerShell")
    }

    # Update ModifiedBy if null
    if ($null -eq $jsonContent.ModifiedBy) {
        $jsonContent.ModifiedBy = $env:USERNAME.ToUpper()
    }

    # Update Modifications if null or empty
    if (-not $jsonContent.PSObject.Properties['Modifications'] -or $jsonContent.Modifications.Length -eq 0) {
        $defaultModification = @{
            "Date" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
            "ModifiedBy" = $env:USERNAME.ToUpper()
            "Modification" = "Default script created."
            "Version" = "0.1"
        }
        $jsonContent.Modifications = @($defaultModification)
    }

    # Update Keywords if null or not an array, or if it does not contain "POWERSHELL"
    if ($null -eq $jsonContent.Keywords -or $jsonContent.Keywords.GetType().IsArray -eq $false -or "POWERSHELL" -notin $jsonContent.Keywords) {
        $jsonContent.Keywords = @("POWERSHELL")
    }

    # Delete unnecessary properties in the JSON
    $jsonContent.PSObject.Properties | Where-Object { $_.Name -notin $requiredProperties } | ForEach-Object {
        $jsonContent.PSObject.Properties.Remove($_.Name)
    }

    # Save the updated JSON content back to the file
    $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8

    return $jsonContent
}

# function is responsible for loading, modifying, and displaying information related to a selected script
function PopulateFieldsAndTags {
    try {
        # Define an enum for TaskStatus
        enum TaskStatus {
            NotStarted
            InProgress
            Completed
        }

        # Clear-Host
        $selectedScript = $listBoxScripts.SelectedItem
        $soundFilePath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "sounds\selectscript.mp3"
        PreloadCoolSound -SoundFileName "selectscript.mp3"

        if ($selectedScript) {
            $pathInfo = Get-ScriptPaths -selectedScript $selectedScript -baseScriptPath $textBoxScriptsPath.Text
            $scriptPath = $pathInfo.ScriptPath
            $jsonPath = $pathInfo.JsonPath

            # Load JSON content
            #$jsonContent = Get-Content -Path $jsonPath | ConvertFrom-Json
            #Write-Host "Debug: Initial JSON Content: $($jsonContent | ConvertTo-Json)" # Debug statement

            # Load or create JSON content for non-default scripts
            if (Test-Path -Path $jsonPath) {
                $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
            } else {
                $jsonContent = @{
                    "Name" = $selectedScript
                    "ScriptPath" = $scriptPath
                    "ScriptJson" = $jsonPath
                    "Author" = $env:username
                    "Description" = "New Script $selectedScript"
                    "Version" = "0.1"  # Set the default version as 0.1
                    "Tags" = @("📜 PowerShell")
                    "ModifiedBy" = $env:username.ToUpper()
                    "Keywords" = @()
                    "Modifications" = @(
                        @{
                            "Date" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
                            "Version" = "0.1"
                            "ModifiedBy" = $env:username.ToUpper()
                            "Modification" = "Created $($selectedScript) a JSON Buddy"
                        }
                    )
                    "ToDoList" = @(
                        @{
                            "Title" = "Review Script"
                            "Task" = "Review Script for errors and Update!"
                            "DateCompleted" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
                            "Status" = ([TaskStatus]::NotStarted).ToString() # Convert the enum value to a string
                            "CompletedBy" = ""
                        }
                    )
                }
                $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $jsonPath -Encoding UTF8
                Write-Host "JSON file $jsonPath created."
            }

            

            # Remove null values from the Keywords array
            $jsonContent.Keywords = $jsonContent.Keywords | Where-Object { $_ -ne $null }

            # Populate fields with the JSON content
            $fileInfo = Get-Item -Path $scriptPath
            $textBoxName.Text = $jsonContent.Name
            $textBoxDescription.Text = $jsonContent.Description
            $textBoxAuthor.Text = $jsonContent.Author
            $textBoxVersion.Text = $jsonContent.Version
            $textBoxDateCreated.Text = $fileInfo.CreationTime.ToString("MMMM dd, yyyy 'at' hh:mm tt")
            $textBoxDateModified.Text = $fileInfo.LastWriteTime.ToString("MMMM dd, yyyy 'at' hh:mm tt")
            $textBoxModifiedBy.Text = $jsonContent.ModifiedBy.ToUpper()
            $textBoxJsonFile.Text = $jsonPath

            # Count the number of backup files and update the backup counter
            Update-BackupCounter -ScriptName $jsonContent.Name -BackupFolderPath $global:BackupFolderPath -LabelBackupCountValue $labelBackupCountValue


             #Clear and populate $listBoxScriptTags with all available tags
            $listBoxKeywords.BeginUpdate()
            $listBoxKeywords.Items.Clear()

            if ($jsonContent.Keywords -ne $null) {
                Write-Host "Debug: Keywords before updating ListBox: $($jsonContent.Keywords -join ', ')"
                $jsonContent.Keywords | ForEach-Object {
                    $uppercaseKeyword = $_.ToUpper()
                    Write-Host "Debug: Adding keyword: $uppercaseKeyword"
                    $listBoxKeywords.Items.Add($uppercaseKeyword)
                }
            } else {
                Write-Host "Debug: No Keywords property found in JSON content"
            }

            $listBoxKeywords.EndUpdate()

            Write-Host "Debug: ListBox Keywords after updating:"
            $listBoxKeywords.Items | ForEach-Object {
                Write-Host "Debug: - $_"
            }

            # Clear and populate $listBoxScriptTags with all available tags
            $listBoxScriptTags.Items.Clear()
            $tags | ForEach-Object {
                $itemText = "$($_.Icon) $($_.Name)"
                $listBoxScriptTags.Items.Add($itemText)
            }

            # Check the tags that the script has
            $selectedTags = $jsonContent.Tags
            $selectedIndices = @()
            foreach ($selectedTag in $selectedTags) {
                $tagIndex = $listBoxScriptTags.Items.IndexOf($selectedTag)
                if ($tagIndex -ge 0) {
                    $selectedIndices += $tagIndex
                }
            }

            $listBoxScriptTags.ClearSelected()
            $selectedIndices | ForEach-Object {
                $listBoxScriptTags.SetSelected($_, $true)
            }
        } else {
            Write-Host "No script selected in the ListBox."
        }

        # Load requirements
        $requirementsList = $jsonContent.Requirements
        if ($null -eq $requirementsList) {
            # If the Requirements field is missing, initialize it as an empty array
            $requirementsList = @()
        }

        # Bind requirements to the DataGridView
        $requirementsGridView.DataSource = $requirementsList


        # Display a random inspirational quote
        DisplayRandomQuote
        # Refresh the form
        $form.Refresh()
    } catch {
        Write-Host "An error occurred: $_"
    }
}

# Function to populate the ListBox with tags and highlight matching tags
function PopulateTags {
    $listBoxScriptTags.Items.Clear()  # Clear the ListBox items before populating

    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        $jsonPath = ($scriptPath -replace "\.ps1$", ".json")

        if (Test-Path -Path $jsonPath) {
            $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
            $tags = $jsonContent.Tags
            if ($tags) {
                $tags.ForEach({ $listBoxScriptTags.Items.Add($_) })

                # Check the tags that the script has
                foreach ($tag in $tags) {
                    $index = $listBoxScriptTags.FindStringExact($tag)
                    if ($index -ge 0) {
                        $listBoxScriptTags.SetSelected($index, $true)
                    }
                }
            }
        }
    }
}

# Function to clear the form fields
function ClearForm {
    $textBoxName.Text = ""
    $textBoxDescription.Text = ""
    $textBoxVersion.Text = ""
    $textBoxAuthor.Text = ""
    $textBoxDateCreated.Text = ""
    $textBoxDateModified.Text = ""
    $textBoxModifiedBy.Text = ""
    $textBoxJsonFile.Text = ""
    $richTextBoxOutput.Text = ""
    $listBoxScriptTags.ClearSelected()
    $listBoxKeywords.ClearSelected()
}

# Define the function to populate the ListBox with script names
function PopulateScriptsListBox {
    $listBoxScripts.Items.Clear()  # Clear the ListBox items before populating
    $scriptFiles = Get-ChildItem -Path $textBoxScriptsPath.Text -Filter "*.ps1" -File
    $scriptFiles.ForEach({ $listBoxScripts.Items.Add($_.Name) })
}

# Function to show the tag selection form and filter scripts based on selected tags
function ShowTagSelectionForm {
    $tagForm = New-Object System.Windows.Forms.Form
    $tagForm.Text = "Select Tag to Filter Scripts"
    $tagForm.Size = New-Object System.Drawing.Size(265, 995)  # Width is half, height is the same
    $tagForm.FormBorderStyle = "FixedSingle"
    $tagForm.StartPosition = "CenterScreen"
    $tagForm.BackColor = [System.Drawing.Color]::DarkSlateGray  # Set the background color to light pink

    # ListBox for tag selection
    $listBoxTags = New-Object System.Windows.Forms.ListBox
    $listBoxTags.Location = New-Object System.Drawing.Point(20, 20)
    $listBoxTags.Size = New-Object System.Drawing.Size(220, 925)  # Width is half, height is the same
    $listBoxTags.Font = New-Object System.Drawing.Font("Arial", 14)  # Change the font family and size here

    # Populate the ListBox with tags (with emojis)
    $tags | ForEach-Object {
        $listBoxTags.Items.Add($_.Icon + " " + $_.Name)
    }

    # Allow multiple selections in the ListBox
    $listBoxTags.SelectionMode = "MultiSimple"

    # Add the ListBox to the form
    $tagForm.Controls.Add($listBoxTags)

    # Event handler for the SelectedIndexChanged event of the $listBoxTags ListBox
    $listBoxTags.Add_SelectedIndexChanged({
    $selectedTags = $listBoxTags.SelectedItems
    PlayCoolSound -SoundFileName "selectscript.mp3"
    # Update the main form's list box with the filtered scripts
    FilterScriptsAndUpdateMainForm $selectedTags
    })

    # Show the tag selection form
    $tagForm.ShowDialog()
}

# Function to filter scripts based on selected tags and update the main form's script list box
function FilterScriptsAndUpdateMainForm {
    param (
        [string[]]$Tags
    )

    #Write-Host "FilterScriptsAndUpdateMainForm: Start"

    # Get the list of script files in the scripts path
    $scriptFiles = Get-ChildItem -Path $textBoxScriptsPath.Text -Filter "*.ps1" -File

    # Initialize the list of matching script names
    $matchingScripts = @()

    # Iterate through each script file
    foreach ($scriptFile in $scriptFiles) {
        $scriptPathWithoutExtension = $scriptFile.FullName -replace '\.ps1$', ''
        $jsonPath = $scriptPathWithoutExtension + ".json"

        # Check if the corresponding JSON file exists
        if (Test-Path -Path $jsonPath) {
            #Write-Host "FilterScriptsAndUpdateMainForm: JSON file found - $jsonPath"

            # Read the JSON content
            $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

            # Check if the script's JSON content contains all the selected tags
            if ($Tags -eq $null -or ($Tags -is [Array] -and $Tags.Count -eq 0)) {
                # No tags selected, include all scripts
                #Write-Host "FilterScriptsAndUpdateMainForm: All tags selected for $scriptFile"
                $matchingScripts += $scriptFile
            } elseif ($Tags -is [Array]) {
                # Check if all selected tags are present in the JSON content
                $containsAllTags = $true
                foreach ($tag in $Tags) {
                    if (-not $jsonContent.Tags.Contains($tag)) {
                        #Write-Host "FilterScriptsAndUpdateMainForm: Tag '$tag' not found in $jsonPath"
                        $containsAllTags = $false
                        break
                    }
                }

                if ($containsAllTags) {
                    #Write-Host "FilterScriptsAndUpdateMainForm: All tags found for $scriptFile"
                    $matchingScripts += $scriptFile
                }
            }
        } else {
            Write-Host "FilterScriptsAndUpdateMainForm: JSON file not found - $jsonPath"
        }
    }

    # Update the main form's list box
    $listBoxScripts.Items.Clear()
    $matchingScripts | ForEach-Object {
        $scriptName = $_.Name.ToUpper()
        $scriptNameWithoutExtension = $scriptName.Replace('.PS1', '')
        $listBoxScripts.Items.Add($scriptNameWithoutExtension)
    }

    # Update the output message without RichTextBox
    $resultMessage = "Searching for all scripts in $($textBoxScriptsPath.Text) for tags matching: $($Tags -join ', ')`n"
    if ($matchingScripts.Count -gt 0) {
        $resultMessage += "Found $($matchingScripts.Count) Matches"
    } else {
        $resultMessage += "Found 0 Matches"
    }

    #Write-Host $resultMessage

    #Write-Host "FilterScriptsAndUpdateMainForm: End"
}

# Functions for the To Do List Section
function ShowToDoList {
    Write-Host "Debug: Inside ShowToDoList function"

    $selectedScript = $listBoxScripts.SelectedItem
    if (-not $selectedScript) {
        Write-Host "No script selected"
        return
    }

    $scriptPath = Join-Path -Path "C:\SCRIPTS" -ChildPath "$selectedScript.ps1"
    $jsonPath = $scriptPath -replace "\.ps1$", ".json"

    Write-Host "Debug: JSON path created - $jsonPath"

    if (Test-Path -Path $jsonPath) {
        Write-Host "Debug: JSON path exists - $jsonPath"
        $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

        Write-Host "Debug: JsonContent loaded - $($jsonContent | ConvertTo-Json)"

        ShowToDoListForm -JsonPath $jsonPath -ScriptName $selectedScript

    } else {
        Write-Host "Debug: JSON path does not exist - $jsonPath"

        $emptyJsonContent = @{
            ToDoList = @()
        } | ConvertTo-Json -Depth 4
        $emptyJsonContent | Out-File -FilePath $jsonPath -Encoding UTF8

        ShowToDoListForm -JsonPath $jsonPath -ScriptName $selectedScript
    }
}

# Define the enum for TaskStatus
enum TaskStatus {
    NotStarted
    InProgress
    Completed
}

# Show ToDoList Form DataGrid setup
function ShowToDoListForm {
    param (
        [string]$JsonPath,
        [string]$ScriptName
    )
    
    Write-Host "Debug: Inside ShowToDoListForm function"
    $JsonContent = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json

    Write-Host "Debug: Selected Script is: $ScriptName"

    # Create the To-Do List form
    $toDoListForm = New-Object System.Windows.Forms.Form
    $toDoListForm.Text = "To-Do List"
    $toDoListForm.Size = New-Object System.Drawing.Size(1280, 600)
    $toDoListForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Set the background color to match your main form's color
    $toDoListForm.BackColor = [System.Drawing.Color]::CadetBlue


    # Create and place labels
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "TO-DO-LIST FORM FOR SCRIPT --->"
    $label.Font = New-Object System.Drawing.Font("Segoe Print", 18, [System.Drawing.FontStyle]::Bold)
    $label.Location = New-Object System.Drawing.Point(20, 10) # Increase X = Right, y = Down
    $label.Size = New-Object System.Drawing.Point(500, 40)    # Increase W = Wider, H = Taller.
    $label.ForeColor = [System.Drawing.Color]::White
    $toDoListForm.Controls.Add($label)

    $labelScriptName = New-Object System.Windows.Forms.Label
    $labelScriptName.Text = "$selectedScript"
    $labelScriptName.Font = New-Object System.Drawing.Font("Segoe Print", 18) # Optional: larger font
    $labelScriptName.Location = New-Object System.Drawing.Point(520, 10) # Increase X = Right, y = Down
    $labelScriptName.Size = New-Object System.Drawing.Size(600, 40) # Increase W = Wider, H = Taller.
    $labelScriptName.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#F4C67A") # Sandy Beach Color
    $labelScriptName.BackColor = [System.Drawing.Color]::CadetBlue
    $toDoListForm.Controls.Add($labelScriptName) # Add the label to the TO-DO-LIST form


    # Assume $selectedScript variable is being updated somewhere in your code
    #$selectedScript = "Your Script Name Here" # Or from the JSON's Name property

    # Update the label's text
    


    $dataGridViewToDoList = New-Object System.Windows.Forms.DataGridView
    $dataGridViewToDoList.Location = New-Object System.Drawing.Point(20, 50)
    $dataGridViewToDoList.Size = New-Object System.Drawing.Size(1220, 300)
    $dataGridViewToDoList.AllowUserToAddRows = $true
    $dataGridViewToDoList.MultiSelect = $false
    $dataGridViewToDoList.BackgroundColor = [System.Drawing.Color]::White

    # Create columns for the DataGridView
    $completeColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $completeColumn.HeaderText = "Mark Completed" # Updated text
    $completeColumn.Text = "Complete Task"      # Button text
    $completeColumn.UseColumnTextForButtonValue = $true
    $completeColumn.DefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#4CAF50") # Cool green color


    $TitleColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $TitleColumn.HeaderText = "Title"
    $TitleColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $TitleColumn.Width = 140  # Set the width you prefer

    $TaskColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $TaskColumn.HeaderText = "Task"
    $TaskColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $TaskColumn.Width = 580  # Set the width you prefer
    $TaskColumn.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True

    $StatusColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $StatusColumn.HeaderText = "Status"
    $StatusColumn.DataSource = [Enum]::GetNames([TaskStatus])
    $StatusColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $StatusColumn.Width = 90

    $completedByColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $completedByColumn.HeaderText = "Completed By"
    $completedByColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $completedByColumn.Width = 100

    $datecompletedColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $datecompletedColumn.HeaderText = "Date Completed"
    $datecompletedColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $datecompletedColumn.Width = 160
    

    # Add columns to DataGridView
    $dataGridViewToDoList.Columns.Add($completeColumn)
    $dataGridViewToDoList.Columns.Add($TitleColumn)
    $dataGridViewToDoList.Columns.Add($TaskColumn)
    $dataGridViewToDoList.Columns.Add($StatusColumn)
    $dataGridViewToDoList.Columns.Add($completedByColumn)
    $dataGridViewToDoList.Columns.Add($datecompletedColumn)

    # Attach the DataGridView to the form
    $toDoListForm.Controls.Add($dataGridViewToDoList)

    # Read existing tasks from JSON and add them to the DataGridView
    foreach ($task in $JsonContent.ToDoList) {
        $rowIndex = $dataGridViewToDoList.Rows.Add()
        $row = $dataGridViewToDoList.Rows[$rowIndex]
        $row.Cells[$TitleColumn.Index].Value = $task.Title
        $row.Cells[$TaskColumn.Index].Value = $task.Task
        $row.Cells[$StatusColumn.Index].Value = $task.Status
        $row.Cells[$completedByColumn.Index].Value = $task.CompletedBy
        $row.Cells[$datecompletedColumn.Index].Value = $task.DateCompleted
    }


    # Create "Add New Task" Button
    $todoAddTaskButton = New-Object System.Windows.Forms.Button
    $todoAddTaskButton.Text = "Add New Task"
    $todoAddTaskButton.Location = New-Object System.Drawing.Point(20, 350)  # Increase X = Right, y = Down
    $todoAddTaskButton.Size = New-Object System.Drawing.Size(100, 30) # Increase W = Wider, H = Taller.
    $todoAddTaskButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F4C67A") # Sandy Beach Color
    $todoAddTaskButton.Add_Click({
    # Add a new row to the DataGridView with default or empty values
    $rowIndex = $dataGridViewToDoList.Rows.Add()
    $row = $dataGridViewToDoList.Rows[$rowIndex]
    $row.Cells[$TitleColumn.Index].Value = "" # You can set a default title if you want
    $row.Cells[$TaskColumn.Index].Value = ""
    $row.Cells[$StatusColumn.Index].Value = "NotStarted" # Setting default status to NotStarted
    $row.Cells[$completedByColumn.Index].Value = ""
    $row.Cells[$datecompletedColumn.Index].Value = ""
    
    # Optional: Update the JSON file with the new task
    $newTask = New-Object PSObject -property @{
        "Title" = ""
        "Task" = ""
        "Status" = ([TaskStatus]::NotStarted).ToString()
        "CompletedBy" = ""
        "DateCompleted" = ""
    }
    $JsonContent.ToDoList += $newTask
    $JsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8
})

    $toDoListForm.Controls.Add($todoAddTaskButton)

    # Create "Open JSON" Button
    $todoOpenJsonButton = New-Object System.Windows.Forms.Button
    $todoOpenJsonButton.Text = "Open JSON File"
    $todoOpenJsonButton.Location = New-Object System.Drawing.Point(120, 350) # Increase X = Right, y = Down
    $todoOpenJsonButton.Size = New-Object System.Drawing.Size(100, 30) # Increase W = Wider, H = Taller.
    $todoOpenJsonButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F4C67A") # Sandy Beach Color
    $todoOpenJsonButton.Add_Click({
        $scriptPath = Join-Path -Path "C:\SCRIPTS" -ChildPath "$selectedScript.ps1"
        $jsonPath = $scriptPath -replace "\.ps1$", ".json"

        if (Test-Path -Path $jsonPath) {
            Invoke-Item $jsonPath
        } else {
            Write-Host "JSON file not found for $selectedScript"
        }
    })
    $toDoListForm.Controls.Add($todoOpenJsonButton)

  # Custom event handler for the "Complete" button click
$dataGridViewToDoList_CellClick = {
    param (
        $sender,
        [System.Windows.Forms.DataGridViewCellEventArgs]$e
    )

    # Check if the clicked cell is in the "Complete Task" column
    if ($e.ColumnIndex -eq $completeColumn.Index -and $e.RowIndex -ge 0) {
        try {
            # Update the task properties in the JSON content
            $JsonContent.ToDoList[$e.RowIndex].Status = [TaskStatus]::Completed
            $JsonContent.ToDoList[$e.RowIndex].DateCompleted = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
            $JsonContent.ToDoList[$e.RowIndex].CompletedBy = $env:USERNAME.ToUpper()

            # Remove the completed task from the JSON content
            $JsonContent.ToDoList = $JsonContent.ToDoList | Where-Object { $_ -ne $JsonContent.ToDoList[$e.RowIndex] }

            # Save the changes to the JSON file
            $JsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8

           # Remove the row from the DataGridView
            $dataGridViewToDoList.Rows.RemoveAt($e.RowIndex)

            # Check if the ToDoList array is empty
            if ($JsonContent.ToDoList.Count -eq 0) {
                # Set the ToDoList back to the default values
                $JsonContent.ToDoList = @(
                    @{
                        "Title" = "Example Task 1"
                        "Task" = "Description of Task 1"
                        "Status" = ([TaskStatus]::NotStarted).ToString()
                        "CompletedBy" = ""
                        "DateCompleted" = ""
                    }
                )
            }

            # Save the changes to the JSON file
            $JsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8



            Write-Host "Debug: Task marked as completed and row removed"
        } catch {
            Write-Host "Error: Failed to complete task - $_"
        }
    }
}

    # Attach the "Complete" button click event handler
    $dataGridViewToDoList.Add_CellClick($dataGridViewToDoList_CellClick)
    
    # Subscribe to the CellValueChanged event
    $dataGridViewToDoList.add_CellValueChanged($dataGridViewToDoList_CellValueChanged)

    # Add this function to handle the RowLeave event
    $dataGridViewToDoList_RowLeave = {
        param (
            $sender,
            [System.Windows.Forms.DataGridViewCellEventArgs]$e
        )

        # Get the current row
        $row = $dataGridViewToDoList.Rows[$e.RowIndex]

        # Update the JSON object with the current row's values
        $JsonContent.ToDoList[$e.RowIndex].Title = $row.Cells[$TitleColumn.Index].Value
        $JsonContent.ToDoList[$e.RowIndex].Task = $row.Cells[$TaskColumn.Index].Value
        $JsonContent.ToDoList[$e.RowIndex].Status = [TaskStatus]::($row.Cells[$StatusColumn.Index].Value)
        $JsonContent.ToDoList[$e.RowIndex].CompletedBy = $row.Cells[$completedByColumn.Index].Value
        $JsonContent.ToDoList[$e.RowIndex].DateCompleted = $row.Cells[$datecompletedColumn.Index].Value

        # Save the changes to the JSON file
        $JsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8
    }

    # Subscribe to the RowLeave event
    $dataGridViewToDoList.add_RowLeave($dataGridViewToDoList_RowLeave)

    # Show the To-Do List form
    $toDoListForm.ShowDialog()
}

function Add-TaskToToDoList {
    param (
        [System.Windows.Forms.DataGridView]$dataGridViewToDoList,
        [System.Management.Automation.PSCustomObject]$jsonContent
    )

    # Create a new task object
    $newTask = [PSCustomObject]@{
        Complete = $false
        Task = "New Task"
        Status = [TaskStatus]::NotStarted
        CompletedBy = $env:USERNAME.ToUpper()
        DateCompleted = $null
    }

    # Add the new task to the JSON ToDoList
    $jsonContent.ToDoList += $newTask

    # Add the new task to the DataGridView
    $rowIndex = $dataGridViewToDoList.Rows.Add()
    $row = $dataGridViewToDoList.Rows[$rowIndex]
    $row.Cells[$completeColumn.Index].Value = $newTask.Complete
    $row.Cells[$titleColumn.Index].Value = $newTask.Task
    $row.Cells[$taskColumn.Index].Value = $newTask.Task
    $row.Cells[$StatusColumn.Index].Value = $newTask.Status
    $row.Cells[$completedByColumn.Index].Value = $newTask.CompletedBy
    $row.Cells[$datecompletedColumn.Index].Value = $newTask.DateCompleted
}

# Function Get Highest Version Number from Modifications for Incrementing Log
function Get-HighestVersionFromModifications {
    param (
        [array]$modifications,
        [string]$jsonPath
    )

    $versions = @()
    if ($null -ne $modifications) {
        $versions = $modifications | ForEach-Object { $_.Version }
    }

    $highestVersion = "0.1" # Default value
    if ($versions.Count -gt 0) {
        $highestVersion = [version]($versions | Sort-Object -Descending | Select-Object -First 1)
    }

    return $highestVersion
}


# Function to set up the DataGridView event handler for CellContentClick
function SetupDataGridViewEventHandler {
    param (
        $JsonPath,
        $tableView,
        $textBoxVersion  # Add the version TextBox as a parameter
    )

    Write-Host "Inside SetupDataGridViewEventHandler function"

    # Inside the event handler setup:
    $tableView_CellContentClick = {
        param ($sender, $e)

        Write-Host "Inside tableView_CellContentClick event handler"
        Write-Host "Column index: $($e.ColumnIndex)"

        # Check if the clicked column is the delete column (column index 0)
        if ($e.ColumnIndex -eq 0) {
            $rowIndex = $e.RowIndex
            Write-Host "Row index for delete button: $rowIndex"

            if ($rowIndex -ge 0) {
                HandleDeleteButtonClick $JsonPath $tableView $rowIndex $textBoxVersion
            } else {
                Write-Host "Invalid row index: $rowIndex"
            }
        } else {
            Write-Host "Not a delete button click."
        }
    }


    $tableView.add_CellContentClick($tableView_CellContentClick)
}

# Function to handle the CellContentClick event for the "Delete" button column
function HandleDeleteButtonClick {
    param (
        $JsonPath,
        $tableView,
        $rowIndex,
        $textBoxVersion  # Add the version TextBox as a parameter
    )

    try {
        Write-Host "Inside HandleDeleteButtonClick function"

        if ($rowIndex -ge 0) {
            $selectedVersion = $tableView.Rows[$rowIndex].Cells["Version"].Value
            if ($selectedVersion -eq $null) {
                $selectedVersion = ""  # Handle empty version value
            }

            Write-Host "Selected version to delete: $selectedVersion"

            # Read the JSON content
            $jsonContent = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json

            Write-Host "JSON content before deletion:"
            $jsonContent | ConvertTo-Json -Depth 4

            # Find and remove the entry from the JSON content
            $jsonContent.Modifications = $jsonContent.Modifications | Where-Object { $_.Version -ne $selectedVersion }

            # Update the version property in the JSON content
            if ($jsonContent.Modifications.Count -gt 0) {
                $latestModification = $jsonContent.Modifications | Sort-Object -Property Version | Select-Object -Last 1
                $jsonContent.Version = $latestModification.Version  # Update to the latest version
            } else {
                $jsonContent.Version = "0.0"  # Reset to 0.0 if no modifications are left
            }

            # Save the updated JSON content back to the file
            $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8

            # Remove the selected row from the DataGridView
            $tableView.Rows.RemoveAt($rowIndex)

            # Get the highest version from the modified JSON content
            Write-Host "Before calling Get-HighestVersionFromModifications: JsonPath = $JsonPath"
            #$highestVersion, $_ = Get-HighestVersionFromModifications $jsonContent.Modifications
            $highestVersion, $_ = Get-HighestVersionFromModifications -Modifications $jsonContent.Modifications -JsonPath $JsonPath


            Write-Host "Highest version after deletion: $highestVersion"

            # Update the version TextBox in the main form
            $textBoxVersion.Invoke([Action]{
                $textBoxVersion.Text = $highestVersion
            })

            # Refresh the form
            $form.Refresh()

            Write-Host "JSON content after deletion:"
            $jsonContent | ConvertTo-Json -Depth 4

            Write-Host "Modification log entry deleted successfully."
        } else {
            Write-Host "Invalid row index: $rowIndex"
        }
    } catch {
        Write-Host "An error occurred: $_"
    }
}

# Function to show a blank modification log form
function ShowBlankModificationLogForm {
    param (
        [string]$JsonPath,
        [System.Windows.Forms.Form]$mainForm,
        [string]$currentVersion
    )

    Write-Host "Inside ShowBlankModificationLogForm function"
    Write-Host "JSON file path: $JsonPath"

    # Create the form and ListView
    $modLogForm = New-Object System.Windows.Forms.Form
    $modLogForm.Text = "Modification Log"
    $modLogForm.Size = New-Object System.Drawing.Size(820, 1000)  # Increase W for wider, H for taller
    $modLogForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $modLogForm.BackColor = [System.Drawing.Color]::SaddleBrown
    $modLogForm.ForeColor = [System.Drawing.Color]::White

    # Table View
    $tableView = New-Object System.Windows.Forms.DataGridView
    $tableView.Location = New-Object System.Drawing.Point(20, 20) # Increase X = Right, y = Down
    $tableView.Size = New-Object System.Drawing.Size(760, 800)  # Increase W for wider, H for taller
    $tableView.ColumnHeadersVisible = $true
    $tableView.RowHeadersVisible = $false
    $tableView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $tableView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
    $tableView.AllowUserToAddRows = $false
    $tableView.AllowUserToDeleteRows = $false
    $tableView.AllowUserToResizeRows = $false
    $tableView.MultiSelect = $false
    $tableView.BackgroundColor = [System.Drawing.Color]::BurlyWood
    $tableView.GridColor = [System.Drawing.Color]::Navy
    $tableView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Red
    $tableView.DefaultCellStyle.ForeColor = [System.Drawing.Color]::SaddleBrown
    #$tableviewfont = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    #$tableViewfont.DefaultCellStyle.Font = $font

    # Create "Delete" button column
    $deleteColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $deleteColumn.HeaderText = "Delete"
    $deleteColumn.Text = "Delete"
    $deleteColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $deleteColumn.UseColumnTextForButtonValue = $true

    # Customize the button cell style
    $buttonCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $buttonCellStyle.ForeColor = [System.Drawing.Color]::Red

    # Set the FlatStyle property to Popup to change the button appearance
    $deleteColumn.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup

    # Assign the custom style to the button column
    $deleteColumn.DefaultCellStyle = $buttonCellStyle

    $tableView.Columns.Add($deleteColumn)

    $dateColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $dateColumn.Name = "Date"
    $dateColumn.HeaderText = "Date"
    $dateColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $tableView.Columns.Add($dateColumn)

    $versionColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $versionColumn.Name = "Version"
    $versionColumn.HeaderText = "Version"
    $versionColumn.Width = 50
    $versionColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $versionColumn.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $versionColumn.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
    $tableView.Columns.Add($versionColumn)

    $modifiedByColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $modifiedByColumn.Name = "ModifiedBy"
    $modifiedByColumn.HeaderText = "Modified By"
    $modifiedByColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $modifiedByColumn.Width = 60
    $tableView.Columns.Add($modifiedByColumn)

    $modificationColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $modificationColumn.Name = "Modification"
    $modificationColumn.HeaderText = "Modification"
    $modificationColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $modificationColumn.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    $tableView.Columns.Add($modificationColumn)

    # Read the JSON content
    $jsonContent = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json

    # Ensure the "Modifications" field exists and is either an array or null
    if (-not $jsonContent.PSObject.Properties.Name -contains "Modifications" -or ($jsonContent.Modifications -ne $null -and $jsonContent.Modifications -isnot [System.Collections.IEnumerable])) {
        # Initialize "Modifications" as an empty array if it is missing or not an array
        $jsonContent.Modifications = @()
    }

    # Populate the table view with existing modifications
    foreach ($entry in $jsonContent.Modifications) {
        # Convert the Version value to a string with one decimal place
        $versionString = if ($entry.Version -ne $null) { [string]::Format("{0:F1}", $entry.Version) } else { "" }

        # Log the entry being added to the table view
        Write-Host "Adding entry to the table view:"
        Write-Host "Date: $($entry.Date), ModifiedBy: $($entry.ModifiedBy), Version: $($entry.Version), Modification: $($entry.Modification)"
        $tableView.Rows.Add("Delete", $entry.Date, $versionString, $entry.ModifiedBy, $entry.Modification)
    }

    # Ensure the "Version" field exists and is in the correct format
    if ($jsonContent.Version -eq $null -or ([Version]::TryParse($jsonContent.Version, [ref]$null) -eq $false)) {
        $jsonContent.Version = "0.0"  # Set the default version to 0.0
    }

    # Add this line to call the event handler setup function
    SetupDataGridViewEventHandler $JsonPath $tableView $textBoxVersion

    $modLogForm.Controls.Add($tableView)

    # Label and TextBox for "Modified Date"
    $labelModifiedDate = New-Object System.Windows.Forms.Label
    $labelModifiedDate.Text = "Modified Date:"
    $labelModifiedDate.Location = New-Object System.Drawing.Point(20, 840) # Adjust coordinates as needed
    $labelModifiedDate.Size = New-Object System.Drawing.Size(100, 20)
    $modLogForm.Controls.Add($labelModifiedDate)

    # TextBox for "Modified Date Text"
    $textBoxModifiedDate = New-Object System.Windows.Forms.TextBox
    $textBoxModifiedDate.Location = New-Object System.Drawing.Point(140, 840) # Adjust coordinates as needed
    $textBoxModifiedDate.Size = New-Object System.Drawing.Size(280, 20) # Increase width for longer date format
    $textBoxModifiedDate.BackColor = [System.Drawing.Color]::White
    $textBoxModifiedDate.Text = Get-Date -Format "MMMM d, yyyy 'at' h:mm tt" # Prepopulate with current date and time
    $modLogForm.Controls.Add($textBoxModifiedDate)

    # Label and TextBox for "Modified By"
    $labelAddModifiedBy = New-Object System.Windows.Forms.Label
    $labelAddModifiedBy.Text = "Modified By:"
    $labelAddModifiedBy.Location = New-Object System.Drawing.Point(425, 845) # Move right for additional field
    $labelAddModifiedBy.Size = New-Object System.Drawing.Size(70, 20)
    $modLogForm.Controls.Add($labelAddModifiedBy)

    $textBoxAddModifiedBy = New-Object System.Windows.Forms.TextBox
    $textBoxAddModifiedBy.Location = New-Object System.Drawing.Point(500, 840) # Move right for additional field
    $textBoxAddModifiedBy.Size = New-Object System.Drawing.Size(280, 20) # Adjust width as needed
    $textBoxAddModifiedBy.BackColor = [System.Drawing.Color]::White
    $textBoxAddModifiedBy.Text = $env:USERNAME.ToUpper() # Prepopulate with the current user in uppercase
    $modLogForm.Controls.Add($textBoxAddModifiedBy)

    # Label for "Modification"
    $labelModification = New-Object System.Windows.Forms.Label
    $labelModification.Text = "Modification:"
    $labelModification.Location = New-Object System.Drawing.Point(20, 870) # Move down for additional field
    $labelModification.Size = New-Object System.Drawing.Size(100, 20)
    $modLogForm.Controls.Add($labelModification)

    # TextBox for "Modification Text"
    $textBoxAddModificationText = New-Object System.Windows.Forms.TextBox
    $textBoxAddModificationText.Multiline = $true # Allow multiple lines
    $textBoxAddModificationText.ScrollBars = "Vertical" # Show vertical scroll bar for long text
    $textBoxAddModificationText.Location = New-Object System.Drawing.Point(140, 870) # Move down for additional field
    $textBoxAddModificationText.Size = New-Object System.Drawing.Size(640, 80) # Adjust height and width as needed
    $textBoxAddModificationText.BackColor = [System.Drawing.Color]::White
    $modLogForm.Controls.Add($textBoxAddModificationText)

    # Create "Add Modification" button
    $buttonSaveModification = New-Object System.Windows.Forms.Button
    $buttonSaveModification.Text = "Add Modification" # Change button text
    $buttonSaveModification.Location = New-Object System.Drawing.Point(15, 920) # Move down for additional field
    $buttonSaveModification.Size = New-Object System.Drawing.Size(120, 30) # Adjust width as needed
    $buttonSaveModification.BackColor = [System.Drawing.Color]::DodgerBlue  # Set the background color of the button
    $buttonSaveModification.ForeColor = [System.Drawing.Color]::White      # Set the text color of the button

    # Save Modification Event Handler
    $buttonSaveModification.Add_Click({
        try {
            $date = $textBoxModifiedDate.Text
            $modifiedBy = $textBoxAddModifiedBy.Text.ToUpper() # Convert to uppercase
            $modificationText = $textBoxAddModificationText.Text.ToUpper() # Convert to uppercase

            Write-Host "Before accessing tableView: Count = $($tableView.Rows.Count)"
            # Get the latest version from the DataGridView (table view)
            if ($tableView.Rows.Count -gt 0) {
                $latestVersion = $tableView.Rows[$tableView.Rows.Count - 1].Cells["Version"].Value
            } else {
                $latestVersion = [Version]::new(0, 0)
            }
            Write-Host "After accessing tableView: Latest Version = $latestVersion"

            $versionColumn = $tableView.Columns["Version"]
            Write-Host "Version column exists: $($versionColumn -ne $null)"

            # Increment the version by 0.1 for each new modification
            if ($latestVersion -ne $null) {
                $currentVersion = [Version]$latestVersion
                $minorVersion = $currentVersion.Minor + 1

                if ($minorVersion -gt 9) {
                    $majorVersion = $currentVersion.Major + 1
                    $minorVersion = 0
                } else {
                    $majorVersion = $currentVersion.Major
                }

                $incrementedVersion = [Version]::new($majorVersion, $minorVersion)
            } else {
                $incrementedVersion = [Version]::new(0, 1)
            }

            # Create a hashtable for the new modification
            $newModification = @{
                Date = $date
                ModifiedBy = $modifiedBy
                Version = $incrementedVersion.ToString()  # Update the version here
                Modification = $modificationText
            }

            # Read the JSON content
            $jsonContent = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json

            # Ensure the "Modifications" field exists and is either an array or null
            if (-not $jsonContent.PSObject.Properties.Name -contains "Modifications" -or ($jsonContent.Modifications -ne $null -and $jsonContent.Modifications -isnot [System.Collections.IEnumerable])) {
                # Initialize "Modifications" as an empty array if it is missing or not an array
                $jsonContent.Modifications = @()
            }

            # Add the new modification to the "Modifications" array
            $jsonContent.Modifications += $newModification

            # Save the updated JSON content back to the file
            $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8

            # Update the ListView with the modified JSON content
            $tableView.Rows.Add("Delete", $date, $incrementedVersion.ToString(), $modifiedBy, $modificationText)

            # Update the highest version and get modified JSON content
            Write-Host "JSON path before calling Get-HighestVersionFromModifications: $JsonPath"
            $highestVersion, $jsonContent = Get-HighestVersionFromModifications $jsonContent.Modifications $JsonPath
            Write-Host "Highest version after Get-HighestVersionFromModifications: $highestVersion"

            # Update the version TextBox on the main form
            $textBoxVersion.Text = $highestVersion.ToString()

            if ($jsoncontent.description -match "PSM Master Script"){
                UpdateFormTitle -form $form -textBoxVersion $textBoxVersion -newVersion $highestVersion
            }

            # Clear the fields
            $textBoxModifiedDate.Text = Get-Date -Format "MMMM d, yyyy 'at' h:mm tt"
            $textBoxAddModifiedBy.Text = $env:USERNAME.ToUpper()
            $textBoxAddModificationText.Text = ""

            #PopulateFieldsAndTags
            } catch {
                Write-Host "An error occurred: $_"
            }
        })


    $modLogForm.Controls.Add($buttonSaveModification)

    # Show the form for adding a modification
    $modLogForm.ShowDialog()
    
    #Write-Host "Finished ShowBlankModificationLogForm function"
}

# Function to show the form for adding a new modification
function ShowAddModificationForm {
    param (
        [string]$JsonPath
    )

    # Create the form for adding a modification
    $addModForm = New-Object System.Windows.Forms.Form
    $addModForm.Text = "Add Modification"
    $addModForm.Size = New-Object System.Drawing.Size(400, 400) # Adjust size as needed
    $addModForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Label and TextBox for "Version"
    $labelAddVersion = New-Object System.Windows.Forms.Label
    $labelAddVersion.Text = "Version:"
    $labelAddVersion.Location = New-Object System.Drawing.Point(20, 20) # Increase X = Right, y = Down
    $labelAddVersion.Size = New-Object System.Drawing.Size(80, 20) # Increase W = Wider, H = Taller.
    $addModForm.Controls.Add($labelAddVersion)

    $textBoxAddVersion = New-Object System.Windows.Forms.TextBox
    $textBoxAddVersion.Location = New-Object System.Drawing.Point(140, 20) # Increase X = Right, y = Down
    $textBoxAddVersion.Size = New-Object System.Drawing.Size(220, 20) # Increase W = Wider, H = Taller.
    $textBoxAddVersion.BackColor = [System.Drawing.Color]::White
    $textBoxAddVersion.Text = $jsonContent.Version  # Prepopulate with the current version from JSON content
    $textBoxAddVersion.Enabled = $false # Disable editing of the version field
    $addModForm.Controls.Add($textBoxAddVersion)

    # Label and TextBox for "Modified By"
    $labelAddModifiedBy = New-Object System.Windows.Forms.Label
    $labelAddModifiedBy.Text = "Modified By:"
    $labelAddModifiedBy.Location = New-Object System.Drawing.Point(20, 50) # Increase X = Right, y = Down
    $labelAddModifiedBy.Size = New-Object System.Drawing.Size(100, 20)      # Increase W for wider, H for taller
    $addModForm.Controls.Add($labelAddModifiedBy)

    $textBoxAddModifiedBy = New-Object System.Windows.Forms.TextBox
    $textBoxAddModifiedBy.Location = New-Object System.Drawing.Point(140, 50) # Increase X = Right, y = Down
    $textBoxAddModifiedBy.Size = New-Object System.Drawing.Size(220, 20)      # Increase W = Wider, H = Taller.
    $textBoxAddModifiedBy.BackColor = [System.Drawing.Color]::White
    $textBoxAddModifiedBy.Text = $env:USERNAME.ToUpper()  # Populate with the current user's name
    $addModForm.Controls.Add($textBoxAddModifiedBy)

    # Label and TextBox for "Modification Text"
    $labelAddModificationText = New-Object System.Windows.Forms.Label
    $labelAddModificationText.Text = "Modification Text:"
    $labelAddModificationText.Location = New-Object System.Drawing.Point(20, 80) # Increase X = Right, y = Down
    $labelAddModificationText.Size = New-Object System.Drawing.Size(100, 20)      # Increase W for wider, H for taller
    $addModForm.Controls.Add($labelAddModificationText)

    $textBoxAddModificationText = New-Object System.Windows.Forms.TextBox
    $textBoxAddModificationText.Location = New-Object System.Drawing.Point(140, 80)# Increase X = Right, y = Down
    $textBoxAddModificationText.Size = New-Object System.Drawing.Size(220, 180)      # Increase W = Wider, H = Taller.
    $textBoxAddModificationText.BackColor = [System.Drawing.Color]::White
    $addModForm.Controls.Add($textBoxAddModificationText)

    # Create "Save" button for saving the new modification
    $buttonSaveModification = New-Object System.Windows.Forms.Button
    $buttonSaveModification.Text = "Save"
    $buttonSaveModification.Location = New-Object System.Drawing.Point(200, 260)  # Increase X = Right, y = Down
    $buttonSaveModification.Size = New-Object System.Drawing.Size(100, 30)        # Increase W = Wider, H = Taller.
    
    # Save Modifications Event Handler
$buttonSaveModification.Add_Click({
    try {
        Write-Host "Button Save Modification Click Event Handler Triggered"

        $date = Get-Date
        $modifiedBy = $textBoxAddModifiedBy.Text
        $modificationText = $textBoxAddModificationText.Text

        # Get the latest version from the Modifications array
        $latestVersion = Get-HighestVersionFromModifications $jsonContent.Modifications

        # Increment the version by 0.1 for each new modification
        $newVersion = [Version]::new($latestVersion.Major, $latestVersion.Minor + 1)

        Write-Host "Version Incremented: New Version = $($newVersion)"

        # Create a hashtable for the new modification
        $newModification = @{
            Date = $date
            ModifiedBy = $modifiedBy
            Version = $newVersion.ToString()
            Modification = $modificationText
        }

        Write-Host "New Modification:"
        $newModification | Format-Table -AutoSize

        # Add the new modification to the "Modifications" array
        $jsonContent.Modifications += $newModification

        Write-Host "New Modification Added to Modifications Array"

        # Save the updated JSON content back to the file
        $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8

        Write-Host "JSON Content Updated and Saved"

        # Update the ListView in the main form with the new modification
        Update-ListView

        Write-Host "ListView Updated with New Modification"

        $textBoxAddModificationText.Text = ""  # Clear the modification text box

        Write-Host "Modification TextBox Cleared"

        # Close the "Add Modification" form
        $addModForm.Close()

        Write-Host "Add Modification Form Closed"

    } catch {
        Write-Host "An error occurred: $_"
    }
})


    $addModForm.Controls.Add($buttonSaveModification)

    # Show the form for adding a modification
    $addModForm.ShowDialog()
}

# Update ListView
function Update-ListView {

    # Clear the current list view
    $modLogView.Items.Clear()

    # Read the JSON content again
    $jsonContent = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json

    # Populate the list view with existing modifications
    foreach ($entry in $jsonContent.Modifications) {
        $item = New-Object System.Windows.Forms.ListViewItem
        $item.Text = $entry.Date
        $item.SubItems.Add($entry.ModifiedBy)
        $item.SubItems.Add($entry.Version)
        $item.SubItems.Add($entry.Modification)
        $modLogView.Items.Add($item)
    }
}

# Function Show Modification Log
function ShowModificationLog {
    #Clear-Host
    # Check if a script is selected in the ListBox
    $selectedScript = $listBoxScripts.SelectedItem
    if (-not $selectedScript) {
        Write-Host "No script selected"
        return
    }

    # Get the script and JSON file paths
    $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$selectedScript.ps1"
    $jsonPath = $scriptPath -replace "\.ps1$", ".json"
    #Write-Host "JSON Path: $jsonPath"

    if (Test-Path -Path $jsonPath) {
        # Read the JSON content
        $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
        #Write-Host "JSON Content:"
        $jsonContent | ConvertTo-Json -Depth 1

        # Ensure the "Modifications" field exists and is either an array or null
        if (-not $jsonContent.PSObject.Properties.Name -contains "Modifications" -or ($jsonContent.Modifications -ne $null -and $jsonContent.Modifications -isnot [System.Collections.IEnumerable])) {
            Write-Host "Modifications field not found or not an array: $jsonPath"

            # Initialize "Modifications" as an empty array if it is missing or not an array
            $jsonContent.Modifications = @()
            $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $jsonPath -Encoding UTF8
        }

        # Show the modification log form
        ShowBlankModificationLogForm -JsonPath $jsonPath -JsonContent $jsonContent
    } else {
        Write-Host "JSON file not found: $jsonPath"

        # Create a blank JSON file with an empty "Modifications" array
        $emptyJsonContent = @{
            Modifications = @()
        } | ConvertTo-Json -Depth 4
        $emptyJsonContent | Out-File -FilePath $jsonPath -Encoding UTF8

        # Show the modification log form
        ShowBlankModificationLogForm -JsonPath $jsonPath -JsonContent $emptyJsonContent
    }
}

# Inspirtaional Quotes
$inspirationalQuotes = @(
    "The only way to do great work is to love what you do. -Steve Jobs",
    "Don't watch the clock; do what it does. Keep going. -Sam Levenson",
    "Believe you can and you're halfway there. -Theodore Roosevelt",
    "Your time is limited, don't waste it living someone else's life. -Steve Jobs",
    "You miss 100% of the shots you don't take. -Wayne Gretzky",
    "Success is not the key to happiness. Happiness is the key to success. If you love what you are doing, you will be successful. -Albert Schweitzer",
    "The future belongs to those who believe in the beauty of their dreams. -Eleanor Roosevelt",
    "The only person you are destined to become is the person you decide to be. -Ralph Waldo Emerson",
    "Happiness is not something ready-made. It comes from your own actions. -Dalai Lama",
    "Believe in yourself, take on your challenges, dig deep within yourself to conquer fears. -Chantal Sutherland",
    "Don't be afraid to give up the good to go for the great. -John D. Rockefeller",
    "The secret of getting ahead is getting started. -Mark Twain",
    "The only limit to our realization of tomorrow will be our doubts of today. -Franklin D. Roosevelt",
    "The best way to predict the future is to create it. -Peter Drucker",
    "The only way to achieve the impossible is to believe it is possible. -Charles Kingsleigh",
    "It does not matter how slowly you go as long as you do not stop. -Confucius",
    "Your work is going to fill a large part of your life, and the only way to be truly satisfied is to do what you believe is great work. And the only way to do great work is to love what you do. If you haven't found it yet, keep looking. Don't settle. As with all matters of the heart, you'll know when you find it. -Steve Jobs",
    "The only thing standing between you and your goal is the story you keep telling yourself as to why you can't achieve it. -Jordan Belfort",
    "The only way to achieve the impossible is to believe it is possible. -Charles Kingsleigh",
    "The more you praise and celebrate your life, the more there is in life to celebrate. -Oprah Winfrey",
    "Believe you can and you're halfway there. -Theodore Roosevelt",
    "Success is not final, failure is not fatal: It is the courage to continue that counts. -Winston Churchill",
    "You are never too old to set another goal or to dream a new dream. -C.S. Lewis",
    "Don't watch the clock; do what it does. Keep going. -Sam Levenson",
    "The future belongs to those who believe in the beauty of their dreams. -Eleanor Roosevelt",
    "The only person you are destined to become is the person you decide to be. -Ralph Waldo Emerson",
    "Happiness is not something ready-made. It comes from your own actions. -Dalai Lama",
    "Believe in yourself, take on your challenges, dig deep within yourself to conquer fears. -Chantal Sutherland",
    "Your time is limited, don't waste it living someone else's life. -Steve Jobs",
    "The best way to predict the future is to create it. -Peter Drucker",
    "Don't be afraid to give up the good to go for the great. -John D. Rockefeller",
    "The secret of getting ahead is getting started. -Mark Twain",
    "The only limit to our realization of tomorrow will be our doubts of today. -Franklin D. Roosevelt",
    "The best way to predict the future is to create it. -Peter Drucker",
    "The only way to achieve the impossible is to believe it is possible. -Charles Kingsleigh",
    "It does not matter how slowly you go as long as you do not stop. -Confucius",
    "Your work is going to fill a large part of your life, and the only way to be truly satisfied is to do what you believe is great work. And the only way to do great work is to love what you do. If you haven't found it yet, keep looking. Don't settle. As with all matters of the heart, you'll know when you find it. -Steve Jobs",
    "The only thing standing between you and your goal is the story you keep telling yourself as to why you can't achieve it. -Jordan Belfort",
    "The only way to achieve the impossible is to believe it is possible. -Charles Kingsleigh",
    "The more you praise and celebrate your life, the more there is in life to celebrate. -Oprah Winfrey",
    "Your time is limited, don't waste it living someone else's life. -Steve Jobs",
    "The future belongs to those who believe in the beauty of their dreams. -Eleanor Roosevelt",
    "The best way to predict the future is to create it. -Peter Drucker",
    "The only person you are destined to become is the person you decide to be. -Ralph Waldo Emerson",
    "Happiness is not something ready-made. It comes from your own actions. -Dalai Lama",
    "Believe in yourself, take on your challenges, dig deep within yourself to conquer fears. -Chantal Sutherland",
    "Don't be afraid to give up the good to go for the great. -John D. Rockefeller",
    "The secret of getting ahead is getting started. -Mark Twain",
    "The only limit to our realization of tomorrow will be our doubts of today. -Franklin D. Roosevelt",
    "The best way to predict the future is to create it. -Peter Drucker",
    "The only way to achieve the impossible is to believe it is possible. -Charles Kingsleigh",
    "It does not matter how slowly you go as long as you do not stop. -Confucius",
    "Your work is going to fill a large part of your life, and the only way to be truly satisfied is to do what you believe is great work. And the only way to do great work is to love what you do. If you haven't found it yet, keep looking. Don't settle. As with all matters of the heart, you'll know when you find it. -Steve Jobs",
    "The only thing standing between you and your goal is the story you keep telling yourself as to why you can't achieve it. -Jordan Belfort",
    "The only way to achieve the impossible is to believe it is possible. -Charles Kingsleigh",
    "The more you praise and celebrate your life, the more there is in life to celebrate. -Oprah Winfrey",
    "Don't watch the clock; do what it does. Keep going. -Sam Levenson",
    "Your time is limited, don't waste it living someone else's life. -Steve Jobs",
    "The future belongs to those who believe in the beauty of their dreams. -Eleanor Roosevelt",
    "The best way to predict the future is to create it. -Peter Drucker",
    "The only person you are destined to become is the person you decide to be. -Ralph Waldo Emerson",
    "Happiness is not something ready-made. It comes from your own actions. -Dalai Lama",
    "Believe in yourself, take on your challenges, dig deep within yourself to conquer fears. -Chantal Sutherland",
    "Don't be afraid to give up the good to go for the great. -John D. Rockefeller",
    "The secret of getting ahead is getting started. -Mark Twain",
    "The only limit to our realization of tomorrow will be our doubts of today. -Franklin D. Roosevelt",
    "The best way to predict the future is to create it. -Peter Drucker",
    "The only way to achieve the impossible is to believe it is possible. -Charles Kingsleigh",
    "It does not matter how slowly you go as long as you do not stop. -Confucius",
    "Your work is going to fill a large part of your life, and the only way to be truly satisfied is to do what you believe is great work. And the only way to do great work is to love what you do. If you haven't found it yet, keep looking. Don't settle. As with all matters of the heart, you'll know when you find it. -Steve Jobs",
    "The only thing standing between you and your goal is the story you keep telling yourself as to why you can't achieve it. -Jordan Belfort",
    "The only way to achieve the impossible is to believe it is possible. -Charles Kingsleigh",
    "The more you praise and celebrate your life, the more there is in life to celebrate. -Oprah Winfrey"
)

# Updated Tags
$tags = @(
    @{ Name = "Active Directory"; Icon = "👥" },
    @{ Name = "ARI Storage"; Icon = "🔍" },
    @{ Name = "Automation"; Icon = "🤖" },
    @{ Name = "Azure"; Icon = "☁️" },
    @{ Name = "Backed Up"; Icon = "💾" },
    @{ Name = "Bitlocker"; Icon = "🔒" },
    @{ Name = "Broken"; Icon = "🔨" },
    @{ Name = "Certificates"; Icon = "📜" },
    @{ Name = "Citrix"; Icon = "🖥️" },
    @{ Name = "Cloud"; Icon = "☁️" },
    @{ Name = "Commands"; Icon = "⌨️" },
    @{ Name = "Configuration"; Icon = "⚙️" },
    @{ Name = "Counters"; Icon = "🔢" },
    @{ Name = "Databases"; Icon = "🗃️" },
    @{ Name = "DCAP"; Icon = "🔒" },
    @{ Name = "Deployment"; Icon = "🚀" },
    @{ Name = "Design"; Icon = "✏️" },
    @{ Name = "Development"; Icon = "👩‍💻" },
    @{ Name = "Engineering"; Icon = "👷" },
    @{ Name = "Error Handling"; Icon = "❌" },
    @{ Name = "Excel"; Icon = "📊" },
    @{ Name = "Favorite"; Icon = "⭐" },
    @{ Name = "Finalized"; Icon = "🏁" },
    @{ Name = "Firewall"; Icon = "🔥" },
    @{ Name = "Fixing"; Icon = "🚧" },
    @{ Name = "Functions"; Icon = "🔧" },
    @{ Name = "Has Keywords"; Icon = "🔠" },
    @{ Name = "High Priority"; Icon = "🔝" },
    @{ Name = "Hypervisor"; Icon = "🛠️" },
    @{ Name = "iDRAC"; Icon = "🖥️" },
    @{ Name = "Integration"; Icon = "🔗" },
    @{ Name = "Junk"; Icon = "🗑️" },
    @{ Name = "Loops"; Icon = "➰" },
    @{ Name = "Logging"; Icon = "📝" },
    @{ Name = "Maintenance"; Icon = "🔧" },
    @{ Name = "Management"; Icon = "👨‍💼" },
    @{ Name = "Mine"; Icon = "🏴" },
    @{ Name = "Monitoring"; Icon = "👀" },
    @{ Name = "Modules"; Icon = "📦" },
    @{ Name = "Needs Approval"; Icon = "👍" },
    @{ Name = "Networking"; Icon = "🌐" },
    @{ Name = "New Script"; Icon = "🆕" },
    @{ Name = "Not Mine"; Icon = "🏳️" },
    @{ Name = "Open Tasks"; Icon = "📝" },
    @{ Name = "Parameters"; Icon = "🛠️" },
    @{ Name = "PowerShell"; Icon = "📜" },
    @{ Name = "Profiles"; Icon = "👤" },
    @{ Name = "PS Dashboard"; Icon = "📊" },
    @{ Name = "Quality Control"; Icon = "✅" },
    @{ Name = "Requirements"; Icon = "📋" },
    @{ Name = "Reviewed"; Icon = "🔍" },
    @{ Name = "Remoting"; Icon = "🔗" },
    @{ Name = "Reports"; Icon = "📊" },
    @{ Name = "Security"; Icon = "🛡️" },
    @{ Name = "Secrets"; Icon = "🔐" },
    @{ Name = "SSH"; Icon = "🔑" },
    @{ Name = "Storage"; Icon = "🗄️" },
    @{ Name = "Switch"; Icon = "🔀" },
    @{ Name = "SYSAdmin"; Icon = "💻" },
    @{ Name = "Tape Library"; Icon = "📼" },
    @{ Name = "Templates"; Icon = "📄" },
    @{ Name = "Tools"; Icon = "🛠️" },
    @{ Name = "User Management"; Icon = "👤" },
    @{ Name = "Virtual Machines"; Icon = "🖥️" },
    @{ Name = "Working"; Icon = "✔️" },
    @{ Name = "WTF"; Icon = "⁉️" }
)


# Set the script directory
#$scriptDirectory = "C:\curtis\PSM"

# Find all script files in the directory matching the pattern "PSM*.ps1"
$scriptFiles = Get-ChildItem -Path $global:BaseScriptPath -File -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Where-Object { $_.Name -match '^PSM\d+\.ps1$' }

# Initialize variables
$highestScriptVersion = -1
$highestJsonContent = $null
$highestVersion = -1

# Loop through the script files
foreach ($scriptFile in $scriptFiles) {
    # Extract the script version from the filename
    $scriptVersion = [int]($scriptFile.Name -replace '^PSM(\d+)\.ps1$', '$1')

    # Check if the extracted version is higher than the current highest
    if ($scriptVersion -gt $highestScriptVersion) {
        $highestScriptVersion = $scriptVersion

        # Construct the JSON file path
        $jsonFilePath = Join-Path -Path $global:BaseScriptPath -ChildPath ($scriptFile.Name -replace '\.ps1$', '.json')

        # Get the JSON content and convert it to a PowerShell object
        $jsonContent = Get-Content -Path $jsonFilePath -Raw | ConvertFrom-Json

        # Find the highest version within the JSON modifications
        $highestVersionObject = $jsonContent.Modifications | Sort-Object -Property Version | Select-Object -Last 1

        # Check if the version string is empty
        if (-not [string]::IsNullOrWhiteSpace($highestVersionObject.Version)) {
            $highestVersion = [Version]::new($highestVersionObject.Version)
            $highestJsonContent = $jsonContent
        }
    }
}

# Extract the highest version string from the JSON content
$versionString = "v$($highestVersion.Major).$($highestVersion.Minor)"

# Get the highest version from the JSON modifications
$scriptVersionString = $highestJsonContent.Version

# Initializing the Main PowerShell Script Manager Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell Script Manager $scriptVersionString"
$form.Size = New-Object System.Drawing.Size(1080, 1000) # Increase W = Wider, H = Taller.
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::Wheat

## Create a button to open the summer colors form
#$buttonViewColors = New-Object System.Windows.Forms.Button
#$buttonViewColors.Text = "View Summer Colors"
#$buttonViewColors.Location = New-Object System.Drawing.Point(275, 480)
#$buttonViewColors.Size = New-Object System.Drawing.Size(150, 30)
#$buttonViewColors.Add_Click({
    #Start-Process powershell -ArgumentList "-File 'c:\curits\psm\colors.ps1'"
#})

#$mainForm.Controls.Add($buttonViewColors)

#$buttonSearchMode.Location = New-Object System.Drawing.Point(55, 510) # Increase X = Right, Y = Down
#$buttonSearchMode.Size = New-Object System.Drawing.Size(100, 30) # Increase W = Wider, H = Taller

# Output RichTextBox
$richTextBoxOutput = New-Object System.Windows.Forms.RichTextBox
$richTextBoxOutput.Location = New-Object System.Drawing.Point(15, 760) # Increase X = Right, Y = Down
$richTextBoxOutput.Size = New-Object System.Drawing.Size(1035, 180)     # Increase W = Wider, H = Taller.
$richTextBoxOutput.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9) # Set font size
$richTextBoxOutput.BackColor = [System.Drawing.Color]::SaddleBrown
$richTextBoxOutput.ForeColor = [System.Drawing.Color]::WhiteSmoke
$richTextBoxOutput.ReadOnly = $true
$form.Controls.Add($richTextBoxOutput)

# Scripts Path Label
$labelScriptsPath = New-Object System.Windows.Forms.Label
$labelScriptsPath.Text = "Script Functions"
$labelScriptsPath.Location = New-Object System.Drawing.Point(20, 25) # Increase X = Right, y = Down
$labelScriptsPath.Size = New-Object System.Drawing.Size(103, 15)     # Increase W = Wider, H = Taller.
$labelScriptsPath.Font = New-Object System.Drawing.Font("Mistral", 10)
$form.Controls.Add($labelScriptsPath)

# Script Path Text Box
$textBoxScriptsPath = New-Object System.Windows.Forms.TextBox
$textBoxScriptsPath.Location = New-Object System.Drawing.Point(0, 7) # Increase X = Right, y = Down
$textBoxScriptsPath.Size = New-Object System.Drawing.Size(1035, 30)      # Increase W = Wider, H = Taller.
$textBoxScriptsPath.Font = New-Object System.Drawing.Font("Agency FB", 16, [System.Drawing.FontStyle]::Bold)
$textBoxScriptsPath.BorderStyle = "None"
$textBoxScriptsPath.BackColor = $form.BackColor
$textBoxScriptsPath.ForeColor = [System.Drawing.Color]::Black
$textBoxScriptsPath.ReadOnly = $true
$textBoxScriptsPath.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center # Center alignment
$textBoxScriptsPath.Text = Get-ConfigScriptPath
$textBoxScriptsPath.CharacterCasing = "Upper"
$textBoxScriptsPath.Add_Click({
    Set-ScriptsPath
})
$form.Controls.Add($textBoxScriptsPath)

# Change Path Button
#$buttonChangePath = New-Object System.Windows.Forms.Button
#$buttonChangePath.Text = "Click here to change the Scripts Path"
#$buttonChangePath.Location = New-Object System.Drawing.Point(285, 30) # Increase X = Right, y = Down
#$buttonChangePath.Size = New-Object System.Drawing.Size(200, 15)      # Increase W = Wider, H = Taller.
#$buttonChangePath.BackColor = [System.Drawing.Color]::Gray
#$buttonChangePath.ForeColor = [System.Drawing.Color]::WhiteSmoke
#$buttonChangePath.FlatStyle = [Windows.Forms.FlatStyle]::System
# Change Path Event Handler
#$buttonChangePath.Add_Click({
    #Set-ScriptsPath
#})

#$form.Controls.Add($buttonChangePath)

#$buttonOpenFolder = New-Object System.Windows.Forms.Button
#$buttonOpenFolder.Text = "Open Folder"
#$buttonOpenFolder.Location = New-Object System.Drawing.Point(745, 10) # Modified location: Adjusted for longer textbox
#$buttonOpenFolder.Size = New-Object System.Drawing.Size(80, 25) # Modified size: 10% longer
#$buttonOpenFolder.BackColor = [System.Drawing.Color]::Gray
#$buttonOpenFolder.ForeColor = [System.Drawing.Color]::WhiteSmoke
#$buttonOpenFolder.FlatStyle = [Windows.Forms.FlatStyle]::System

# Open Folder Event Handler
#$buttonOpenFolder.Add_Click({
    #Invoke-Item -Path $textBoxScriptsPath.Text
#})

#$form.Controls.Add($buttonOpenFolder)

# Function to draw a rectangle
$form.Add_Paint({
    $graphics = $_.Graphics
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray, 1)
    $rectangle = New-Object System.Drawing.Rectangle(15, 35, 1035, 50) #Right,Down,Wide,Tall
    $graphics.DrawRectangle($pen, $rectangle)
})

# Change button sizes
$buttonsize = 90
# Define the gap between buttons
$buttonGap = 0
$buttonHeight = 30

# Load Button
$buttonLoadScripts = New-Object System.Windows.Forms.Button
$buttonLoadScripts.Text = "Load"
$buttonLoadScripts.Location = New-Object System.Drawing.Point(45, 46) # Increase X = Right, y = Down
$buttonLoadScripts.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)     # Increase W = Wider, H = Taller. 
#$buttonLoadScripts.BackColor = [System.Drawing.Color]::DarkBlue
$buttonLoadScripts.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonLoadScripts.FlatAppearance.BorderSize = 0
$buttonLoadScripts.BackColor = [System.Drawing.Color]::NavajoWhite
$buttonLoadScripts.ForeColor = [System.Drawing.Color]::Black

# Event handler for the Load Scripts button
$buttonLoadScripts.Add_Click({
    # Get the full path of the script directory
    $scriptDirectory = $textBoxScriptsPath.Text
    Clear-Host
    $listBoxScripts.Items.Clear()

    # Play the cool sound while loading scripts
    PlayCoolSound -SoundFileName "loadscripts.mp3"

    $richTextBoxOutput.Clear()
    $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Arial", 18, [System.Drawing.FontStyle]::Bold)
    #$richTextBoxOutput.SelectionColor = [System.Drawing.Color]::SaddleBrown
    $richTextBoxOutput.BackColor = [System.Drawing.Color]::SaddleBrown
    $richTextBoxOutput.ForeColor = [System.Drawing.Color]::CornSilk
    $richTextBoxOutput.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Center  # Center the text
    $richTextBoxOutput.AppendText("Welcome to PowerShell Script Manager $scriptVersionString`r`n`r`n")
    $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Helvetica", 12)
    #$richTextBoxOutput.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Left  # Left align the text
 
    $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::White
    $richTextBoxOutput.AppendText("PSM $scriptVersionString empowers users with various operations for script management. Each script will get it's own JSON, capturing essential details, including modifications, authors, versions, and timestamps.

With a built-in tagging system and keyword search, finding scripts is easier than ever. Efficiently search using keywords for streamlined script management and retrieval.")

    # Check if the script path exists
    $scriptPath = $textBoxScriptsPath.Text
    if (-not (Test-Path -Path $scriptPath)) {
        # Create a FolderBrowserDialog to let the user select the script path
        $scriptPath = Set-ScriptsPath
        if (-not (Test-Path -Path $scriptPath)) {
            return
        }
    }

    # Create a backup subdirectory if it doesn't exist
    $backupPath = Join-Path -Path $scriptPath -ChildPath "backup"
    if (-not (Test-Path -Path $backupPath)) {
        New-Item -Path $backupPath -ItemType Directory | Out-Null
    }

    # Create a default script and JSON file if they don't already exist
    Write-Host "The Scriptpath is $Scriptpath" -ForegroundColor Green
    CreateDefaultScriptAndJson -scriptPath $scriptPath

    # Load the scripts into the list box
    $listBoxScripts.Items.Clear()
    try {
        $scriptFiles = Get-ChildItem -Path $scriptPath -Filter "*.ps1" | Select-Object -ExpandProperty Name | Sort-Object
        foreach ($scriptFile in $scriptFiles) {
            $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($scriptFile).ToUpper()
            $listBoxScripts.Items.Add($scriptName)
        }
    } catch {
        Write-Error $_
        [System.Windows.Forms.MessageBox]::Show("An error occurred while loading the scripts: `n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($buttonLoadScripts)



################# keyword search area ###################

# Create a label for the search text box
#$searchLabel = New-Object System.Windows.Forms.Label
#$searchLabel.Location = New-Object System.Drawing.Point(865, 120) # Position it above the search text box
#$searchLabel.Size = New-Object System.Drawing.Size(180, 20) # Set the size
#$searchLabel.Font = New-Object System.Drawing.Font("Arial", 10)
#$searchLabel.Text = "Script Keyword Search"
#$searchLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
#$searchLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
#$form.Controls.Add($searchLabel)





################# keyword search area ###################

$labelDevelopedBy = New-Object System.Windows.Forms.Label
$labelDevelopedBy.Text = "Developed By: Curtis Dove, MCSA MCSE VCP CASP"
$labelDevelopedBy.Location = New-Object System.Drawing.Point(18, 715) # Increase X = Right, y = Down
$labelDevelopedBy.Size = New-Object System.Drawing.Size(375, 20) # Increase W = Wider, H = Taller.
$labelDevelopedBy.Font = New-Object System.Drawing.Font("Georgia", 8, ([System.Drawing.FontStyle]::Italic))
$labelDevelopedBy.ForeColor = [System.Drawing.Color]::Silver
$labelDevelopedBy.BackColor = [System.Drawing.Color]::MintCream
$labelDevelopedBy.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($labelDevelopedBy)

# Calculate the position of the new buttons
$newButtonX = $buttonLoadScripts.Left + $buttonLoadScripts.Width + $buttonGap
$newButtonY = $buttonLoadScripts.Top

# Edit Script Button
$buttonEditScript = New-Object System.Windows.Forms.Button
$buttonEditScript.Text = "Edit"
#$buttonEditScript.Location = New-Object System.Drawing.Point(132, 46) # Increase X = Right, y = Down
$buttonEditScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY) # Increase X = Right, y = Down
$buttonEditScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)      # Increase W = Wider, H = Taller.
#$buttonEditScript.BackColor = [System.Drawing.Color]::Orange
$buttonEditScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonEditScript.FlatAppearance.BorderSize = 0
$buttonEditScript.BackColor = [System.Drawing.Color]::PeachPuff
$buttonEditScript.ForeColor = [System.Drawing.Color]::Black

# Edit Script Event Handler
$buttonEditScript.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ($selectedScript + ".ps1")
        Start-Process powershell_ise -ArgumentList "`"$scriptPath`""
    }
})
$form.Controls.Add($buttonEditScript)

# Calculate the position of the new buttons
$newButtonX = $buttonEditScript.Left + $buttonEditScript.Width + $buttonGap
$newButtonY = $buttonEditScript.Top

# Copy Script Button
$buttoncopyScript = New-Object System.Windows.Forms.Button
$buttoncopyScript.Text = "Copy"
$buttoncopyScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY) # Increase X = Right, y = Down
$buttoncopyScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)      # Increase W = Wider, H = Taller.
#$buttoncopyScript.BackColor = [System.Drawing.Color]::LightBlue
$buttoncopyScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttoncopyScript.FlatAppearance.BorderSize = 0
$buttoncopyScript.BackColor = [System.Drawing.Color]::BurlyWood
$buttoncopyScript.ForeColor = [System.Drawing.Color]::Black

# Event handler for the Copy Script button
$buttonCopyScript.Add_Click({
    # Check if a script is selected in the list box
    if ($listBoxScripts.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select a script to perform a copy!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Get the selected script name
    $selectedScriptName = $listBoxScripts.SelectedItem.ToString()

    # Construct the source and destination paths based on the selected script
    $selectedScriptPath = Get-ScriptFilePath -scriptName $selectedScriptName
    $selectedJsonPath = Get-JsonFilePath -scriptName $selectedScriptName

    # Define the new script name
    $newScriptName = $selectedScriptName + "_Copy"
    $copyNumber = 2
    while (Test-Path (Get-ScriptFilePath -scriptName $newScriptName)) {
        $newScriptName = "$selectedScriptName" + "_Copy$copyNumber"
        $copyNumber++
    }

    # Construct the new script and JSON paths
    $newScriptPath = Get-ScriptFilePath -scriptName $newScriptName
    $newJsonPath = Get-JsonFilePath -scriptName $newScriptName

    # Copy the selected script and JSON to the new paths
    Copy-Item -Path $selectedScriptPath -Destination $newScriptPath
    Copy-Item -Path $selectedJsonPath -Destination $newJsonPath

    # Refresh the scripts in the list box
    Refresh-ScriptsInListBox
})

$form.Controls.Add($buttonCopyScript)

# Calculate the position of the new buttons
$newButtonX = $buttonCopyScript.Left + $buttonCopyScript.Width + $buttonGap
$newButtonY = $buttonCopyScript.Top

# Create Script Button
$buttonCreateScript = New-Object System.Windows.Forms.Button
$buttonCreateScript.Text = "Create"
#$buttonCreateScript.Location = New-Object System.Drawing.Point(254, 46) # Increase X = Right, y = Down
$buttonCreateScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY) # Increase X = Right, y = Down
$buttonCreateScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)      # Increase W = Wider, H = Taller.
#$buttonCreateScript.BackColor = [System.Drawing.Color]::LimeGreen
$buttonCreateScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonCreateScript.FlatAppearance.BorderSize = 0
$buttonCreateScript.BackColor = [System.Drawing.Color]::Peru
$buttonCreateScript.ForeColor = [System.Drawing.Color]::White
# Create Script Event Handler
$buttonCreateScript.Add_Click({

    $newScriptForm = New-Object System.Windows.Forms.Form
    $newScriptForm.Text = "Create New Script"
    $newScriptForm.Size = New-Object System.Drawing.Size(400, 760)        # Increase W = Wider, H = Taller.
    $newScriptForm.StartPosition = "CenterScreen"
    $newScriptForm.BackColor = [System.Drawing.Color]::Peru
    $newScriptForm.ForeColor = [System.Drawing.Color]::White

    $labelNewScriptName = New-Object System.Windows.Forms.Label
    $labelNewScriptName.Text = "Script Name:"
    $labelNewScriptName.Location = New-Object System.Drawing.Point(10, 10) # Increase X = Right, y = Down
    $labelNewScriptName.Size = New-Object System.Drawing.Size(120, 22)     # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($labelNewScriptName)

    $textBoxNewScriptName = New-Object System.Windows.Forms.TextBox
    $textBoxNewScriptName.Location = New-Object System.Drawing.Point(132, 10) # Increase X = Right, y = Down
    $textBoxNewScriptName.Size = New-Object System.Drawing.Size(230, 22)      # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($textBoxNewScriptName)

    # Add a ListBox to select tags
    $labelTags = New-Object System.Windows.Forms.Label
    $labelTags.Text = "Tags:"
    $labelTags.Location = New-Object System.Drawing.Point(35, 40) # Increase X = Right, y = Down
    $labelTags.Size = New-Object System.Drawing.Size(60, 22)     # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($labelTags)

    $tagsListBox = New-Object System.Windows.Forms.ListBox
    $tagsListBox.Name = "tagsListBox"
    $tagsListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiSimple  # Set to MultiSimple for easy multiple selection
    $tagsListBox.Location = New-Object System.Drawing.Point(132, 40) # Increase X = Right, y = Down
    $tagsListBox.Size = New-Object System.Drawing.Size(230, 540)     # Increase W = Wider, H = Taller.
    $tagsListBox.MultiColumn = $true
    $tagsListBox.ColumnWidth = 100
    $tagsListBox.BackColor = [System.Drawing.Color]::PapayaWhip
    $tagsListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $tagsListBox.Font = New-Object System.Drawing.Font("Arial", 10)

    # Add tags to the ListBox with custom formatting
    $tagsListBox.Items.AddRange(($tags | ForEach-Object { "$($_.Icon) $($_.Name)" }))


    $newScriptForm.Controls.Add($tagsListBox)

    # Add a TextBox for entering custom tags
    $labelCustomTag = New-Object System.Windows.Forms.Label
    $labelCustomTag.Text = "Custom Tag:"
    $labelCustomTag.Location = New-Object System.Drawing.Point(10, 600) # Increase X = Right, y = Down
    $labelCustomTag.Size = New-Object System.Drawing.Size(120, 22)      # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($labelCustomTag)

    $textBoxCustomTag = New-Object System.Windows.Forms.TextBox
    $textBoxCustomTag.Location = New-Object System.Drawing.Point(132, 600) # Increase X = Right, y = Down
    $textBoxCustomTag.Size = New-Object System.Drawing.Size(230, 22)       # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($textBoxCustomTag)

    # Add a Button to add the custom tag to the ListBox
    $buttonAddCustomTag = New-Object System.Windows.Forms.Button
    $buttonAddCustomTag.Text = "Add Tag"
    $buttonAddCustomTag.Location = New-Object System.Drawing.Point(10, 630) # Increase X = Right, y = Down
    $buttonAddCustomTag.Size = New-Object System.Drawing.Size(120, 33)      # Increase W = Wider, H = Taller.
    $buttonAddCustomTag.Add_Click({
    $customTag = $textBoxCustomTag.Text
        if (-not [string]::IsNullOrWhiteSpace($customTag)) {
            $tagsListBox.Items.Add($customTag)
            $textBoxCustomTag.Text = ""
        }
    })
    $newScriptForm.Controls.Add($buttonAddCustomTag)

    $buttonCreateNewScript = New-Object System.Windows.Forms.Button
    $buttonCreateNewScript.Text = "Create"
    $buttonCreateNewScript.Location = New-Object System.Drawing.Point(10, 670) # Increase X = Right, y = Down
    $buttonCreateNewScript.Size = New-Object System.Drawing.Size($buttonsize, 33)      # Increase W = Wider, H = Taller.
    $buttonCreateNewScript.Add_Click({
    $newScriptName = $textBoxNewScriptName.Text

    if (-not [string]::IsNullOrWhiteSpace($newScriptName)) {
        $newScriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ($newScriptName + ".ps1")
        $newJsonPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ($newScriptName + ".json")

        if (-not (Test-Path -Path $newScriptPath) -and -not (Test-Path -Path $newJsonPath)) {
            $newScriptContent = @"
# $newScriptName.ps1

Write-Host "Hello, $newScriptName!"
"@
            $newScriptContent | Out-File -FilePath $newScriptPath -Encoding UTF8

            $selectedTags = $tagsListBox.SelectedItems | ForEach-Object { $_.ToString() -replace '^.*\s', '' }  # Extract tag name from the custom format

            # Update the JSON content for the new script
            $newScriptJson = @{
                "Name" = $selectedScript
                "ScriptPath" = $newScriptPath
                "ScriptJson" = $newJsonPath
                "Author" = $env:username.ToUpper()
                "Description" = "New Script $selectedScript"
                "Version" = "0.1"
                "Tags" = @("📜 PowerShell")
                "ModifiedBy" = $env:username.ToUpper()
                "Keywords" = @()
                "Modifications" = @(
                    @{
                        "Date" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
                        "Version" = "0.1"
                        "ModifiedBy" = $env:username.ToUpper()
                        "Modification" = "Created $($selectedScript) a JSON Buddy"
                    }
                )
                "ToDoList" = @(
                    @{
                        "Title" = "Review Script"
                        "Task" = "Review Script for errors and Update!"
                        "DateCompleted" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
                        "Status" = ([TaskStatus]::NotStarted).ToString()
                        "CompletedBy" = ""
                    }
                )
            } | ConvertTo-Json -Depth 4

            $newScriptJson | Out-File -FilePath $newJsonPath -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show("New script '$newScriptName' created successfully!", "Create Script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $newScriptForm.Close()
            $listBoxScripts.Items.Add($newScriptName.ToUpper())
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("A script with the name '$newScriptName' already exists.", "Create Script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a script name.", "Create Script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})




    $newScriptForm.Controls.Add($buttonCreateNewScript)

    $newScriptForm.ShowDialog()
})
$form.Controls.Add($buttonCreateScript)

# Calculate the position of the new buttons
$newButtonX = $buttonCreateScript.Left + $buttonCreateScript.Width + $buttonGap
$newButtonY = $buttonCreateScript.Top

# Rename Script Button
$buttonRenameScript = New-Object System.Windows.Forms.Button
$buttonRenameScript.Text = "Rename"
#$buttonRenameScript.Location = New-Object System.Drawing.Point(377, 46)
$buttonRenameScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY)
$buttonRenameScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)
$buttonRenameScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRenameScript.FlatAppearance.BorderSize = 0
$buttonRenameScript.BackColor = [System.Drawing.Color]::Chocolate
$buttonRenameScript.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($buttonRenameScript)

# Rename Script Event Handler
$buttonRenameScript.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ($selectedScript + ".ps1")
        Write-Host "Selected Script: $selectedScript"
        Write-Host "Script Path: $scriptPath"

        # Show a MessageBox with an input field
        $newName = ""
        $inputBoxForm = New-Object System.Windows.Forms.Form
        $inputBoxForm.Width = 275
        $inputBoxForm.Height = 175
        $inputBoxForm.Text = "Rename Script"
        $inputBoxForm.Font = New-Object System.Drawing.Font("Segoe UI", 12)  # Set the font and size
        $inputBoxForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $inputBoxForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $inputBoxForm.BackColor = [System.Drawing.Color]::BurlyWood   # Set the background color
        $inputBoxForm.ForeColor = [System.Drawing.Color]::Black  # Set the foreground (text) color

        $inputLabel = New-Object System.Windows.Forms.Label
        $inputLabel.Text = "Enter the new name for the script:"
        $inputLabel.Location = New-Object System.Drawing.Point(10, 20) # Increase X = Right, y = Down
        $inputLabel.Size = New-Object System.Drawing.Size(250, 22)     # Increase W = Wider, H = Taller.
        $inputBoxForm.Controls.Add($inputLabel)

        $inputTextBox = New-Object System.Windows.Forms.TextBox
        $inputTextBox.Location = New-Object System.Drawing.Point(10, 50) # Increase X = Right, y = Down
        $inputTextBox.Size = New-Object System.Drawing.Size(250, 30)     # Increase W = Wider, H = Taller.
        $inputTextBox.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper  # Force text to be uppercase
        $inputTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None 
        $inputBoxForm.Controls.Add($inputTextBox)

        $inputButtonOK = New-Object System.Windows.Forms.Button
        $inputButtonOK.Text = "OK"
        $inputButtonOK.Location = New-Object System.Drawing.Point(10, 80) # Increase X = Right, y = Down
        $inputButtonOK.Size = New-Object System.Drawing.Size(250, 30)     # Increase W = Wider, H = Taller.
        $inputButtonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $inputButtonOK.BackColor = [System.Drawing.Color]::SaddleBrown
        $inputButtonOK.ForeColor = [System.Drawing.Color]::White
        $inputBoxForm.Controls.Add($inputButtonOK)

        $inputBoxForm.AcceptButton = $inputButtonOK

        $result = $inputBoxForm.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $newName = $inputTextBox.Text
        }

        Write-Host "New Name: $newName"

        if (-not [string]::IsNullOrWhiteSpace($newName)) {
            $scriptDirectory = Split-Path -Path $scriptPath -Parent
            $newScriptPath = Join-Path -Path $scriptDirectory -ChildPath ($newName + ".ps1")
            $newJsonPath = Join-Path -Path $scriptDirectory -ChildPath ($newName + ".json")

            Write-Host "Script Path: $scriptPath"
            Write-Host "Old JSON Path: $oldJsonPath"
            Write-Host "New Script Path: $newScriptPath"
            Write-Host "New JSON Path: $newJsonPath"

            if (-not (Test-Path -Path $newScriptPath) -and -not (Test-Path -Path $newJsonPath)) {
                Write-Host "Renaming Script..."
                Rename-Item -Path $scriptPath -NewName ($newName + ".ps1") -ErrorAction SilentlyContinue

                # Update corresponding JSON file
                $oldJsonPath = ($scriptPath -replace "\.ps1$", ".json")
                if (Test-Path -Path $oldJsonPath) {
                    Write-Host "Renaming JSON..."
                    Rename-Item -Path $oldJsonPath -NewName ($newName + ".json") -ErrorAction SilentlyContinue

                    # Update the script name in the JSON file
                    Write-Host "Updating JSON Content..."
                    $jsonContent = Get-Content -Path $newJsonPath | ConvertFrom-Json

                    # Increment the version number by 0.1
                    # Assuming the version number is in the Version property of the JSON content
                    $currentVersion = [double]::Parse($jsonContent.Version)
                    $newVersion = $currentVersion + 0.1
                    $jsonContent.Version = $newVersion.ToString("F1")  # Formats it with one decimal place

                    # Update the script name within the Modifications array
                    $currentDate = Get-Date
                    $currentDateFormatted = $currentDate.ToString("MMMM d, yyyy 'at' h:mm tt")
                    $newModification = @{
                        "Date" = $currentDateFormatted
                        "Modification" = "Renamed Script and JSON from $selectedScript to $newName"
                        "ModifiedBy" = $env:USERNAME.ToUpper()
                        "Version" = $newVersion  # Include the new version number in the modification entry
                    }
                    $jsonContent.Modifications += $newModification

                    $jsonContent.Name = $newName
                    $jsonContent | ConvertTo-Json -Depth 4 | Set-Content -Path $newJsonPath -Encoding UTF8
                }

                $listBoxScripts.Items[$listBoxScripts.SelectedIndex] = $newName.ToUpper()
                [System.Windows.Forms.MessageBox]::Show("Script renamed successfully!", "Rename Script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                # Refresh the form
                #PopulateFieldsAndTags
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("A script with the name '$newName' already exists.", "Rename Script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    }
})
$form.Controls.Add($buttonRenameScript)

# Calculate the position of the new buttons
$newButtonX = $buttonRenameScript.Left + $buttonRenameScript.Width + $buttonGap
$newButtonY = $buttonRenameScript.Top

# Delete Script Button
$buttonDeleteScript = New-Object System.Windows.Forms.Button
$buttonDeleteScript.Text = "Delete"
#$buttonDeleteScript.Location = New-Object System.Drawing.Point(499, 46) # Increase X = Right, y = Down
$buttonDeleteScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY) # Increase X = Right, y = Down
$buttonDeleteScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)      # Increase W = Wider, H = Taller.
$buttonDeleteScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonDeleteScript.FlatAppearance.BorderSize = 0
$buttonDeleteScript.BackColor = [System.Drawing.Color]::FireBrick
$buttonDeleteScript.ForeColor = [System.Drawing.Color]::White

# Event handler for the Delete Script button
$buttonDeleteScript.Add_Click({
    # Get the name of the selected script from the ListBox
    $selectedScript = $listBoxScripts.SelectedItem

    if ($selectedScript) {
        # Construct the paths for the script and JSON file based on the selected script name
        $scriptPathWithoutExtension = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        $scriptPath = $scriptPathWithoutExtension + ".ps1"
        $jsonPath = $scriptPathWithoutExtension + ".json"

        # Delete the script file
        if (Test-Path -Path $scriptPath) {
            Remove-Item -Path $scriptPath -Force
        }

        # Delete the JSON file
        if (Test-Path -Path $jsonPath) {
            Remove-Item -Path $jsonPath -Force
        }

         # Play a cool sound
        $soundFilePath = "$Scriptspath\sounds\deletescript.mp3"
        #PlayCoolSound -SoundFile $soundFilePath
        PreloadCoolSound -SoundFile "deletescript.mp3"
        # Load the scripts again to update the ListBox
        $listBoxScripts.Items.Clear()
        try {
            $scriptFiles = Get-ChildItem -Path $textBoxScriptsPath.Text -Filter "*.ps1" | Select-Object -ExpandProperty Name | Sort-Object
            foreach ($scriptFile in $scriptFiles) {
                $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($scriptFile).ToUpper()
                $listBoxScripts.Items.Add($scriptName)
            }
        } catch {
            Write-Error $_
            [System.Windows.Forms.MessageBox]::Show("An error occurred while loading the scripts: `n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }

        # Clear the form fields
        ClearForm

        # Clear the selected tags in the ListBox
        $listBoxScriptTags.ClearSelected()
    }
})
$form.Controls.Add($buttonDeleteScript)

# Calculate the position of the new buttons
$newButtonX = $buttonDeleteScript.Left + $buttonDeleteScript.Width + $buttonGap
$newButtonY = $buttonDeleteScript.Top

# Run Script Button
$buttonRunScript = New-Object System.Windows.Forms.Button
$buttonRunScript.Text = "Run"
#$buttonRunScript.Location = New-Object System.Drawing.Point(622, 46) # Increase X = Right, y = Down
$buttonRunScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY) # Increase X = Right, y = Down
$buttonRunScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)      # Increase W = Wider, H = Taller.
$buttonRunScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRunScript.FlatAppearance.BorderSize = 0
$buttonRunScript.BackColor = [System.Drawing.Color]::MediumAquaMarine
$buttonRunScript.ForeColor = [System.Drawing.Color]::Black

# Run Script Event Handler
$buttonRunScript.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptName = $selectedScript -replace '\.ps1$',''
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ("{0}.ps1" -f $selectedScript)
        Write-Host "Starting Script In Powershell 7: $ScriptPath" -ForegroundColor Green
        Start-Process pwsh -ArgumentList "-File `"$scriptPath`""
    }
    else {
        Write-Host "No Script found at: $ScriptPath" -ForegroundColor Yellow
    }
})
$form.Controls.Add($buttonRunScript)

# Create the backup counter label
$labelBackupCounter = New-Object System.Windows.Forms.Label
$labelBackupCounter.Text = "Backups:"
$labelBackupCounter.Location = New-Object System.Drawing.Point(330, 125) # Increase X = Right, y = Down
$labelBackupCounter.size = New-Object System.Drawing.Size(50, 15)
#$labelBackupCounter.Size = $true
$labelBackupCounter.ForeColor = [System.Drawing.Color]::Black
# Creating on on-click affect to out-grid the backup data
$labelBackupCounter.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptName = $selectedScript -replace '\.ps1$',''
        $backupFolderPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ("Backup")
        
        if (-not (Test-Path -Path $backupFolderPath)) {
            [System.Windows.Forms.MessageBox]::Show("Backup folder does not exist.", "Backup Folder", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $backupFiles = Get-ChildItem -Path $backupFolderPath -File | Where-Object { $_.Name -like "${scriptName}_backup*.*" } | Select-Object FullName, LastWriteTime, Length, @{Name="BackupNumber"; Expression={if ($_.Name -match "_backup(\d+)") {[int]$Matches[1]}}}

        if ($backupFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Sorry, there are no backups for this script", "No Backups", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $formattedFiles = $backupFiles | Select-Object @{Name="Path"; Expression={ $_.FullName }},
                                                      @{Name="Date"; Expression={ $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm") }},
                                                      @{Name="Size (KB)"; Expression={ [math]::Round($_.Length / 1KB, 2) }},
                                                      BackupNumber

        $selectedBackup = $formattedFiles | Out-GridView -Title "Backup Files for $scriptName" -PassThru

        if ($selectedBackup) {
            $selectedPath = $selectedBackup.Path
            $restorePath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ([System.IO.Path]::GetFileName($selectedPath))
            Move-Item -Path $selectedPath -Destination $restorePath -Force
            
            # Restore JSON file if exists
            $jsonPath = $selectedPath -replace [regex]::Escape(".ps1"), ".json"
            if (Test-Path -Path $jsonPath) {
                $restoreJsonPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ([System.IO.Path]::GetFileName($jsonPath))
                Move-Item -Path $jsonPath -Destination $restoreJsonPath -Force
            }
            
            [System.Windows.Forms.MessageBox]::Show("Backup restored successfully!", "Restore Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No script selected.", "Backup Script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})


# Backup Counter Value
$labelBackupCountValue = New-Object System.Windows.Forms.Label
$labelBackupCountValue.Text = "0"
$labelBackupCountValue.Location = New-Object System.Drawing.Point(380, 125)
$labelBackupCountValue.Size = New-Object System.Drawing.Size(30, 15)
$labelBackupCountValue.ForeColor = [System.Drawing.Color]::Red

# Calculate the position of the new buttons
$newButtonX = $buttonRunScript.Left + $buttonRunScript.Width + $buttonGap
$newButtonY = $buttonRunScript.Top

# Backup Script Button
$buttonBackupScript = New-Object System.Windows.Forms.Button
$buttonBackupScript.Text = "Backup"
#$buttonBackupScript.Location = New-Object System.Drawing.Point(745, 46)
$buttonBackupScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY)
$buttonBackupScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)
$buttonBackupScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonBackupScript.FlatAppearance.BorderSize = 0
$buttonBackupScript.BackColor = [System.Drawing.Color]::SteelBlue
$buttonBackupScript.ForeColor = [System.Drawing.Color]::White

# Backup Script Event Handler
$buttonBackupScript.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptName = $selectedScript -replace '\.ps1$',''
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ("{0}.ps1" -f $selectedScript)
        $backupFolderPath = Join-Path -Path (Split-Path -Path $scriptPath -Parent) -ChildPath "Backup"
        
        if (-not (Test-Path -Path $backupFolderPath)) {
            New-Item -ItemType Directory -Path $backupFolderPath | Out-Null
        }
        if (-not (Test-Path -Path $scriptPath)) {
            Write-Host "Script path does not exist:" -NoNewline -ForegroundColor Green
            Write-Host " $scriptPath" -ForegroundColor Cyan
            return
        }

        $scriptExtension = ".ps1"
        $jsonExtension = ".json"

        # Get the list of backup files
        $backupFiles = Get-ChildItem -Path $backupFolderPath -File | Where-Object { $_.Name -like "$scriptName_backup*.*" }

        # Find the maximum backup number
        $maxBackupNumber = 0
        foreach ($file in $backupFiles) {
            if ($file.Name -match "_backup(\d+)") {
                $backupNumber = [int]$Matches[1]
                if ($backupNumber -gt $maxBackupNumber) {
                    $maxBackupNumber = $backupNumber
                }
            }
        }
        $maxBackupNumber += 1
        $backupScriptPath = Join-Path -Path $backupFolderPath -ChildPath ($scriptName + "_backup" + $maxBackupNumber.ToString("D2") + $scriptExtension)
        Copy-Item -Path $scriptPath -Destination $backupScriptPath

        # Check if a corresponding JSON file exists and copy it if it does
        $jsonPath = ($scriptPath -replace [regex]::Escape($scriptExtension), $jsonExtension)
        if (Test-Path -Path $jsonPath) {
            $backupJsonPath = Join-Path -Path $backupFolderPath -ChildPath ($scriptName + "_backup" + $maxBackupNumber.ToString("D2") + $jsonExtension)
            Copy-Item -Path $jsonPath -Destination $backupJsonPath
        }
        
        #[System.Windows.Forms.MessageBox]::Show("Script and JSON file backed up successfully!", "Backup Script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        # Refresh the backup counter
        Update-BackupCounter -ScriptName $scriptName -BackupFolderPath $backupFolderPath -LabelBackupCountValue $labelBackupCountValue

        # Fetch the latest backup files after the new backups are created
        $backupFiles = Get-ChildItem -Path $backupFolderPath -File | Where-Object { $_.Name -match "^$scriptName`_backup\d+\..+" }


        # Display the backup information in the specified format
        $backupCount = $backupFiles.Count
        if ($backupCount -gt 0) {
            Write-Host "$($env:username.toupper()) backed up the following scripts:" -ForegroundColor Green
            $latestBackupFiles = $backupFiles | Sort-Object Name -Descending | Select-Object -First 2
            $latestBackupFiles | ForEach-Object {
                $backupScriptName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                $backupScriptExt = [System.IO.Path]::GetExtension($_.Name)
                Write-Host "  - $backupScriptName$backupScriptExt" -ForegroundColor Cyan
            }
        } else {
            Write-Host "$($env:username.toupper()) did not back up any scripts." -ForegroundColor Yellow
        }

        $soundFilePath = "$Scriptspath\sounds\backupscript.mp3"
        PreloadCoolSound -SoundFileName "backupscript.mp3"

        $form.Refresh()
    } else {
        [System.Windows.Forms.MessageBox]::Show("No script selected.", "Backup Script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$form.Controls.Add($buttonBackupScript)

# Calculate the position of the new buttons
$newButtonX = $buttonBackupScript.Left + $buttonBackupScript.Width + $buttonGap
$newButtonY = $buttonBackupScript.Top

# Archive Script Button
$buttonArchiveScript = New-Object System.Windows.Forms.Button
$buttonArchiveScript.Text = "Archive"
#$buttonArchiveScript.Location = New-Object System.Drawing.Point(745, 46)
$buttonArchiveScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY)
$buttonArchiveScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)
$buttonArchiveScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonArchiveScript.FlatAppearance.BorderSize = 0
$buttonArchiveScript.BackColor = [System.Drawing.Color]::DarkSlateBlue
$buttonArchiveScript.ForeColor = [System.Drawing.Color]::White
$buttonArchiveScript.Add_Click({
    # Check if a script is selected
    if ($listBoxScripts.SelectedIndex -eq -1) {
        [System.Windows.Forms.MessageBox]::Show("Please select a script to perform an archive!", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $selectedScriptName = $listBoxScripts.SelectedItem.ToString()
    $selectedScriptPath = Get-ScriptFilePath -scriptName $selectedScriptName
    $selectedJsonPath = Get-JsonFilePath -scriptName $selectedScriptName
    $backupFolderPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "backup"
    
    $newScriptPath = Join-Path -Path $backupFolderPath -ChildPath "$selectedScriptName.ps1"
    $newJsonPath = Join-Path -Path $backupFolderPath -ChildPath "$selectedScriptName.json"

    # Move the selected script and JSON to the backup folder
    try {
        Move-Item -Path $selectedScriptPath -Destination $newScriptPath -Force
        Move-Item -Path $selectedJsonPath -Destination $newJsonPath -Force

        [System.Windows.Forms.MessageBox]::Show("Script '$selectedScriptName' has been archived.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Refresh the scripts in the list box
        Refresh-ScriptsInListBox
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred while archiving the script: `n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.Controls.Add($buttonArchiveScript)

# Calculate the position of the new buttons
$newButtonX = $buttonArchiveScript.Left + $buttonArchiveScript.Width + $buttonGap
$newButtonY = $buttonArchiveScript.Top

# Move Script Button
$buttonMoveScript = New-Object System.Windows.Forms.Button
$buttonMoveScript.Text = "Move"
#$buttonMoveScript.Location = New-Object System.Drawing.Point(745, 46)
$buttonMoveScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY)
$buttonMoveScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)
$buttonMoveScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonMoveScript.FlatAppearance.BorderSize = 0
$buttonMoveScript.BackColor = [System.Drawing.Color]::MidnightBlue
$buttonMoveScript.ForeColor = [System.Drawing.Color]::White
# Event Handler to Move Script & JSON to new location
$buttonMoveScript.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem # Get selected item
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ($selectedScript + '.ps1')

        # Create a new FolderBrowserDialog and set the SelectedPath to the root of the drive where the files are located
        $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowserDialog.SelectedPath = [System.IO.Path]::GetPathRoot($scriptPath) # Set to the root of the drive
        $folderBrowserDialog.Description = "Select the destination folder"

        $dialogResult = $folderBrowserDialog.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $destinationPath = $folderBrowserDialog.SelectedPath
            $jsonPath = $scriptPath -replace '\.ps1$', '.json'

            # Check if paths exist
            if ((Test-Path -Path $scriptPath) -and (Test-Path -Path $jsonPath)) {
                Move-Item -Path $scriptPath -Destination $destinationPath
                Move-Item -Path $jsonPath -Destination $destinationPath
            } else {
                [System.Windows.Forms.MessageBox]::Show("The script or JSON file was not found. Script Path: $scriptPath, JSON Path: $jsonPath")
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a script to move.")
    }
})





$form.Controls.Add($buttonMoveScript)

# Calculate the position of the new buttons
$newButtonX = $buttonMoveScript.Left + $buttonMoveScript.Width + $buttonGap
$newButtonY = $buttonMoveScript.Top

# Open Path Script Button
$buttonOpenPathScript = New-Object System.Windows.Forms.Button
$buttonOpenPathScript.Text = "Path"
#$buttonOpenPathScript.Location = New-Object System.Drawing.Point(745, 46)
$buttonOpenPathScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY)
$buttonOpenPathScript.Size = New-Object System.Drawing.Size($buttonsize, $buttonHeight)
$buttonOpenPathScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonOpenPathScript.FlatAppearance.BorderSize = 0
$buttonOpenPathScript.BackColor = [System.Drawing.Color]::Black
$buttonOpenPathScript.ForeColor = [System.Drawing.Color]::White
# Open Folder Event Handler
$buttonOpenPathScript.Add_Click({
    Invoke-Item -Path $textBoxScriptsPath.Text
})
$form.Controls.Add($buttonOpenPathScript)

# Update the backup counter initially
Update-BackupCounter

# Form Click Event Handler
$form.Add_Click({
    Update-BackupCounter
    #PopulateFieldsAndTags
    $form.Refresh()
})

# Scripts ListBox label
$labelScripts = New-Object System.Windows.Forms.Label
$labelScripts.Text = "Scripts Listbox:"
$labelScripts.Location = New-Object System.Drawing.Point(13, 125) # Increase X = Right, y = Down
$labelScripts.Size = New-Object System.Drawing.Size(150, 15)      # Increase W = Wider, H = Taller.
$labelScripts.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($labelScripts)

# Create the scripts ListBox
$listBoxScripts = New-Object System.Windows.Forms.ListBox
$listBoxScripts.Name = "listBoxScripts"
$listBoxScripts.Location = New-Object System.Drawing.Point(15, 140) # Increase X = Right, Y = Down
$listBoxScripts.Size = New-Object System.Drawing.Size(380, 630)   # Increase W = Wider, H = Taller.
$listBoxScripts.SelectionMode = "One"
$listBoxScripts.Font = New-Object System.Drawing.Font("Arial", 10)
$listBoxScripts.backcolor = [System.Drawing.Color]::MintCream
$form.Controls.Add($listBoxScripts)

# Create a progress bar
#$progressBar = New-Object Windows.Forms.ProgressBar
#$progressBar.Location = New-Object Drawing.Point(15, 460)
#$progressBar.Size = New-Object Drawing.Size(320, 20)
#$form.Controls.Add($progressBar)

# Commit Script & JSON to GIT Button
$buttonCommitToGit = New-Object System.Windows.Forms.Button
$buttonCommitToGit.Text = "Commit Script and JSON to GIT"
$buttonCommitToGit.Location = New-Object System.Drawing.Point(15,737)  # Increase X = Right, Y = Down
$buttonCommitToGit.Size = New-Object System.Drawing.Size(380, 25)       # Increase W = Wider, H = Taller.
$buttonCommitToGit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonCommitToGit.FlatAppearance.BorderSize = 1
$buttonCommitToGit.BackColor = [System.Drawing.Color]::BurlyWood
$buttonCommitToGit.ForeColor = [System.Drawing.Color]::Black
$buttonCommitToGit.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Italic)

# Commit Script & JSON to GIT Event Handler
$buttonCommitToGit.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPathWithoutExtension = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ("$selectedScript.ps1")
        $jsonPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath ("$selectedScript.json")

        # Ensure both the script and JSON file exist
        if (-not (Test-Path -Path $scriptPath)) {
            Write-Host "Script path does not exist: $scriptPath" -ForegroundColor Red
            return
        }

        if (-not (Test-Path -Path $jsonPath)) {
            Write-Host "JSON file path does not exist: $jsonPath" -ForegroundColor Red
            return
        }
        $gitRepoPath = $global:BaseScriptPath
        Set-Location -Path $gitRepoPath

        # Initialize the Git repository if it does not exist
        if (-not (Test-Path -Path (Join-Path -Path $gitRepoPath -ChildPath ".git"))) {
            git init
        }

                # Set the remote repository if not already set
        $remoteUrl = "https://github.com/dm10169/PSM.git"
        git remote add origin $remoteUrl -m "master"

        # Fetch the latest changes from the remote repository
        git fetch origin

        # Ensure the repository is checked out to the master branch
        git checkout master

        # Add the specific files to the staging area
        git add -A -- $scriptPath $jsonPath

        # Commit the changes
        $commitMessage = "Committed script $selectedScript and its JSON file"
        git commit -m $commitMessage

        # Push the changes to the remote repository on the master branch
        $pushOutput = git push origin master 2>&1

        Start-Process "https://github.com/dm10169/PSM"

        # Check if the push was successful or if everything is up-to-date
        if ($pushOutput -match "Everything up-to-date") {
            Write-Host "No new changes to push. Repository is up-to-date." -ForegroundColor Green
        } elseif ($pushOutput -match "pushed") {
            Write-Host "Changes pushed to remote repository on GitHub." -ForegroundColor Green
        } else {
            Write-Host "Failed to push changes to remote repository on GitHub." -ForegroundColor Red
        }
    } else {
        Write-Host "No script selected." -ForegroundColor Yellow
    }
})
$Form.Controls.Add($buttonCommitToGit)



## Create a horizontal separator line
##$horizontalLine = New-Object System.Windows.Forms.Label
#$horizontalLine.Location = New-Object System.Drawing.Point(15, 470)
#$horizontalLine.Size = New-Object System.Drawing.Size(835, 2)     # Set the width and height of the line.
#$horizontalLine.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
#$horizontalLine.ForeColor = [System.Drawing.Color]::LightGray 
#$form.Controls.Add($horizontalLine)

# Create the tags ListBox
$tagsListBox = New-Object System.Windows.Forms.ListBox
$tagsListBox.Name = "tagsListBox"
$tagsListBox.SelectionMode = "MultiExtended"
$tagsListBox.Location = New-Object System.Drawing.Point(132, 40) # Increase X = Right, y = Down
$tagsListBox.Size = New-Object System.Drawing.Size(320, 540)     # Increase W = Wider, H = Taller.
$tagsListBox.MultiColumn = $true
$tagsListBox.ColumnWidth = 100
$tagsListBox.BackColor = [System.Drawing.Color]::LightGray
$tagsListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$tagsListBox.Font = New-Object System.Drawing.Font("Arial", 12)

try {
    $tagsListBox.Items.AddRange($tags)
} catch {
    Write-Error $_
}

################################ FIELDS ON RIGHT HAND SIDE OF FARM START HERE #############################

###################################### NAME LABEL & TEXT BOX  #############################################

# Name Label
#$labelName = New-Object System.Windows.Forms.Label
#$labelName.Text = "Script Name:"
#$labelName.Location = New-Object System.Drawing.Point(405, 100)   # Increase X = Right, y = Down
#$labelName.Size = New-Object System.Drawing.Size(100, 20)         # Increase W = Wider, H = Taller.
#$form.Controls.Add($labelName)

# Name Text Box
$textBoxName = New-Object System.Windows.Forms.TextBox
$textBoxName.Location = New-Object System.Drawing.Point(405, 95) # Increase X = Right, y = Down
$textBoxName.Size = New-Object System.Drawing.Size(330, 160)      # Increase W = Wider, H = Taller.
$textBoxName.Font = New-Object System.Drawing.Font("Agency FB", 11, [System.Drawing.FontStyle]::Bold)
$textBoxName.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Left # Center the text
#$textBoxName.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
#$textBoxName.FlatAppearance.BorderSize = 0
$textBoxName.BorderStyle = "None"
#$textBoxName.BorderStyle = "Fixed3D"
$textBoxName.BackColor = $form.BackColor
#$textBoxName.BackColor = "MintCream"
$textBoxName.ForeColor = [System.Drawing.Color]::SaddleBrown
$textBoxName.CharacterCasing = "Upper"
$form.Controls.Add($textBoxName)

###################################### NAME LABEL & TEXT BOX  #############################################

################################# DESCRIPTION LABEL & TEXT BOX  ###########################################

# Label Description Label
$labelDescription = New-Object System.Windows.Forms.Label
$labelDescription.Text = "|Description|"
$labelDescription.Location = New-Object System.Drawing.Point(405, 125) # Increase X = Right, y = Down
$labelDescription.Size = New-Object System.Drawing.Size(75, 15)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelDescription)

# Label Description TextBox
$textBoxDescription = New-Object System.Windows.Forms.TextBox
$textBoxDescription.Location = New-Object System.Drawing.Point(405, 140) # Increase X = Right, y = Down
$textBoxDescription.Size = New-Object System.Drawing.Size(330, 80)       # Increase W = Wider, H = Taller.
$textBoxDescription.Multiline = $true
$textBoxDescription.CharacterCasing = "Upper"
$textBoxDescription.Backcolor = "MintCream"
$textBoxDescription.Add_Leave({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$selectedScript.ps1"
        $jsonPath = ($scriptPath -replace "\.ps1$", ".json")

        if (Test-Path -Path $jsonPath) {
            try {
                $jsonContent = Get-Content -Path $jsonPath -Raw -ErrorAction Stop | ConvertFrom-Json

                # Save the previous value of Description for comparison
                $previousDescription = $jsonContent.Description

                # Update the Description field with the new value
                $jsonContent.Description = $textBoxDescription.Text

                # Save the updated JSON content back to the file
                $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $jsonPath -Encoding UTF8

                # Clear the RichTextBoxOutput and display the updated Description field data
                $richTextBoxOutput.Clear()
                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::PapayaWhip
                $richTextBoxOutput.AppendText("Description for script '$selectedScript' updated successfully!" + [Environment]::NewLine)

                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Silver

                $descriptionData = @{
                    Description = @{
                        OldValue = $previousDescription
                        NewValue = $textBoxDescription.Text
                    }
                } | ConvertTo-Json -Depth 4

                $richTextBoxOutput.AppendText($descriptionData)
            }
            catch {
                $errorMessage = "Failed to read or update the JSON file: $jsonPath`nError: $_"
                Add-Content -Path "error.log" -Value $errorMessage

                # Show an error message to the user
                [System.Windows.Forms.MessageBox]::Show("An error occurred while updating the Description field: `n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        else {
            $warningMessage = "No JSON file exists for the selected script: $jsonPath"
            Add-Content -Path "error.log" -Value $warningMessage

            # Show a warning message to the user
            [System.Windows.Forms.MessageBox]::Show("No JSON file exists for the selected script.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    else {
        $infoMessage = "Please select a Script to save its JSON file."

        # Show an info message to the user
        [System.Windows.Forms.MessageBox]::Show("Please select a Script to save its JSON file.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$form.Controls.Add($textBoxDescription)

# Label Version
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "|Ver|"
$labelVersion.Location = New-Object System.Drawing.Point(405, 230) # Increase X = Right, y = Down
$labelVersion.Size = New-Object System.Drawing.Size(30, 15)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelVersion)

# TextBox Version
$textBoxVersion = New-Object System.Windows.Forms.TextBox
$textBoxVersion.Location = New-Object System.Drawing.Point(405, 245) # Increase X = Right, y = Down
$textBoxVersion.Size = New-Object System.Drawing.Size(30, 20)       # Increase W = Wider, H = Taller.
$textBoxVersion.Backcolor = "MintCream"
$textBoxVersion.Add_Leave({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$selectedScript.ps1"
        $jsonPath = ($scriptPath -replace "\.ps1$", ".json")

        if ( Test-Path -Path $jsonPath ) {
            try {
                $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

                # Save the previous value of Version for comparison
                $previousVersion = $jsonContent.Version

                # Update the Version field with the new value
                $jsonContent.Version = $textBoxVersion.Text

                # Save the updated JSON content back to the file
                $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $jsonPath -Encoding UTF8

                # Clear the RichTextBoxOutput and display the updated Version field data
                $richTextBoxOutput.Clear()
                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::PapayaWhip
                $richTextBoxOutput.AppendText("Version for script '$selectedScript' updated successfully!" + [Environment]::NewLine)

                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Silver

                $versionData = @{
                    Version = @{
                        OldValue = $previousVersion
                        NewValue = $textBoxVersion.Text
                    }
                } | ConvertTo-Json -Depth 4

                $richTextBoxOutput.AppendText($versionData)
            }
            catch {
                $errorMessage = "Failed to read or update the JSON file: $jsonPath`nError: $_"
                Add-Content -Path "error.log" -Value $errorMessage

                # Show an error message to the user
                [System.Windows.Forms.MessageBox]::Show("An error occurred while updating the Version field: `n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        else {
            $warningMessage = "No JSON file exists for the selected script: $jsonPath"
            Add-Content -Path "error.log" -Value $warningMessage

            # Show a warning message to the user
            [System.Windows.Forms.MessageBox]::Show("No JSON file exists for the selected script.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    else {
        $infoMessage = "Please select a Script to save its JSON file."

        # Show an info message to the user
        [System.Windows.Forms.MessageBox]::Show("Please select a Script to save its JSON file.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$form.Controls.Add($textBoxVersion)

# Author Label
$labelAuthor = New-Object System.Windows.Forms.Label
$labelAuthor.Text = "|Author|"
$labelAuthor.Location = New-Object System.Drawing.Point(435, 230) # Increase X = Right, y = Down
$labelAuthor.Size = New-Object System.Drawing.Size(150, 15)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelAuthor)

# Author TextBox
$textBoxAuthor = New-Object System.Windows.Forms.TextBox
$textBoxAuthor.Location = New-Object System.Drawing.Point(435, 245) # Increase X = Right, y = Down
$textBoxAuthor.Size = New-Object System.Drawing.Size(150, 20)       # Increase W = Wider, H = Taller.
$textBoxAuthor.BackColor = "MintCream"
$textBoxAuthor.CharacterCasing = "Upper"
$textBoxAuthor.Add_Leave({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$selectedScript.ps1"
        $jsonPath = ($scriptPath -replace "\.ps1$", ".json")

        if (Test-Path -Path $jsonPath) {
            try {
                $jsonContent = Get-Content -Path $jsonPath -Raw -ErrorAction Stop | ConvertFrom-Json

                # Save the previous value of Author for comparison
                $previousAuthor = $jsonContent.Author

                # Update the Author field with the new value
                $jsonContent.Author = $textBoxAuthor.Text

                # Save the updated JSON content back to the file
                $jsonContent | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8

                # Clear the RichTextBoxOutput and display the updated Author field data
                $richTextBoxOutput.Clear()
                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::PapayaWhip
                $richTextBoxOutput.AppendText("Author for script '$selectedScript' updated successfully!" + [Environment]::NewLine)

                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Silver

                $authorData = @{
                    Author = @{
                        OldValue = $previousAuthor
                        NewValue = $textBoxAuthor.Text
                    }
                } | ConvertTo-Json -Depth 4

                $richTextBoxOutput.AppendText($authorData)
            }
            catch {
                $errorMessage = "Failed to read or update the JSON file: $jsonPath`nError: $_"
                Add-Content -Path "error.log" -Value $errorMessage

                # Show an error message to the user
                [System.Windows.Forms.MessageBox]::Show("An error occurred while updating the Author field: `n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        else {
            $warningMessage = "No JSON file exists for the selected script: $jsonPath"
            Add-Content -Path "error.log" -Value $warningMessage

            # Show a warning message to the user
            [System.Windows.Forms.MessageBox]::Show("No JSON file exists for the selected script.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    else {
        $infoMessage = "Please select a Script to save its JSON file."

        # Show an info message to the user
        [System.Windows.Forms.MessageBox]::Show("Please select a Script to save its JSON file.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$form.Controls.Add($textBoxAuthor)

# Modified By Label
$labelModifiedBy = New-Object System.Windows.Forms.Label
$labelModifiedBy.Text = "|Modified By|"
$labelModifiedBy.Location = New-Object System.Drawing.Point(585, 230) # Increase X = Right, y = Down
$labelModifiedBy.Size = New-Object System.Drawing.Size(150, 15)      # Increase W for wider, H for talle
$form.Controls.Add($labelModifiedBy)

# Modified By TextBox
$textBoxModifiedBy = New-Object System.Windows.Forms.TextBox
$textBoxModifiedBy.Location = New-Object System.Drawing.Point(585, 245) # Increase X = Right, y = Down
$textBoxModifiedBy.Size = New-Object System.Drawing.Size(150, 20)       # Increase W for wider, H for talle
$textBoxModifiedBy.BackColor = [System.Drawing.Color]::Mintcream
$textBoxModifiedBy.Add_Leave({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$selectedScript.ps1"
        $jsonPath = ($scriptPath -replace "\.ps1$", ".json")

        if (Test-Path -Path $jsonPath) {
            try {
                $jsonContent = Get-Content -Path $jsonPath -Raw -ErrorAction Stop | ConvertFrom-Json

                # Save the previous value of ModifiedBy for comparison
                $previousModifiedBy = $jsonContent.ModifiedBy

                # Update the ModifiedBy field with the new value
                $jsonContent.ModifiedBy = $textBoxModifiedBy.Text

                # Save the updated JSON content back to the file
                $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $jsonPath -Encoding UTF8

                # Clear the RichTextBoxOutput and display the updated ModifiedBy field data
                $richTextBoxOutput.Clear()
                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::PapayaWhip
                $richTextBoxOutput.AppendText("ModifiedBy for script '$selectedScript' updated successfully!" + [Environment]::NewLine)

                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Silver

                $modifiedByData = @{
                    ModifiedBy = @{
                        OldValue = $previousModifiedBy
                        NewValue = $textBoxModifiedBy.Text
                    }
                } | ConvertTo-Json -Depth 4

                $richTextBoxOutput.AppendText($modifiedByData)
            }
            catch {
                $errorMessage = "Failed to read or update the JSON file: $jsonPath`nError: $_"
                Add-Content -Path "error.log" -Value $errorMessage

                # Show an error message to the user
                [System.Windows.Forms.MessageBox]::Show("An error occurred while updating the ModifiedBy field: `n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        else {
            $warningMessage = "No JSON file exists for the selected script: $jsonPath"
            Add-Content -Path "error.log" -Value $warningMessage

            # Show a warning message to the user
            [System.Windows.Forms.MessageBox]::Show("No JSON file exists for the selected script.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    else {
        $infoMessage = "Please select a Script to save its JSON file."

        # Show an info message to the user
        [System.Windows.Forms.MessageBox]::Show("Please select a Script to save its JSON file.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$form.Controls.Add($textBoxModifiedBy)

# Date Created Label
$labelDateCreated = New-Object System.Windows.Forms.Label
$labelDateCreated.Text = "|Created|"
$labelDateCreated.Location = New-Object System.Drawing.Point(405, 270) # Increase X = Right, y = Down
$labelDateCreated.Size = New-Object System.Drawing.Size(100, 15)       # Increase W for wider, H for talle
$form.Controls.Add($labelDateCreated)

# Date Created Textbox
$textBoxDateCreated = New-Object System.Windows.Forms.TextBox
$textBoxDateCreated.Location = New-Object System.Drawing.Point(405, 285) # Increase X = Right, y = Down
$textBoxDateCreated.Size = New-Object System.Drawing.Size(165, 20)       # Increase W for wider, H for talle
$textBoxDateCreated.Backcolor = "MintCream"
$textBoxDateCreated.ReadOnly = $true
$form.Controls.Add($textBoxDateCreated)

# Date Modified Label
$labelDateModified = New-Object System.Windows.Forms.Label
$labelDateModified.Text = "|Modified|"
$labelDateModified.Location = New-Object System.Drawing.Point(575, 270) # Increase X = Right, y = Down
$labelDateModified.Size = New-Object System.Drawing.Size(100, 15)      # Increase W for wider, H for taller
$form.Controls.Add($labelDateModified)

# Date Modified TextBox
$textBoxDateModified = New-Object System.Windows.Forms.TextBox
$textBoxDateModified.Location = New-Object System.Drawing.Point(570, 285) # Increase X = Right, y = Down
$textBoxDateModified.Size = New-Object System.Drawing.Size(165, 20)      # Increase W for wider, H for talle
$textBoxDateModified.Backcolor = "MintCream"
$textBoxDateModified.ReadOnly = $true
$form.Controls.Add($textBoxDateModified)


# JSON File Label
$labelJsonFile = New-Object System.Windows.Forms.Label
$labelJsonFile.Text = "|JSON File|"
$labelJsonFile.Location = New-Object System.Drawing.Point(405, 313) # Increase X = Right, y = Down
$labelJsonFile.Size = New-Object System.Drawing.Size(65, 15)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelJsonFile)

# JSON File Textbox
$textBoxJsonFile = New-Object System.Windows.Forms.TextBox
$textBoxJsonFile.Location = New-Object System.Drawing.Point(405, 328) # Increase X = Right, y = Down
$textBoxJsonFile.Size = New-Object System.Drawing.Size(330, 15)       # Increase W = Wider, H = Taller.
$textBoxJsonFile.CharacterCasing = "Upper"
$textBoxJsonFile.BackColor = [System.Drawing.Color]::Mintcream
$textBoxJsonFile.ForeColor = [System.Drawing.Color]::Black
$textBoxJsonFile.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$textBoxJsonFile.ReadOnly = $true
$form.Controls.Add($textBoxJsonFile)

# Open JSON Button
$buttonOpenJson = New-Object System.Windows.Forms.Button
$buttonOpenJson.Text = "OPEN VSC"
$buttonOpenJson.Location = New-Object System.Drawing.Point(490, 305) # Increase X = Right, Y = Down
#$buttonOpenJson.Size = New-Object System.Drawing.Size(60, 15)       # Increase W = Wider, H = Taller.
$buttonOpenJson.Autosize = $true      # Increase W = Wider, H = Taller.
$buttonOpenJson.FlatStyle = [System.Windows.Forms.FlatStyle]::System
#$buttonOpenJson.FlatAppearance.BorderSize = 0
#$buttonOpenJson.BackColor = [System.Drawing.Color]::BurlyWood
#$buttonOpenJson.ForeColor = [System.Drawing.Color]::Black
# Event handler for the "Open JSON" button
$buttonOpenJson.Add_Click({
    # Get the selected script from the ListBox
    $selectedScript = $listBoxScripts.SelectedItem

    if ($selectedScript) {
        # Construct the script and JSON file paths
        $scriptPathWithoutExtension = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        $jsonPath = $scriptPathWithoutExtension + ".json"

        if (Test-Path -Path $jsonPath) {
            try {
                # Open the JSON file with Visual Studio Code
                Start-Process "code" -ArgumentList $jsonPath
            }
            catch {
                $errorMessage = "Failed to open the JSON file: $jsonPath`nError: $_"
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Add-Content -Path "error.log" -Value $errorMessage
                Write-Host $errorMessage
            }
        }
        else {
            $warningMessage = "No JSON file exists for the selected script: $jsonPath"
            [System.Windows.Forms.MessageBox]::Show($warningMessage, "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Add-Content -Path "error.log" -Value $warningMessage
            Write-Host $warningMessage
        }
    }
    else {
        $infoMessage = "Please select a Script to open its JSON file."
        [System.Windows.Forms.MessageBox]::Show($infoMessage, "No Script Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Add-Content -Path "error.log" -Value $infoMessage
        Write-Host $infoMessage
    }
})

# Add the "Open JSON" button to the form
$form.Controls.Add($buttonOpenJson)

# Open JSON with Notepad Button
$buttonOpenJsonNotepad = New-Object System.Windows.Forms.Button
$buttonOpenJsonNotepad.Text = "OPEN NOTEPAD"
$buttonOpenJsonNotepad.Location = New-Object System.Drawing.Point(562, 305)
$buttonOpenJsonNotepad.Autosize = $true
$buttonOpenJsonNotepad.FlatStyle = [System.Windows.Forms.FlatStyle]::System

# Event handler for the "Open JSON with Notepad" button
$buttonOpenJsonNotepad.Add_Click({
    # Get the selected script from the ListBox
    $selectedScript = $listBoxScripts.SelectedItem

    if ($selectedScript) {
        # Construct the script and JSON file paths
        $scriptPathWithoutExtension = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        $jsonPath = $scriptPathWithoutExtension + ".json"

        if (Test-Path -Path $jsonPath) {
            try {
                # Open the JSON file with Notepad
                Start-Process "notepad.exe" -ArgumentList $jsonPath
            }
            catch {
                $errorMessage = "Failed to open the JSON file: $jsonPath`nError: $_"
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Add-Content -Path "error.log" -Value $errorMessage
                Write-Host $errorMessage
            }
        }
        else {
            $warningMessage = "No JSON file exists for the selected script: $jsonPath"
            [System.Windows.Forms.MessageBox]::Show($warningMessage, "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Add-Content -Path "error.log" -Value $warningMessage
            Write-Host $warningMessage
        }
    }
    else {
        $infoMessage = "Please select a Script to open its JSON file."
        [System.Windows.Forms.MessageBox]::Show($infoMessage, "No Script Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Add-Content -Path "error.log" -Value $infoMessage
        Write-Host $infoMessage
    }
})

# Add the "Open JSON with Notepad" button to the form
$form.Controls.Add($buttonOpenJsonNotepad)


# Delete JSON Button
$buttonDeleteJson = New-Object System.Windows.Forms.Button
$buttonDeleteJson.Text = "DELETE"
$buttonDeleteJson.Location = New-Object System.Drawing.Point(660, 305) # Increase X = Right, y = Down
#$buttonDeleteJson.Size = New-Object System.Drawing.Size(60, 22)       # Increase W = Wider, H = Taller.
$buttonDeleteJson.AutoSize = $true       # Increase W = Wider, H = Taller.
$buttonDeleteJson.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$buttonDeleteJson.FlatAppearance.BorderSize = 0
#$buttonDeleteJson.BackColor = [System.Drawing.Color]::BurlyWood
#$buttonDeleteJson.ForeColor = [System.Drawing.Color]::Black
# Delete JSON Button Event Handler
$buttonDeleteJson.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$selectedScript.ps1"
        $jsonPath = $scriptPath -replace "\.ps1$", ".json"

        if (Test-Path -Path $jsonPath) {
            $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the JSON file?", "Confirm Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($result -eq 'Yes') {
                Remove-Item -Path $jsonPath -Force
                # Clear the fields
                $textBoxName.Text = ""
                $textBoxDescription.Text = ""
                $textBoxVersion.Text = ""
                $textBoxAuthor.Text = ""
                $textBoxDateCreated.Text = ""
                $textBoxJsonFile.Text = ""
                # Display a message box to indicate successful deletion
                [System.Windows.Forms.MessageBox]::Show("JSON file deleted successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
        else {
            # Display a message box to indicate that the JSON file does not exist
            [System.Windows.Forms.MessageBox]::Show("No JSON file exists for the selected script.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    else {
        # Display a message box to indicate that no script is selected
        [System.Windows.Forms.MessageBox]::Show("Please select a Script to delete its JSON file.", "No Script Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$form.Controls.Add($buttonDeleteJson)

# Start of Requirements Area
$requirementsList = New-Object System.Collections.ArrayList

$requirementsGridView = New-Object Windows.Forms.DataGridView
$requirementsGridView.Location = New-Object Drawing.Point(405, 360)
$requirementsGridView.Size = New-Object Drawing.Size(330, 140)
$requirementsGridView.BackgroundColor = [System.Drawing.Color]::MintCream
$requirementsGridView.AutoGenerateColumns = $false

$numberColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$numberColumn.HeaderText = "Number"
$numberColumn.DataPropertyName = "Number"
$numberColumn.ReadOnly = $true
$numberColumn.Width = 40
$requirementsGridView.Columns.Add($numberColumn)

$requirementColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$requirementColumn.HeaderText = "Requirement"
$requirementColumn.DataPropertyName = "Requirement"
$requirementColumn.Width = 150
$requirementsGridView.Columns.Add($requirementColumn)

$fulfilledColumn = New-Object Windows.Forms.DataGridViewCheckBoxColumn
$fulfilledColumn.HeaderText = "Fulfilled"
$fulfilledColumn.DataPropertyName = "Fulfilled"
$fulfilledColumn.TrueValue = $true
$fulfilledColumn.FalseValue = $false
$fulfilledColumn.Width = 60
$requirementsGridView.Columns.Add($fulfilledColumn)
$form.Controls.Add($requirementsGridView)

# Add Requirement button
$addRequirementButton = New-Object Windows.Forms.Button
$addRequirementButton.Text = "Add Requirement"
$addRequirementButton.Location = New-Object Drawing.Point(735, 360) # Increase X = Right, y = Down
$addRequirementButton.Size = New-Object Drawing.Size(130, 25)       # Increase W for wider, H for taller
# Event handler for adding a new requirement
$addRequirementButton.Add_Click({
    # Create a new form to enter the requirement description
    $addRequirementForm = New-Object Windows.Forms.Form
    $addRequirementForm.Text = "Add Requirement"
    $addRequirementForm.Size = New-Object Drawing.Size(300, 150)
    $addRequirementForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen # This line centers the form

    $requirementTextbox = New-Object Windows.Forms.TextBox
    $requirementTextbox.Location = New-Object Drawing.Point(10, 10)
    $requirementTextbox.Size = New-Object Drawing.Size(250, 20)
    $addRequirementForm.Controls.Add($requirementTextbox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(60, 60)
    $okButton.Size = New-Object System.Drawing.Size(75, 25)
    $okButton.Text = "OK"
    $okButton.Add_Click({
        $addRequirementForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })
    $addRequirementForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150,60)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $addRequirementForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    })
    $addRequirementForm.Controls.Add($cancelButton)

    $addRequirementForm.AcceptButton = $okButton
    $addRequirementForm.CancelButton = $cancelButton

    $result = $addRequirementForm.ShowDialog()


    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $requirementDescription = $requirementTextbox.Text

        $newRequirement = New-Object PSObject -Property @{
            "Number" = ($requirementsList.Count + 1) # Assuming you want to increment the Number
            "Requirement" = $requirementDescription
            "Fulfilled" = $false
        }
    
        # Add the new requirement to the requirements list
         # Add the new requirement to the requirements list
        $requirementsList.Add($newRequirement) # Use this if you have a List
        # Or
        # $requirementsList += $newRequirement # Use this if you have an array # Use this if you have a List
        # Or
        # $requirementsList += $newRequirement # Use this if you have an array
    
        # Refresh the DataGridView to display the updated data
        $requirementsGridView.DataSource = $null
        $requirementsGridView.DataSource = $requirementsList
    }


    $addRequirementForm.Dispose()
})
$form.Controls.Add($addRequirementButton)

# Remove Requirement button
$removeRequirementButton = New-Object Windows.Forms.Button
$removeRequirementButton.Text = "Remove Requirement"
$removeRequirementButton.Location = New-Object Drawing.Point(735, 385) # Increase X = Right, y = Down
$removeRequirementButton.Size = New-Object Drawing.Size(130, 25)       # Increase W for wider, H for taller
# Event handler for removing a requirement
$removeRequirementButton.Add_Click({
    if ($requirementsGridView.CurrentRow -ne $null) {
        $selectedIndex = $requirementsGridView.CurrentRow.Index
        $requirementsList.RemoveAt($selectedIndex)
        $requirementsGridView.DataSource = $null
        $requirementsGridView.DataSource = $requirementsList
    }
})
$form.Controls.Add($removeRequirementButton)

# Save Requirement button
$saveButton = New-Object Windows.Forms.Button
$saveButton.Text = "Save"
$saveButton.Location = New-Object Drawing.Point(735, 410) # Increase X = Right, Y = Down
$saveButton.Size = New-Object Drawing.Size(130, 25)       # Increase W for wider, H for taller
$saveButton.Add_Click({
    if ($global:selectedScript -eq $null) {
        Write-Host "Error: No script selected."
        return
    }

    $scriptPaths = Get-ScriptPaths -selectedScript $global:selectedScript -baseScriptPath $global:BaseScriptPath
    Write-Host ("Script Path: $($scriptPaths.ScriptPath)") -ForegroundColor Green
    Write-Host ("JSON Path: $($scriptPaths.JsonPath)") -ForegroundColor Green

    # Load the JSON content
    if (Test-Path -Path $scriptPaths.JsonPath) {
        
        
        Write-Host "JSON content BEFORE saving:"
        $jsonContentBefore | ConvertTo-Json -Depth 4

        $jsonContentBefore = Get-Content -Path $scriptPaths.JsonPath -Raw | ConvertFrom-Json

        # Perform the modifications and updates
        Write-Host "Updating JSON content..."
        # ... your code for updating requirements and other fields ...
        Write-Host "Modifications and updates completed."

        # Save the updated JSON content back to the file
        $jsonContent = UpdateJsonContent -jsonContent $jsonContentBefore -selectedScript $scriptPaths.ScriptName -scriptPath $scriptPaths.ScriptPath -jsonPath $scriptPaths.JsonPath

        $jsonContentAfter = Get-Content -Path $scriptPaths.JsonPath -Raw | ConvertFrom-Json

        Write-Host "JSON content BEFORE saving:"
        Write-Host $jsonContentBefore

        Write-Host "JSON content AFTER saving:"
        Write-Host $jsonContentAfter

        # Save the updated JSON content back to the file
        Write-Host "Saving JSON content..."
        $jsonContent | ConvertTo-Json -Depth 4 | Set-Content -Path $scriptPaths.JsonPath -Encoding UTF8
        Write-Host "JSON content saved."

        Write-Host "JSON content saved to $($scriptPaths.JsonPath)"
    } else {
        Write-Host "JSON Path does not exist: $($scriptPaths.JsonPath)"
        # Handle error as needed
    }
})
$form.Controls.Add($saveButton)










# Create a button to initiate the search
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(400, 510)    # Increase X = Right, y = Down
$searchButton.Size = New-Object System.Drawing.Size(185, 30)          # Increase W for wider, H for taller
$searchButton.backcolor = [System.Drawing.Color]::SaddleBrown
$searchButton.forecolor = [System.Drawing.Color]::White
$searchButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$searchButton.FlatAppearance.BorderSize = 0
$searchButton.Text = "FILTER BY KEYWORDS"
$searchButton.Add_Click({
    # Create a new form for selecting keywords
    $keywordsForm = New-Object System.Windows.Forms.Form
    $keywordsForm.Text = "Select Keywords"
    $keywordsForm.Size = New-Object System.Drawing.Size(300, 1000)
    $keywordsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $keywordsForm.BackColor = "SaddleBrown"

    # Collect unique keywords from all JSON files
    $keywords = Get-ChildItem -Path $global:BaseScriptPath -Filter '*.json' | ForEach-Object {
        $jsonContent = Get-Content -Path $_.FullName | ConvertFrom-Json
        $jsonContent.Keywords
    } | Sort-Object | Select-Object -Unique

    # Create a ListBox for available Keywords, sorted alphabetically
    $listBoxKeywords = New-Object System.Windows.Forms.ListBox
    $listBoxKeywords.Location = New-Object System.Drawing.Point(30, 50)  # Increase X = Right, Y = Down
    $listBoxKeywords.Size = New-Object System.Drawing.Size(225, 880)     # Increase W = Wider, H = Taller
    $listBoxKeywords.Font = New-Object System.Drawing.Font("Arial Narrow", 10)
    $listBoxKeywords.BackColor = "LightYellow"
    $listBoxKeywords.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiSimple
    $keywords | ForEach-Object { $listBoxKeywords.Items.Add($_) }

    $listBoxKeywords.Add_SelectedIndexChanged({
        # Clear the scripts ListBox
        $listBoxScripts.Items.Clear()

        # Get selected keywords
        $selectedKeywords = $listBoxKeywords.SelectedItems

        # Iterate through the scripts and JSON files, filtering by selected keywords
        $scripts = Get-ChildItem -Path $global:BaseScriptPath -Filter '*.ps1'
        foreach ($script in $scripts) {
            $jsonPath = [System.IO.Path]::ChangeExtension($script.FullName, 'json')
            if (Test-Path -Path $jsonPath) {
                $jsonContent = Get-Content -Path $jsonPath | ConvertFrom-Json

                # Check if JSON content contains all selected keywords
                if ($selectedKeywords | Where-Object { $jsonContent.Keywords -notcontains $_ }) {
                    continue
                }
            }

            # Add the script to the scripts ListBox without extension
            $listBoxScripts.Items.Add($script.BaseName)
        }
    })

    $keywordsForm.Controls.Add($listBoxKeywords)

    # Show the keywords form
    $keywordsForm.ShowDialog()
})


$form.Controls.Add($searchButton)

<#
$clearFilterButton = New-Object System.Windows.Forms.Button
$clearFilterButton.Location = New-Object System.Drawing.Point(865, 230)  # Increase X = Right, Y = Down
$clearFilterButton.Size = New-Object System.Drawing.Size(185, 25)        # Increase W = Wider, H = Taller
$clearFilterButton.Text = "Clear Filter"
$clearFilterButton.Add_Click({
    $listBoxScripts.Items.Clear() # Clear the current list

    # Load all scripts without filtering
    $scripts = Get-ChildItem -Path $global:BaseScriptPath -Filter '*.ps1'
    foreach ($script in $scripts) {
        $listBoxScripts.Items.Add($script.BaseName) # Add scripts without extension
    }
})

$form.Controls.Add($clearFilterButton)


# Create a multiline text box for search input
$searchTextBox = New-Object System.Windows.Forms.TextBox
$searchTextBox.Location = New-Object System.Drawing.Point(865, 140) # Increase X = Right, Y = Down
$searchTextBox.Size = New-Object System.Drawing.Size(185, 110)     # Increase W = Wider, H = Taller
$searchTextBox.BackColor = [System.Drawing.Color]::LightYellow
$searchTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
$searchTextBox.Multiline = $true
$searchTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$searchTextBox.Text = 'Enter Keywords'
$searchTextBox.ForeColor = [System.Drawing.Color]::Gray

$searchTextBox.Add_Enter({
    if ($this.Text -eq 'Enter Keywords') {
        $this.Text = ''
        $this.ForeColor = [System.Drawing.Color]::Black
    }
})

$searchTextBox.Add_Leave({
    if ($this.Text -eq '') {
        $this.Text = 'Enter Keywords'
        $this.ForeColor = [System.Drawing.Color]::Gray
    }
})

$searchTextBox.Add_KeyPress({
    # Check if Enter key was pressed
    if ($_.KeyChar -eq [System.Windows.Forms.Keys]::Enter) {
        $keywords = $searchTextBox.Text.Trim().Split(' ')

        # Clear the scripts ListBox
        $listBoxScripts.Items.Clear()

        # Iterate through the scripts and JSON files
        $scripts = Get-ChildItem -Path $global:BaseScriptPath -Filter '*.ps1'
        foreach ($script in $scripts) {
            $jsonPath = [System.IO.Path]::ChangeExtension($script.FullName, 'json')
            if (Test-Path -Path $jsonPath) {
                $jsonContent = Get-Content -Path $jsonPath | ConvertFrom-Json

                # Check if JSON content contains all keywords
                if ($keywords | Where-Object { $jsonContent.Keywords -notcontains $_ }) {
                    continue
                }
            }

            # Add the script to the scripts ListBox without extension
            $listBoxScripts.Items.Add($script.BaseName)
        }

        # Clear the search TextBox
        $searchTextBox.Text = ''
    }
})

$form.Controls.Add($searchTextBox)


#=================================================================== CHECK BOX AREA

# Create a CheckedListBox for available Keywords
$checkedListBoxKeywords = New-Object System.Windows.Forms.CheckedListBox
$checkedListBoxKeywords.Location = New-Object System.Drawing.Point(865, 530)
$checkedListBoxKeywords.Size = New-Object System.Drawing.Size(185, 235)
$checkedListBoxKeywords.Font = New-Object System.Drawing.Font("Arial Narrow", 10)
$checkedListBoxKeywords.BackColor = "LightYellow"
$checkedListBoxKeywords.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$form.Controls.Add($checkedListBoxKeywords)

# Collect unique keywords from all JSON files
$keywords = Get-ChildItem -Path $global:BaseScriptPath -Filter '*.json' | ForEach-Object {
    $jsonContent = Get-Content -Path $_.FullName | ConvertFrom-Json
    $jsonContent.Keywords
} | Select-Object -Unique

# Add unique keywords to the CheckedListBox
$keywords | ForEach-Object { $checkedListBoxKeywords.Items.Add($_) }

# Create a button to apply filters
$filterButton = New-Object System.Windows.Forms.Button
$filterButton.Text = 'Filter Scripts'
$filterButton.Location = New-Object System.Drawing.Point(865, 770)
$filterButton.Size = New-Object System.Drawing.Size(185, 40)
$filterButton.Add_Click({
    # Clear the scripts ListBox
    $listBoxScripts.Items.Clear()

    # Get selected keywords
    $selectedKeywords = $checkedListBoxKeywords.CheckedItems

    # Iterate through the scripts and JSON files, filtering by selected keywords
    $scripts = Get-ChildItem -Path $global:BaseScriptPath -Filter '*.ps1'
    foreach ($script in $scripts) {
        $jsonPath = [System.IO.Path]::ChangeExtension($script.FullName, 'json')
        if (Test-Path -Path $jsonPath) {
            $jsonContent = Get-Content -Path $jsonPath | ConvertFrom-Json

            # Check if JSON content contains all selected keywords
            if ($selectedKeywords | Where-Object { $jsonContent.Keywords -notcontains $_ }) {
                continue
            }
        }

        # Add the script to the scripts ListBox without extension
        $listBoxScripts.Items.Add($script.BaseName)
    }
})
$form.Controls.Add($filterButton)

#>



# Button for Search Mode (Default button)
$buttonSearchMode = New-Object System.Windows.Forms.Button
$buttonSearchMode.Text = "SEARCH BY TAGS"
$buttonSearchMode.Location = New-Object System.Drawing.Point(585,510) # Increase X = Right, Y = Down
$buttonSearchMode.Size = New-Object System.Drawing.Size(185, 30)      # Increase W = Wider, H = Taller
$buttonSearchMode.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonSearchMode.FlatAppearance.BorderSize = 0
$buttonSearchMode.BackColor = [System.Drawing.Color]::SaddleBrown  # Light blue color
$buttonSearchMode.ForeColor = [System.Drawing.Color]::White  # Black text color
# Event handler for the "Search" button click
$buttonSearchMode.Add_Click({
    ShowTagSelectionForm
})
$form.Controls.Add($buttonSearchMode)

# Create View To-Do List button
$buttonViewToDoList = New-Object System.Windows.Forms.Button
$buttonViewToDoList.Text = "TO DO LIST"
$buttonViewToDoList.Location = New-Object System.Drawing.Point(770, 510)  # Increase X = Right, Y = Down
$buttonViewToDoList.Size = New-Object System.Drawing.Size(150, 30)        # Increase W = Wider, H = Taller.
$buttonViewToDoList.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonViewToDoList.FlatAppearance.BorderSize = 0
$buttonViewToDoList.BackColor = [System.Drawing.Color]::SaddleBrown
$buttonViewToDoList.ForeColor = [System.Drawing.Color]::White  # Set text color
$buttonViewToDoList.Add_Click({ ShowToDoList })
$form.Controls.Add($buttonViewToDoList)

# Create Modification Log button
$buttonModificationLog = New-Object System.Windows.Forms.Button
$buttonModificationLog.Text = "CHANGE LOG"
$buttonModificationLog.Location = New-Object System.Drawing.Point(900, 510)  # Increase X = Right, Y = Down
$buttonModificationLog.Size = New-Object System.Drawing.Size(150, 30)        # Increase W = Wider, H = Taller.
$buttonModificationLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonModificationLog.FlatAppearance.BorderSize = 0
$buttonModificationLog.BackColor = [System.Drawing.Color]::SaddleBrown
$buttonModificationLog.ForeColor = [System.Drawing.Color]::White
$buttonModificationLog.Add_Click({ ShowModificationLog })
$form.Controls.Add($buttonModificationLog)

# Define the function to add keywords
function AddKeywords {
    param (
        [string]$selectedScript,
        [string]$baseScriptPath,
        [System.Windows.Forms.TextBox]$newKeywordTextBox,
        [System.Windows.Forms.ListBox]$listBoxKeywords
    )

    Write-Host "Base Script Path: $baseScriptPath"
    Write-Host "Selected Script: $selectedScript"

    if ([string]::IsNullOrEmpty($baseScriptPath) -or [string]::IsNullOrEmpty($selectedScript)) {
        Write-Host "Error: baseScriptPath or selectedScript is null or empty."
        return
    }

    $scriptPath = Join-Path -Path $baseScriptPath -ChildPath "$selectedScript.ps1"
    $jsonPath = Join-Path -Path $baseScriptPath -ChildPath "$selectedScript.json"

    $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

    # Ensure "Keywords" property is initialized as an array
    if ($null -eq $jsonContent.Keywords) {
        $jsonContent.Keywords = @()
    }

    $newKeyword = $newKeywordTextBox.Text.Trim().ToUpper() # Converting to uppercase

    if ($newKeyword -ne '' -and $newKeyword -ne 'Add New Keyword' -and $jsonContent.Keywords -notcontains $newKeyword) {
        # Append the new keyword to the array
        $jsonContent.Keywords += $newKeyword

        # Sort and remove duplicates from the array
        $jsonContent.Keywords = $jsonContent.Keywords | Sort-Object -Unique

        # Update the ListBox with the updated keywords array
        $listBoxKeywords.BeginUpdate()
        $listBoxKeywords.Items.Clear()
        $jsonContent.Keywords | ForEach-Object {
            $listBoxKeywords.Items.Add($_)
        }
        $listBoxKeywords.EndUpdate()

        # Save the updated JSON content back to the file
        $jsonContent | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8

        $newKeywordTextBox.Clear()

        Write-Host "Adding keyword to ListBox." -ForegroundColor Green
        Write-Host "JSON content updated and written to file." -ForegroundColor Green
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid and unique keyword.")
    }
    Write-Host "End of AddKeywords function" -ForegroundColor Green
}

function RemoveKeywords {
    param (
        [System.Windows.Forms.ListBox]$listBoxKeywords
    )

    $selectedScript = $global:selectedScript
    $scriptPath = Get-ScriptFilePath -scriptName $selectedScript
    $jsonPath = Get-JsonFilePath -scriptName $selectedScript

    if (Test-Path -Path $jsonPath) {
        $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

        if ($null -eq $jsonContent.Keywords) {
            Write-Host "Warning: 'Keywords' property not found in JSON content." -ForegroundColor Yellow
            return
        }

        # Get selected keywords to remove
        $selectedKeywordsToRemove = $listBoxKeywords.SelectedItems

        # Remove selected keywords
        $jsonContent.Keywords = $jsonContent.Keywords | Where-Object { $_ -notin $selectedKeywordsToRemove }

        # Check if there are no remaining keywords
        if ($null -eq $jsonContent.Keywords -or $jsonContent.Keywords.Count -eq 0) {
            $jsonContent.Keywords = @() # Initialize Keywords as an empty array
        } else {
            # Update the ListBox with the updated keywords array
            $listBoxKeywords.BeginUpdate()
            $listBoxKeywords.Items.Clear()
            $jsonContent.Keywords | ForEach-Object {
                $listBoxKeywords.Items.Add($_)
            }
            $listBoxKeywords.EndUpdate()
        }

        # Save the updated JSON content back to the file
        $jsonContent | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8

        Write-Host "Selected keywords removed from ListBox and JSON content." -ForegroundColor Green
    } else {
        Write-Host "JSON Path does not exist: $jsonPath"
    }
}

# Create a TextBox for new keywords input
$newKeywordTextBox = New-Object System.Windows.Forms.TextBox
$newKeywordTextBox.Location = New-Object System.Drawing.Point(865, 140) # Increase X = Right, Y = Down
$newKeywordTextBox.Size = New-Object System.Drawing.Size(185, 30)      # Increase W = Wider, H = Taller
$newKeywordTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
$newKeywordTextBox.BackColor = "MintCream"
$newKeywordTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center # Center the text
$newKeywordTextBox.Text = "Enter New Keyword"
$newKeywordTextBox.ForeColor = [System.Drawing.Color]::Gray
$newKeywordTextBox.Add_Enter({
    if ($this.Text -eq 'Add New Keyword') {
        $this.Text = '';
        $this.ForeColor = [System.Drawing.Color]::Black;
    }
})
$newKeywordTextBox.Add_Leave({
    if ($this.Text -eq '') {
        $this.Text = 'Add New Keyword';
        $this.ForeColor = [System.Drawing.Color]::Gray;
    }
})
$form.Controls.Add($newKeywordTextBox)

# Create Add Keywords button
$buttonAddKeywords = New-Object System.Windows.Forms.Button
$buttonAddKeywords.Text = "| ADD |"
$buttonAddKeywords.Location = New-Object System.Drawing.Point(960, 165)  # Increase X = Right, Y = Down
$buttonAddKeywords.Size = New-Object System.Drawing.Size(90, 20)        # Increase W = Wider, H = Taller.
$buttonAddKeywords.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonAddKeywords.FlatAppearance.BorderSize = 0
$buttonAddKeywords.BackColor = [System.Drawing.Color]::DarkSeaGreen
$buttonAddKeywords.ForeColor = [System.Drawing.Color]::White
$buttonAddKeywords.Add_Click({
    Write-Host "Executing AddKeywords function..."
    Write-Host "ScriptPath: $global:BaseScriptPath"
    Write-Host "SelectedScript: $selectedScript"
    Write-Host "NewKeywordTextBox.Text: $($newKeywordTextBox.Text)"
    AddKeywords -baseScriptPath $global:BaseScriptPath -selectedScript $selectedScript -newKeywordTextBox $newKeywordTextBox -listBoxKeywords $listBoxKeywords
    Write-Host "AddKeywords function executed."
})

$form.Controls.Add($buttonAddKeywords)

# Remove Selected Keywords Button
$buttonRemoveKeywords = New-Object System.Windows.Forms.Button
$buttonRemoveKeywords.Text = "| REMOVE |"
$buttonRemoveKeywords.Location = New-Object System.Drawing.Point(865, 165) # Increase X = Right, Y = Down
$buttonRemoveKeywords.Size = New-Object System.Drawing.Size(95, 20) # Increase W = Wider, H = Taller
$buttonRemoveKeywords.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRemoveKeywords.FlatAppearance.BorderSize = 0
$buttonRemoveKeywords.BackColor = [System.Drawing.Color]::LightSlateGray
$buttonRemoveKeywords.ForeColor = [System.Drawing.Color]::White
$buttonRemoveKeywords.Add_Click({
    RemoveKeywords -defaultJsonPath $defaultJsonPath -listBoxKeywords $listBoxKeywords
})
$form.Controls.Add($buttonRemoveKeywords)

# Create a ListBox for Keywords
$listBoxKeywords = New-Object System.Windows.Forms.ListBox
$listBoxKeywords.Location = New-Object System.Drawing.Point(865, 185) # Increase X = Right, Y = Down
$listBoxKeywords.Size = New-Object System.Drawing.Size(185, 335)     # Increase W = Wider, H = Taller
$listBoxKeywords.Font = New-Object System.Drawing.Font("Arial Narrow", 10)
$listBoxKeywords.BackColor = "LightYellow"
$listBoxKeywords.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$listBoxKeywords.SelectionMode = "MultiSimple"
$form.Controls.Add($listBoxKeywords)



# Add controls to the form
$form.Controls.Add($labelBackupCounter)
$form.Controls.Add($labelBackupCountValue)

# Save JSON Button
#$buttonSave = New-Object System.Windows.Forms.Button
#$buttonSave.Text = "SAVE JSON"
#$buttonSave.Location = New-Object System.Drawing.Point(720, 435) # Increase X = Right, y = Down
#$buttonSave.Size = New-Object System.Drawing.Size(130, 22)       # Increase W = Wider, H = Taller.
#$buttonSave.BackColor = [System.Drawing.Color]::Pink
#$buttonSave.ForeColor = [System.Drawing.Color]::Black
#$buttonSave.Visible = $false
<# 
$buttonSave.Add_Click({
    if (-not $textBoxJsonFile.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Script to Save its JSON file.", "No Script Selected", "OK", [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPathWithoutExtension = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        $jsonPath = $scriptPathWithoutExtension + ".json"

        $jsonContent = Get-Content -Path $jsonPath -Raw -ErrorAction SilentlyContinue
        if (-not $jsonContent) {
            [System.Windows.Forms.MessageBox]::Show("Failed to read the JSON file: $jsonPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        try {
            $jsonContent = $jsonContent | ConvertFrom-Json
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Invalid JSON format in the file: $jsonPath`nError: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Check and add missing properties
        $propertiesToAdd = @(
            [PSCustomObject]@{ Name = "Name"; Value = "" },
            [PSCustomObject]@{ Name = "Author"; Value = "" },
            [PSCustomObject]@{ Name = "Description"; Value = "" },
            [PSCustomObject]@{ Name = "Version"; Value = "" },
            [PSCustomObject]@{ Name = "Tags"; Value = @() },
            [PSCustomObject]@{ Name = "ModifiedBy"; Value = "" },
            [PSCustomObject]@{ Name = "Modifications"; Value = @() },
            [PSCustomObject]@{ Name = "ToDoList"; Value = @() }
        )

        $updated = $false
        foreach ($property in $propertiesToAdd) {
            $propertyName = $property.Name
            $propertyValue = $property.Value

            if (-not $jsonContent.PSObject.Properties.Name -contains $propertyName) {
                $jsonContent | Add-Member -Type NoteProperty -Name $propertyName -Value $propertyValue
                $updated = $true
            }
        }

        if ($updated) {
            $jsonContent | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8
        }

        # Update the properties with new values
        $jsonContent.Name = $textBoxName.Text
        $jsonContent.Description = $textBoxDescription.Text
        $jsonContent.Version = $textBoxVersion.Text
        $jsonContent.Author = $textBoxAuthor.Text
        $jsonContent.ModifiedBy = $textBoxModifiedBy.Text

        # Convert selected tags to a regular array of strings and update the "Tags" property
        $selectedTags = @()
        foreach ($tag in $tagsListBox.SelectedItems) {
            $selectedTags += $tag.ToString()
        }
        $jsonContent.Tags = $selectedTags

        $jsonContent.Modifications = $jsonContent.Modifications -join ','

        $jsonContent | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show("JSON file saved successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$form.Controls.Add($buttonSave)
#>
# Create a label for the Tags section
#$labelTags = New-Object System.Windows.Forms.Label
#$labelTags.Text = ""
#$labelTags.Location = New-Object System.Drawing.Point(115, 505 )  # Increase X = Right, y = Down
#$labelTags.Size = New-Object System.Drawing.Size(100, 20) # Increase W = Wider, H = Taller.
#$form.Controls.Add($labelTags)

# ListBox for Script SizeTags
$listBoxScriptTags = New-Object System.Windows.Forms.ListBox
$listBoxScriptTags.Location = New-Object System.Drawing.Point(400, 538)  # Increase X = Right, y = Down
$listBoxScriptTags.Size = New-Object System.Drawing.Size(650, 240)       # Increase W = Wider, H = Taller.
$listBoxScriptTags.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$listBoxScriptTags.SelectionMode = "MultiSimple"
$listBoxScriptTags.DisplayMember = "Text"
$listBoxScriptTags.ValueMember = "Tag"
$listBoxScriptTags.ColumnWidth = 130
$listBoxScriptTags.MultiColumn = $true
$listBoxScriptTags.BackColor = "MintCream"
$listBoxScriptTags.ForeColor = "Black"
$listBoxScriptTags.Font = New-Object System.Drawing.Font("Arial Narrow", 9)  # Change the font family and size her
$form.Controls.Add($listBoxScriptTags)

# Event handler for the SelectedIndexChanged event of the $listBoxScriptTags ListBox
$listBoxScriptTags.Add_SelectedIndexChanged({
    $selectedScript = $listBoxScripts.SelectedItem

    if ($selectedScript) {
        $scriptPathWithoutExtension = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        $jsonPath = $scriptPathWithoutExtension + ".json"

        if (Test-Path -Path $jsonPath) {
            $jsonContent = Get-Content -Path $jsonPath -Raw -ErrorAction SilentlyContinue
            if (-not $jsonContent) {
                [System.Windows.Forms.MessageBox]::Show("Failed to read the JSON file: $jsonPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            try {
                $jsonContent = $jsonContent | ConvertFrom-Json
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Invalid JSON format in the file: $jsonPath`nError: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            # Update the tags in the JSON content based on the selected items
            $selectedTags = $listBoxScriptTags.SelectedItems
            $jsonContent.Tags = $selectedTags

            $jsonString = $jsonContent | ConvertTo-Json -Depth 4
            $jsonString | Set-Content -Path $jsonPath -Encoding UTF8
            
        }
    }
})

# Selecting a Script
$listBoxScripts.Add_SelectedIndexChanged({
    if ($null -ne $listBoxScripts.SelectedItem) {
        $global:selectedScript = $listBoxScripts.SelectedItem.ToString()
        PopulateFieldsAndTags
        $scriptPath = Get-ScriptFilePath -scriptName $global:selectedScript
        $jsonPath = Get-JsonFilePath -scriptName $global:selectedScript

        # Check if the jsonPath exists before attempting to read content
        if (Test-Path -Path $jsonPath) {
            Write-Host "JSON Path exists: $jsonPath"

            $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
            $jsonContent = UpdateJsonContent -jsonContent $jsonContent -selectedScript $global:selectedScript -scriptPath $scriptPath -jsonPath $jsonPath
            
            # Add debugging output to show the contents of $jsonContent before populating the DataGridView
            Write-Host "JSON Content before populating DataGridView:"
            $jsonContent | ConvertTo-Json -Depth 4

            # Populate the DataGridView with requirements
            $requirementsList = $jsonContent.Requirements
            $requirementsGridView.DataSource = $requirementsList

            # Add debugging output to show the contents of $requirementsList after populating the DataGridView
            Write-Host "Requirements List after populating DataGridView:"
            $requirementsList | Format-Table -AutoSize
        } else {
            Write-Host "JSON Path does not exist: $jsonPath"
            # Handle error as needed
        }
    } else {
        Write-Host "No item selected in listBoxScripts"
        # Handle the error as needed
    }
})


#Testing to see if this is needed
$form.Controls.Add($listBoxScripts)

<# Loop through each button and update its color
foreach ($button in $form.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] }) {
    # Select a color from the "Summer" color set
    $selectedColor = $summerColors[(Get-Random -Minimum 0 -Maximum $summerColors.Count)]

    # Update the button's BackColor property
    $button.BackColor = $selectedColor
}
#>
[void]$form.ShowDialog()