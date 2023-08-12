clear-host

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define the base script path
$global:BaseScriptPath = Get-ConfigScriptPath
$global:BackupFolderPath = Join-Path -Path $global:BaseScriptPath -ChildPath "backup"
$global:preloadedMediaPlayer = $null
$global:textBoxBackupCount = $null

$colorList = @()

[System.Enum]::GetValues([System.Drawing.KnownColor]) | ForEach-Object {
    $color = [System.Drawing.Color]::FromKnownColor($_)
    $colorList += [PSCustomObject]@{
        Name = $_
        Color = $color
    }
}

function UpdateFormTitle {
    param (
        [System.Windows.Forms.Form]$form,
        [System.Windows.Forms.TextBox]$textBoxVersion,  # Pass the TextBox object as a parameter
        [string]$newVersion
    )

    $form.Text = "PowerShell Script Manager v$newVersion"
    $textBoxVersion.Text = $newVersion
}

function Get-ConfigScriptPath {
    $configPath = Join-Path -Path $global:DefaultPath -ChildPath 'config.json'
    if (Test-Path -Path $configPath) {
        $config = Get-Content -Path $configPath | ConvertFrom-Json
        return $config.ScriptPath
    }
    else {
        return $null
    }
}

function Save-Config {
    param ($path)
    $config = @{
        scriptPath = $path
    }
    $configFile = Join-Path $PSScriptRoot "config.json"
    $config | ConvertTo-Json | Set-Content -Path $configFile
}

function Set-ConfigScriptPath {
    param (
        [string]$path
    )
    $config = @{
        ScriptPath = $path
    } | ConvertTo-Json

    $configPath = Join-Path -Path $global:DefaultPath -ChildPath 'config.json'
    $config | Set-Content -Path $configPath
}

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

function Set-ScriptsPath {
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
    }
}

function Copy-ScriptsAndConfig {
    param (
        [string]$destinationPath
    )
    # Define the source paths for the files you want to copy
    $psmPath = Join-Path $PSScriptRoot "PSM.ps1"
    $jsonPath = Join-Path $PSScriptRoot "PSM.json"
    $configPath = Join-Path $PSScriptRoot "config.json"

    # Copy each file to the destination path
    Copy-Item -Path $psmPath -Destination $destinationPath
    Copy-Item -Path $jsonPath -Destination $destinationPath
    Copy-Item -Path $configPath -Destination $destinationPath
}

# Function to Get Scripts File Path by joining the user-defined scripts path with the provided script name and appending the .ps1 extension.
function Get-ScriptFilePath {
    param (
        [string]$scriptName
    )
    return Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$scriptName.ps1"
}

# Set Default Path
$global:DefaultPath = [System.Environment]::GetFolderPath('MyDocuments')

# Retrieve or Set Base Script Path
$global:BaseScriptPath = Get-ConfigScriptPath
if (-not $global:BaseScriptPath) {
    Set-ScriptsPath
}

# Function to Get JSON File Path by joining the user-defined scripts path with the provided script name and appending the .json extension.
function Get-JsonFilePath {
    param (
        [string]$scriptName
    )
    return Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$scriptName.json"
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

# Function to toggle between Tagging Mode and Search Mode
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

# Define the font size for tags and images in the ListBox
$fontSize = 11
$cellHeight = 40

function DisplayRandomQuote {
    $randomQuote = Get-Random -InputObject $inspirationalQuotes
    $richTextBoxOutput.Clear()
    $richTextBoxOutput.BackColor = [System.Drawing.Color]::LightGray
    $richTextBoxOutput.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 14, [System.Drawing.FontStyle]::Italic)
    $richTextBoxOutput.AppendText($randomQuote)
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

# Function to create a default script and JSON file
function CreateDefaultScriptAndJson {
    param (
        [string]$scriptPath
    )

    $defaultScriptPath = Join-Path -Path $scriptPath -ChildPath "DefaultScript.ps1"
    $defaultJsonPath = Join-Path -Path $scriptPath -ChildPath "DefaultScript.json"

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
                "Version" = "1.0" 
                "Tags" = @("🆕 New Script", "📜 PowerShell", "👩‍💻 Development", "🗑️ Junk")
                "ModifiedBy" = $env:username.ToUpper()
                "Modifications" = @(
                    @{
                        "Date" = (Get-Date -Format "MMMM d, yyyy 'at' h:mm tt")
                        "ModifiedBy" = $env:username.ToUpper()
                        "Modification" = "Default script created."
                        "Version" = "0.0"  
                    }
                )
            }
            
            # Calculate the initial version value based on existing modification entries
            $modifications = $defaultJsonContent.Modifications
            $highestVersion = Get-HighestVersionFromModifications $modifications
            
            # Update the initial version value in the JSON content
            $defaultJsonContent.Version = $highestVersion

            $defaultJsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $defaultJsonPath -Encoding UTF8
        }
    } catch {
        Write-Error $_
    }
}

# Helper function for validating JSON content
function ValidateJsonContent {
    param (
        $jsonContent,
        $requiredProperties,
        $selectedScript,
        $ScriptsPath
    )

    $defaultValues = @{
        'Name'          = $selectedScript.ToUpper()
        'Author'        = $env:USERNAME.ToUpper()
        'Description'   = "Please update this description"
        'Version'       = "0.0"
        'Tags'          = @("📜 PowerShell")
        'ModifiedBy'    = $env:USERNAME.ToUpper()
        'Modifications' = @()
        'ScriptJson'    = Join-Path -Path $ScriptsPath -ChildPath "$selectedScript.json"
        'ScriptPath'    = Join-Path -Path $ScriptsPath -ChildPath "$selectedScript.ps1"
    }

    # Check if all required properties exist, if not add them with a default value
    foreach ($property in $requiredProperties) {
        if ($property -notin $jsonContent.PSObject.Properties.Name) {
            $defaultValue = $defaultValues[$property]
            Add-Member -InputObject $jsonContent -NotePropertyName $property -NotePropertyValue $defaultValue
        }
    }

    # Check if any properties exist that aren't in the required list, if so remove them
    foreach ($property in $jsonContent.PSObject.Properties.Name) {
        if ($property -notin $requiredProperties) {
            $jsonContent.PSObject.Properties.Remove($property)
        }
    }

    # Validate that each property has an appropriate value. This will depend on what values you're expecting for each property.

    return $jsonContent
}

# Helper function for Updating JSON content
function UpdateJsonContent {
    param (
        $jsonContent,
        $selectedScript,
        $scriptPath,
        $jsonPath
    )


    $requiredProperties = @("Name", "ScriptPath", "ScriptJson", "Author", "Description", "Version", "Tags", "ModifiedBy", "Modifications")

    # Validate and clean up $jsonContent
    $jsonContent = ValidateJsonContent -jsonContent $jsonContent -requiredProperties $requiredProperties -selectedScript $selectedScript -ScriptsPath $jsonPath

    # Ensure $jsonContent is not null and is an object
    if ($null -eq $jsonContent) {
        $jsonContent = New-Object PSObject
    }

    # Ensure properties exist on $jsonContent
    if ($null -eq $jsonContent.ScriptPath) {
        Add-Member -InputObject $jsonContent -NotePropertyName ScriptPath -NotePropertyValue $scriptPath
    } else {
        $jsonContent.ScriptPath = $scriptPath
    }

    if ($null -eq $jsonContent.ScriptJson) {
        Add-Member -InputObject $jsonContent -NotePropertyName ScriptJson -NotePropertyValue $jsonPath
    } else {
        $jsonContent.ScriptJson = $jsonPath
    }

    $jsonContent.Name = $selectedScript.ToUpper()

    $jsonContent.Author = if ($jsonContent.Author -ne $null) { $jsonContent.Author } else { $env:USERNAME.ToUpper() }

    $jsonContent.Description = if ($jsonContent.Description -ne $null) { $jsonContent.Description } else { "New Script $selectedScript" }

    if ($null -eq $jsonContent.Version) {
        $jsonContent.Version = "0.1"  # Set the default version as 1.0 if the Version property is null
    }
    
    # Get the highest version from the modifications
    $highestVersion = Get-HighestVersionFromModifications $tableView

    # If the version is not set or is "1.0", set it to "0.1" for the first modification
    if ($jsonContent.Version -eq "1.0" -or [Version]$jsonContent.Version -eq [Version]::new(0, 0)) {
        $jsonContent.Version = "0.1"
    }

    # Update the version property
    $jsonContent.Version = $highestVersion.ToString()

    # Save the updated JSON content back to the file
    $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $JsonPath -Encoding UTF8

    if ($null -eq $jsonContent.Tags) {
        Add-Member -InputObject $jsonContent -NotePropertyName Tags -NotePropertyValue @("📜 PowerShell")
    } elseif ($jsonContent.Tags.Count -eq 0) {
        $jsonContent.Tags = @("📜 PowerShell")
    }

    if ($null -eq $jsonContent.ModifiedBy) {
        Add-Member -InputObject $jsonContent -NotePropertyName ModifiedBy -NotePropertyValue $env:USERNAME.ToUpper()
    } else {
        $jsonContent.ModifiedBy = $env:USERNAME.ToUpper()
    }

    if ($null -eq $jsonContent.Modifications) {
        Add-Member -InputObject $jsonContent -NotePropertyName Modifications -NotePropertyValue @()
    }

    # Delete unnecessary properties in the JSON
    $requiredProperties = @("Name", "ScriptPath", "ScriptJson", "Author", "Description", "Version", "Tags", "ModifiedBy", "Modifications")
    foreach ($property in $jsonContent.PSObject.Properties.Name) {
        if ($property -notin $requiredProperties) {
            $jsonContent.PSObject.Properties.Remove($property)
        }
    }

    return $jsonContent
}

# Function for Populating Fields and Tags
function PopulateFieldsAndTags {
    try {
        #Clear-Host
        $selectedScript = $listBoxScripts.SelectedItem
        $soundFilePath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "sounds\selectscript.mp3"
        PreloadCoolSound -SoundFileName "selectscript.mp3"

        if ($selectedScript) {
            
            $scriptPathWithoutExtension = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
            $scriptPath = $scriptPathWithoutExtension + ".ps1"
            $jsonPath = $scriptPathWithoutExtension + ".json"

            # Check if the script path exists
            if (-not (Test-Path -Path $scriptPath)) {
                Write-Host "Script path does not exist: $scriptPath"
                return
            }

            # Load or create json content
            if (Test-Path -Path $jsonPath) {
                $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
            } else {
                $jsonContent = @{
                    "Name" = $selectedScript
                    "ScriptPath" = $scriptPath
                    "ScriptJson" = $jsonPath
                    "Author" = $env:username
                    "Description" = "New Script $selectedScript"
                    "Version" = "0.0"  # Set the default version as 0.0
                    "Tags" = @("📜 PowerShell")
                    "ModifiedBy" = $env:username.ToUpper()
                    "Modifications" = @()
                }
            }

            # Call the helper function to update JSON content
            #$jsonContent = UpdateJsonContent -jsonContent $jsonContent -selectedScript $selectedScript -scriptPath $scriptPath -jsonPath $jsonPath

            # Save updated json content
            #$jsonContent | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8

            # Populate fields with the JSON content
            $fileInfo = Get-Item -Path $scriptPath
            $textBoxName.Text = $jsonContent.Name
            $textBoxDescription.Text = $jsonContent.Description
            $textBoxAuthor.Text = $jsonContent.Author
            $textBoxVersion.Text = $jsonContent.Version  # Update the version field
            $textBoxDateCreated.Text = $fileInfo.CreationTime.ToString("MMMM dd, yyyy 'at' hh:mm tt")
            $textBoxDateModified.Text = $fileInfo.LastWriteTime.ToString("MMMM dd, yyyy 'at' hh:mm tt")
            $textBoxModifiedBy.Text = $jsonContent.ModifiedBy.ToUpper()
            $textBoxJsonFile.Text = $jsonPath

            # Count the number of backup files and update the backup counter
            Update-BackupCounter -ScriptName $jsonContent.Name -BackupFolderPath $global:BackupFolderPath -LabelBackupCountValue $labelBackupCountValue

            # Clear and populate $listBoxScriptTags with all available tags
            $listBoxScriptTags.Items.Clear()
            $tags | ForEach-Object {
                $itemText = "$($_.Icon) $($_.Name)"
                $listBoxScriptTags.Items.Add($itemText)
            }

            # Set the font size for the ListBox
            $listBoxScriptTags.Font = New-Object System.Drawing.Font("Agency FB", $fontSize)
            $cellHeight = $cellHeight

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
    @{ Name = "Cloud"; Icon = "☁️" },
    @{ Name = "Configuration"; Icon = "⚙️" },
    @{ Name = "Counters"; Icon = "🔢" },
    @{ Name = "Databases"; Icon = "🗃️" },
    @{ Name = "DCAP"; Icon = "🔒" },
    @{ Name = "Deployment"; Icon = "🚀" },
    @{ Name = "Development"; Icon = "👩‍💻" },
    @{ Name = "Error Handling"; Icon = "❌" },
    @{ Name = "Excel"; Icon = "📊" },
    @{ Name = "Favorite"; Icon = "⭐" },
    @{ Name = "Finalized"; Icon = "🏁" },
    @{ Name = "Firewall"; Icon = "🔥" },
    @{ Name = "Fixing"; Icon = "🚧" },
    @{ Name = "Functions"; Icon = "🔧" },
    @{ Name = "High Priority"; Icon = "🔝" },
    @{ Name = "Hypervisor"; Icon = "🛠️" },
    @{ Name = "iDRAC"; Icon = "🖥️" },
    @{ Name = "Integration"; Icon = "🔗" },
    @{ Name = "Junk"; Icon = "🗑️" },
    @{ Name = "Loops"; Icon = "➰" },
    @{ Name = "Logging"; Icon = "📝" },
    @{ Name = "Maintenance"; Icon = "🔧" },
    @{ Name = "Mine"; Icon = "🏴" },
    @{ Name = "Monitoring"; Icon = "👀" },
    @{ Name = "Modules"; Icon = "📦" },
    @{ Name = "Networking"; Icon = "🌐" },
    @{ Name = "New Script"; Icon = "🆕" },
    @{ Name = "Not Mine"; Icon = "🏳️" },
    @{ Name = "Parameters"; Icon = "🛠️" },
    @{ Name = "PowerShell"; Icon = "📜" },
    @{ Name = "PS Dashboard"; Icon = "📊" },
    @{ Name = "Profiles"; Icon = "👤" },
    @{ Name = "Quality Control"; Icon = "✅" },
    @{ Name = "Remoting"; Icon = "🔗" },
    @{ Name = "Reports"; Icon = "📊" },
    @{ Name = "Security"; Icon = "🛡️" },
    @{ Name = "Secrets"; Icon = "🔐" },
    @{ Name = "SSH"; Icon = "🔑" },
    @{ Name = "Storage"; Icon = "🗄️" },
    @{ Name = "Switch"; Icon = "🔀" },
    @{ Name = "Tape Library"; Icon = "📼" },
    @{ Name = "Templates"; Icon = "📄" },
    @{ Name = "Tools"; Icon = "🛠️" },
    @{ Name = "User Management"; Icon = "👤" },
    @{ Name = "Virtual Machines"; Icon = "🖥️" },
    @{ Name = "Working"; Icon = "✔️" }
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


$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell Script Manager v$scriptVersionString"
$form.Size = New-Object System.Drawing.Size(880, 1000) # Increase W = Wider, H = Taller.
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::PowderBlue

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
$richTextBoxOutput.Location = New-Object System.Drawing.Point(15, 740) # Increase X = Right, Y = Down
$richTextBoxOutput.Size = New-Object System.Drawing.Size(835, 200)     # Increase W = Wider, H = Taller.
$richTextBoxOutput.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9) # Set font size
$richTextBoxOutput.ReadOnly = $true
$form.Controls.Add($richTextBoxOutput)

# Scripts Path Label
$labelScriptsPath = New-Object System.Windows.Forms.Label
$labelScriptsPath.Text = "Scripts Path:"
$labelScriptsPath.Location = New-Object System.Drawing.Point(15, 15) # Increase X = Right, y = Down
$labelScriptsPath.Size = New-Object System.Drawing.Size(110, 30)     # Increase W = Wider, H = Taller.
$labelScriptsPath.Font = New-Object System.Drawing.Font("Agency FB", 12)
$form.Controls.Add($labelScriptsPath)

# Script Path Text Box
$textBoxScriptsPath = New-Object System.Windows.Forms.TextBox
$textBoxScriptsPath.Location = New-Object System.Drawing.Point(135, 15) # Increase X = Right, y = Down
$textBoxScriptsPath.Size = New-Object System.Drawing.Size(475, 30)      # Increase W = Wider, H = Taller.
$textBoxScriptsPath.Font = New-Object System.Drawing.Font("Agency FB", 14, [System.Drawing.FontStyle]::Bold)
$textBoxScriptsPath.BorderStyle = "None"
$textBoxScriptsPath.BackColor = $form.BackColor
$textBoxScriptsPath.ForeColor = [System.Drawing.Color]::CadetBlue
$textBoxScriptsPath.ReadOnly = $true
$textBoxScriptsPath.Text = Get-ConfigScriptPath
$textBoxScriptsPath.CharacterCasing = "Upper"
$form.Controls.Add($textBoxScriptsPath)

# Change Path Button
$buttonChangePath = New-Object System.Windows.Forms.Button
$buttonChangePath.Text = "Change Path"
$buttonChangePath.Location = New-Object System.Drawing.Point(622, 10) # Increase X = Right, y = Down
$buttonChangePath.Size = New-Object System.Drawing.Size(110, 25)      # Increase W = Wider, H = Taller.
$buttonChangePath.BackColor = [System.Drawing.Color]::Gray
$buttonChangePath.ForeColor = [System.Drawing.Color]::WhiteSmoke
# Change Path Event Handler
$buttonChangePath.Add_Click({
    Set-ScriptsPath
})

$form.Controls.Add($buttonChangePath)

$buttonOpenFolder = New-Object System.Windows.Forms.Button
$buttonOpenFolder.Text = "Open Folder"
$buttonOpenFolder.Location = New-Object System.Drawing.Point(745, 10) # Modified location: Adjusted for longer textbox
$buttonOpenFolder.Size = New-Object System.Drawing.Size(110, 25) # Modified size: 10% longer
$buttonOpenFolder.BackColor = [System.Drawing.Color]::Gray
$buttonOpenFolder.ForeColor = [System.Drawing.Color]::WhiteSmoke

# Open Folder Event Handler
$buttonOpenFolder.Add_Click({
    Invoke-Item -Path $textBoxScriptsPath.Text
})

$form.Controls.Add($buttonOpenFolder)

# Button Load Scripts + ListBox Scripts Event Handler
$buttonLoadScripts = New-Object System.Windows.Forms.Button
$buttonLoadScripts.Text = "Load"
$buttonLoadScripts.Location = New-Object System.Drawing.Point(10, 46) # Increase X = Right, y = Down
$buttonLoadScripts.Size = New-Object System.Drawing.Size(80, 33)     # Increase W = Wider, H = Taller. 
#$buttonLoadScripts.BackColor = [System.Drawing.Color]::DarkBlue
$buttonLoadScripts.BackColor = [System.Drawing.Color]::NavajoWhite
$buttonLoadScripts.ForeColor = [System.Drawing.Color]::Black

# Event handler for the Load Scripts button
$buttonLoadScripts.Add_Click({
    # Get the full path of the script directory
    $scriptDirectory = $textBoxScriptsPath.Text
    #Clear-Host
    $listBoxScripts.Items.Clear()

    # Play the cool sound while loading scripts
    PlayCoolSound -SoundFileName "loadscripts.mp3"

    $richTextBoxOutput.Clear()
    $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Blue
    $richTextBoxOutput.AppendText("Welcome to PowerShell Script Manager $version`r`n`r`n")

    $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Arial", 10)
    $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Black
    $richTextBoxOutput.AppendText("PSM 2.4 empowers users with various operations for script management. Create, edit, rename, delete, execute, and back up scripts effortlessly. Each script is accompanied by a JSON file, capturing essential details, including modifications, authors, versions, and timestamps.

With a built-in tagging system, finding scripts is easier than ever. Efficiently search using keywords for streamlined script management and retrieval.")

    # Check if the script path exists
    $scriptPath = $textBoxScriptsPath.Text
    if (-not (Test-Path -Path $scriptPath)) {
        # Create a FolderBrowserDialog to let the user select the script path
        $scriptPath = Set-ScriptPath
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

# Button for Search Mode (Default button)
$buttonSearchMode = New-Object System.Windows.Forms.Button
$buttonSearchMode.Text = "Search"
$buttonSearchMode.Location = New-Object System.Drawing.Point(55, 510) # Increase X = Right, Y = Down
$buttonSearchMode.Size = New-Object System.Drawing.Size(100, 30) # Increase W = Wider, H = Taller
$buttonSearchMode.BackColor = [System.Drawing.Color]::GhostWhite  # Light blue color
$buttonSearchMode.ForeColor = [System.Drawing.Color]::CadetBlue  # Black text color
# Event handler for the "Search" button click
$buttonSearchMode.Add_Click({
    ShowTagSelectionForm
})
$form.Controls.Add($buttonSearchMode)

$labelDevelopedBy = New-Object System.Windows.Forms.Label
$labelDevelopedBy.Text = "Developed By: Curtis Dove"
$labelDevelopedBy.Location = New-Object System.Drawing.Point(18, 439) # Increase X = Right, y = Down
$labelDevelopedBy.Size = New-Object System.Drawing.Size(300, 20) # Increase W = Wider, H = Taller.
$labelDevelopedBy.Font = New-Object System.Drawing.Font("Georgia", 8, ([System.Drawing.FontStyle]::Italic))
$labelDevelopedBy.ForeColor = [System.Drawing.Color]::Gray
$labelDevelopedBy.BackColor = [System.Drawing.Color]::White
$labelDevelopedBy.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($labelDevelopedBy)

# Define the gap between buttons
$buttonGap = 5

# Calculate the position of the new buttons
$newButtonX = $buttonLoadScripts.Left + $buttonLoadScripts.Width + $buttonGap
$newButtonY = $buttonLoadScripts.Top

# Edit Script Button
$buttonEditScript = New-Object System.Windows.Forms.Button
$buttonEditScript.Text = "Edit"
#$buttonEditScript.Location = New-Object System.Drawing.Point(132, 46) # Increase X = Right, y = Down
$buttonEditScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY) # Increase X = Right, y = Down
$buttonEditScript.Size = New-Object System.Drawing.Size(80, 33)      # Increase W = Wider, H = Taller.
#$buttonEditScript.BackColor = [System.Drawing.Color]::Orange
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
$buttoncopyScript.Size = New-Object System.Drawing.Size(80, 33)      # Increase W = Wider, H = Taller.
#$buttoncopyScript.BackColor = [System.Drawing.Color]::LightBlue
$buttoncopyScript.BackColor = [System.Drawing.Color]::SandyBrown
$buttoncopyScript.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($buttoncopyScript)

# Calculate the position of the new buttons
$newButtonX = $buttonCopyScript.Left + $buttonCopyScript.Width + $buttonGap
$newButtonY = $buttonCopyScript.Top

# Create Script Button
$buttonCreateScript = New-Object System.Windows.Forms.Button
$buttonCreateScript.Text = "Create"
#$buttonCreateScript.Location = New-Object System.Drawing.Point(254, 46) # Increase X = Right, y = Down
$buttonCreateScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY) # Increase X = Right, y = Down
$buttonCreateScript.Size = New-Object System.Drawing.Size(80, 33)      # Increase W = Wider, H = Taller.
#$buttonCreateScript.BackColor = [System.Drawing.Color]::LimeGreen
$buttonCreateScript.BackColor = [System.Drawing.Color]::CadetBlue
$buttonCreateScript.ForeColor = [System.Drawing.Color]::Black
# Create Script Event Handler
$buttonCreateScript.Add_Click({

    $newScriptForm = New-Object System.Windows.Forms.Form
    $newScriptForm.Text = "Create New Script"
    $newScriptForm.Size = New-Object System.Drawing.Size(500, 760)        # Increase W = Wider, H = Taller.
    $newScriptForm.StartPosition = "CenterScreen"

    $labelNewScriptName = New-Object System.Windows.Forms.Label
    $labelNewScriptName.Text = "Script Name:"
    $labelNewScriptName.Location = New-Object System.Drawing.Point(10, 10) # Increase X = Right, y = Down
    $labelNewScriptName.Size = New-Object System.Drawing.Size(110, 22)     # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($labelNewScriptName)


    $textBoxNewScriptName = New-Object System.Windows.Forms.TextBox
    $textBoxNewScriptName.Location = New-Object System.Drawing.Point(132, 10) # Increase X = Right, y = Down
    $textBoxNewScriptName.Size = New-Object System.Drawing.Size(220, 22)      # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($textBoxNewScriptName)

    # Add a ListBox to select tags
    $labelTags = New-Object System.Windows.Forms.Label
    $labelTags.Text = "Tags:"
    $labelTags.Location = New-Object System.Drawing.Point(10, 40) # Increase X = Right, y = Down
    $labelTags.Size = New-Object System.Drawing.Size(110, 22)     # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($labelTags)

    $tagsListBox = New-Object System.Windows.Forms.ListBox
    $tagsListBox.Name = "tagsListBox"
    $tagsListBox.SelectionMode = "MultiExtended"
    $tagsListBox.Location = New-Object System.Drawing.Point(132, 40) # Increase X = Right, y = Down
    $tagsListBox.Size = New-Object System.Drawing.Size(320, 540)     # Increase W = Wider, H = Taller.
    $tagsListBox.MultiColumn = $true
    $tagsListBox.ColumnWidth = 100
    $tagsListBox.BackColor = [System.Drawing.Color]::LightGray
    $tagsListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $tagsListBox.Font = New-Object System.Drawing.Font("Arial", 10)

    # Add tags to the ListBox
    $tagsListBox.Items.AddRange($tags)

    $newScriptForm.Controls.Add($tagsListBox)

    # Add a TextBox for entering custom tags
    $labelCustomTag = New-Object System.Windows.Forms.Label
    $labelCustomTag.Text = "Custom Tag:"
    $labelCustomTag.Location = New-Object System.Drawing.Point(10, 600) # Increase X = Right, y = Down
    $labelCustomTag.Size = New-Object System.Drawing.Size(110, 22)      # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($labelCustomTag)

    $textBoxCustomTag = New-Object System.Windows.Forms.TextBox
    $textBoxCustomTag.Location = New-Object System.Drawing.Point(132, 600) # Increase X = Right, y = Down
    $textBoxCustomTag.Size = New-Object System.Drawing.Size(220, 22)       # Increase W = Wider, H = Taller.
    $newScriptForm.Controls.Add($textBoxCustomTag)

    # Add a Button to add the custom tag to the ListBox
    $buttonAddCustomTag = New-Object System.Windows.Forms.Button
    $buttonAddCustomTag.Text = "Add Tag"
    $buttonAddCustomTag.Location = New-Object System.Drawing.Point(10, 630) # Increase X = Right, y = Down
    $buttonAddCustomTag.Size = New-Object System.Drawing.Size(110, 33)      # Increase W = Wider, H = Taller.
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
    $buttonCreateNewScript.Size = New-Object System.Drawing.Size(110, 33)      # Increase W = Wider, H = Taller.
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

            $newScriptJson = @{
                Name = $newScriptName
                Description = "This is a new script."
                Version = 0.0
                Author = "Your Name"
                ModifiedBy = "Your Name"
                Tags = $tagsListBox.SelectedItems
                Modifications = @()
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
$buttonRenameScript.Size = New-Object System.Drawing.Size(80, 33)
$buttonRenameScript.BackColor = [System.Drawing.Color]::Coral
$buttonRenameScript.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($buttonRenameScript)

# Define colors
$primaryColor = [System.Drawing.Color]::FromArgb(30, 30, 30) # Dark Gray
$accentColor = [System.Drawing.Color]::FromArgb(255, 100, 100) # Light Red

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
        $inputBoxForm.BackColor = $primaryColor  # Set the background color
        $inputBoxForm.ForeColor = [System.Drawing.Color]::Black  # Set the foreground (text) color

        $inputLabel = New-Object System.Windows.Forms.Label
        $inputLabel.Text = "Enter the new name for the script:"
        $inputLabel.Location = New-Object System.Drawing.Point(10, 20)
        $inputLabel.Size = New-Object System.Drawing.Size(250, 22)
        $inputBoxForm.Controls.Add($inputLabel)

        $inputTextBox = New-Object System.Windows.Forms.TextBox
        $inputTextBox.Location = New-Object System.Drawing.Point(10, 50)
        $inputTextBox.Size = New-Object System.Drawing.Size(250, 22)
        $inputTextBox.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper  # Force text to be uppercase
        $inputBoxForm.Controls.Add($inputTextBox)

        $inputButtonOK = New-Object System.Windows.Forms.Button
        $inputButtonOK.Text = "OK"
        $inputButtonOK.Location = New-Object System.Drawing.Point(85, 80)
        $inputButtonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $inputButtonOK.BackColor = $accentColor  # Set button color
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

                    # Update the script name within the Modifications array
                    $currentDate = Get-Date
                    $currentDateFormatted = $currentDate.ToString("MMMM d, yyyy 'at' h:mm tt")
                    $newModification = @{
                        "Date" = $currentDateFormatted
                        "Modification" = "Renamed Script and JSON from $selectedScript to $newName"
                        "ModifiedBy" = $newName
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
$buttonDeleteScript.Size = New-Object System.Drawing.Size(80, 33)      # Increase W = Wider, H = Taller.
$buttonDeleteScript.BackColor = [System.Drawing.Color]::Red
$buttonDeleteScript.ForeColor = [System.Drawing.Color]::MistyRose

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
$buttonRunScript.Text = "Run Script"
#$buttonRunScript.Location = New-Object System.Drawing.Point(622, 46) # Increase X = Right, y = Down
$buttonRunScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY) # Increase X = Right, y = Down
$buttonRunScript.Size = New-Object System.Drawing.Size(80, 33)      # Increase W = Wider, H = Taller.
$buttonRunScript.BackColor = [System.Drawing.Color]::Green
$buttonRunScript.ForeColor = [System.Drawing.Color]::LightGreen

# Run Script Event Handler
$buttonRunScript.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        Start-Process pwsh -ArgumentList "-File `"$scriptPath`""
    }
})
$form.Controls.Add($buttonRunScript)

# Create the backup counter label
$labelBackupCounter = New-Object System.Windows.Forms.Label
$labelBackupCounter.Text = "Backups:"
$labelBackupCounter.Location = New-Object System.Drawing.Point(260, 100) # Increase X = Right, y = Down
$labelBackupCounter.AutoSize = $true
$labelBackupCounter.ForeColor = [System.Drawing.Color]::Black

# Backup Counter Value
$labelBackupCountValue = New-Object System.Windows.Forms.Label
$labelBackupCountValue.Text = "0"
$labelBackupCountValue.Location = New-Object System.Drawing.Point(320, 100)
$labelBackupCountValue.Size = New-Object System.Drawing.Size(40, 20)
$labelBackupCountValue.ForeColor = [System.Drawing.Color]::Red

# Calculate the position of the new buttons
$newButtonX = $buttonRunScript.Left + $buttonRunScript.Width + $buttonGap
$newButtonY = $buttonRunScript.Top

# Backup Script Button
$buttonBackupScript = New-Object System.Windows.Forms.Button
$buttonBackupScript.Text = "Backup"
#$buttonBackupScript.Location = New-Object System.Drawing.Point(745, 46)
$buttonBackupScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY)
$buttonBackupScript.Size = New-Object System.Drawing.Size(80, 33)
$buttonBackupScript.BackColor = [System.Drawing.Color]::Indigo
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
        $backupFiles = Get-ChildItem -Path $backupFolderPath -File | Where-Object { $_.Name -like "$scriptName*_backup*.*" }

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
        $backupFiles = Get-ChildItem -Path $backupFolderPath -File | Where-Object { $_.Name -like "$scriptName*_backup*.*" }

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

# Future Script Button
$buttonFutureScript = New-Object System.Windows.Forms.Button
$buttonFutureScript.Text = "Future"
#$buttonFutureScript.Location = New-Object System.Drawing.Point(745, 46)
$buttonFutureScript.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY)
$buttonFutureScript.Size = New-Object System.Drawing.Size(80, 33)
$buttonFutureScript.BackColor = [System.Drawing.Color]::GhostWhite
$buttonFutureScript.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($buttonFutureScript)

# Calculate the position of the new buttons
$newButtonX = $buttonFutureScript.Left + $buttonFutureScript.Width + $buttonGap
$newButtonY = $buttonFutureScript.Top

# Future2 Script Button
$buttonFuture2Script = New-Object System.Windows.Forms.Button
$buttonFuture2Script.Text = "Future2"
#$buttonFuture2Script.Location = New-Object System.Drawing.Point(745, 46)
$buttonFuture2Script.Location = New-Object System.Drawing.Point($newButtonX, $newButtonY)
$buttonFuture2Script.Size = New-Object System.Drawing.Size(80, 33)
$buttonFuture2Script.BackColor = [System.Drawing.Color]::GhostWhite
$buttonFuture2Script.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($buttonFuture2Script)


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
$labelScripts.Location = New-Object System.Drawing.Point(10, 100) # Increase X = Right, y = Down
$labelScripts.Size = New-Object System.Drawing.Size(150, 15)      # Increase W = Wider, H = Taller.
$labelScripts.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($labelScripts)

# Create the scripts ListBox
$listBoxScripts = New-Object System.Windows.Forms.ListBox
$listBoxScripts.Name = "listBoxScripts"
$listBoxScripts.Location = New-Object System.Drawing.Point(15, 120) # Increase X = Right, Y = Down
$listBoxScripts.Size = New-Object System.Drawing.Size(320, 355)   # Increase W = Wider, H = Taller.
$listBoxScripts.SelectionMode = "One"
$listBoxScripts.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($listBoxScripts)

# Commit Script & JSON to GIT Button
$buttonCommitToGit = New-Object System.Windows.Forms.Button
$buttonCommitToGit.Text = "Commit Script and JSON to GIT"
$buttonCommitToGit.Location = New-Object System.Drawing.Point(15, 458)
$buttonCommitToGit.Size = New-Object System.Drawing.Size(320, 25)
$buttonCommitToGit.BackColor = [System.Drawing.Color]::Teal
$buttonCommitToGit.ForeColor = [System.Drawing.Color]::White

# Commit Script & JSON to GIT Event Handler
$buttonCommitToGit.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPathWithoutExtension = Join-Path -Path $textBoxScriptsPath.Text -ChildPath $selectedScript
        $scriptPath = $scriptPathWithoutExtension + ".ps1"
        $jsonPath = $scriptPathWithoutExtension + ".json"

        # Ensure both the script and JSON file exist
        if (-not (Test-Path -Path $scriptPath)) {
            Write-Host "Script path does not exist: $scriptPath" -ForegroundColor Red
            return
        }

        if (-not (Test-Path -Path $jsonPath)) {
            Write-Host "JSON file path does not exist: $jsonPath" -ForegroundColor Red
            return
        }

        Write-Host "Script Path: $scriptPath" -ForegroundColor Yellow
        Write-Host "JSON Path: $jsonPath" -ForegroundColor Yellow

        $gitRepoPath = $global:BaseScriptPath
        Set-Location -Path $gitRepoPath


        # Check if the repository is already initialized
        $gitDir = Join-Path -Path $gitRepoPath -ChildPath ".git"
        $repoExists = Test-Path -Path $gitDir -PathType Container

        if (-not $repoExists) {
            # Initialize a new Git repository
            git init $repoPath
        }

        # Check repository status
        $status = & git status
        Write-Host "Git Status: $status"

        # Add and commit the script and JSON file to Git
        & git add $scriptPath
        & git add $jsonPath

        if ($LASTEXITCODE -eq 0) {
            Write-Host "Files added to Git staging area." -ForegroundColor Green

            $commitMessage = "Committed script $selectedScript and its JSON file"
            git commit -m $commitMessage

            if ($LASTEXITCODE -eq 0) {
                Write-Host "Changes committed successfully!" -ForegroundColor Green

                # Push changes to the remote repository on GitHub
                git push origin master

                if ($LASTEXITCODE -eq 0) {
                    Write-Host "Changes pushed to remote repository on GitHub." -ForegroundColor Green
                } else {
                    Write-Host "Failed to push changes to remote repository on GitHub." -ForegroundColor Red
                }
            } else {
                Write-Host "Failed to commit changes to Git." -ForegroundColor Red
            }
        } else {
            Write-Host "Failed to add files to Git staging area." -ForegroundColor Red
        }
    } else {
        Write-Host "No script selected." -ForegroundColor Yellow
    }
})





$form.Controls.Add($buttonCommitToGit)


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
$tagsListBox.Font = New-Object System.Drawing.Font("Arial", 10)

try {
    $tagsListBox.Items.AddRange($tags)
} catch {
    Write-Error $_
}

################################ FIELDS ON RIGHT HAND SIDE OF FARM START HERE #############################

###################################### NAME LABEL & TEXT BOX  #############################################

# Name Label
$labelName = New-Object System.Windows.Forms.Label
$labelName.Text = "Name:"
$labelName.Location = New-Object System.Drawing.Point(350, 100)   # Increase X = Right, y = Down
$labelName.Size = New-Object System.Drawing.Size(100, 20)         # Increase W = Wider, H = Taller.
$form.Controls.Add($labelName)

# Name Text Box
$textBoxName = New-Object System.Windows.Forms.TextBox
$textBoxName.Location = New-Object System.Drawing.Point(450, 100) # Increase X = Right, y = Down
$textBoxName.Size = New-Object System.Drawing.Size(400, 100)      # Increase W = Wider, H = Taller.
$textBoxName.Font = New-Object System.Drawing.Font("Agency FB", 16, [System.Drawing.FontStyle]::Bold)
#$textBoxName.BorderStyle = "None"
$textBoxName.BackColor = $form.BackColor
$textBoxName.ForeColor = [System.Drawing.Color]::CadetBlue
$textBoxName.CharacterCasing = "Upper"
$form.Controls.Add($textBoxName)

###################################### NAME LABEL & TEXT BOX  #############################################

################################# DESCRIPTION LABEL & TEXT BOX  ###########################################

# Label Description Label
$labelDescription = New-Object System.Windows.Forms.Label
$labelDescription.Text = "Description:"
$labelDescription.Location = New-Object System.Drawing.Point(350, 140) # Increase X = Right, y = Down
$labelDescription.Size = New-Object System.Drawing.Size(100, 20)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelDescription)

# Label Description TextBox
$textBoxDescription = New-Object System.Windows.Forms.TextBox
$textBoxDescription.Location = New-Object System.Drawing.Point(450, 140)
$textBoxDescription.Size = New-Object System.Drawing.Size(400, 110)
$textBoxDescription.Multiline = $true
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
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Green
                $richTextBoxOutput.AppendText("Description for script '$selectedScript' updated successfully!" + [Environment]::NewLine)

                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Blue

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

################################### DESCRIPTION LABEL & TEXT BOX  #########################################

##################################### VERSION LABEL & TEXT BOX  ###########################################

# Label Version
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "Version:"
$labelVersion.Location = New-Object System.Drawing.Point(350, 259) # Increase X = Right, y = Down
$labelVersion.Size = New-Object System.Drawing.Size(100, 20)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelVersion)

# TextBox Version
$textBoxVersion = New-Object System.Windows.Forms.TextBox
$textBoxVersion.Location = New-Object System.Drawing.Point(450, 259) # Increase X = Right, y = Down
$textBoxVersion.Size = New-Object System.Drawing.Size(400, 20)       # Increase W = Wider, H = Taller.
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
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Green
                $richTextBoxOutput.AppendText("Version for script '$selectedScript' updated successfully!" + [Environment]::NewLine)

                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Blue

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


##################################### VERSION LABEL & TEXT BOX  ###########################################

##################################### AUTHOR LABEL & TEXT BOX  ###########################################

# Author Label
$labelAuthor = New-Object System.Windows.Forms.Label
$labelAuthor.Text = "Author:"
$labelAuthor.Location = New-Object System.Drawing.Point(350, 289) # Increase X = Right, y = Down
$labelAuthor.Size = New-Object System.Drawing.Size(100, 20)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelAuthor)

# Author TextBox
$textBoxAuthor = New-Object System.Windows.Forms.TextBox
$textBoxAuthor.Location = New-Object System.Drawing.Point(450, 289) # Increase X = Right, y = Down
$textBoxAuthor.Size = New-Object System.Drawing.Size(400, 20)       # Increase W = Wider, H = Taller.
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
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Green
                $richTextBoxOutput.AppendText("Author for script '$selectedScript' updated successfully!" + [Environment]::NewLine)

                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Blue

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


##################################### AUTHOR LABEL & TEXT BOX  ###########################################

# Date Created Label
$labelDateCreated = New-Object System.Windows.Forms.Label
$labelDateCreated.Text = "Date Created:"
$labelDateCreated.Location = New-Object System.Drawing.Point(350, 319) # Increase X = Right, y = Down
$labelDateCreated.Size = New-Object System.Drawing.Size(100, 20)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelDateCreated)

# Date Created Textbox
$textBoxDateCreated = New-Object System.Windows.Forms.TextBox
$textBoxDateCreated.Location = New-Object System.Drawing.Point(450, 319) # Increase X = Right, y = Down
$textBoxDateCreated.Size = New-Object System.Drawing.Size(400, 20)       # Increase W = Wider, H = Taller.
#$textBoxDateCreated.BackColor = [System.Drawing.Color]::White
$textBoxDateCreated.ReadOnly = $true
$form.Controls.Add($textBoxDateCreated)

# Date Modified Label
$labelDateModified = New-Object System.Windows.Forms.Label
$labelDateModified.Text = "Date Modified:"
$labelDateModified.Location = New-Object System.Drawing.Point(350, 349) # Increase X = Right, y = Down
$labelDateModified.Size = New-Object System.Drawing.Size(100, 20)      # Increase W for wider, H for taller
$form.Controls.Add($labelDateModified)

# Date Modified TextBox
$textBoxDateModified = New-Object System.Windows.Forms.TextBox
$textBoxDateModified.Location = New-Object System.Drawing.Point(450, 349) # Increase X = Right, y = Down
$textBoxDateModified.Size = New-Object System.Drawing.Size(200, 20)      # Same size as author and created.
#$textBoxDateModified.BackColor = [System.Drawing.Color]::White
$textBoxDateModified.ReadOnly = $true
$form.Controls.Add($textBoxDateModified)

# Modified By Label
$labelModifiedBy = New-Object System.Windows.Forms.Label
$labelModifiedBy.Text = "Modified By:"
$labelModifiedBy.Location = New-Object System.Drawing.Point(350, 379) # Increase X = Right, y = Down
$labelModifiedBy.Size = New-Object System.Drawing.Size(100, 20)      # Increase W for wider, H for taller
$form.Controls.Add($labelModifiedBy)

# Modified By TextBox
$textBoxModifiedBy = New-Object System.Windows.Forms.TextBox
$textBoxModifiedBy.Location = New-Object System.Drawing.Point(450, 379) # Increase X = Right, y = Down
$textBoxModifiedBy.Size = New-Object System.Drawing.Size(400, 20)       # Increase W for wider, H for taller
$textBoxModifiedBy.BackColor = [System.Drawing.Color]::White
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
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Green
                $richTextBoxOutput.AppendText("ModifiedBy for script '$selectedScript' updated successfully!" + [Environment]::NewLine)

                $richTextBoxOutput.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
                $richTextBoxOutput.SelectionColor = [System.Drawing.Color]::Blue

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

# Get highest version from Modification Log
function Get-HighestVersionFromModifications {
    param (
        [array]$modifications,
        [string]$jsonPath
    )

    Write-Host "Inside Get-HighestVersionFromModifications function"
    Write-Host "JSON path received: $jsonPath"
    Write-Host "Number of modifications received: $($modifications.Count)"

    $versions = $modifications | ForEach-Object { $_.Version }
    $highestVersion = [version]($versions | Sort-Object -Descending | Select-Object -First 1)
    
    # Read the JSON content
    $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
    
    # Update the version in the JSON content
    $jsonContent.Version = $highestVersion.ToString()
    
    # Save the updated JSON content back to the file
    $jsonContent | ConvertTo-Json -Depth 4 | Out-File -FilePath $jsonPath -Encoding UTF8
    
    return $highestVersion, $jsonContent
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
    $modLogForm.Size = New-Object System.Drawing.Size(820, 640)  # Increase W for wider, H for taller
    $modLogForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Table View
    $tableView = New-Object System.Windows.Forms.DataGridView
    $tableView.Location = New-Object System.Drawing.Point(20, 20) # Increase X = Right, y = Down
    $tableView.Size = New-Object System.Drawing.Size(760, 400)  # Increase W for wider, H for taller
    $tableView.ColumnHeadersVisible = $true
    $tableView.RowHeadersVisible = $false
    $tableView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $tableView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
    $tableView.AllowUserToAddRows = $false
    $tableView.AllowUserToDeleteRows = $false
    $tableView.AllowUserToResizeRows = $false
    $tableView.MultiSelect = $false

    # Create "Delete" button column
    $deleteColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $deleteColumn.HeaderText = "Delete"
    $deleteColumn.Text = "Delete"
    $deleteColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $deleteColumn.UseColumnTextForButtonValue = $true
    $tableView.Columns.Add($deleteColumn)

    $dateColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $dateColumn.Name = "Date"
    $dateColumn.HeaderText = "Date"
    $dateColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $tableView.Columns.Add($dateColumn)

    $versionColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $versionColumn.Name = "Version"
    $versionColumn.HeaderText = "Version"
    $versionColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $versionColumn.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    $versionColumn.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
    $tableView.Columns.Add($versionColumn)

    $modifiedByColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $modifiedByColumn.Name = "ModifiedBy"
    $modifiedByColumn.HeaderText = "Modified By"
    $modifiedByColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
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
    $labelModifiedDate.Location = New-Object System.Drawing.Point(20, 430) # Adjust coordinates as needed
    $labelModifiedDate.Size = New-Object System.Drawing.Size(100, 20)
    $modLogForm.Controls.Add($labelModifiedDate)

    # TextBox for "Modified Date Text"
    $textBoxModifiedDate = New-Object System.Windows.Forms.TextBox
    $textBoxModifiedDate.Location = New-Object System.Drawing.Point(140, 430) # Adjust coordinates as needed
    $textBoxModifiedDate.Size = New-Object System.Drawing.Size(220, 20) # Increase width for longer date format
    $textBoxModifiedDate.BackColor = [System.Drawing.Color]::White
    $textBoxModifiedDate.Text = Get-Date -Format "MMMM d, yyyy 'at' h:mm tt" # Prepopulate with current date and time
    $modLogForm.Controls.Add($textBoxModifiedDate)

    # Label and TextBox for "Modified By"
    $labelAddModifiedBy = New-Object System.Windows.Forms.Label
    $labelAddModifiedBy.Text = "Modified By:"
    $labelAddModifiedBy.Location = New-Object System.Drawing.Point(380, 430) # Move right for additional field
    $labelAddModifiedBy.Size = New-Object System.Drawing.Size(100, 20)
    $modLogForm.Controls.Add($labelAddModifiedBy)

    $textBoxAddModifiedBy = New-Object System.Windows.Forms.TextBox
    $textBoxAddModifiedBy.Location = New-Object System.Drawing.Point(500, 430) # Move right for additional field
    $textBoxAddModifiedBy.Size = New-Object System.Drawing.Size(100, 20) # Adjust width as needed
    $textBoxAddModifiedBy.BackColor = [System.Drawing.Color]::White
    $textBoxAddModifiedBy.Text = $env:USERNAME.ToUpper() # Prepopulate with the current user in uppercase
    $modLogForm.Controls.Add($textBoxAddModifiedBy)

    # Label for "Modification"
    $labelModification = New-Object System.Windows.Forms.Label
    $labelModification.Text = "Modification:"
    $labelModification.Location = New-Object System.Drawing.Point(20, 460) # Move down for additional field
    $labelModification.Size = New-Object System.Drawing.Size(100, 20)
    $modLogForm.Controls.Add($labelModification)

    # TextBox for "Modification Text"
    $textBoxAddModificationText = New-Object System.Windows.Forms.TextBox
    $textBoxAddModificationText.Multiline = $true # Allow multiple lines
    $textBoxAddModificationText.ScrollBars = "Vertical" # Show vertical scroll bar for long text
    $textBoxAddModificationText.Location = New-Object System.Drawing.Point(140, 460) # Move down for additional field
    $textBoxAddModificationText.Size = New-Object System.Drawing.Size(640, 80) # Adjust height and width as needed
    $textBoxAddModificationText.BackColor = [System.Drawing.Color]::White
    $modLogForm.Controls.Add($textBoxAddModificationText)

    # Create "Add Modification" button
    $buttonSaveModification = New-Object System.Windows.Forms.Button
    $buttonSaveModification.Text = "Add Modification" # Change button text
    $buttonSaveModification.Location = New-Object System.Drawing.Point(15, 560) # Move down for additional field
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
    $labelAddVersion.Location = New-Object System.Drawing.Point(20, 20) # Adjust coordinates as needed
    $labelAddVersion.Size = New-Object System.Drawing.Size(100, 20)
    $addModForm.Controls.Add($labelAddVersion)

    $textBoxAddVersion = New-Object System.Windows.Forms.TextBox
    $textBoxAddVersion.Location = New-Object System.Drawing.Point(140, 20) # Adjust coordinates as needed
    $textBoxAddVersion.Size = New-Object System.Drawing.Size(220, 20) # Adjust size as needed
    $textBoxAddVersion.BackColor = [System.Drawing.Color]::White
    $textBoxAddVersion.Text = $jsonContent.Version  # Prepopulate with the current version from JSON content
    $textBoxAddVersion.Enabled = $false # Disable editing of the version field
    $addModForm.Controls.Add($textBoxAddVersion)

    # Label and TextBox for "Modified By"
    $labelAddModifiedBy = New-Object System.Windows.Forms.Label
    $labelAddModifiedBy.Text = "Modified By:"
    $labelAddModifiedBy.Location = New-Object System.Drawing.Point(20, 50) # Adjust coordinates as needed
    $labelAddModifiedBy.Size = New-Object System.Drawing.Size(100, 20)      # Increase W for wider, H for taller
    $addModForm.Controls.Add($labelAddModifiedBy)

    $textBoxAddModifiedBy = New-Object System.Windows.Forms.TextBox
    $textBoxAddModifiedBy.Location = New-Object System.Drawing.Point(140, 50) # Adjust coordinates as needed
    $textBoxAddModifiedBy.Size = New-Object System.Drawing.Size(220, 20)      # Adjust size as needed
    $textBoxAddModifiedBy.BackColor = [System.Drawing.Color]::White
    $textBoxAddModifiedBy.Text = $env:USERNAME.ToUpper()  # Populate with the current user's name
    $addModForm.Controls.Add($textBoxAddModifiedBy)

    # Label and TextBox for "Modification Text"
    $labelAddModificationText = New-Object System.Windows.Forms.Label
    $labelAddModificationText.Text = "Modification Text:"
    $labelAddModificationText.Location = New-Object System.Drawing.Point(20, 80) # Adjust coordinates as needed
    $labelAddModificationText.Size = New-Object System.Drawing.Size(100, 20)      # Increase W for wider, H for taller
    $addModForm.Controls.Add($labelAddModificationText)

    $textBoxAddModificationText = New-Object System.Windows.Forms.TextBox
    $textBoxAddModificationText.Location = New-Object System.Drawing.Point(140, 80) # Adjust coordinates as needed
    $textBoxAddModificationText.Size = New-Object System.Drawing.Size(220, 180)      # Adjust size as needed
    $textBoxAddModificationText.BackColor = [System.Drawing.Color]::White
    $addModForm.Controls.Add($textBoxAddModificationText)

    # Create "Save" button for saving the new modification
    $buttonSaveModification = New-Object System.Windows.Forms.Button
    $buttonSaveModification.Text = "Save"
    $buttonSaveModification.Location = New-Object System.Drawing.Point(200, 260)  # Adjust coordinates as needed
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


# Create Modification Log button
$buttonModificationLog = New-Object System.Windows.Forms.Button
$buttonModificationLog.Text = "View Modification Log"
$buttonModificationLog.Location = New-Object System.Drawing.Point(655, 345)  # Increase X = Right, Y = Down
$buttonModificationLog.Size = New-Object System.Drawing.Size(195, 31)        # Increase W = Wider, H = Taller.
$buttonModificationLog.BackColor = [System.Drawing.Color]::LightPink
$buttonModificationLog.Add_Click({ ShowModificationLog })
$form.Controls.Add($buttonModificationLog)

# JSON File Label
$labelJsonFile = New-Object System.Windows.Forms.Label
$labelJsonFile.Text = "JSON File:"
$labelJsonFile.Location = New-Object System.Drawing.Point(350, 410) # Increase X = Right, y = Down
$labelJsonFile.Size = New-Object System.Drawing.Size(100, 20)       # Increase W = Wider, H = Taller.
$form.Controls.Add($labelJsonFile)

# JSON File Textbox
$textBoxJsonFile = New-Object System.Windows.Forms.TextBox
$textBoxJsonFile.Location = New-Object System.Drawing.Point(450, 410) # Increase X = Right, y = Down
$textBoxJsonFile.Size = New-Object System.Drawing.Size(400, 20)       # Increase W = Wider, H = Taller.
$textBoxJsonFile.CharacterCasing = "Upper"
#$textBoxJsonFile.BackColor = [System.Drawing.Color]::Pink
$textBoxJsonFile.ReadOnly = $true
$form.Controls.Add($textBoxJsonFile)

# Open JSON Button
$buttonOpenJson = New-Object System.Windows.Forms.Button
$buttonOpenJson.Text = "OPEN JSON"
$buttonOpenJson.Location = New-Object System.Drawing.Point(450, 435) # Increase X = Right, Y = Down
$buttonOpenJson.Size = New-Object System.Drawing.Size(130, 22)       # Increase W = Wider, H = Taller.
$buttonOpenJson.BackColor = [System.Drawing.Color]::MintCream
$buttonOpenJson.ForeColor = [System.Drawing.Color]::Black

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

# Add the "Open JSON" button to the form
$form.Controls.Add($buttonOpenJson)

# Add controls to the form
$form.Controls.Add($labelBackupCounter)
$form.Controls.Add($labelBackupCountValue)

# Delete JSON Button
$buttonDeleteJson = New-Object System.Windows.Forms.Button
$buttonDeleteJson.Text = "DELETE JSON"
$buttonDeleteJson.Location = New-Object System.Drawing.Point(585, 435) # Increase X = Right, y = Down
$buttonDeleteJson.Size = New-Object System.Drawing.Size(130, 22)       # Increase W = Wider, H = Taller.
$buttonDeleteJson.BackColor = [System.Drawing.Color]::MintCream
$buttonDeleteJson.ForeColor = [System.Drawing.Color]::Black
# Delete JSON Button Event Handler
$buttonDeleteJson.Add_Click({
    $selectedScript = $listBoxScripts.SelectedItem
    if ($selectedScript) {
        $scriptPath = Join-Path -Path $textBoxScriptsPath.Text -ChildPath "$selectedScript.ps1"
        $jsonPath = $scriptPath -replace "\.ps1$", ".json"

        if (Test-Path -Path $jsonPath) {
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

# Save JSON Button
$buttonSave = New-Object System.Windows.Forms.Button
$buttonSave.Text = "SAVE JSON"
$buttonSave.Location = New-Object System.Drawing.Point(720, 435) # Increase X = Right, y = Down
$buttonSave.Size = New-Object System.Drawing.Size(130, 22)       # Increase W = Wider, H = Taller.
#$buttonSave.BackColor = [System.Drawing.Color]::Pink
#$buttonSave.ForeColor = [System.Drawing.Color]::Black
$buttonSave.Visible = $false 
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
            [PSCustomObject]@{ Name = "Modifications"; Value = @() }
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

# Create a label for the Tags section
$labelTags = New-Object System.Windows.Forms.Label
$labelTags.Text = "Tags:"
$labelTags.Location = New-Object System.Drawing.Point(10, 515 )  # Increase X = Right, y = Down
$labelTags.Size = New-Object System.Drawing.Size(100, 20) # Increase W = Wider, H = Taller.
$form.Controls.Add($labelTags)

# ListBox for Script Tags
$listBoxScriptTags = New-Object System.Windows.Forms.ListBox
$listBoxScriptTags.Location = New-Object System.Drawing.Point(15, 545)  # Increase X = Right, y = Down
$listBoxScriptTags.Size = New-Object System.Drawing.Size(835, 180)       # Increase W = Wider, H = Taller.
$listBoxScriptTags.SelectionMode = "MultiSimple"
$listBoxScriptTags.DisplayMember = "Text"
$listBoxScriptTags.ValueMember = "Tag"
$listBoxScriptTags.ColumnWidth = 136
$listBoxScriptTags.MultiColumn = $true
$listBoxScriptTags.Font = New-Object System.Drawing.Font("Arial", 16)  # Change the font family and size her
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

$listBoxScripts.Add_SelectedIndexChanged({
    PopulateFieldsAndTags
})

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