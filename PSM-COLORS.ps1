# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Color Palette'
$form.Width = 1620
$form.Height = 920

# Set font size
$fontSize = 12

# Determine button size
$buttonWidth = 200
$buttonHeight = 40

# Create buttons for each color
$row = 0
$col = 0

# Iterate through known colors
[System.Enum]::GetValues([System.Drawing.KnownColor]) | ForEach-Object {
    $color = [System.Drawing.Color]::FromKnownColor($_)

    # Create button
    $button = New-Object System.Windows.Forms.Button
    $button.Width = $buttonWidth
    $button.Height = $buttonHeight
    $button.BackColor = $color
    $button.Text = $color.Name
    $button.Font = New-Object System.Drawing.Font('Arial', $fontSize)

    # Check if the color is dark or light
    $brightness = $color.GetBrightness()
    if ($brightness -lt 0.5) {
        $button.ForeColor = [System.Drawing.Color]::White
    } else {
        $button.ForeColor = [System.Drawing.Color]::Black
    }

    # Calculate position
    $x = $col * $buttonWidth
    $y = $row * $buttonHeight

    # Position button
    $button.Location = New-Object System.Drawing.Point($x, $y)
    $form.Controls.Add($button)

    # Update row and column, limiting columns to 8
    $col++
    if ($col -ge 8) {
        $col = 0
        $row++
    }
}

# Show Form
$form.ShowDialog()