<#
.SYNOPSIS
    Automated PowerPoint & Certificate Generator (CSV to PPTX).

.DESCRIPTION
    Originally designed for a school environment to automate "Time to Shine" awards. 
    This script utilizes the PowerPoint COM object to map CSV data onto custom 
    backgrounds with intelligent font-scaling and automated transitions.

.TECHNICAL NOTES
    - Column Headers: Defaults to 'Student' and 'Staff'. Adjust the foreach loop 
      logic if your schema differs.
    - Dependencies: Requires local installation of Microsoft PowerPoint.

.DISCLAIMER & LIMITATION OF LIABILITY
    *** USE AT YOUR OWN RISK ***
    This script is provided "as is" without warranty of any kind, express or implied.  
    Always test with sample data before deploying to staff or production machines.

.LICENSE
    Distributed under the MIT License. (See below)
#>

# Copyright (c) 2026 [Neil Crofts/Grimbarian]
# Licensed under the MIT License.
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Main Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "PPT Awards Generator"
$form.Size = New-Object System.Drawing.Size(600, 480)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# --- Data File Section ---
$btnData = New-Object System.Windows.Forms.Button
$btnData.Text = "1. Select Data (CSV)"
$btnData.Location = New-Object System.Drawing.Point(20, 20)
$btnData.Size = New-Object System.Drawing.Size(150, 30)

$txtData = New-Object System.Windows.Forms.TextBox
$txtData.Location = New-Object System.Drawing.Point(180, 24)
$txtData.Size = New-Object System.Drawing.Size(380, 30)
$txtData.ReadOnly = $true

$btnData.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Data Files (*.csv)|*.csv"
    if ($dialog.ShowDialog() -eq "OK") { $txtData.Text = $dialog.FileName }
})

# --- Image Section ---
$btnImage = New-Object System.Windows.Forms.Button
$btnImage.Text = "2. Select Background"
$btnImage.Location = New-Object System.Drawing.Point(20, 70)
$btnImage.Size = New-Object System.Drawing.Size(150, 30)

$txtImage = New-Object System.Windows.Forms.TextBox
$txtImage.Location = New-Object System.Drawing.Point(180, 74)
$txtImage.Size = New-Object System.Drawing.Size(380, 30)
$txtImage.ReadOnly = $true

$btnImage.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Image Files (*.png;*.jpg)|*.png;*.jpg"
    if ($dialog.ShowDialog() -eq "OK") { $txtImage.Text = $dialog.FileName }
})

# --- Custom Labels & Checkboxes ---

# Title 1
$chkTitle1 = New-Object System.Windows.Forms.CheckBox
$chkTitle1.Text = "Award Heading? Type  your certificate heading below if needed."
$chkTitle1.Location = New-Object System.Drawing.Point(20, 115)
$chkTitle1.Size = New-Object System.Drawing.Size(540, 20)
$chkTitle1.Checked = $true

$txtTitle = New-Object System.Windows.Forms.Textbox
$txtTitle.Location = New-Object System.Drawing.Point(20, 140)
$txtTitle.Size = New-Object System.Drawing.Size(540, 20)
$txtTitle.Text = "Congratulations on your Time to Shine nomination!"

# Title 2 (Subtext)
$chkTitle2 = New-Object System.Windows.Forms.CheckBox
$chkTitle2.Text = "Award SubText? (e.g Nominated by, Date of Award, Awarded For, etc)"
$chkTitle2.Location = New-Object System.Drawing.Point(20, 165)
$chkTitle2.Size = New-Object System.Drawing.Size(540, 20)
$chkTitle2.Checked = $true

$txtTitle2 = New-Object System.Windows.Forms.Textbox
$txtTitle2.Location = New-Object System.Drawing.Point(20, 190)
$txtTitle2.Size = New-Object System.Drawing.Size(540, 20)
$txtTitle2.Text = "Nominated by"

# Image Export Checkbox
$chkExport = New-Object System.Windows.Forms.CheckBox
$chkExport.Text = "Check this box to Export all slides as individual images.(PNG)"
$chkExport.Location = New-Object System.Drawing.Point(20, 220)
$chkExport.Size = New-Object System.Drawing.Size(540, 20)
$chkExport.Checked = $false

# --- Save Section ---
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "3. Save Location"
$btnSave.Location = New-Object System.Drawing.Point(20, 260)
$btnSave.Size = New-Object System.Drawing.Size(150, 30)

$txtSave = New-Object System.Windows.Forms.TextBox
$txtSave.Location = New-Object System.Drawing.Point(180, 264)
$txtSave.Size = New-Object System.Drawing.Size(380, 30)
$txtSave.ReadOnly = $true

$btnSave.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "PowerPoint File (*.pptx)|*.pptx"
    $dialog.FileName = "Awards_$(Get-Date -Format 'dd_MM_yy').pptx"
    if ($dialog.ShowDialog() -eq "OK") { $txtSave.Text = $dialog.FileName }
})

# --- RUN BUTTON ---
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "CREATE POWERPOINT"
$btnRun.Location = New-Object System.Drawing.Point(20, 310)
$btnRun.Size = New-Object System.Drawing.Size(540, 60)
$btnRun.BackColor = [System.Drawing.Color]::LightGreen

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 380)
$progressBar.Size = New-Object System.Drawing.Size(540, 20)
$progressBar.Style = "Blocks"

function CreatePowerPoint {
    try {
        if ([string]::IsNullOrWhiteSpace($txtData.Text) -or [string]::IsNullOrWhiteSpace($txtImage.Text) -or [string]::IsNullOrWhiteSpace($txtSave.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please ensure you have selected a CSV file, a Background Image, and a Save Location.", "Missing Information")
        return
        }

        $nomins = Import-Csv -Path $txtData.Text
        $BackgroundImagePath = $txtImage.Text
        $saveTitle = $txtSave.Text

        $progressBar.Value = 0
        $progressBar.Maximum = $nomins.Count
        $stepCount = 0

        $PowerPoint = New-Object -ComObject PowerPoint.Application
        $Presentation = $PowerPoint.Presentations.Add()
        $SlideWidth = $Presentation.PageSetup.SlideWidth

        # Prepare Image Export Folder
        if ($chkExport.Checked) {
            $imgFolder = Join-Path (Split-Path $saveTitle) "Award_Images"
            if (!(Test-Path $imgFolder)) { New-Item -ItemType Directory -Path $imgFolder }
        }

        foreach ($row in $nomins) {
            $stepCount++
            $name = $row.NomineeName
            $subtext = $row.SubText
            
            # Add Slide (12 = Blank Layout for full control)
            $Stm = $Presentation.Slides.Add($stepCount, 12) 
            
            # --- 1. Congrats Heading ---
            if ($chkTitle1.Checked) {
                $txtCongrats = $Stm.Shapes.AddTextBox(1, 100, 220, 550, 50)
                $txtCongrats.TextFrame.TextRange.Text = $txtTitle.Text
                $txtCongrats.TextFrame.TextRange.Font.Size = 24
                $txtCongrats.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                $txtCongrats.Left = ($SlideWidth - 550) / 2
            }

            # --- 2. Student Name (With Autosize) ---
            $txtName = $Stm.Shapes.AddTextBox(1, 100, 260, 550, 80)
            $nameRange = $txtName.TextFrame.TextRange
            $nameRange.Text = $name
            $nameRange.Font.Bold = $true
            
            # Autosize logic
            if ($name.Length -gt 25) { $nameRange.Font.Size = 36 }
            elseif ($name.Length -gt 15) { $nameRange.Font.Size = 44 }
            else { $nameRange.Font.Size = 54 }

            $txtName.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $txtName.Left = ($SlideWidth - 550) / 2

            # --- 3. Nominated By ---
            if ($chkTitle2.Checked) {
                $txtNom = $Stm.Shapes.AddTextBox(1, 100, 350, 550, 50)
                $txtNom.TextFrame.TextRange.Text = "$($txtTitle2.Text) $subtext"
                $txtNom.TextFrame.TextRange.Font.Size = 24
                $txtNom.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                $txtNom.Left = ($SlideWidth - 550) / 2
            }

            # Set Background
            $Stm.FollowMasterBackground = $false
            $Stm.Background.Fill.UserPicture($BackgroundImagePath)

            # Export Image if selected
            if ($chkExport.Checked) {
                $safeName = $name -replace '[\\\/\:\*\?\"\\<\>\|]', ''
                $Stm.Export((Join-Path $imgFolder "$safeName.png"), "PNG")
            }

            # --- Slide Transitions ---
            $Stm.SlideShowTransition.EntryEffect = [Microsoft.Office.Interop.PowerPoint.PpEntryEffect]::ppEffectFlyThroughInBounce
            $Stm.SlideShowTransition.AdvanceOnTime = [Microsoft.Office.Core.MsoTriState]::msoTrue
            $Stm.SlideShowTransition.AdvanceOnClick = [Microsoft.Office.Core.MsoTriState]::msoFalse
            $Stm.SlideShowTransition.AdvanceTime = 10
            $Stm.SlideShowTransition.Duration = 1.5 # How long the "Fly In" takes to finish

            $progressBar.Value = $stepCount
            $form.Refresh()
        }

        # Presentation Settings
        $Presentation.SlideShowSettings.LoopUntilStopped = $true
        $Presentation.SaveAs($saveTitle)
        $Presentation.Close()
        $PowerPoint.Quit()
        
        [System.Windows.Forms.MessageBox]::Show("Success! Files saved to save location.", "Finished")
        explorer.exe /select,$saveTitle
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)")
        if ($PowerPoint) { $PowerPoint.Quit() }
    }
}

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    CreatePowerPoint
    $btnRun.Enabled = $true
})

# Add all remaining controls
$form.Controls.AddRange(@($btnData, $txtData, $btnImage, $txtImage, $btnSave, $txtSave, $btnRun, $chkTitle1, $txtTitle, $chkTitle2, $txtTitle2, $chkExport, $progressBar))
$form.ShowDialog()