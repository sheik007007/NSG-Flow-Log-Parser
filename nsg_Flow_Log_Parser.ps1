Add-Type -AssemblyName System.Windows.Forms

Function Show-GUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Sheik's NSG Flow Log Parser"
    $form.Size = New-Object System.Drawing.Size(420, 350)
    $form.StartPosition = "CenterScreen"

    $lblInput = New-Object System.Windows.Forms.Label
    $lblInput.Location = New-Object System.Drawing.Point(10, 20)
    $lblInput.Size = New-Object System.Drawing.Size(200, 20)
    $lblInput.Text = "Select Input NSG JSON Log:"
    $form.Controls.Add($lblInput)

    $txtInput = New-Object System.Windows.Forms.TextBox
    $txtInput.Location = New-Object System.Drawing.Point(10, 50)
    $txtInput.Size = New-Object System.Drawing.Size(270, 20)
    $form.Controls.Add($txtInput)

    $btnInput = New-Object System.Windows.Forms.Button
    $btnInput.Location = New-Object System.Drawing.Point(290, 50)
    $btnInput.Size = New-Object System.Drawing.Size(100, 20)
    $btnInput.Text = "Browse"
    $btnInput.Add_Click({
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        if ($OpenFileDialog.ShowDialog() -eq "OK") {
            $txtInput.Text = $OpenFileDialog.FileName
        }
    })
    $form.Controls.Add($btnInput)

    $lblOutput = New-Object System.Windows.Forms.Label
    $lblOutput.Location = New-Object System.Drawing.Point(10, 90)
    $lblOutput.Size = New-Object System.Drawing.Size(200, 20)
    $lblOutput.Text = "Select Output CSV File:"
    $form.Controls.Add($lblOutput)

    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Location = New-Object System.Drawing.Point(10, 120)
    $txtOutput.Size = New-Object System.Drawing.Size(270, 20)
    $form.Controls.Add($txtOutput)

    $btnOutput = New-Object System.Windows.Forms.Button
    $btnOutput.Location = New-Object System.Drawing.Point(290, 120)
    $btnOutput.Size = New-Object System.Drawing.Size(100, 20)
    $btnOutput.Text = "Browse"
    $btnOutput.Add_Click({
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $txtOutput.Text = $SaveFileDialog.FileName
        }
    })
    $form.Controls.Add($btnOutput)

    # Progress Bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 160)
    $progressBar.Size = New-Object System.Drawing.Size(380, 20)
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $form.Controls.Add($progressBar)

    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Location = New-Object System.Drawing.Point(160, 200)
    $btnStart.Size = New-Object System.Drawing.Size(100, 30)
    $btnStart.Text = "Convert"
    $btnStart.Add_Click({
        psParseNsgFlowData -InputFile $txtInput.Text -OutputFile $txtOutput.Text -ProgressBar $progressBar
    })
    $form.Controls.Add($btnStart)

    $form.ShowDialog()
}

Function psParseNsgFlowData {
    param (
        [string]$InputFile,
        [string]$OutputFile,
        [System.Windows.Forms.ProgressBar]$ProgressBar
    )

    if (!(Test-Path $InputFile)) {
        [System.Windows.Forms.MessageBox]::Show("Error: Input file not found.", "Error", "OK", "Error")
        return
    }

    try {
        $d = Get-Content $InputFile | ConvertFrom-Json -ErrorAction Stop
        if (!$d.records) {
            [System.Windows.Forms.MessageBox]::Show("Error: No records found in JSON file.", "Error", "OK", "Error")
            return
        }

        $arrFlowMap = @("UnixEpoch","sourceIP","destIP","sourcePort","destPort","proto","trafficFlow","action","flowState","packetsSrcToDest","bytesSrcToDest","packetsDstToSrc","bytesDestToSrc")
        $arrResults = @()
        $recordCount = $d.records.Count
        $currentRecord = 0

        foreach ($r in $d.records) {
            $currentRecord++
            $ProgressBar.Value = [math]::Round(($currentRecord / $recordCount) * 100)

            if (!$r.properties.flows.flows.flowtuples) { continue }

            foreach ($tuple in $r.properties.flows.flows.flowtuples) {
                $arrData = $tuple -split ","
                $objResult = New-Object PSCustomObject
                $objResult | Add-Member -MemberType NoteProperty -Name "timestamp" -Value $r.time
                $objResult | Add-Member -MemberType NoteProperty -Name "rulename" -Value ($r.properties.flows.rule -join " ~~ ")

                for ($i = 0; $i -lt $arrData.Count; $i++) {
                    $objResult | Add-Member -MemberType NoteProperty -Name $arrFlowMap[$i] -Value $arrData[$i]
                }

                $arrResults += $objResult
            }
        }

        if ($arrResults.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Error: No valid data to save.", "Error", "OK", "Error")
            return
        }

       # Save in proper CSV format (comma-separated)
$arrResults | Export-Csv -Path $OutputFile -NoTypeInformation -Force -UseCulture

[System.Windows.Forms.MessageBox]::Show("File successfully saved as CSV!", "Success", "OK", "Information")

# Open CSV in Excel if available
try {
    Start-Process "EXCEL.EXE" $OutputFile
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Conversion successful! Open the CSV file manually.", "Info", "OK", "Information")
}

    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error processing file: $_", "Error", "OK", "Error")
    }
}

Show-GUI