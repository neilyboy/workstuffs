# Define the inventory, prices, and default taxable status
$inventory = [ordered]@{
    "Mileage (Per Mile)" = @{
        Price = 1.25
    }
    "Backhoe Use (Per Hour)" = @{
        Price = 100
    }
    "Yanmar Use (Per Hour)" = @{
        Price = 50
    }
    "Boring Rig 24x40 (Per Hour)" = @{
        Price = 150
    }
    "Fiber Splice Case (Each)" = @{
        Price = 450
    }
    "Vault (Each)" = @{
        Price = 1972
    }
    "2 Fiber (Per Foot)" = @{
        Price = 0.35
    }
    "4 Fiber (Per Foot)" = @{
        Price = 0.33
    }
    "6 Fiber (Per Foot)" = @{
        Price = 0.38
    }
    "12 Fiber (Per Foot)" = @{
        Price = 0.41
    }
    "24 Fiber (Per Foot)" = @{
        Price = 0.43
    }
    "48 Fiber (Per Foot)" = @{
        Price = 0.59
    }
    "72 Fiber (Per Foot)" = @{
        Price = 0.67
    }
    "96 Fiber (Per Foot)" = @{
        Price = 0.75
    }
    "144 Fiber (Per Foot)" = @{
        Price = 1.26
    }
    "16MM (Per Foot)" = @{
        Price = 0.16
    }
    "16MM [With Tracer] (Per Foot)" = @{
        Price = 0.25
    }
    "2-Way Duct (Per Foot)" = @{
        Price = 0.75
    }
    "4-Way Duct (Per Foot)" = @{
        Price = 1.19
    }
    "8-Way Duct (Per Foot)" = @{
        Price = 1.60
    }
}

# Load the necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Display GUI for user input
$inventoryForm = New-Object System.Windows.Forms.Form
$inventoryForm.Text = "MTCO Fiber Construction Quote Calculator"
$inventoryForm.Size = New-Object System.Drawing.Size(630, 720)

# Create a hashtable to store input controls dynamically
$inputControls = @{}

# Header labels
$labelItemHeader = New-Object System.Windows.Forms.Label
$labelItemHeader.Location = New-Object System.Drawing.Point(10, 10)
$labelItemHeader.Text = "Item"
$labelItemHeader.AutoSize = $true
$inventoryForm.Controls.Add($labelItemHeader)

$labelAmountHeader = New-Object System.Windows.Forms.Label
$labelAmountHeader.Location = New-Object System.Drawing.Point(220, 10)
$labelAmountHeader.Text = "Amount"
$labelAmountHeader.AutoSize = $true
$inventoryForm.Controls.Add($labelAmountHeader)

$labelPriceHeader = New-Object System.Windows.Forms.Label
$labelPriceHeader.Location = New-Object System.Drawing.Point(300, 10)
$labelPriceHeader.Text = "Price"
$labelPriceHeader.AutoSize = $true
$inventoryForm.Controls.Add($labelPriceHeader)

# Create controls for output file name and calculate button
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Size = New-Object System.Drawing.Size(63, 20)
$outputLabel.Text = "Job Name:"
$inventoryForm.Controls.Add($outputLabel)
# $outputLabelLocationX = $inventoryForm.ClientSize.Width - $outputLabel.Width - 130
# $outputLabelLocationY = $inventoryForm.ClientSize.Height - $outputLabel.Height - 95
$outputLabel.Location = New-Object System.Drawing.Point(420, 570)

$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Size = New-Object System.Drawing.Size(180, 20)
$inventoryForm.Controls.Add($outputTextBox)
# $outputTextBoxLocationX = $inventoryForm.ClientSize.Width - $outputTextBox.Width - 13
# $outputTextBoxLocationY = $inventoryForm.ClientSize.Height - $outputTextBox.Height - 67
$outputTextBox.Location = New-Object System.Drawing.Point(420, 600)


# Create controls for additional input fields
$labelGuysOnJob = New-Object System.Windows.Forms.Label
$labelGuysOnJob.Location = New-Object System.Drawing.Point(420, 340)
$labelGuysOnJob.Text = "Guys On Job:"
$labelGuysOnJob.AutoSize = $true
$inventoryForm.Controls.Add($labelGuysOnJob)

$inputGuysOnJob = New-Object System.Windows.Forms.TextBox
$inputGuysOnJob.Location = New-Object System.Drawing.Point(550, 340)
$inputGuysOnJob.Size = New-Object System.Drawing.Size(50, 20)
$inputGuysOnJob.Text = "5"
$inventoryForm.Controls.Add($inputGuysOnJob)

$labelHourlyRate = New-Object System.Windows.Forms.Label
$labelHourlyRate.Location = New-Object System.Drawing.Point(420, 410)
$labelHourlyRate.Text = "Hourly Rate:"
$labelHourlyRate.AutoSize = $true
$inventoryForm.Controls.Add($labelHourlyRate)

$inputHourlyRate = New-Object System.Windows.Forms.TextBox
$inputHourlyRate.Location = New-Object System.Drawing.Point(550, 410)
$inputHourlyRate.Size = New-Object System.Drawing.Size(50, 20)
$inputHourlyRate.Text = "100"
$inventoryForm.Controls.Add($inputHourlyRate)

$labelFeetPerDay = New-Object System.Windows.Forms.Label
$labelFeetPerDay.Location = New-Object System.Drawing.Point(420, 480)
$labelFeetPerDay.Text = "Feet Per Day:"
$labelFeetPerDay.AutoSize = $true
$inventoryForm.Controls.Add($labelFeetPerDay)

$inputFeetPerDay = New-Object System.Windows.Forms.TextBox
$inputFeetPerDay.Location = New-Object System.Drawing.Point(550, 480)
$inputFeetPerDay.Size = New-Object System.Drawing.Size(50, 20)
$inputFeetPerDay.Text = "600"
$inventoryForm.Controls.Add($inputFeetPerDay)


# Create controls for each inventory item
$top = 32
foreach ($item in $inventory.Keys) {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, $top)
    $label.Text = $item
    $label.AutoSize = $true
    $inventoryForm.Controls.Add($label)

    $amountTextBox = New-Object System.Windows.Forms.TextBox
    $amountTextBox.Location = New-Object System.Drawing.Point(220, $top)
    $amountTextBox.Size = New-Object System.Drawing.Size(70, 20)
    $amountTextBox.Text = "0"
    $inventoryForm.Controls.Add($amountTextBox)

    $priceTextBox = New-Object System.Windows.Forms.TextBox
    $priceTextBox.Location = New-Object System.Drawing.Point(300, $top)
    $priceTextBox.Size = New-Object System.Drawing.Size(90, 20)
    $priceTextBox.Text = $inventory[$item].Price.ToString('F2')
    $inventoryForm.Controls.Add($priceTextBox)

    $inputControls[$item] = @{
        AmountTextBox = $amountTextBox
        PriceTextBox = $priceTextBox
    }

    $top += 30

}

# Add a border around the entire group of inventory items

$border = New-Object System.Windows.Forms.Label
$border.Location = New-Object System.Drawing.Point(5, 18)
$border.Size = New-Object System.Drawing.Size(400, 615)
$border.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$inventoryForm.Controls.Add($border)

# Add a border around the entire group of inventory items

#$border2 = New-Object System.Windows.Forms.Label
#$border2.Location = New-Object System.Drawing.Point(410, 538)
#$border2.Size = New-Object System.Drawing.Size(200, 45)
#$border2.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
#$inventoryForm.Controls.Add($border2)

# Add a border around the entire group of inventory items

$border3 = New-Object System.Windows.Forms.Label
$border3.Location = New-Object System.Drawing.Point(410, 580)
$border3.Size = New-Object System.Drawing.Size(200, 53)
$border3.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$inventoryForm.Controls.Add($border3)

# Add a border around the entire group of inventory items

$border4 = New-Object System.Windows.Forms.Label
$border4.Location = New-Object System.Drawing.Point(410, 330)
$border4.Size = New-Object System.Drawing.Size(200, 45)
$border4.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$inventoryForm.Controls.Add($border4)

# Add a border around the entire group of inventory items

$border5 = New-Object System.Windows.Forms.Label
$border5.Location = New-Object System.Drawing.Point(410, 400)
$border5.Size = New-Object System.Drawing.Size(200, 45)
$border5.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$inventoryForm.Controls.Add($border5)

# Add a border around the entire group of inventory items

$border6 = New-Object System.Windows.Forms.Label
$border6.Location = New-Object System.Drawing.Point(410, 470)
$border6.Size = New-Object System.Drawing.Size(200, 45)
$border6.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$inventoryForm.Controls.Add($border6)




$calculateButton = New-Object System.Windows.Forms.Button
$calculateButton.Size = New-Object System.Drawing.Size(150, 30)
$calculateButton.Text = "Quote Job"
$calculateButton.Add_Click({
    $outputFileName = "$($outputTextBox.Text).txt"  # Append ".txt" to the file name
    if (-not $outputFileName) {
        $outputFileName = (Get-Date).ToString("yyyyMMdd_HHmmss") + ".txt"
    }

    $outputFilePath = Join-Path -Path $PSScriptRoot -ChildPath $outputFileName

    # Calculate subtotal, total sales tax, and total cost
    $subtotal = 0

    foreach ($item in $inputControls.Keys) {
        $amountUsed = [double]$inputControls[$item].AmountTextBox.Text
        $pricePerUnit = [double]$inputControls[$item].PriceTextBox.Text

        $subtotal += $amountUsed * $pricePerUnit
    }

    $totalCost = $subtotal

    # Additional functionality: Calculate totalDuctLength
    $totalDuctLength = 0
    foreach ($item in @("2-Way Duct (Per Foot)", "4-Way Duct (Per Foot)", "8-Way Duct (Per Foot)")) {
        $totalDuctLength += [double]$inputControls[$item].AmountTextBox.Text
    }

    # Additional functionality: Calculate totalWorkDays, laborTotal
    $feetPerDay = [double]$inputFeetPerDay.Text
    $totalWorkDays = [math]::Ceiling($totalDuctLength / $feetPerDay)  # Round up to the nearest whole day
    $guysOnJob = [double]$inputGuysOnJob.Text
    $totalHoursWorked = $totalWorkDays * 8 * $guysOnJob  # Assuming 8 hours per day per guy
    $hourlyRate = [double]$inputHourlyRate.Text
    $laborTotal = $totalHoursWorked * $hourlyRate

    # Update subtotal, total sales tax, and total cost
    $subtotal += $laborTotal
    $totalCost = $subtotal + $laborTotal

    "||| MTCO Fiber Project Quote For $($outputTextBox.Text) |||" | Out-File -FilePath $outputFilePath
    "`n" | Out-File -FilePath $outputFilePath -Append
    "Financial Costs For Job" | Out-File -FilePath $outputFilePath -Append
    "===================================" | Out-File -FilePath $outputFilePath -Append
    "| Subtotal: $($subtotal.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append
    "| Total Labor Cost: $($laborTotal.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append
    "| Total Cost (including tax and labor): $($totalCost.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append
    "===================================" | Out-File -FilePath $outputFilePath -Append
    "`n" | Out-File -FilePath $outputFilePath -Append
    "Project Details" | Out-File -FilePath $outputFilePath -Append
    "===================================" | Out-File -FilePath $outputFilePath -Append
    "| Total Duct Footage: $totalDuctLength'" | Out-File -FilePath $outputFilePath -Append
    "| Total Days To Complete: $totalWorkDays" | Out-File -FilePath $outputFilePath -Append
    "===================================" | Out-File -FilePath $outputFilePath -Append
    "`n" | Out-File -FilePath $outputFilePath -Append
    "--------------------------------------" | Out-File -FilePath $outputFilePath -Append
    "| Price Breakdown (Individual Items) |" | Out-File -FilePath $outputFilePath -Append
    "--------------------------------------" | Out-File -FilePath $outputFilePath -Append

    # Iterate through input controls and calculate totals
    foreach ($item in $inputControls.Keys) {
        $amountUsed = [double]$inputControls[$item].AmountTextBox.Text
        $pricePerUnit = [double]$inputControls[$item].PriceTextBox.Text

        # Add a condition to check if the amount used is greater than zero
        if ($amountUsed -gt 0) {
            $totalPrice = $amountUsed * $pricePerUnit
            "----------------------------------" | Out-File -FilePath $outputFilePath -Append
            "$item" | Out-File -FilePath $outputFilePath -Append
            "----------------------------------" | Out-File -FilePath $outputFilePath -Append
            "Total Price: $($totalPrice.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append
            "Amount Used: $amountUsed, Price Per Unit: $($pricePerUnit.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append
            "`n" | Out-File -FilePath $outputFilePath -Append
        }
    }

    "------------END OF QUOTE------------" | Out-File -FilePath $outputFilePath -Append
    "`n" | Out-File -FilePath $outputFilePath -Append

    [System.Windows.Forms.MessageBox]::Show("Calculation complete. Output saved to: $outputFilePath", "Success")
})

$inventoryForm.Controls.Add($calculateButton)
$calculateButtonLocationX = $inventoryForm.ClientSize.Width - $calculateButton.Width - 10
$calculateButtonLocationY = $inventoryForm.ClientSize.Height - $calculateButton.Height - 10
$calculateButton.Location = New-Object System.Drawing.Point($calculateButtonLocationX, $calculateButtonLocationY)





$inventoryForm.ActiveControl = $inputControls[$($inventory.Keys[0])].AmountTextBox

# Show the form
$inventoryForm.ShowDialog()
