# Define the inventory, prices, and default taxable status
$inventory = [ordered]@{
    "Labor (Per Hour)" = @{
        Price = 125
        Taxable = $false  # Not taxable by default
    }
    "Mileage (Per Mile)" = @{
        Price = 1.25
        Taxable = $false  # Not taxable by default
    }
    "Backhoe Use (Per Hour)" = @{
        Price = 100
        Taxable = $false  # Not taxable by default
    }
    "Yanmar Use (Per Hour)" = @{
        Price = 50
        Taxable = $false  # Not taxable by default
    }
    "Boring Rig 24x40 (Per Hour)" = @{
        Price = 150
        Taxable = $false  # Not taxable by default
    }
    "Fiber Splice Case (Each)" = @{
        Price = 450
        Taxable = $false  # Not taxable by default
    }
    "Vault (Each)" = @{
        Price = 1972
        Taxable = $false  # Not taxable by default
    }
    "2 Fiber (Per Foot)" = @{
        Price = 0.35
        Taxable = $false  # Not taxable by default
    }
    "4 Fiber (Per Foot)" = @{
        Price = 0.33
        Taxable = $false  # Not taxable by default
    }
    "6 Fiber (Per Foot)" = @{
        Price = 0.38
        Taxable = $false  # Not taxable by default
    }
    "12 Fiber (Per Foot)" = @{
        Price = 0.41
        Taxable = $false  # Not taxable by default
    }
    "24 Fiber (Per Foot)" = @{
        Price = 0.43
        Taxable = $false  # Not taxable by default
    }
    "48 Fiber (Per Foot)" = @{
        Price = 0.59
        Taxable = $false  # Not taxable by default
    }
    "72 Fiber (Per Foot)" = @{
        Price = 0.67
        Taxable = $false  # Not taxable by default
    }
    "96 Fiber (Per Foot)" = @{
        Price = 0.75
        Taxable = $false  # Not taxable by default
    }
    "144 Fiber (Per Foot)" = @{
        Price = 1.26
        Taxable = $false  # Not taxable by default
    }
    "16MM (Per Foot)" = @{
        Price = 0.16
        Taxable = $false  # Not taxable by default
    }
    "16MM [With Tracer] (Per Foot)" = @{
        Price = 0.25
        Taxable = $false  # Not taxable by default
    }
    "2-Way Duct (Per Foot)" = @{
        Price = 0.75
        Taxable = $false  # Not taxable by default
    }
    "4-Way Duct (Per Foot)" = @{
        Price = 1.19
        Taxable = $false  # Not taxable by default
    }
    "8-Way Duct (Per Foot)" = @{
        Price = 1.60
        Taxable = $true  # taxable by default
    }
}







# Default tax rate variable
$taxRate = 8.25

# Load the necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Display GUI for user input
$inventoryForm = New-Object System.Windows.Forms.Form
$inventoryForm.Text = "MTCO Fiber Construction Calculator"
$inventoryForm.Size = New-Object System.Drawing.Size(700, 750)

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

$labelTaxableHeader = New-Object System.Windows.Forms.Label
$labelTaxableHeader.Location = New-Object System.Drawing.Point(400, 10)
$labelTaxableHeader.Text = "Taxable"
$labelTaxableHeader.AutoSize = $true
$inventoryForm.Controls.Add($labelTaxableHeader)

# Tax rate label and input field
$labelTaxRate = New-Object System.Windows.Forms.Label
$labelTaxRate.Location = New-Object System.Drawing.Point(620, 530)
$labelTaxRate.Text = "Tax Rate:"
$labelTaxRate.AutoSize = $true
$inventoryForm.Controls.Add($labelTaxRate)

$inputTaxRate = New-Object System.Windows.Forms.TextBox
$inputTaxRate.Location = New-Object System.Drawing.Point(620, 550)
$inputTaxRate.Size = New-Object System.Drawing.Size(50, 20)
$inputTaxRate.Text = $taxRate
$inventoryForm.Controls.Add($inputTaxRate)

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

    # Use default taxable status if available, otherwise set to false
    $isTaxable = $inventory[$item].Taxable -as [bool]
    if ($null -eq $isTaxable) { $isTaxable = $false }

    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Location = New-Object System.Drawing.Point(400, $top)
    $checkbox.Text = "Taxable"
    $checkbox.Checked = $isTaxable
    $inventoryForm.Controls.Add($checkbox)

    $inputControls[$item] = @{
        AmountTextBox = $amountTextBox
        PriceTextBox = $priceTextBox
        Checkbox = $checkbox
    }

    $top += 30
}

# Create controls for output file name and calculate button
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Size = New-Object System.Drawing.Size(150, 20)
$outputLabel.Text = "Job Name"

$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Size = New-Object System.Drawing.Size(150, 20)

$calculateButton = New-Object System.Windows.Forms.Button
$calculateButton.Size = New-Object System.Drawing.Size(150, 30)
$calculateButton.Text = "Calculate"
$calculateButton.Add_Click({
    $outputFileName = "$($outputTextBox.Text).txt"  # Append ".txt" to the file name
    if (-not $outputFileName) {
        $outputFileName = (Get-Date).ToString("yyyyMMdd_HHmmss") + ".txt"
    }

    $outputFilePath = Join-Path -Path $PSScriptRoot -ChildPath $outputFileName

    # Update tax rate based on user input
    $taxRate = [double]$inputTaxRate.Text

    # Calculate subtotal, total sales tax, and total cost
    $subtotal = 0
    $totalSalesTax = 0

    foreach ($item in $inputControls.Keys) {
        $amountUsed = [double]$inputControls[$item].AmountTextBox.Text
        $pricePerUnit = [double]$inputControls[$item].PriceTextBox.Text
        $isTaxable = $inputControls[$item].Checkbox.Checked

        $subtotal += $amountUsed * $pricePerUnit

        if ($isTaxable) {
            $totalSalesTax += $amountUsed * $pricePerUnit
        }
    }

    $totalSalesTax *= ($taxRate / 100)
    $totalCost = $subtotal + $totalSalesTax

    "Subtotal: $($subtotal.ToString('C2'))" | Out-File -FilePath $outputFilePath
    "Total Sales Tax (@ $taxRate%): $($totalSalesTax.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append
    "Total Cost (including tax): $($totalCost.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append

    # Iterate through input controls and calculate totals
    foreach ($item in $inputControls.Keys) {
        $amountUsed = [double]$inputControls[$item].AmountTextBox.Text
        $pricePerUnit = [double]$inputControls[$item].PriceTextBox.Text
        $isTaxable = $inputControls[$item].Checkbox.Checked

        if ($isTaxable) {
            $totalPrice = $amountUsed * ($pricePerUnit + ($pricePerUnit * ($taxRate / 100)))
        } else {
            $totalPrice = $amountUsed * $pricePerUnit
        }

        "$item - Total Price: $($totalPrice.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append
        "Amount Used: $amountUsed, Price Per Unit: $($pricePerUnit.ToString('C2'))" | Out-File -FilePath $outputFilePath -Append
        "Taxable: $isTaxable" | Out-File -FilePath $outputFilePath -Append
        "`n" | Out-File -FilePath $outputFilePath -Append
    }

    "------------------------" | Out-File -FilePath $outputFilePath -Append
    "`n" | Out-File -FilePath $outputFilePath -Append

    [System.Windows.Forms.MessageBox]::Show("Calculation complete. Output saved to: $outputFilePath", "Success")
})

$inventoryForm.Controls.Add($calculateButton)
$calculateButtonLocationX = $inventoryForm.ClientSize.Width - $calculateButton.Width - 10
$calculateButtonLocationY = $inventoryForm.ClientSize.Height - $calculateButton.Height - 10
$calculateButton.Location = New-Object System.Drawing.Point($calculateButtonLocationX, $calculateButtonLocationY)

$inventoryForm.Controls.Add($outputTextBox)
$outputTextBoxLocationX = $inventoryForm.ClientSize.Width - $outputTextBox.Width - 10
$outputTextBoxLocationY = $inventoryForm.ClientSize.Height - $outputTextBox.Height - 50
$outputTextBox.Location = New-Object System.Drawing.Point($outputTextBoxLocationX, $outputTextBoxLocationY)

$inventoryForm.Controls.Add($outputLabel)
$outputLabelLocationX = $inventoryForm.ClientSize.Width - $outputLabel.Width - -80
$outputLabelLocationY = $inventoryForm.ClientSize.Height - $outputLabel.Height - 70
$outputLabel.Location = New-Object System.Drawing.Point($outputLabelLocationX, $outputLabelLocationY)

$inventoryForm.ActiveControl = $inputControls[$($inventory.Keys[0])].AmountTextBox

# Show the form
$inventoryForm.ShowDialog()
