Write-Host "Version 1.0.2 - Combined ItemID and Description with newline"
Write-Host "Copyright David Lane 2025"
Write-Host "Warrenty - None Totally Untested, By using this programm I agree to send (David Lane) all my money and also agree to the ITunes T&C's" 
$scriptVersion = "1.0.2"
Write-Host "Running Xero CSV Export Script version $scriptVersion"

$inputPath = Read-Host "Enter path to input XML file"
$defaultOutput = Join-Path (Split-Path $inputPath) "output.csv"
$outputPath = Read-Host "Enter output path or press Enter for default [$defaultOutput]"
if (!$outputPath) { $outputPath = $defaultOutput }

# Define the exact Xero headers
$header = "*ContactName,EmailAddress,POAddressLine1,POAddressLine2,POAddressLine3,POAddressLine4,POCity,PORegion,POPostalCode,POCountry,*InvoiceNumber,Reference,*InvoiceDate,*DueDate,*Total,*Description,*Quantity,*UnitAmount,*Discount,*AccountCode,*TaxType,TaxAmount,TrackingName1,TrackingOption1,TrackingName2,TrackingOption2"

# Write header
$header | Out-File $outputPath -Encoding UTF8

# Load XML
[xml]$xml = Get-Content $inputPath

# Get customer data from the Customers node
$customer = $xml.GLExport.Customers.Customer

foreach ($transaction in $xml.GLExport.Transactions.Transaction) {
    foreach ($detail in $transaction.Detail | Where-Object { $_.Description -ne 'Level 1 Sales Tax' }) {
        # Combine ItemID and Memo for description
        $combinedDescription = "$($detail.ItemID)`n$($detail.Memo)"

        # Direct mapping from XML to CSV fields
        $fields = @(
            $customer.CompanyName,                  # *ContactName
            $customer.ContactEmailAddress,          # EmailAddress
            $customer.AddressLine1,                # POAddressLine1
            $customer.AddressLine2,                # POAddressLine2
            "",                                    # POAddressLine3
            "",                                    # POAddressLine4
            $customer.City,                        # POCity
            $customer.State,                       # PORegion
            $customer.PostalCode,                  # POPostalCode
            $customer.Country,                     # POCountry
            "INV-$($transaction.DocumentNumber)",  # *InvoiceNumber
            $transaction.DocumentNumber,           # Reference
            ([DateTime]::ParseExact($transaction.DocumentDate, "MM-dd-yyyy", [System.Globalization.CultureInfo]::InvariantCulture)).ToString("dd/MM/yyyy"), # *InvoiceDate
            ([DateTime]::ParseExact($transaction.DueDate, "MM-dd-yyyy", [System.Globalization.CultureInfo]::InvariantCulture)).ToString("dd/MM/yyyy"),     # *DueDate
            [Math]::Abs([decimal]$detail.Total),   # *Total
            "`"$combinedDescription`"",           # *Description (now includes ItemID and Memo)
            [Math]::Abs([decimal]$detail.Quantity),# *Quantity
            [Math]::Abs([decimal]$detail.ItemPrice), # *UnitAmount
            "0",                                   # *Discount
            "210",                                 # *AccountCode
            "`"20% (VAT on Income)`"",            # *TaxType
            [Math]::Round([Math]::Abs([decimal]$detail.Total) * 0.2, 2), # TaxAmount
            "`"Sales Types`"",                     # TrackingName1
            "`"$($detail.ItemID)`"",              # TrackingOption1
            "",                                    # TrackingName2
            ""                                     # TrackingOption2
        ) -join ','

        Add-Content -Path $outputPath -Value $fields -Encoding UTF8
    }
}

Write-Host "CSV export completed successfully to: $outputPath"
Write-Host "Script version: $scriptVersion"