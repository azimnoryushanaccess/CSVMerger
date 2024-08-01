try {
    # Define paths
    $inputTxtPath = (Get-Location).Path + "\Input\JPMAccessWire.txt"
    $inputXlsxPath = (Get-Location).Path + "\Input\WIRE Bank Details.xlsx"
    $outputDirectory = (Get-Location).Path + "\Output\" 
    # $testOutput = (Get-Location).Path + "\Output\test_processed.txt" 

    # Function to check if a file is in use
    function Test-FileOpen {
        param (
            [string]$filePath
        )
        try {
            $fileStream = [System.IO.File]::Open($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $fileStream.Close()
            return $false
        }
        catch {
            return $true
        }
    }

    # Check if the .txt file exists
    if (-Not (Test-Path $inputTxtPath)) {
        throw "The input text file 'JPMAccessWire.txt' does not exist."
    }

    # Check if the Wire Bank Details file exists
    if (-Not (Test-Path $inputXlsxPath)) {
        throw "The input Excel file 'WIRE Bank Details.xlsx' does not exist."
    }

    if (Test-FileOpen -filePath $inputXlsxPath) {
        throw "The input Excel file 'WIRE Bank Details.xlsx' is already open. Please close the excel file and retry..."
    }

    # Create an instance of Excel
    $excelObj = New-Object -ComObject Excel.Application

    # Capture the process ID of the Excel instance
    $excelProcess = Get-Process -Name Excel | Sort-Object StartTime | Select-Object -Last 1

    Write-Host "Opening Wire Bank Details Excel workbook..."
    $ExcelWorkbook = $excelObj.Workbooks.Open($inputXlsxPath)
    Write-Host "Accessing specific sheets in Wire Bank Details Excel workbook..."
    $ExcelWorkSheet = $ExcelWorkbook.Sheets.Item(1)
    Write-Host "Getting row count from worksheets..."
    $wirebank_rowCount = $ExcelWorkSheet.UsedRange.Rows.Count
    Write-Host "Reading data from Wire Bank Details Excel workbooks..."
    $wirebank_data = @()
    for ($j = 1; $j -le $wirebank_rowCount; $j++) {
        $row = @()
        for ($col = 3; $col -le 20; $col++) {
            $row += $ExcelWorkSheet.Cells.Item($j, $col).Value2
        }
        $wirebank_data += , $row
    }

    Write-Host "Reading data from JPMAccessWire.txt"
    $headers = 1..125
    $csvContent = Import-Csv -Path $inputTxtPath -Header $headers

    $outputFileNamePrefix = $csvContent[0].2
    # Construct the output file path including the prefix from Header Date
    $outputPath = Join-Path -Path $outputDirectory -ChildPath ($outputFileNamePrefix + "_output.txt")

    # Initialize an array to hold modified rows
    $modifiedCsvContent = @()

    # Process each row except the first and last
    for ($i = 0; $i -lt $csvContent.Count; $i++) {
        if ($i -eq 0 -or $i -eq $csvContent.Count - 1) {
            # Add the first and last rows without modification
            $modifiedCsvContent += $csvContent[$i]
        }
        else {
            $row = $csvContent[$i]
            foreach ($row_data in $wirebank_data) {
                # Convert bank acc from scientific notation to whole number
                if ($row_data[1] -match '^[+-]?\d+(\.\d+)?[Ee][+-]?\d+$') {
                    [decimal]$row_data[1] = $row_data[1] 
                }

                $wirebankacc_value = $row_data[1]

                if (-not [string]::IsNullOrWhiteSpace($wirebankacc_value) -and $row.15 -eq $wirebankacc_value.ToString()) {
                    $row.3 = "CHASHKHH"
                    if ($row.15 -eq "GB74LOYD30166332775601") {
                        $row.6 = "GBP"
                    } 
                    else {
                        $row.6 = "USD"
                    }
                    $row.14 = "ACCT"
                    $row.17 = ""
                    $row.18 = ""
                    $row.19 = ""
                    $row.20 = ""
                    $row.21 = ""

                    #Beneficiary Bank Section
                    $BeneficiaryBankAddr = $row_data[7] -replace "`r`n", "," -replace "`n", "," -replace "`r", ","
                    $BeneficiaryBankAddr3 = ""

                    if (-not [string]::IsNullOrEmpty($BeneficiaryBankAddr)) {

                        # Remove IFSC Code from Beneficiary bank address
                        if ($BeneficiaryBankAddr -match "IFSC CODE:\s*(\w+),\s*(.*)") {
                            $BeneficiaryBankAddr = $matches[2]
                        }

                        # Check if the length of the BeneficiaryBankAddr is more than 35
                        if ($BeneficiaryBankAddr.Length -gt 35) {
                            # Split the string by commas
                            $splitAddr = $BeneficiaryBankAddr -split ","

                            # Check if there are more than two splits
                            if ($splitAddr.Length -gt 2) {
                                # Get the last two values from the split array
                                $BeneficiaryBankAddr3 = ($splitAddr[-2..-1] -join ",").Trim()
            
                                # Remove the last two values and join the remaining parts back into a string
                                $BeneficiaryBankAddr = ($splitAddr[0..($splitAddr.Length - 3)] -join ",").Trim()    
                            }
                            else {
                                # Get the last values from the split array
                                $BeneficiaryBankAddr3 = $splitAddr[-1].Trim()
            
                                # Remove the last two values and join the remaining parts back into a string
                                $BeneficiaryBankAddr = $splitAddr[0].Trim()
                            }
                            
                        }
                    }

                    $BeneficiaryBankCountryCode = if (-not [string]::IsNullOrEmpty($row_data[8])) { $row_data[8].Substring($row_data[8].Length - 2) } else { $row_data[8] } # Get last 2 chars

                    $row.27 = $BeneficiaryBankAddr
                    $row.29 = $BeneficiaryBankAddr3
                    $row.30 = $BeneficiaryBankCountryCode
                    
                    #Intermediary Bank Section
                    $SwiftID = $row_data[10]
                    $IntermediaryBankName = $row_data[11] -replace "`r`n", " " -replace "`n", " " -replace "`r", " "

                    $row.40 = "SWIFT"
                    $row.41 = $SwiftID
                    $row.42 = $IntermediaryBankName

                    #Address Section
                    $AddressCellValue = $row_data[12] -replace "`r`n", "," -replace "`n", "," -replace "`r", "," #240 GREENWICH STREET, Address 2, NEW YORK,NY
                    $Address3 = ""
                    $IntermediaryBankCountryCode = if (-not [string]::IsNullOrEmpty($row_data[13])) { $row_data[13].Substring($row_data[13].Length - 2) } else { $row_data[13] } # Get last 2 chars

                    function Split-Address {
                        param (
                            [string]$address
                        )

                        # Initialize output variables
                        $Address1 = ""
                        $Address2 = ""
                        $Address3 = ""

                        # Split the address by commas
                        $parts = $address -split ','

                        # Trim each part
                        $parts = $parts | ForEach-Object { $_.Trim() }

                        # Handle different cases based on number of parts
                        switch ($parts.Length) {
                            1 {
                                $Address3 = $parts[0]
                            }
                            2 {
                                $Address1 = ""
                                $Address2 = ""
                                $Address3 = "$($parts[0]), $($parts[1])"
                            }
                            3 {
                                if ($parts[2] -match '^\d+$') {
                                    $Address1 = ""
                                    $Address2 = ""
                                    $Address3 = "$($parts[0]), $($parts[1]), $($parts[2])"
                                }
                                else {
                                    $Address1 = $parts[0]
                                    $Address2 = ""
                                    $Address3 = "$($parts[1]), $($parts[2])"
                                }
                            }
                            default {
                                if ($parts[$parts.Length - 1] -match '^\w{2}$' -and $parts[$parts.Length - 2] -match '^\w+$') {
                                    $Address3 = "$($parts[$parts.Length - 2]), $($parts[$parts.Length - 1])"
                                    $Address1 = ($parts[0..($parts.Length - 3)] -join ', ')
                                    $Address2 = $parts[$parts.Length - 3]
                                }
                                else {
                                    $Address3 = "$($parts[$parts.Length - 2]), $($parts[$parts.Length - 1])"
                                    $Address1 = ($parts[0..($parts.Length - 3)] -join ', ')
                                    $Address2 = ""
                                }
                            }
                        }

                        # Output the results
                        return @{
                            Address1 = $Address1
                            Address2 = $Address2
                            Address3 = $Address3
                        }
                    }

                    $AddressSplitResult = Split-Address -address $AddressCellValue

                    if (-not [string]::IsNullOrEmpty($AddressCellValue)) {
                        # Check if the length of the AddressCellValue is more than 35
                        if ($AddressCellValue.Length -gt 35) {
                            # Split the string by commas
                            $splitIAddr = $AddressCellValue -split ","

                            # Check if there are more than two splits
                            if ($splitIAddr.Length -gt 2) {
                                # Get the last two values from the split array
                                $Address3 = ($splitIAddr[-2..-1] -join ",").Trim()
            
                                # Remove the last two values and join the remaining parts back into a string
                                $AddressCellValue = ($splitIAddr[0..($splitIAddr.Length - 3)] -join ",").Trim()
                            }
                            else {
                                $Address3 = $splitIAddr[-1].Trim()
                                $AddressCellValue = $splitIAddr[0].Trim()
                            }
                            
                        }
                    }

                    # Split the value on the comma
                    $Address1 = $AddressSplitResult.Address1
                    $Address2 = $AddressSplitResult.Address2
                    # $Address3 = $AddressSplitResult.Address3

                    $row.43 = $AddressCellValue
                    $row.44 = ""
                    $row.45 = $Address3
                    $row.46 = $IntermediaryBankCountryCode

                    #Internal ref
                    $InternalRef = $row_data[15]
                    $row.75 = $InternalRef

                    #Transaction Detail Section
                    $transDetail = ""
                    if (-not [string]::IsNullOrEmpty($row_data[14])) {
                        # Split the string by newline characters and get the first line
                        $firstLine = $row_data[14] -split '\r?\n' | Select-Object -First 1
                        # Set $row.77 with the first line
                        $transDetail = $firstLine
                    }
                    else {
                        $transDetail = $row_data[14]
                    }
                    $row.77 = $transDetail
                    $row.78 = ""
                    $row.79 = ""
                    $row.80 = ""

                    #Bank to Bank Detail Section
                    $BankToBankLines = if (-not [string]::IsNullOrEmpty($row_data[16])) { $row_data[16] -split "`n" } else { $row_data[16] }
                    $startColumn = 100

                    foreach ($line in $BankToBankLines) {
                        if ($line -match "^(ACC |IFSC CODE|FND|INT:)\s*(.*)") {
                            $code = $matches[1].TrimEnd(' :')
                            $details = $matches[2]
                            Write-Host "Copying code '$code' and details '$details' to rows $startColumn and $($startColumn + 1) in output worksheet"
                            $row.$startColumn = $code
                            $row.($startColumn + 1) = $details
                        }
                        else {
                            Write-Host "No match for line '$line'"
                            $row.($startColumn + 1) = $line
                            # $ExcelWorkSheet_Output.Cells.Item($outputRow, $startColumn + 1).Value2 = ""
                        }
                        $startColumn += 2
                    }

                    #Additional Info
                    if (-not [string]::IsNullOrEmpty($row_data[4])) {
                        $paymentCode = $row_data[4] -split '\s+' | Select-Object -First 1
                        $row.116 = "PurposeOfPaymentDestination=$paymentCode" #$row_data[4]
                    }
                    else {
                        $row.116 = ""
                    }

                }
            }

            # Add the modified row to the array
            $modifiedCsvContent += $row
        }
    }

    $csvContent | Export-Csv -Path $outputPath -NoTypeInformation
    (Get-Content $outputPath | Select-Object -Skip 1) | Set-Content $outputPath

    Write-Host "Closing workbook sessions..."
    $ExcelWorkbook.Close($false)
    Write-Host "Quitting Excel application..."
    # Quit the Excel application
    $excelObj.Quit()

    Write-Host "Releasing COM objects..."
    # Release the COM objects to free up memory
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkSheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelObj) | Out-Null

    Write-Host "Collecting garbage..."
    # Collect garbage
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Write-Host "Excel automation script completed successfully."

} 
catch {
    Write-Error "An error occurred: $_"
    Write-Error "$($_.InvocationInfo.ScriptName)($($_.InvocationInfo.ScriptLineNumber)): $($_.InvocationInfo.Line)"

    # Ensure the specific Excel process is terminated to avoid zombies
    if ($excelProcess -and $excelProcess.HasExited -eq $false) {
        Write-Host "Terminating the Excel process created by this script..."
        Stop-Process -Id $excelProcess.Id -Force -ErrorAction SilentlyContinue
    }

    Write-Host "Press Enter to exit..."
    [void][System.Console]::ReadLine() # Keeps the terminal open
} 