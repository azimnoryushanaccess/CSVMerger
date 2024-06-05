try {
    Write-Host "Starting Excel automation script..."

    # Define paths
    $inputTxtPath = (Get-Location).Path + "\Input\JPMAccessWire.txt"
    $inputCsvPath = (Get-Location).Path + "\Input\JPMAccessWire.csv"
    $inputXlsxPath = (Get-Location).Path + "\Input\WIRE Bank Details.xlsx"
    $outputDirectory = (Get-Location).Path + "\Output\"

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

    Write-Host "Converting JPMAccessWire.txt to JPMAccessWire.csv..."

    $txtContent = Get-Content -Path $inputTxtPath

    # Modify column O to add a leading single quote for account numbers
    $csvContent = @()

    # Iterate over each line in txtContent
    foreach ($line in $txtContent) {
        # Split the line by comma
        $columns = $line -split ","
    
        # Check if the number of columns is greater than 14
        if ($columns.Length -gt 14) {
            # Add "#" to the 15th column (index 14)
            $columns[14] = "#" + $columns[14]
        }
    
        # Join the modified columns with commas and add to csvContent
        $csvContent += ($columns -join ",")
    }

    # Read the .txt file and save it as .csv
    # Get-Content $inputTxtPath | Set-Content $inputCsvPath

    $csvContent | Set-Content $inputCsvPath

    # Check if the input files are in use
    if (Test-FileOpen -filePath $inputCsvPath) {
        throw "The input CSV file 'JPMAccessWire.csv' is already open. Please close the excel file and retry..."
    }

    # Create an instance of Excel
    $excelObj = New-Object -ComObject Excel.Application

    # Capture the process ID of the Excel instance
    $excelProcess = Get-Process -Name Excel | Sort-Object StartTime | Select-Object -Last 1

    Write-Host "Opening the converted CSV file and the existing Excel workbook..."
    # Open the converted CSV file and the existing Excel workbook
    $ExcelWorkbook_a = $excelObj.Workbooks.Open($inputCsvPath)
    $ExcelWorkbook_b = $excelObj.Workbooks.Open($inputXlsxPath)

    Write-Host "Accessing specific sheets in each workbook..."
    # Access the specific sheets in each workbook
    $ExcelWorkSheet_a = $ExcelWorkbook_a.Sheets.Item(1)
    $ExcelWorkSheet_b = $ExcelWorkbook_b.Sheets.Item(1)

    # Get the value of cell C1 from the first worksheet of the input CSV
    $outputFileNamePrefix = $ExcelWorkSheet_a.Cells.Item(1, 2).Value2.ToString()

    # Construct the output file path including the prefix from cell C1
    $outputPath = Join-Path -Path $outputDirectory -ChildPath ($outputFileNamePrefix + "_output.csv")
    # $outputCSVPath = Join-Path -Path $outputDirectory -ChildPath ($outputFileNamePrefix + "_output.csv")

    Write-Host "Creating a new workbook for output..."
    # Add a new workbook for output
    $ExcelWorkbook_Output = $excelObj.Workbooks.Add()
    $ExcelWorkSheet_Output = $ExcelWorkbook_Output.Sheets.Item(1)

    Write-Host "Getting row count from worksheets..."
    # Get the count of rows in each worksheet
    $rowCount_a = $ExcelWorkSheet_a.UsedRange.Rows.Count
    $rowCount_b = $ExcelWorkSheet_b.UsedRange.Rows.Count

    Write-Host "Reading data from input workbooks..."
    # Read data from worksheets into arrays for faster processing
    $data_a = @()
    for ($i = 2; $i -le $rowCount_a; $i++) {
        $row = @()
        for ($col = 0; $col -le 125; $col++) {
            if ($col -eq 14 -and -not [string]::IsNullOrWhiteSpace($ExcelWorkSheet_a.Cells.Item($i, $col + 1).Value2)) {
                # $trim_row = $ExcelWorkSheet_a.Cells.Item($i, $col + 1).Value2.Substring(1).ToString()
                $trim_row = $ExcelWorkSheet_a.Cells.Item($i, $col + 1).Value2 -replace '[#"]', ''
                $row += $trim_row
            }
            else {
                $row += $ExcelWorkSheet_a.Cells.Item($i, $col + 1).Value2
            }
        }
        $data_a += , $row
    }

    $data_b = @()
    for ($j = 1; $j -le $rowCount_b; $j++) {
        $row = @()
        for ($col = 3; $col -le 20; $col++) {
            $row += $ExcelWorkSheet_b.Cells.Item($j, $col).Value2
        }
        $data_b += , $row
    }

    Write-Host "Copying first and last rows from input CSV to output workbook..."

    # Copy the first row from input CSV to the first row of the output workbook
    for ($col = 1; $col -le 3; $col++) {
        $cellValue = $ExcelWorkSheet_a.Cells.Item(1, $col).Value2
        $cellValue = if ($null -ne $cellValue) { $cellValue.ToString() } else { "" }
        Write-Host "Copying value '$cellValue' to column $col in output workbook first row"
        $ExcelWorkSheet_Output.Cells.Item(1, $col).Value2 = $cellValue
    }

    # Copy the last row from input CSV to the last row of the output workbook
    for ($col = 1; $col -le 3; $col++) {
        $cellValue = $ExcelWorkSheet_a.Cells.Item($rowCount_a, $col).Value2
        $cellValue = if ($null -ne $cellValue) { $cellValue.ToString() } else { "" }
        Write-Host "Copying value '$cellValue' to column $col in output workbook last row"
        $ExcelWorkSheet_Output.Cells.Item($rowCount_a, $col).Value2 = $cellValue
    }

    Write-Host "Copying matched data from input workbooks to output workbook..."
    # Initialize row index for the output worksheet after the first row
    $outputRow = 2

    # Loop through each row in worksheet a
    foreach ($row_a in $data_a) {
        $value_a = $row_a[14]
 
        # Loop through each row in worksheet b to find a match
        foreach ($row_b in $data_b) {
            $value_b = $row_b[1]

            # Write-Host "Value A: $value_a"
            # Write-Host "Value B: $value_b"

            # if (-not [string]::IsNullOrWhiteSpace($value_a) -and -not [string]::IsNullOrWhiteSpace($value_b)) {
            #     Write-Host $value_a
            #     Write-Host $value_b
            # }

            if (-not [string]::IsNullOrWhiteSpace($value_a) -and -not [string]::IsNullOrWhiteSpace($value_b) -and $value_a -eq $value_b.ToString()) {

                # Copy the data to the output worksheet
                for ($col = 0; $col -le 125; $col++) {
                    $cellValue = $row_a[$col]

                    if ($null -ne $cellValue) {
                        if ($col -eq 14) {
                            # Add a leading single quote to preserve leading zeros in account numbers
                            $cellValue = "'" + $cellValue
                        }
                        else {
                            $cellValue = $cellValue.ToString()
                        }
                    }
                    else {
                        $cellValue = ""
                    }

                    Write-Host "Copying value '$cellValue' to column $col in output worksheet"
                    $ExcelWorkSheet_Output.Cells.Item($outputRow, $col + 1).Value2 = $cellValue
                }

                # Clear Beneficiary Details
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 17).Value2 = ""
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 18).Value2 = ""
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 19).Value2 = ""
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 20).Value2 = ""
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 21).Value2 = ""

                #Beneficiary Bank Section
                $BeneficiaryBankAddr = $row_b[7]
                $BeneficiaryBankCountryCode = $row_b[8].Substring($row_b[8].Length - 2) # Get last 2 chars

                $ExcelWorkSheet_Output.Cells.Item($outputRow, 27).Value2 = $BeneficiaryBankAddr
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 30).Value2 = $BeneficiaryBankCountryCode
                
                #Intermediary Bank Section
                $SwiftID = $row_b[10]
                $IntermediaryBankName = $row_b[11]

                $ExcelWorkSheet_Output.Cells.Item($outputRow, 40).Value2 = "SWIFT"
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 41).Value2 = $SwiftID
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 42).Value2 = $IntermediaryBankName
                
                #Address Section
                $AddressCellValue = $row_b[12] #240 GREENWICH STREET, Address 2, NEW YORK,NY
                $IntermediaryBankCountryCode = $row_b[13].Substring($row_b[13].Length - 2) # Get last 2 chars

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

                # Split the value on the comma
                $Address1 = $AddressSplitResult.Address1
                $Address2 = $AddressSplitResult.Address2
                $Address3 = $AddressSplitResult.Address3

                $ExcelWorkSheet_Output.Cells.Item($outputRow, 43).Value2 = $AddressCellValue
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 44).Value2 = ""
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 45).Value2 = ""
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 46).Value2 = $IntermediaryBankCountryCode

                # Not needed remove this part
                # $ExcelWorkSheet_Output.Cells.Item($outputRow, 47).Value2 = "USABA"
                # $ExcelWorkSheet_Output.Cells.Item($outputRow, 48).Value2 = "021000018"

                #Internal ref
                $InternalRef = $row_b[15]
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 75).Value2 = $InternalRef

                #Transaction Detail Section
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 77).Value2 = $row_b[14]
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 78).Value2 = ""
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 79).Value2 = ""
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 80).Value2 = ""

                #Bank to Bank Detail Section
                $BankToBankLines = $row_b[16] -split "`n"
                $startColumn = 100

                foreach ($line in $BankToBankLines) {
                    if ($line -match "^(ACC:|IFSC CODE|FND|INT:)\s*(.*)") {
                        $code = $matches[1].TrimEnd(' :')
                        $details = $matches[2]
                        Write-Host "Copying code '$code' and details '$details' to rows $startColumn and $($startColumn + 1) in output worksheet"
                        $ExcelWorkSheet_Output.Cells.Item($outputRow, $startColumn).Value2 = $code
                        $ExcelWorkSheet_Output.Cells.Item($outputRow, $startColumn + 1).Value2 = $details
                    }
                    else {
                        Write-Host "No match for line '$line'"
                        $ExcelWorkSheet_Output.Cells.Item($outputRow, $startColumn).Value2 = $line
                        $ExcelWorkSheet_Output.Cells.Item($outputRow, $startColumn + 1).Value2 = ""
                    }
                    $startColumn += 2
                }

                #Charges
                # $ExcelWorkSheet_Output.Cells.Item($outputRow, 114).Value2 = $row_b[17]

                #Additional Info
                $ExcelWorkSheet_Output.Cells.Item($outputRow, 116).Value2 = $row_b[4]

                # Move to the next row in the output worksheet
                $outputRow++
                break
            }
        }
    }

    Write-Host "Auto-fitting columns in the output workbook..."
    # Auto-fit the columns in the output workbook
    $ExcelWorkSheet_Output.Columns.AutoFit()

    Write-Host "Saving the output workbook..."
    try {
        # Save the output workbook
        $ExcelWorkbook_Output.SaveAs($outputPath, 6)
    }
    catch {
        throw "Failed to save the output file. The file might be open or locked."
    }

    Write-Host "Closing workbook sessions..."
    # Close the workbook sessions
    $ExcelWorkbook_a.Close($false)  # Do not save changes to the CSV
    $ExcelWorkbook_b.Close($true)  # Save changes
    $ExcelWorkbook_Output.Close($true)  # Save changes

    Write-Host "Quitting Excel application..."
    # Quit the Excel application
    $excelObj.Quit()

    Write-Host "Releasing COM objects..."
    # Release the COM objects to free up memory
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkSheet_a) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkSheet_b) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkSheet_Output) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkbook_a) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkbook_b) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkbook_Output) | Out-Null
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
