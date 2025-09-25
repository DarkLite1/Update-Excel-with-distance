<# 
    .SYNOPSIS
        Update Excel file with distance between 2 coordinates and the time
        to travel between them.

    .LINK
        https://project-osrm.org/
        https://router.project-osrm.org/route/v1/driving/52.4675669420078,4.6330354579344;52.17393,7.88456
#>

param (
    [string]$ExcelFilePath = 'T:\Test\Brecht\PowerShell\Logs\Distance tracker.xlsx',
    [string]$WorksheetName = 'Distances',
    [hashtable]$ColumnLetterHeader = @{
        startDestination = 'F'
        coordinate       = 'G'
        distance         = 'H'
        duration         = 'I'
    }
)

begin {
    try {
        $ErrorActionPreference = 'Stop'

        #region Open Excel worksheet
        Write-Verbose "Open Excel file '$ExcelFilePath'"

        $excelPackage = Open-ExcelPackage -Path $ExcelFilePath

        if (-not $excelPackage) {
            throw "Excel file '$ExcelFilePath' not found"
        }
        #endregion

        #region Get sheet data
        Write-Verbose "Get data in worksheet '$WorksheetName'"

        $sheet = $excelPackage.Workbook.Worksheets[$WorksheetName]

        if (-not $sheet) {
            throw "Sheet name '$WorksheetName' not found"
        }
        #endregion
    }
    catch {
        throw "Failed to open Excel file '$ExcelFilePath': $_"
    }
}

process {
    try {
        #region Get data from Excel sheet
        $results = @()
        $startCoordinate = $null
        
        foreach (
            $row in
            $sheet.Cells | Group-Object -Property { $_.Start.Row }
        ) {
            $rowNumber = $row.Name

            Write-Verbose "Row '$rowNumber'"

            $rowCells = $row.Group

            #region Create cell addresses
            $cellAddress = @{
                startDestination = '{0}{1}' -f $ColumnLetterHeader.startDestination, $rowNumber
                coordinate       = '{0}{1}' -f 
                $ColumnLetterHeader.coordinate, $rowNumber
                distance         = '{0}{1}' -f 
                $ColumnLetterHeader.distance, $rowNumber
                duration         = '{0}{1}' -f 
                $ColumnLetterHeader.duration, $rowNumber
            }
            #endregion

            #region Test if row is a start or destination row
            $isStartCoordinateRow = $rowCells.Where(
                {
                    ($_.Start.Address -eq $cellAddress.startDestination) -and
                    ($_.Value -eq 'S') 
                }
            )

            $isDestinationCoordinateRow = $rowCells.Where(
                {
                    ($_.Start.Address -eq $cellAddress.startDestination) -and
                    ($_.Value -eq 'D') 
                }
            )
            #endregion

            #region Get coordinate
            $coordinate = (
                $rowCells.Where(
                    { ($_.Start.Address -eq $cellAddress.coordinate) }
                )
            ).Value

            if ($isStartCoordinateRow) {
                Write-Verbose "Start coordinate '$coordinate'"
                
                $startCoordinate = $coordinate
            }
            elseif ($isDestinationCoordinateRow -and $startCoordinate) {
                Write-Verbose "Destination coordinate '$coordinate'"

                $results += @{
                    coordinate  = @{
                        start       = $startCoordinate
                        destination = $coordinate
                    }
                    cellAddress = @{
                        distance = $cellAddress.distance
                        duration = $cellAddress.duration
                    }
                    apiResponse = $null
                    errors      = @()
                }
            }
            #endregion
        }

        Write-Verbose "Found $($results.Count) start and destination pairs"
        #endregion

        #region Get distance and duration from OSRM API
        foreach ($pair in $results) {
            try {
                $params = @{
                    Uri     = (
                        'https://router.project-osrm.org/route/v1/driving/{0};{1}' -f
                        $pair.coordinate.start, $pair.coordinate.destination
                    ) -replace '\s'
                    Verbose = $false
                }

                Write-Verbose "Call API endpoint '$($params.Uri)'"

                $pair.apiResponse = Invoke-RestMethod @params
            }
            catch {
                $pair.errors += "Failed API request: $_"
            }
        }
        #endregion

        #region Update Excel sheet
        foreach (
            $pair in $results.Where({ -not $_.errors })
        ) {
            #region Set distance
            try {
                $distanceCell = $pair.cellAddress.distance
                $distanceValue = $pair.apiResponse.routes[0].distance

                Write-Verbose "Set distance in cell '$distanceCell' value '$distanceValue'"

                $sheet.Cells[$distanceCell].Value = $distanceValue
            }
            catch {
                $pair.errors += "Failed to set distance in cell '$distanceCell' with value '$distanceValue': $_"
            }
            #endregion
            
            #region Set duration
            try {
                $durationCell = $pair.cellAddress.duration
                $durationValue = $pair.apiResponse.routes[0].duration

                Write-Verbose "Set distance in cell '$durationCell' value '$durationValue'"

                $sheet.Cells[$durationCell].Value = $durationValue
            }
            catch {
                $pair.errors += "Failed to set duration in cell '$durationCell' with value '$durationValue': $_"
            }
            #endregion
        }
        #endregion
    }
    catch {
        throw "Failed processing Excel file '$ExcelFilePath': $_"
    }
    finally {
        if ($excelPackage) {
            Write-Verbose 'Close Excel file'
            Close-ExcelPackage -ExcelPackage $excelPackage -EA Ignore
        }
    }
}