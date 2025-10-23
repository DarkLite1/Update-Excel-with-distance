#Requires -Version 7

<#
    .SYNOPSIS
        Update Excel file with distance and travel time between 2 coordinates.

    .DESCRIPTION
        This script reads an Excel file to find the source and destination
        coordinates. It sends a request to the Open Street Maps API to
        gather the travel time and distance and updates the Excel file.

    .PARAMETER ConfigurationJsonFile
        Contains all the parameters used by the script.
        See 'Example.json' for a detailed explanation of parameters.

    .LINK
        https://project-osrm.org/
        https://router.project-osrm.org/route/v1/driving/52.4675669420078,4.6330354579344;52.17393,7.88456
#>

param (
    [Parameter(Mandatory)]
    [string]$ConfigurationJsonFile
)

begin {
    $ErrorActionPreference = 'stop'

    $systemErrors = @()
    $logFileData = [System.Collections.Generic.List[PSObject]]::new()
    $eventLogData = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date

    try {
        $eventLogData.Add(
            [PSCustomObject]@{
                Message   = 'Script started'
                DateTime  = $scriptStartTime
                EntryType = 'Information'
                EventID   = '100'
            }
        )

        #region Import .json file
        Write-Verbose "Import .json file '$ConfigurationJsonFile'"

        $jsonFileItem = Get-Item -LiteralPath $ConfigurationJsonFile -ErrorAction Stop

        $jsonFileContent = Get-Content $jsonFileItem -Raw -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        @(
            'Excel', 'DropFolder'
        ).where(
            { -not $jsonFileContent.$_ }
        ).foreach(
            { throw "Property '$_' not found" }
        )

        @(
            'Path'
        ).where(
            { -not $jsonFileContent.DropFolder.$_ }
        ).foreach(
            { throw "Property 'DropFolder.$_' not found" }
        )

        @(
            'WorksheetName', 'Column'
        ).where(
            { -not $jsonFileContent.Excel.$_ }
        ).foreach(
            { throw "Property 'Excel.$_' not found" }
        )

        @(
            'startDestination', 'coordinate', 'distance', 'duration'
        ).where(
            { -not $jsonFileContent.Excel.Column.$_ }
        ).foreach(
            { throw "Property 'Excel.Column.$_' not found" }
        )
        #endregion

        $WorksheetName = $jsonFileContent.Excel.WorksheetName
        $StartDestinationColumn = $jsonFileContent.Excel.Column.StartDestination
        $CoordinateColumn = $jsonFileContent.Excel.Column.coordinate
        $DistanceColumn = $jsonFileContent.Excel.Column.distance
        $DurationColumn = $jsonFileContent.Excel.Column.duration

        $DropFolderPath = $jsonFileContent.DropFolder.Path
        $ArchiveFolderPath = $jsonFileContent.DropFolder.ArchivePath

        #region Test drop folder exists
        if (-not (Test-Path -LiteralPath $DropFolderPath -PathType Container)) {
            throw "DropFolder.Path '$DropFolderPath' not found"
        }
        #endregion

        #region Create archive folder
        if (
            $ArchiveFolderPath -and
            (-not (Test-Path -LiteralPath $ArchiveFolderPath -PathType Container))
        ) {
            try {
                Write-Verbose "Create archive folder '$ArchiveFolderPath'"
                $null = New-Item -Path $ArchiveFolderPath -ItemType Directory
            }
            catch {
                throw "Failed creating archive folder '$ArchiveFolderPath': $_"
            }
        }
        #endregion

        #region Get .xlsx files in drop folder
        $excelFiles = Get-ChildItem -Path $DropFolderPath -Filter '*.xlsx' -File

        if (-not $excelFiles) {
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = "No Excel files found in drop folder '$DropFolderPath'"
                    EntryType = 'Information'
                    EventID   = '4'
                }
            )
            Write-Verbose $eventLogData[-1].Message

            return
        }

        $eventLogData.Add(
            [PSCustomObject]@{
                Message   = "Found $($excelFiles.Count) Excel file(s) to process in '$DropFolderPath'"
                EntryType = 'Information'
                EventID   = '4'
            }
        )
        Write-Verbose $eventLogData[-1].Message
        #endregion
    }
    catch {
        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Message  = "Input file '$ConfigurationJsonFile': $_"
        }

        Write-Warning $systemErrors[-1].Message

        return
    }
}

process {
    if ($systemErrors) { return }

    function Add-EventLogMessageHC {
        param (
            [Parameter(Mandatory)]
            [String]$Message
        )

        $eventLogData.Add(
            [PSCustomObject]@{
                DateTime  = Get-Date
                Message   = $Message
                EntryType = 'Information'
                EventID   = '4'
            }
        )
        Write-Verbose $Message
    }

    function ConvertTo-LongitudeLatitudeCoordinateHC {
        param (
            [Parameter(Mandatory)]
            [String]$LatitudeLongitudeCoordinate
        )

        try {
            $parts = $LatitudeLongitudeCoordinate -split ','
            $longitudeLatitudeCoordinate = ($parts[1].Trim() + ',' + $parts[0].Trim())
            $longitudeLatitudeCoordinate
        }
        catch {
            $Error.RemoveAt(0)
        }
    }

    $results = @()

    foreach ($excelFile in $excelFiles) {
        try {
            $result = @{
                File            = $excelFile
                CoordinatePairs = @()
            }

            #region Open and get Excel sheet data
            try {
                #region Open Excel file
                Add-EventLogMessageHC "Excel file '$excelFile': Open file"

                $excelPackage = Open-ExcelPackage -Path $excelFile.FullName

                if (-not $excelPackage) {
                    throw "Excel file '$excelFile' not found"
                }
                #endregion

                #region Get sheet data
                Add-EventLogMessageHC "Excel file '$excelFile': Get data in worksheet '$WorksheetName'"

                $sheet = $excelPackage.Workbook.Worksheets[$WorksheetName]

                if (-not $sheet) {
                    throw "Sheet name '$WorksheetName' not found"
                }
                #endregion
            }
            catch {
                throw "Failed to open Excel file '$excelFile': $_"
            }
            #endregion

            #region Get coordinates from sheet
            Add-EventLogMessageHC "Excel file '$excelFile': Get coordinates"

            $startCoordinate = $null

            foreach (
                $row in
                $sheet.Cells | Group-Object -Property { $_.Start.Row }
            ) {
                $rowNumber = $row.Name

                $rowCells = $row.Group

                #region Create cell addresses
                $cellAddress = @{
                    startDestination = '{0}{1}' -f $StartDestinationColumn, $rowNumber
                    coordinate       = '{0}{1}' -f $CoordinateColumn, $rowNumber
                    distance         = '{0}{1}' -f $DistanceColumn, $rowNumber
                    duration         = '{0}{1}' -f $DurationColumn, $rowNumber
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
                    Write-Verbose "Row '$rowNumber' start coordinate '$coordinate'"

                    $startCoordinate = $coordinate
                }
                elseif ($isDestinationCoordinateRow -and $startCoordinate) {
                    Write-Verbose "Row '$rowNumber' destination coordinate '$coordinate'"

                    $result.CoordinatePairs += @{
                        dateTime    = Get-Date
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
                else {
                    Write-Verbose "Row '$rowNumber' ignored, not a start or destination row"
                }
                #endregion
            }

            Add-EventLogMessageHC "Excel file '$excelFile': Found $($result.CoordinatePairs.Count) start and destination pairs"
            #endregion

            #region Get distance and duration from OSRM API
            if ($result.CoordinatePairs) {
                Add-EventLogMessageHC "Excel file '$excelFile': Get distance and duration from OSRM API for $($result.CoordinatePairs.Count) coordinate pairs"
            }

            $i = 0

            foreach ($pair in $result.CoordinatePairs) {
                try {
                    $i++

                    $startCoordinate = ConvertTo-LongitudeLatitudeCoordinateHC -LatitudeLongitudeCoordinate $pair.coordinate.start

                    $destinationCoordinate = ConvertTo-LongitudeLatitudeCoordinateHC -LatitudeLongitudeCoordinate $pair.coordinate.destination

                    $params = @{
                        Uri     = (
                            'https://router.project-osrm.org/route/v1/driving/{0};{1}' -f
                            $startCoordinate, $destinationCoordinate
                        ) -replace '\s'
                        Verbose = $false
                    }

                    Write-Verbose "$i/$($result.CoordinatePairs.Count): call API endpoint '$($params.Uri)'"

                    $pair.apiResponse = Invoke-RestMethod @params
                }
                catch {
                    $pair.errors += "Failed API request: $_"
                }
            }
            #endregion

            #region Update Excel sheet
            $coordinatePairsWithoutErrors = $result.CoordinatePairs.Where(
                { -not $_.errors }
            )

            if ($coordinatePairsWithoutErrors) {
                Add-EventLogMessageHC "Excel file '$excelFile': Update Excel sheet '$WorksheetName'"
            }

            foreach (
                $pair in $coordinatePairsWithoutErrors
            ) {
                #region Set distance
                try {
                    $distanceCell = $pair.cellAddress.distance
                    $distanceValue = $pair.apiResponse.routes[0].distance

                    Write-Verbose "Set distance in cell '$distanceCell' value '$distanceValue'"

                    # distance is returned in meters, we need km
                    $sheet.Cells[$distanceCell].Value = $distanceValue / 1000

                    $sheet.Cells[$distanceCell].Style.NumberFormat.Format = '0.00 \ \k\m'
                }
                catch {
                    $pair.errors += "Failed to set distance in cell '$distanceCell' with value '$distanceValue': $_"
                }
                #endregion

                #region Set duration
                try {
                    $durationCell = $pair.cellAddress.duration
                    $durationValue = $pair.apiResponse.routes[0].duration

                    Write-Verbose "Set duration in cell '$durationCell' value '$durationValue'"

                    # duration is returned in seconds, we need minutes
                    $sheet.Cells[$durationCell].Value = $durationValue / 60

                    $sheet.Cells[$durationCell].Style.NumberFormat.Format = '0\ \m\i\n'
                }
                catch {
                    $pair.errors += "Failed to set duration in cell '$durationCell' with value '$durationValue': $_"
                }
                #endregion
            }
            #endregion
        }
        catch {
            $systemErrors += [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Excel file '$excelFile': Failed processing file: $_"
            }

            Write-Warning $systemErrors[-1].Message
        }
        finally {
            #region Save Excel file updates
            if ($excelPackage) {
                try {
                    Add-EventLogMessageHC "Excel file '$excelFile': Save updates in Excel and close file"

                    Close-ExcelPackage -ExcelPackage $excelPackage
                }
                catch {
                    $systemErrors += [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Excel file '$excelFile': Failed to save updates in Excel file: $_"
                    }

                    Write-Warning $systemErrors[-1].Message
                }
            }
            #endregion

            #region Move Excel file to archive folder
            if ($ArchiveFolderPath) {
                try {
                    Add-EventLogMessageHC "Excel file '$excelFile': Move Excel file to archive folder '$ArchiveFolderPath'"

                    $params = @{
                        LiteralPath = $excelFile.FullName
                        Destination = $ArchiveFolderPath
                    }
                    Move-Item @params -Force
                }
                catch {
                    $systemErrors += [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Failed to move file from '$($params.LiteralPath)' to '$($params.Destination)': $_"
                    }

                    Write-Warning $systemErrors[-1].Message
                }
            }
            #endregion

            $results += $result
        }
    }
}

end {
    function Out-LogFileHC {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [PSCustomObject[]]$DataToExport,
            [Parameter(Mandatory)]
            [String]$PartialPath,
            [Parameter(Mandatory)]
            [String[]]$FileExtensions,
            [hashtable]$ExcelFile = @{
                SheetName = 'Overview'
                TableName = 'Overview'
                CellStyle = $null
            },
            [Switch]$Append
        )

        $allLogFilePaths = @()

        foreach (
            $fileExtension in
            $FileExtensions | Sort-Object -Unique
        ) {
            try {
                $logFilePath = "$PartialPath{0}" -f $fileExtension

                $M = "Export {0} object{1} to '$logFilePath'" -f
                $DataToExport.Count,
                $(if ($DataToExport.Count -ne 1) { 's' })
                Write-Verbose $M

                switch ($fileExtension) {
                    '.csv' {
                        $params = @{
                            LiteralPath       = $logFilePath
                            Append            = $Append
                            Delimiter         = ';'
                            NoTypeInformation = $true
                        }
                        $DataToExport | Export-Csv @params

                        break
                    }
                    '.json' {
                        #region Convert error object to error message string
                        $convertedDataToExport = foreach (
                            $exportObject in
                            $DataToExport
                        ) {
                            foreach ($property in $exportObject.PSObject.Properties) {
                                $name = $property.Name
                                $value = $property.Value
                                if (
                                    $value -is [System.Management.Automation.ErrorRecord]
                                ) {
                                    if (
                                        $value.Exception -and $value.Exception.Message
                                    ) {
                                        $exportObject.$name = $value.Exception.Message
                                    }
                                    else {
                                        $exportObject.$name = $value.ToString()
                                    }
                                }
                            }
                            $exportObject
                        }
                        #endregion

                        if (
                            $Append -and
                            (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                        ) {
                            $params = @{
                                LiteralPath = $logFilePath
                                Raw         = $true
                                Encoding    = 'UTF8'
                            }
                            $jsonFileContent = Get-Content @params | ConvertFrom-Json

                            $convertedDataToExport = [array]$convertedDataToExport + [array]$jsonFileContent
                        }

                        $convertedDataToExport |
                        ConvertTo-Json -Depth 7 |
                        Out-File -LiteralPath $logFilePath

                        break
                    }
                    '.txt' {
                        $params = @{
                            LiteralPath = $logFilePath
                            Append      = $Append
                        }

                        $DataToExport | Format-List -Property * -Force |
                        Out-File @params

                        break
                    }
                    '.xlsx' {
                        if (
                            (-not $Append) -and
                            (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                        ) {
                            $logFilePath | Remove-Item
                        }

                        $excelParams = @{
                            LiteralPath   = $logFilePath
                            Append        = $true
                            AutoNameRange = $true
                            AutoSize      = $true
                            FreezeTopRow  = $true
                            WorksheetName = $ExcelFile.SheetName
                            TableName     = $ExcelFile.TableName
                            Verbose       = $false
                        }
                        if ($ExcelFile.CellStyle) {
                            $excelParams.CellStyleSB = $ExcelFile.CellStyle
                        }
                        $DataToExport | Export-Excel @excelParams

                        break
                    }
                    default {
                        throw "Log file extension '$_' not supported. Supported values are '.csv', '.json', '.txt' or '.xlsx'."
                    }
                }

                $allLogFilePaths += $logFilePath
            }
            catch {
                Write-Warning "Failed creating log file '$logFilePath': $_"
            }
        }

        $allLogFilePaths
    }

    function Get-LogFolderHC {
        <#
        .SYNOPSIS
            Ensures that a specified LiteralPath exists, creating it if it doesn't.
            Supports absolute paths and paths relative to $PSScriptRoot. Returns
            the full LiteralPath of the folder.

        .DESCRIPTION
            This function takes a LiteralPath as input and checks if it exists. If
            the LiteralPath does not exist, it attempts to create the folder. It
            handles both absolute paths and paths relative to the location of
            the currently running script ($PSScriptRoot).

        .PARAMETER LiteralPath
            The LiteralPath to ensure exists. This can be an absolute LiteralPath (ex.
            C:\MyFolder\SubFolder) or a LiteralPath relative to the script's
            directory (ex. Data\Logs).

        .EXAMPLE
            Get-LogFolderHC -Path 'C:\MyData\Output'
            # Ensures the directory 'C:\MyData\Output' exists.

        .EXAMPLE
            Get-LogFolderHC -Path 'Logs\Archive'
            # If the script is in 'C:\Scripts', this ensures 'C:\Scripts\Logs\Archive' exists.
        #>

        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if ($Path -match '^[a-zA-Z]:\\' -or $Path -match '^\\') {
            $fullPath = $Path
        }
        else {
            $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $Path
        }

        if (-not (Test-Path -Path $fullPath -PathType Container)) {
            try {
                Write-Verbose "Create log folder '$fullPath'"
                $null = New-Item -Path $fullPath -ItemType Directory -Force
            }
            catch {
                throw "Failed creating log folder '$fullPath': $_"
            }
        }

        $fullPath
    }

    function Send-MailKitMessageHC {
        <#
            .SYNOPSIS
                Send an email using MailKit and MimeKit assemblies.

            .DESCRIPTION
                This function sends an email using the MailKit and MimeKit
                assemblies. It requires the assemblies to be installed before
                calling the function:

                $params = @{
                    Source           = 'https://www.nuget.org/api/v2'
                    SkipDependencies = $true
                    Scope            = 'AllUsers'
                }
                Install-Package @params -Name 'MailKit'
                Install-Package @params -Name 'MimeKit'

            .PARAMETER MailKitAssemblyPath
                The LiteralPath to the MailKit assembly.

            .PARAMETER MimeKitAssemblyPath
                The LiteralPath to the MimeKit assembly.

            .PARAMETER SmtpServerName
                The name of the SMTP server.

            .PARAMETER SmtpPort
                The port of the SMTP server.

            .PARAMETER SmtpConnectionType
                The connection type for the SMTP server.

                Valid values are:
                - 'None'
                - 'Auto'
                - 'SslOnConnect'
                - 'StartTlsWhenAvailable'
                - 'StartTls'

            .PARAMETER Credential
                The credential object containing the username and password.

            .PARAMETER From
                The sender's email address.

            .PARAMETER FromDisplayName
            The display name to show for the sender.

            Email clients may display this differently. It is most likely
            to be shown if the sender's email address is not recognized
                (e.g., not in the address book).

            .PARAMETER To
                The recipient's email address.

            .PARAMETER Body
            The body of the email, HTML is supported.

            .PARAMETER Subject
            The subject of the email.

            .PARAMETER Attachments
            An array of file paths to attach to the email.

            .PARAMETER Priority
            The email priority.

            Valid values are:
            - 'Low'
            - 'Normal'
            - 'High'

            .EXAMPLE
            # Send an email with StartTls and credential

            $SmtpUserName = 'smtpUser'
            $SmtpPassword = 'smtpPassword'

            $securePassword = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($SmtpUserName, $securePassword)

            $params = @{
                SmtpServerName = 'SMT_SERVER@example.com'
                SmtpPort = 587
                SmtpConnectionType = 'StartTls'
                Credential = $credential
                from = 'm@example.com'
                To = '007@example.com'
                Body = '<p>Mission details in attachment</p>'
                Subject = 'For your eyes only'
                Priority = 'High'
                Attachments = @('c:\Mission.ppt', 'c:\ID.pdf')
                MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
            }

            Send-MailKitMessageHC @params

            .EXAMPLE
            # Send an email without authentication

            $params = @{
                SmtpServerName      = 'SMT_SERVER@example.com'
                SmtpPort            = 25
                From                = 'hacker@example.com'
                FromDisplayName     = 'White hat hacker'
                Bcc                 = @('james@example.com', 'mike@example.com')
                Body                = '<h1>You have been hacked</h1>'
                Subject             = 'Oops'
                MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
            }

            Send-MailKitMessageHC @params
            #>

        [CmdletBinding()]
        param (
            [parameter(Mandatory)]
            [string]$MailKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$MimeKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$SmtpServerName,
            [parameter(Mandatory)]
            [ValidateSet(25, 465, 587, 2525)]
            [int]$SmtpPort,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
            [parameter(Mandatory)]
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [string]$FromDisplayName,
            [string[]]$To,
            [string[]]$Bcc,
            [int]$MaxAttachmentSize = 20MB,
            [ValidateSet(
                'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
            )]
            [string]$SmtpConnectionType = 'None',
            [ValidateSet('Normal', 'Low', 'High')]
            [string]$Priority = 'Normal',
            [string[]]$Attachments,
            [PSCredential]$Credential
        )

        begin {
            function Test-IsAssemblyLoaded {
                param (
                    [String]$Name
                )
                foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                    if ($assembly.FullName -like "$Name, Version=*") {
                        return $true
                    }
                }
                return $false
            }

            function Add-Attachments {
                param (
                    [string[]]$Attachments,
                    [MimeKit.Multipart]$BodyMultiPart
                )

                $attachmentList = New-Object System.Collections.ArrayList($null)

                foreach (
                    $attachmentPath in
                    $Attachments | Sort-Object -Unique
                ) {
                    try {
                        #region Test if file exists
                        try {
                            $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                            if ($attachmentItem.PSIsContainer) {
                                Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                                continue
                            }
                        }
                        catch {
                            Write-Warning "Attachment '$attachmentPath' not found"
                            continue
                        }
                        #endregion

                        $totalSizeAttachments += $attachmentItem.Length

                        $null = $attachmentList.Add($attachmentItem)

                        #region Check size of attachments
                        if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                            $M = 'The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details.' -f
                            ([math]::Round(($MaxAttachmentSize / 1MB))),
                            ([math]::Round(($totalSizeAttachments / 1MB), 2))

                            Write-Warning $M

                            return [PSCustomObject]@{
                                AttachmentLimitExceededMessage = $M
                            }
                        }
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentPath': $_"
                    }
                }
                #endregion

                foreach (
                    $attachmentItem in
                    $attachmentList
                ) {
                    try {
                        Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                        $attachment = New-Object MimeKit.MimePart

                        #region Create a MemoryStream to hold the file content
                        $memoryStream = New-Object System.IO.MemoryStream

                        try {
                            $fileStream = [System.IO.File]::OpenRead($attachmentItem.FullName)
                            $fileStream.CopyTo($memoryStream)
                        }
                        finally {
                            if ($fileStream) {
                                $fileStream.Dispose()
                            }
                        }

                        $memoryStream.Position = 0
                        #endregion

                        $attachment.Content = New-Object MimeKit.MimeContent($memoryStream)

                        $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                        $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                        $attachment.FileName = $attachmentItem.Name

                        $bodyMultiPart.Add($attachment)
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentItem': $_"
                    }
                }
            }

            try {
                #region Test To or Bcc required
                if (-not ($To -or $Bcc)) {
                    throw "Either 'To' to 'Bcc' is required for sending emails"
                }
                #endregion

                #region Test To
                foreach ($email in $To) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "To email address '$email' not valid."
                    }
                }
                #endregion

                #region Test Bcc
                foreach ($email in $Bcc) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "Bcc email address '$email' not valid."
                    }
                }
                #endregion

                #region Load MimeKit assembly
                if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                    try {
                        Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                        Add-Type -Path $MimeKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                    }
                }
                #endregion

                #region Load MailKit assembly
                if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                    try {
                        Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                        Add-Type -Path $MailKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                    }
                }
                #endregion
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
        }

        process {
            try {
                $message = New-Object -TypeName 'MimeKit.MimeMessage'

                #region Create body with attachments
                $bodyPart = New-Object MimeKit.TextPart('html')
                $bodyPart.Text = $Body

                $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
                $bodyMultiPart.Add($bodyPart)

                if ($Attachments) {
                    $params = @{
                        Attachments   = $Attachments
                        BodyMultiPart = $bodyMultiPart
                    }
                    $addAttachments = Add-Attachments @params

                    if ($addAttachments.AttachmentLimitExceededMessage) {
                        $bodyPart.Text += '<p><i>{0}</i></p>' -f
                        $addAttachments.AttachmentLimitExceededMessage
                    }
                }

                $message.Body = $bodyMultiPart
                #endregion

                $fromAddress = New-Object MimeKit.MailboxAddress(
                    $FromDisplayName, $From
                )
                $message.From.Add($fromAddress)

                foreach ($email in $To) {
                    $message.To.Add($email)
                }

                foreach ($email in $Bcc) {
                    $message.Bcc.Add($email)
                }

                $message.Subject = $Subject

                #region Set priority
                switch ($Priority) {
                    'Low' {
                        $message.Headers.Add('X-Priority', '5 (Lowest)')
                        break
                    }
                    'Normal' {
                        $message.Headers.Add('X-Priority', '3 (Normal)')
                        break
                    }
                    'High' {
                        $message.Headers.Add('X-Priority', '1 (Highest)')
                        break
                    }
                    default {
                        throw "Priority type '$_' not supported"
                    }
                }
                #endregion

                $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

                try {
                    $smtp.Connect(
                        $SmtpServerName, $SmtpPort,
                        [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
                    )
                }
                catch {
                    throw "Failed to connect to SMTP server '$SmtpServerName' on port '$SmtpPort' with connection type '$SmtpConnectionType': $_"
                }

                if ($Credential) {
                    try {
                        $smtp.Authenticate(
                            $Credential.UserName,
                            $Credential.GetNetworkCredential().Password
                        )
                    }
                    catch {
                        throw "Failed to authenticate with user name '$($Credential.UserName)' to SMTP server '$SmtpServerName': $_"
                    }
                }

                Write-Verbose "Send mail to '$To' with subject '$Subject'"

                $null = $smtp.Send($message)
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
            finally {
                if ($smtp) {
                    $smtp.Disconnect($true)
                    $smtp.Dispose()
                }
                if ($message) {
                    $message.Dispose()
                }
            }
        }
    }

    function Write-EventsToEventLogHC {
        <#
        .SYNOPSIS
            Write events to the event log.

        .DESCRIPTION
            The use of this function will allow standardization in the Windows
            Event Log by using the same EventID's and other properties across
            different scripts.

            Custom Windows EventID's based on the PowerShell standard streams:

            PowerShell Stream     EventIcon    EventID   EventDescription
            -----------------     ---------    -------   ----------------
            [i] Info              [i] Info     100       Script started
            [4] Verbose           [i] Info     4         Verbose message
            [1] Output/Success    [i] Info     1         Output on success
            [3] Warning           [w] Warning  3         Warning message
            [2] Error             [e] Error    2         Fatal error message
            [i] Info              [i] Info     199       Script ended successfully

        .PARAMETER Source
            Specifies the script name under which the events will be logged.

        .PARAMETER LogName
            Specifies the name of the event log to which the events will be
            written. If the log does not exist, it will be created.

        .PARAMETER Events
            Specifies the events to be written to the event log. This should be
            an array of PSCustomObject with properties: Message, EntryType, and
            EventID.

        .PARAMETER Events.xxx
            All properties that are not 'EntryType' or 'EventID' will be used to
            create a formatted message.

        .PARAMETER Events.EntryType
            The type of the event.

            The following values are supported:
            - Information
            - Warning
            - Error
            - SuccessAudit
            - FailureAudit

            The default value is Information.

        .PARAMETER Events.EventID
            The ID of the event. This should be a number.
            The default value is 4.

        .EXAMPLE
            $eventLogData = [System.Collections.Generic.List[PSObject]]::new()

            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script started'
                    EntryType = 'Information'
                    EventID   = '100'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message  = 'Failed to read the file'
                    FileName = 'C:\Temp\test.txt'
                    DateTime = Get-Date
                    EntryType = 'Error'
                    EventID   = '2'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message  = 'Created file'
                    FileName = 'C:\Report.xlsx'
                    FileSize = 123456
                    DateTime = Get-Date
                    EntryType = 'Information'
                    EventID   = '1'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script finished'
                    EntryType = 'Information'
                    EventID   = '199'
                }
            )

            $params = @{
                Source  = 'Test (Brecht)'
                LogName = 'HCScripts'
                Events  = $eventLogData
            }
            Write-EventsToEventLogHC @params
        #>

        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [String]$Source,
            [Parameter(Mandatory)]
            [String]$LogName,
            [PSCustomObject[]]$Events
        )

        try {
            if (
                -not(
                    ([System.Diagnostics.EventLog]::Exists($LogName)) -and
                    [System.Diagnostics.EventLog]::SourceExists($Source)
                )
            ) {
                Write-Verbose "Create event log '$LogName' and source '$Source'"
                New-EventLog -LogName $LogName -Source $Source -ErrorAction Stop
            }

            foreach ($eventItem in $Events) {
                $params = @{
                    LogName     = $LogName
                    Source      = $Source
                    EntryType   = $eventItem.EntryType
                    EventID     = $eventItem.EventID
                    Message     = ''
                    ErrorAction = 'Stop'
                }

                if (-not $params.EntryType) {
                    $params.EntryType = 'Information'
                }
                if (-not $params.EventID) {
                    $params.EventID = 4
                }

                foreach (
                    $property in
                    $eventItem.PSObject.Properties | Where-Object {
                        ($_.Name -ne 'EntryType') -and ($_.Name -ne 'EventID')
                    }
                ) {
                    $params.Message += "`n- $($property.Name) '$($property.Value)'"
                }

                Write-Verbose "Write event to log '$LogName' source '$Source' message '$($params.Message)'"

                Write-EventLog @params
            }
        }
        catch {
            throw "Failed to write to event log '$LogName' source '$Source': $_"
        }
    }

    function Get-StringValueHC {
        <#
        .SYNOPSIS
            Retrieve a string from the environment variables or a regular string.

        .DESCRIPTION
            This function checks the 'Name' property. If the value starts with
            'ENV:', it attempts to retrieve the string value from the specified
            environment variable. Otherwise, it returns the value directly.

        .PARAMETER Name
            Either a string starting with 'ENV:'; a plain text string or NULL.

        .EXAMPLE
            Get-StringValueHC -Name 'ENV:passwordVariable'

            # Output: the environment variable value of $ENV:passwordVariable
            # or an error when the variable does not exist

        .EXAMPLE
            Get-StringValueHC -Name 'mySecretPassword'

            # Output: mySecretPassword

        .EXAMPLE
            Get-StringValueHC -Name ''

            # Output: NULL
        #>
        param (
            [String]$Name
        )

        if (-not $Name) {
            return $null
        }
        elseif (
            $Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            $envVariableName = $Name.Substring(4).Trim()
            $envStringValue = Get-Item -Path "Env:\$envVariableName" -EA Ignore
            if ($envStringValue) {
                return $envStringValue.Value
            }
            else {
                throw "Environment variable '$envVariableName' not found."
            }
        }
        else {
            return $Name
        }
    }

    try {
        $settings = $jsonFileContent.Settings

        $scriptName = $settings.ScriptName
        $saveInEventLog = $settings.SaveInEventLog
        $sendMail = $settings.SendMail
        $saveLogFiles = $settings.SaveLogFiles

        #region Get script name
        if (-not $scriptName) {
            Write-Warning "No 'Settings.ScriptName' found in import file."
            $scriptName = 'Default script name'
        }
        #endregion

        $logFileExtensions = $saveLogFiles.Where.FileExtensions
        $isLog = @{
            systemErrors     = $saveLogFiles.What.SystemErrors
            allActions       = $saveLogFiles.What.AllActions
            onlyActionErrors = $saveLogFiles.What.OnlyActionErrors
        }

        $allLogFilePaths = @()

        $baseLogName = $null
        $logFolderPath = $null

        #region Create system errors log file
        try {
            $logFolder = Get-StringValueHC $saveLogFiles.Where.Folder

            if ($logFolder -and $logFileExtensions) {
                #region Get log folder
                try {
                    $logFolderPath = Get-LogFolderHC -Path $logFolder

                    Write-Verbose "Log folder '$logFolderPath'"

                    $baseLogName = Join-Path -Path $logFolderPath -ChildPath (
                        '{0} - {1} ({2})' -f
                        $scriptStartTime.ToString('yyyy_MM_dd_HHmmss'),
                        $ScriptName,
                        $jsonFileItem.BaseName
                    )
                }
                catch {
                    throw "Failed creating log folder '$LogFolder': $_"
                }
                #endregion

                #region Create system errors log file
                if ($isLog.SystemErrors -and $systemErrors) {
                    $params = @{
                        DataToExport   = $systemErrors
                        PartialPath    = "$baseLogName - System errors log"
                        FileExtensions = $logFileExtensions
                    }
                    $allLogFilePaths += Out-LogFileHC @params
                }
                #endregion
            }
        }
        catch {
            $systemErrors += [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Failed creating log file in folder '$($saveLogFiles.Where.Folder)': $_"
            }

            Write-Warning $systemErrors[-1].Message
        }
        #endregion

        #region Remove old log files
        if ($saveLogFiles.DeleteLogsAfterDays -gt 0 -and $logFolderPath) {
            $cutoffDate = (Get-Date).AddDays(-$saveLogFiles.DeleteLogsAfterDays)

            Write-Verbose "Removing log files older than $cutoffDate from '$logFolderPath'"

            Get-ChildItem -Path $logFolderPath -File |
            Where-Object { $_.LastWriteTime -lt $cutoffDate } |
            ForEach-Object {
                try {
                    $fileToRemove = $_
                    Write-Verbose "Deleting old log file '$_''"
                    Remove-Item -Path $_.FullName -Force
                }
                catch {
                    $systemErrors += [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Failed to remove file '$fileToRemove': $_"
                    }

                    Write-Warning $systemErrors[-1].Message

                    if ($baseLogName -and $isLog.systemErrors) {
                        $params = @{
                            DataToExport   = $systemErrors[-1]
                            PartialPath    = "$baseLogName - Errors"
                            FileExtensions = $logFileExtensions
                        }
                        $allLogFilePaths += Out-LogFileHC @params -EA Ignore
                    }
                }
            }
        }
        #endregion

        #region Write events to event log
        try {
            $saveInEventLog.LogName = Get-StringValueHC $saveInEventLog.LogName

            if ($saveInEventLog.Save -and $saveInEventLog.LogName) {
                $systemErrors | ForEach-Object {
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            DateTime  = $_.DateTime
                            Error     = $_.Message
                            EntryType = 'Error'
                            EventID   = '2'
                        }
                    )
                }

                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = 'Script ended'
                        EntryType = 'Information'
                        EventID   = '199'
                    }
                )

                $params = @{
                    Source  = $scriptName
                    LogName = $saveInEventLog.LogName
                    Events  = $eventLogData
                }
                Write-EventsToEventLogHC @params

            }
            elseif ($saveInEventLog.Save -and (-not $saveInEventLog.LogName)) {
                throw "Both 'Settings.SaveInEventLog.Save' and 'Settings.SaveInEventLog.LogName' are required to save events in the event log."
            }
        }
        catch {
            $systemErrors += [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Failed writing events to event log: $_"
            }

            Write-Warning $systemErrors[-1].Message

            if ($baseLogName -and $isLog.systemErrors) {
                $params = @{
                    DataToExport   = $systemErrors[-1]
                    PartialPath    = "$baseLogName - Errors"
                    FileExtensions = $logFileExtensions
                }
                $allLogFilePaths += Out-LogFileHC @params -EA Ignore
            }
        }
        #endregion

        $isSendMail = $false

        $counter = @{
            logFileDataErrors = 0
        }

        foreach ($result in $results) {
            $logFileData = $result.CoordinatePairs
            $logFileDataErrors = $logFileData.Where({ $_.errors })

            $counter.logFileDataErrors += $logFileDataErrors.Count

            if ($baseLogName -and $logFileData) {
                $params = @{
                    PartialPath    = "$baseLogName - $($result.File.BaseName) - Log"
                    FileExtensions = $logFileExtensions
                    DataToExport   = $null
                }

                if ($isLog.allActions) {
                    $params.DataToExport = $logFileData
                }
                elseif ($isLog.onlyActionErrors -and $logFileDataErrors) {
                    $params.DataToExport = $logFileDataErrors
                }

                if ($params.DataToExport) {
                    $params.DataToExport = $params.DataToExport |
                    Select-Object -Property @{
                        Name       = 'dateTime'
                        Expression = { $_.dateTime }
                    },
                    @{
                        Name       = 'startCoordinate'
                        Expression = { $_.coordinate.start }
                    },
                    @{
                        Name       = 'destinationCoordinate'
                        Expression = { $_.coordinate.destination }
                    },
                    @{
                        Name       = 'distanceInMeters'
                        Expression = { $_.apiResponse.routes[0].distance }
                    },
                    @{
                        Name       = 'durationInSeconds'
                        Expression = { $_.apiResponse.routes[0].duration }
                    },
                    @{
                        Name       = 'error'
                        Expression = { $_.errors -join ', ' }
                    }

                    $allLogFilePaths += Out-LogFileHC @params
                }
            }

            #region Test send email
            try {
                switch ($sendMail.When) {
                    'OnError' {
                        if ($logFileDataErrors) {
                            $isSendMail = $true
                        }
                        break
                    }
                    'OnErrorOrAction' {
                        if (
                            $logFileDataErrors -or
                            $logFileData
                        ) {
                            $isSendMail = $true
                        }
                        break
                    }
                }
            }
            catch {
                $systemErrors += [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed sending email: $_"
                }

                Write-Warning $systemErrors[-1].Message

                if ($baseLogName -and $isLog.systemErrors) {
                    $params = @{
                        DataToExport   = $systemErrors[-1]
                        PartialPath    = "$baseLogName - Errors"
                        FileExtensions = $logFileExtensions
                    }
                    $null = Out-LogFileHC @params -EA Ignore
                }
            }
            #endregion
        }

        #region Send email
        try {
            switch ($sendMail.When) {
                'Never' {
                    $isSendMail = $false
                    break
                }
                'Always' {
                    $isSendMail = $true
                    break
                }
                'OnError' {
                    if ($systemErrors) {
                        $isSendMail = $true
                    }
                    break
                }
                'OnErrorOrAction' {
                    if ($systemErrors) {
                        $isSendMail = $true
                    }
                    break
                }
                default {
                    throw "SendMail.When '$($sendMail.When)' not supported. Supported values are 'Never', 'Always', 'OnError' or 'OnErrorOrAction'."
                }
            }

            if ($isSendMail) {
                #region Test mandatory fields
                @{
                    'From'                 = $sendMail.From
                    'Smtp.ServerName'      = $sendMail.Smtp.ServerName
                    'Smtp.Port'            = $sendMail.Smtp.Port
                    'AssemblyPath.MailKit' = $sendMail.AssemblyPath.MailKit
                    'AssemblyPath.MimeKit' = $sendMail.AssemblyPath.MimeKit
                }.GetEnumerator() |
                Where-Object { -not $_.Value } | ForEach-Object {
                    throw "Input file property 'Settings.SendMail.$($_.Key)' cannot be blank"
                }
                #endregion

                $mailParams = @{
                    From                = Get-StringValueHC $sendMail.From
                    Subject             = '{0} trip{1}, {2} file{3}' -f
                    $logFileData.Count,
                    $(if ($logFileData.Count -ne 1) { 's' }),
                    $excelFiles.Count,
                    $(if ($excelFiles.Count -ne 1) { 's' })
                    SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
                    SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
                    MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
                    MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
                }

                $mailParams.Body = @"
<!DOCTYPE html>
<html>
    <head>
        <style type="text/css">
            body {
                font-family:verdana;
                font-size:14px;
                background-color:white;
            }
            h1 {
                margin-bottom: 0;
            }
            h2 {
                margin-bottom: 0;
            }
            h3 {
                margin-bottom: 0;
            }
            p.italic {
                font-style: italic;
                font-size: 12px;
            }
            table {
                border-collapse:collapse;
                border:0px none;
                padding:3px;
                text-align:left;
            }
            td, th {
                border-collapse:collapse;
                border:1px none;
                padding:3px;
                text-align:left;
            }
            #aboutTable th {
                color: rgb(143, 140, 140);
                font-weight: normal;
            }
            #aboutTable td {
                color: rgb(143, 140, 140);
                font-weight: normal;
            }
            base {
                target="_blank"
            }
        </style>
    </head>
    <body>
        <table>
            <h1>$scriptName</h1>
            <hr size="2" color="#06cc7a">

            $($sendMail.Body)

            <table>
                $(
                    if($counter.logFileDataErrors) {
                        "<tr style=`"background-color: #ffe5ec;`">
                            <th>Retrieval or update errors</th>
                            <td>$($counter.logFileDataErrors)</td>
                        </tr>"
                    }
                )
                $(
                    if($systemErrors.Count) {
                        "<tr style=`"background-color: #ffe5ec;`">
                           <th>System errors</th>
                            <td>$($systemErrors.Count)</td>
                        </tr>"
                    }
                )
                <tr>
                    <th>Start and destination pairs</th>
                    <td>$($logFileData.Count)</td>
                </tr>
                <tr>
                    <th>Excel files</th>
                    <td>
                        <ul>
                        $(
                            $excelFiles.foreach(
                                {"<li>$($_.BaseName)</li>" }
                            )
                        )
                        </ul>
                    </td>
                </tr>
                <tr>
                    <th>Drop folder</th>
                    <td>$('<a href="{0}">{0}</a>' -f $DropFolderPath)</td>
                </tr>
                $(
                    if ($ArchiveFolderPath) {
                        '<tr>
                            <th>Archive folder</th>
                            <td><a href="{0}">{0}</a></td>
                        </tr>' -f $ArchiveFolderPath
                    }
                )
            </table>

            $(
                if ($allLogFilePaths) {
                    '<p><i>* Check the attachment(s) for details</i></p>'
                }
            )

            <hr size="2" color="#06cc7a">
            <table id="aboutTable">
                $(
                    if ($scriptStartTime) {
                        '<tr>
                            <th>Start time</th>
                            <td>{0:00}/{1:00}/{2:00} {3:00}:{4:00} ({5})</td>
                        </tr>' -f
                        $scriptStartTime.Day,
                        $scriptStartTime.Month,
                        $scriptStartTime.Year,
                        $scriptStartTime.Hour,
                        $scriptStartTime.Minute,
                        $scriptStartTime.DayOfWeek
                    }
                )
                $(
                    if ($scriptStartTime) {
                        $runTime = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
                        '<tr>
                            <th>Duration</th>
                            <td>{0:00}:{1:00}:{2:00}</td>
                        </tr>' -f
                        $runTime.Hours, $runTime.Minutes, $runTime.Seconds
                    }
                )
                $(
                    if ($logFolderPath) {
                        '<tr>
                            <th>Log files</th>
                            <td><a href="{0}">Open log folder</a></td>
                        </tr>' -f $logFolderPath
                    }
                )
                <tr>
                    <th>Host</th>
                    <td>$($host.Name)</td>
                </tr>
                <tr>
                    <th>PowerShell</th>
                    <td>$($PSVersionTable.PSVersion.ToString())</td>
                </tr>
                <tr>
                    <th>Computer</th>
                    <td>$env:COMPUTERNAME</td>
                </tr>
                <tr>
                    <th>Account</th>
                    <td>$env:USERDNSDOMAIN\$env:USERNAME</td>
                </tr>
            </table>
        </table>
    </body>
</html>
"@

                if ($sendMail.FromDisplayName) {
                    $mailParams.FromDisplayName = Get-StringValueHC $sendMail.FromDisplayName
                }

                if ($sendMail.Subject) {
                    $mailParams.Subject = '{0}, {1}' -f
                    $mailParams.Subject, $sendMail.Subject
                }

                if ($sendMail.To) {
                    $mailParams.To = $sendMail.To
                }

                if ($sendMail.Bcc) {
                    $mailParams.Bcc = $sendMail.Bcc
                }

                if ($systemErrors -or $counter.logFileDataErrors) {
                    $totalErrorCount = $systemErrors.Count + $counter.logFileDataErrors

                    $mailParams.Priority = 'High'
                    $mailParams.Subject = '{0} error{1}, {2}' -f
                    $totalErrorCount,
                    $(if ($totalErrorCount -ne 1) { 's' }),
                    $mailParams.Subject
                }

                if ($allLogFilePaths) {
                    $mailParams.Attachments = $allLogFilePaths |
                    Sort-Object -Unique
                }

                if ($sendMail.Smtp.ConnectionType) {
                    $mailParams.SmtpConnectionType = Get-StringValueHC $sendMail.Smtp.ConnectionType
                }

                #region Create SMTP credential
                $smtpUserName = Get-StringValueHC $sendMail.Smtp.UserName
                $smtpPassword = Get-StringValueHC $sendMail.Smtp.Password

                if ( $smtpUserName -and $smtpPassword) {
                    try {
                        $securePassword = ConvertTo-SecureString -String $smtpPassword -AsPlainText -Force

                        $credential = New-Object System.Management.Automation.PSCredential($smtpUserName, $securePassword)

                        $mailParams.Credential = $credential
                    }
                    catch {
                        throw "Failed to create credential: $_"
                    }
                }
                elseif ($smtpUserName -or $smtpPassword) {
                    throw "Both 'Settings.SendMail.Smtp.Username' and 'Settings.SendMail.Smtp.Password' are required when authentication is needed."
                }
                #endregion

                Write-Verbose "Send email to '$($mailParams.To)' subject '$($mailParams.Subject)'"

                Send-MailKitMessageHC @mailParams
            }
        }
        catch {
            $systemErrors += [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Failed sending email: $_"
            }

            Write-Warning $systemErrors[-1].Message

            if ($baseLogName -and $isLog.systemErrors) {
                $params = @{
                    DataToExport   = $systemErrors[-1]
                    PartialPath    = "$baseLogName - Errors"
                    FileExtensions = $logFileExtensions
                }
                $null = Out-LogFileHC @params -EA Ignore
            }
        }
        #endregion
    }
    catch {
        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Message  = $_
        }

        Write-Warning $systemErrors[-1].Message
    }
    finally {
        if ($systemErrors) {
            $M = 'Found {0} system error{1}' -f
            $systemErrors.Count,
            $(if ($systemErrors.Count -ne 1) { 's' })
            Write-Warning $M

            $systemErrors | ForEach-Object {
                Write-Warning $_.Message
            }

            Write-Warning 'Exit script with error code 1'
            exit 1
        }
        else {
            Write-Verbose 'Script finished successfully'
        }
    }
}
