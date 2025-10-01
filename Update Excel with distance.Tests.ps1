#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $realCmdLet = @{
        OutFile = Get-Command Out-File
    }

    $testInputFile = @{
        Excel    = @{
            FilePath      = (New-Item 'TestDrive:/file.xlsx' -ItemType File).FullName
            WorksheetName = 'sheet1'
            Column        = @{
                startDestination = 'B'
                coordinate       = 'A'
                distance         = 'C'
                duration         = 'D'
            }
        }
        Settings = @{
            ScriptName     = 'Test (Brecht)'
            SendMail       = @{
                When         = 'Always'
                From         = 'm@example.com'
                To           = '007@example.com'
                Subject      = 'Email subject'
                Body         = 'Email body'
                Smtp         = @{
                    ServerName     = 'SMTP_SERVER'
                    Port           = 25
                    ConnectionType = 'StartTls'
                    UserName       = 'bob'
                    Password       = 'pass'
                }
                AssemblyPath = @{
                    MailKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }
            }
            SaveLogFiles   = @{
                What                = @{
                    SystemErrors     = $true
                    AllActions       = $true
                    OnlyActionErrors = $false
                }
                Where               = @{
                    Folder         = (New-Item 'TestDrive:/log' -ItemType Directory).FullName
                    FileExtensions = @('.json', '.csv')
                }
                deleteLogsAfterDays = 1
            }
            SaveInEventLog = @{
                Save    = $true
                LogName = 'Scripts'
            }
        }
    }

    $testOutParams = @{
        FilePath = (New-Item 'TestDrive:/Test.json' -ItemType File).FullName
    }

    $testData = @(
        [PSCustomObject]@{
            Coordinate = 1; Type = 'S'; Distance = $null; Duration = $null 
        }
        [PSCustomObject]@{
            Coordinate = 2; Type = 'D'; Distance = $null ; Duration = $null 
        }
        [PSCustomObject]@{
            Coordinate = 3; Type = 'S'; Distance = $null ; Duration = $null 
        }
        [PSCustomObject]@{
            Coordinate = 4; Type = 'D'; Distance = $null ; Duration = $null 
        }
    )

    $testData | Export-Excel -Path $testInputFile.Excel.FilePath

    $testExportedLogFileData = @(
        [PSCustomObject]@{
            dateTime              = Get-Date
            startCoordinate       = 1
            destinationCoordinate = 2
            distanceInMeters      = 1033101.5
            durationInSeconds     = 143222.4
            error                 = ''
        }
        [PSCustomObject]@{
            dateTime              = Get-Date
            startCoordinate       = 3
            destinationCoordinate = 4
            distanceInMeters      = 55000
            durationInSeconds     = 6000
            error                 = ''
        }
    )

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ConfigurationJsonFile = $testOutParams.FilePath
    }

    function Copy-ObjectHC {
        <#
        .SYNOPSIS
            Make a deep copy of an object using JSON serialization.

        .DESCRIPTION
            Uses ConvertTo-Json and ConvertFrom-Json to create an independent
            copy of an object. This method is generally effective for objects
            that can be represented in JSON format.

        .PARAMETER InputObject
            The object to copy.

        .EXAMPLE
            $newArray = Copy-ObjectHC -InputObject $originalArray
        #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $jsonString = $InputObject | ConvertTo-Json -Depth 100

        $deepCopy = $jsonString | ConvertFrom-Json

        return $deepCopy
    }
    function Send-MailKitMessageHC {
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
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
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
    }
    function Test-GetLogFileDataHC {
        param (
            [String]$FileNameRegex = '* - System errors log.json',
            [String]$LogFolderPath = $testInputFile.Settings.SaveLogFiles.Where.Folder
        )

        $testLogFile = Get-ChildItem -Path $LogFolderPath -File -Filter $FileNameRegex

        if ($testLogFile.count -eq 1) {
            Get-Content $testLogFile | ConvertFrom-Json
        }
        elseif (-not $testLogFile) {
            throw "No log file found in folder '$LogFolderPath' matching '$FileNameRegex'"
        }
        else {
            throw "Found multiple log files in folder '$LogFolderPath' matching '$FileNameRegex'"
        }
    }

    Mock Invoke-RestMethod
    Mock Send-MailKitMessageHC
    Mock New-EventLog
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ConfigurationJsonFile') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'create an error log file when' {
    It 'the log folder cannot be created' {
        Mock Out-File

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = 'x:\notExistingLocation'

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        $LASTEXITCODE | Should -Be 1

        Should -Not -Invoke Out-File
    }
    Context 'the ConfigurationJsonFile' {
        It 'is not found' {
            Mock Out-File

            $testNewParams = $testParams.clone()
            $testNewParams.ConfigurationJsonFile = 'nonExisting.json'

            .$testScript @testNewParams

            $LASTEXITCODE | Should -Be 1

            Should -Not -Invoke Out-File
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'Excel'
            ) {
                Mock Out-File

                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - System errors log.json') -and
                    ($InputObject -like "*Property '$_' not found*")
                }
            }
            It 'Excel.<_> not found' -ForEach @(
                'FilePath', 'WorksheetName', 'Column'
            ) {
                Mock Out-File

                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Excel.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - System errors log.json') -and
                    ($InputObject -like "*Property 'Excel.$_' not found*")
                }
            }
            It 'Excel.Column.<_> not found' -ForEach @(
                'StartDestination', 'Coordinate', 'Distance', 'Duration'
            ) {
                Mock Out-File

                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Excel.Column.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - System errors log.json') -and
                    ($InputObject -like "*Property 'Excel.Column.$_' not found*")
                }
            }
        }
    }
    It 'The Excel.FilePath cannot be found' {
        Mock Out-File

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Excel.FilePath = 'TestDrive:\NotExisting.xslx'

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        $LASTEXITCODE | Should -Be 1

        Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
            ($LiteralPath -like '* - System errors log.json') -and
            ($InputObject -like "*Excel file 'TestDrive:\\NotExisting.xslx' not found*")
        }
    }
}
Describe 'when the script runs successfully' {
    BeforeAll {
        Mock Invoke-RestMethod {
            @{
                code   = 'Ok'
                routes = @(
                    @{
                        'duration' = 143222.4
                        'distance' = 1033101.5
                    }
                )
            }
        } -ParameterFilter {
            $Uri -eq 'https://router.project-osrm.org/route/v1/driving/1;2'
        }
        Mock Invoke-RestMethod {
            @{
                code   = 'Ok'
                routes = @(
                    @{
                        'duration' = 6000
                        'distance' = 55000
                    }
                )
            }
        } -ParameterFilter {
            $Uri -eq 'https://router.project-osrm.org/route/v1/driving/3;4'
        }
    
        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'create a log file' {
        BeforeAll {
            $actual = Test-GetLogFileDataHC -FileNameRegex '* - Log.json'
        }
        It 'in the log folder' {
            $actual | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedLogFileData.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedLogFileData) {
                $actualRow = $actual | Where-Object {
                    $_.startCoordinate -eq $testRow.startCoordinate
                }
                $actualRow.destinationCoordinate | 
                Should -Be $testRow.destinationCoordinate
                $actualRow.distanceInMeters |
                Should -Be $testRow.distanceInMeters
                $actualRow.durationInSeconds | 
                Should -Be $testRow.durationInSeconds
                $actualRow.error | Should -Be $testRow.error
                $actualRow.dateTime.ToString('yyyyMMdd') |
                Should -Be $testRow.dateTime.ToString('yyyyMMdd')
            }
        }
    } -Tag test
    Context 'send an e-mail' {
        It 'with attachment to the user' {
            Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($From -eq 'm@example.com') -and
                ($To -eq '007@example.com') -and
                ($SmtpPort -eq 25) -and
                ($SmtpServerName -eq 'SMTP_SERVER') -and
                ($SmtpConnectionType -eq 'StartTls') -and
                ($Subject -eq '2 moved, Email subject') -and
                ($Credential) -and
                ($Attachments -like '*- Log.json') -and
                ($Body -like '*Email body*Summary of SFTP actions*table*App x*<th>sftp:/sftp.server.com</th>*Source*Destination*Result*\a*sftp:/folder/a/*1 moved*sftp:/folder/b/*\b*1 moved*<th>2 moved on PC1</th>*') -and
                ($MailKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll') -and
                ($MimeKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll')
            }
        }
    }
}