function Install-XlAddIn{
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true)]
            [String]$AddinPath,
        [Parameter(Mandatory=$false)]
            [Switch]$Reinstall,
        [Parameter(Mandatory=$false)]
            [Switch]$NoCopy
    )

    # Ensure that any errors we receive are considered fatal
    $ErrorActionPreference = 'Stop'

    

    try {
        $Excel = New-Object -ComObject Excel.Application
    } catch {
        Write-Error 'Microsoft Excel does not appear to be installed.'
    }
    # The path to the default folder to copy Excel add-ins
    $ExcelAddinsPath = $Excel.UserLibraryPath

    if (Test-Path -Path $AddinPath -PathType Leaf) {
        $Addin = Get-ChildItem -Path $AddinPath
        if ($Addin.Extension -NotIn ('.xla', '.xlam')) {
            Write-Error 'The file does not appear to be an Excel add-in.'
        }
    } else {
        Write-Error 'The add-in file path does not appear to be valid.'
    }

    try {
        $ExcelAddins = $Excel.Addins
        # The Add() method of the AddIns interface will fail if we don't have a workbook!
        $ExcelWorkbook = $Excel.Workbooks.Add()
        $AddinInstalled = $ExcelAddins | ? { $_.Name -eq $Addin.Name }

        if (!$AddinInstalled -or $Reinstall) {
            if (!(Test-Path -Path $ExcelAddinsPath -PathType Container)) {
                New-Item -Path $ExcelAddinsPath -ItemType Directory
            }

            if (!$NoCopy) {
                Copy-Item -Path $Addin.FullName -Destination $ExcelAddinsPath -Force
                $Addin = Get-ChildItem -Path (Join-Path $ExcelAddinsPath $Addin.Name)
                $Addin.IsReadOnly = $true
            }

            $NewAddin = $ExcelAddins.Add($Addin.FullName, $false)
            $NewAddin.Installed = $true
            Write-Host ('Add-in "' + $Addin.BaseName + '" successfully installed!')
        } else {
            Write-Host ('Add-in "' + $Addin.BaseName + '" already installed!')
        }
    } finally {
        $Excel.Quit()
        [Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel);
    }
}

$scriptRoot = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
Push-Location $scriptRoot

Install-XlAddIn -AddinPath "..\Autofit-CSV.xlam"

Pop-Location