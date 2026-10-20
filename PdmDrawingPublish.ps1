[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DrawingPath,

    [Parameter(Mandatory = $true)]
    [string]$Revision,

    [Parameter(Mandatory = $true)]
    [string]$VaultName,

    [string]$SharePath = '\\srvvcad\AutoCAD\Zeichnungen'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:swDocPart = 1
$script:swDocDrawing = 3
$script:swOpenDocOptionsSilent = 1
$script:swSaveAsOptionsSilent = 1
$script:swExportPdfData = 1
$script:swExportDxfDwgData = 2
$script:swExportAllSheets = 0

function Get-SafeFileSegment {
    param([Parameter(Mandatory = $true)][string]$Value)

    $result = [regex]::Replace($Value, '[<>:"/\\|?*\x00-\x1f]', '_').Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($result)) {
        throw "Revision or sheet name '$Value' is not a usable filename."
    }
    return $result
}

function Start-SolidWorks {
    $application = New-Object -ComObject SldWorks.Application
    $application.Visible = $false
    return $application
}

function Stop-SolidWorks {
    param([object]$Application)

    if ($null -ne $Application) {
        try {
            $Application.ExitApp()
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Application)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Close-SolidWorksDocument {
    param(
        [Parameter(Mandatory = $true)][object]$Application,
        [Parameter(Mandatory = $true)][object]$Document
    )

    try {
        try {
            $Application.CloseDoc($Document.GetTitle())
        }
        catch {
        }
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Document)
    }
}

function Open-SolidWorksDocument {
    param(
        [Parameter(Mandatory = $true)][object]$Application,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$DocumentType
    )

    $errors = 0
    $warnings = 0
    $document = $Application.OpenDoc6(
        $Path,
        $DocumentType,
        $script:swOpenDocOptionsSilent,
        '',
        [ref]$errors,
        [ref]$warnings
    )
    if ($null -eq $document) {
        throw "SolidWorks could not open '$Path' (errors: $errors, warnings: $warnings)."
    }
    return $document
}

function Save-SolidWorksDocument {
    param(
        [Parameter(Mandatory = $true)][object]$Document,
        [Parameter(Mandatory = $true)][string]$Path,
        [object]$ExportData
    )

    $errors = 0
    $warnings = 0
    $saved = $Document.Extension.SaveAs3(
        $Path,
        0,
        $script:swSaveAsOptionsSilent,
        $ExportData,
        $null,
        [ref]$errors,
        [ref]$warnings
    )
    if (-not $saved) {
        throw "SolidWorks could not save '$Path' (errors: $errors, warnings: $warnings)."
    }
}

function Get-DrawingSheetNames {
    param([Parameter(Mandatory = $true)][object]$Drawing)

    $names = @($Drawing.GetSheetNames())
    if ($names.Count -eq 0) {
        throw "The drawing contains no sheets."
    }
    return $names
}

function Export-PdfToTemporaryFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $application = $null
    $drawing = $null
    $pdfData = $null
    try {
        $application = Start-SolidWorks
        $drawing = Open-SolidWorksDocument $application $Path $script:swDocDrawing
        $pdfData = $application.GetExportFileData($script:swExportPdfData)
        if ($null -eq $pdfData) {
            throw 'SolidWorks could not create PDF export settings.'
        }
        $pdfData.SetSheets($script:swExportAllSheets, [string[]]@())
        Save-SolidWorksDocument $drawing $OutputPath $pdfData
    }
    finally {
        if ($null -ne $application) {
            if ($null -ne $pdfData) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($pdfData)
            }
            if ($null -ne $drawing) {
                Close-SolidWorksDocument $application $drawing
            }
            Stop-SolidWorks $application
        }
    }
}

function Export-DxfSheet {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $application = $null
    $drawing = $null
    $dxfData = $null
    try {
        $application = Start-SolidWorks
        $drawing = Open-SolidWorksDocument $application $Path $script:swDocDrawing
        if (-not $drawing.ActivateSheet($SheetName)) {
            throw "SolidWorks could not activate drawing sheet '$SheetName'."
        }
        $dxfData = $application.GetExportFileData($script:swExportDxfDwgData)
        if ($null -eq $dxfData) {
            throw 'SolidWorks could not create DXF export settings.'
        }
        $dxfData.ExportAllSheets = $false
        Save-SolidWorksDocument $drawing $OutputPath $dxfData
    }
    finally {
        if ($null -ne $application) {
            if ($null -ne $dxfData) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dxfData)
            }
            if ($null -ne $drawing) {
                Close-SolidWorksDocument $application $drawing
            }
            Stop-SolidWorks $application
        }
    }
}

function Get-ReferencedModel {
    param([Parameter(Mandatory = $true)][string]$DrawingPath)

    $application = $null
    $drawing = $null
    try {
        $application = Start-SolidWorks
        $drawing = Open-SolidWorksDocument $application $DrawingPath $script:swDocDrawing
        $references = @{}

        foreach ($sheetName in (Get-DrawingSheetNames $drawing)) {
            if (-not $drawing.ActivateSheet($sheetName)) {
                throw "SolidWorks could not activate drawing sheet '$sheetName'."
            }
            $view = $drawing.GetFirstView()
            if ($null -ne $view) {
                $sheetView = $view
                $view = $sheetView.GetNextView()
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheetView)
            }
            while ($null -ne $view) {
                $nextView = $view.GetNextView()
                $modelPath = [string]$view.GetReferencedModelName()
                $configuration = [string]$view.ReferencedConfiguration
                if (-not [string]::IsNullOrWhiteSpace($modelPath)) {
                    $key = $modelPath.ToLowerInvariant() + '|' + $configuration.ToLowerInvariant()
                    $references[$key] = [pscustomobject]@{
                        Path = $modelPath
                        Configuration = $configuration
                    }
                }
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($view)
                $view = $nextView
            }
        }

        if ($references.Count -eq 0) {
            throw "No referenced 3D model was found in '$DrawingPath'."
        }
        if ($references.Count -gt 1) {
            throw "The drawing references multiple models or configurations; STEP export is ambiguous."
        }
        return @($references.Values)[0]
    }
    finally {
        if ($null -ne $application) {
            if ($null -ne $view) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($view)
            }
            if ($null -ne $drawing) {
                Close-SolidWorksDocument $application $drawing
            }
            Stop-SolidWorks $application
        }
    }
}

function Export-Step {
    param(
        [Parameter(Mandatory = $true)][string]$ModelPath,
        [Parameter(Mandatory = $true)][string]$Configuration,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $application = $null
    $model = $null
    try {
        $documentType = if ([IO.Path]::GetExtension($ModelPath).Equals('.sldasm', [StringComparison]::OrdinalIgnoreCase)) {
            2
        }
        else {
            $script:swDocPart
        }
        $application = Start-SolidWorks
        $model = Open-SolidWorksDocument $application $ModelPath $documentType
        if ([string]::IsNullOrWhiteSpace($Configuration) -or -not $model.ShowConfiguration2($Configuration)) {
            throw "Referenced configuration '$Configuration' is not available in '$ModelPath'."
        }
        [void]$model.EditRebuild3()
        Save-SolidWorksDocument $model $OutputPath $null
    }
    finally {
        if ($null -ne $application) {
            if ($null -ne $model) {
                Close-SolidWorksDocument $application $model
            }
            Stop-SolidWorks $application
        }
    }
}

function Add-OrCheckInPdf {
    param(
        [Parameter(Mandatory = $true)][object]$Vault,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][string]$TemporaryPdf
    )

    $folder = $null
    $file = $Vault.GetFileFromPath($PdfPath, [ref]$folder)
    if ($null -ne $file) {
        if ($file.IsLocked) {
            throw "The existing PDM PDF is checked out and cannot be updated: '$PdfPath'."
        }
        $file.LockFile(0, $folder.ID)
        try {
            Copy-Item -LiteralPath $TemporaryPdf -Destination $PdfPath -Force
            $file.UnlockFile(0, 'Updated by PDM drawing publish task')
        }
        catch {
            try { $file.UnlockFile(0, 'PDM drawing publish task failed') } catch { }
            throw
        }
        return
    }

    Copy-Item -LiteralPath $TemporaryPdf -Destination $PdfPath -Force
    $folderPath = Split-Path -LiteralPath $PdfPath -Parent
    $folder = $Vault.GetFolderFromPath($folderPath)
    if ($null -eq $folder) {
        throw "The PDF destination is not inside the logged-in PDM vault: '$folderPath'."
    }
    [void]$folder.AddFile(0, $PdfPath, 'Created by PDM drawing publish task', 0)
}

function Publish-Drawing {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$MainRevision,
        [Parameter(Mandatory = $true)][object]$Vault,
        [Parameter(Mandatory = $true)][string]$DestinationShare
    )

    $fullPath = [IO.Path]::GetFullPath($Path)
    if ([IO.Path]::GetExtension($fullPath) -ine '.slddrw' -or -not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Input must be an existing SLDDRW file: '$Path'."
    }
    if ([string]::IsNullOrWhiteSpace($MainRevision)) {
        throw "The PDM main revision is required for '$fullPath'."
    }

    $revisionPart = Get-SafeFileSegment $MainRevision
    $baseName = [IO.Path]::GetFileNameWithoutExtension($fullPath)
    $vaultPdf = Join-Path (Split-Path -LiteralPath $fullPath -Parent) "$baseName-$revisionPart.pdf"
    $sharePdf = Join-Path $DestinationShare "$baseName-$revisionPart.pdf"
    $temporaryPdf = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString() + '.pdf')

    try {
        Export-PdfToTemporaryFile $fullPath $temporaryPdf
        Add-OrCheckInPdf $Vault $vaultPdf $temporaryPdf
        Copy-Item -LiteralPath $temporaryPdf -Destination $sharePdf -Force
    }
    finally {
        Remove-Item -LiteralPath $temporaryPdf -Force -ErrorAction SilentlyContinue
    }

    $sheets = $null
    $application = $null
    try {
        $application = Start-SolidWorks
        $drawing = Open-SolidWorksDocument $application $fullPath $script:swDocDrawing
        $sheets = Get-DrawingSheetNames $drawing
    }
    finally {
        if ($null -ne $application) {
            if ($null -ne $drawing) {
                Close-SolidWorksDocument $application $drawing
            }
            Stop-SolidWorks $application
        }
    }

    foreach ($sheetName in $sheets) {
        $sheetPart = Get-SafeFileSegment $sheetName
        $dxfPath = Join-Path $DestinationShare "$baseName-$revisionPart-$sheetPart.dxf"
        Export-DxfSheet $fullPath $sheetName $dxfPath
    }

    $reference = Get-ReferencedModel $fullPath
    $stepPath = Join-Path $DestinationShare "$baseName-$revisionPart.step"
    Export-Step $reference.Path $reference.Configuration $stepPath
}

if (-not (Test-Path -LiteralPath $SharePath -PathType Container)) {
    throw "The configured network destination is unavailable: '$SharePath'."
}

$vault = $null
try {
    $vault = New-Object -ComObject ConisioLib.EdmVault
    if (-not $vault.IsLoggedIn) {
        $vault.LoginAuto($VaultName, 0)
    }
    if (-not $vault.IsLoggedIn) {
        throw "Could not log in to the PDM vault '$VaultName'."
    }
    Publish-Drawing $DrawingPath $Revision $vault $SharePath
}
finally {
    if ($null -ne $vault) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($vault)
    }
}
