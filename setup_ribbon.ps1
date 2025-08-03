# PowerShell script to set up custom ribbon in AnkiTool.xlsm
# This script will add the custom ribbon to the Excel file

Write-Host "Setting up custom ribbon for AnkiTool.xlsm..." -ForegroundColor Green

# Check if AnkiTool.xlsm exists
$excelFile = "excel\AnkiTool.xlsm"
if (-not (Test-Path $excelFile)) {
    Write-Host "Error: AnkiTool.xlsm not found in excel folder!" -ForegroundColor Red
    exit 1
}

# Check if Ribbon.xml exists
$ribbonFile = "excel\Ribbon.xml"
if (-not (Test-Path $ribbonFile)) {
    Write-Host "Error: Ribbon.xml not found in excel folder!" -ForegroundColor Red
    exit 1
}

try {
    # Create temporary directory
    $tempDir = "temp_ribbon_setup"
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    
    # Copy Excel file to temp directory
    Copy-Item $excelFile "$tempDir\AnkiTool.xlsm"
    
    # Rename to zip
    Rename-Item "$tempDir\AnkiTool.xlsm" "$tempDir\AnkiTool.zip"
    
    # Extract zip
    Expand-Archive "$tempDir\AnkiTool.zip" "$tempDir\extracted" -Force
    
    # Create customUI directory if it doesn't exist
    $customUIDir = "$tempDir\extracted\customUI"
    if (-not (Test-Path $customUIDir)) {
        New-Item -ItemType Directory -Path $customUIDir | Out-Null
    }
    
    # Copy ribbon XML
    Copy-Item $ribbonFile "$customUIDir\customUI.xml"
    
    # Create [Content_Types].xml entry if needed
    $contentTypesFile = "$tempDir\extracted\[Content_Types].xml"
    if (Test-Path $contentTypesFile) {
        $content = Get-Content $contentTypesFile -Raw
        if ($content -notmatch "customUI\.xml") {
            # Add the customUI override
            $overrideEntry = '    <Override PartName="/customUI/customUI.xml" ContentType="application/vnd.ms-office.customui+xml"/>'
            $content = $content -replace '(\s*</Types>)', "$overrideEntry`n`$1"
            Set-Content $contentTypesFile $content -Encoding UTF8
        }
    }
    
    # Create .rels file if needed
    $relsDir = "$tempDir\extracted\_rels"
    if (-not (Test-Path $relsDir)) {
        New-Item -ItemType Directory -Path $relsDir | Out-Null
    }
    
    $relsFile = "$relsDir\.rels"
    if (-not (Test-Path $relsFile)) {
        $relsContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2006/relationships/ui/customui" Target="/customUI/customUI.xml"/>
</Relationships>
"@
        Set-Content $relsFile $relsContent -Encoding UTF8
    } else {
        $relsContent = Get-Content $relsFile -Raw
        if ($relsContent -notmatch "customui") {
            # Add the customUI relationship
            $customUIRel = '    <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2006/relationships/ui/customui" Target="/customUI/customUI.xml"/>'
            $relsContent = $relsContent -replace '(\s*</Relationships>)', "$customUIRel`n`$1"
            Set-Content $relsFile $relsContent -Encoding UTF8
        }
    }
    
    # Recreate zip
    Remove-Item "$tempDir\AnkiTool.zip" -Force
    Compress-Archive -Path "$tempDir\extracted\*" -DestinationPath "$tempDir\AnkiTool.zip" -Force
    
    # Copy back to original location
    Copy-Item "$tempDir\AnkiTool.zip" $excelFile -Force
    
    # Clean up
    Remove-Item $tempDir -Recurse -Force
    
    Write-Host "✅ Custom ribbon setup complete!" -ForegroundColor Green
    Write-Host "📁 Updated: $excelFile" -ForegroundColor Cyan
    Write-Host "🎯 Next: Open AnkiTool.xlsm and you should see the 'Anki Tools' tab in the ribbon" -ForegroundColor Yellow
    
} catch {
    Write-Host "❌ Error setting up ribbon: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force
    }
    exit 1
} 