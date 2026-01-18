$url = "https://github.com/upx/upx/releases/download/v4.2.1/upx-4.2.1-win64.zip"
$zip = "upx.zip"
$extractPath = "upx_temp"

try {
    Write-Host "Downloading UPX from $url..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
    
    if (!(Test-Path $zip)) {
        Write-Error "Download failed."
        exit 1
    }

    Write-Host "Extracting..."
    Expand-Archive -Path $zip -DestinationPath $extractPath -Force

    Write-Host "Locating upx.exe..."
    $upxFile = Get-ChildItem -Path $extractPath -Recurse -Filter "upx.exe" | Select-Object -First 1
    
    if ($upxFile) {
        Copy-Item -Path $upxFile.FullName -Destination ".\upx.exe"
        Write-Host "✅ UPX installed successfully to $(Get-Location)\upx.exe"
        Write-Host "Version info:"
        .\upx.exe --version
    } else {
        Write-Error "❌ Could not find upx.exe in the downloaded archive."
        exit 1
    }
} catch {
    Write-Error "An error occurred: $_"
    exit 1
} finally {
    # Cleanup
    if (Test-Path $zip) { Remove-Item $zip }
    if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse }
}
