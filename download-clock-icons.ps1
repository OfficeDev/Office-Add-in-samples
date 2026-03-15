# Download Fluent UI Clock Icons for World Clock ribbon button
# Run this script from the repository root to download the clock icons

$imagesPath = "Blazor.Excel.AddIn\wwwroot\Images"

# Fluent UI Icons GitHub raw URLs (Regular style)
$baseUrl = "https://raw.githubusercontent.com/microsoft/fluentui-system-icons/main/assets/Clock/PNG"

# Download clock icons in required sizes
$icons = @(
    @{ Size = 16; Output = "ic_fluent_clock_16_regular.png"; Url = "$baseUrl/ic_fluent_clock_16_regular.png" },
    @{ Size = 32; Output = "ic_fluent_clock_32_regular.png"; Url = "$baseUrl/ic_fluent_clock_32_regular.png" },
    @{ Size = 80; Output = "ic_fluent_clock_80_regular.png"; Url = "$baseUrl/ic_fluent_clock_80_regular.png" }
)

foreach ($icon in $icons) {
    $outputPath = Join-Path $imagesPath $icon.Output
    Write-Host "Downloading $($icon.Output)..."
    
    try {
        Invoke-WebRequest -Uri $icon.Url -OutFile $outputPath -ErrorAction Stop
        Write-Host "  Downloaded successfully to $outputPath" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to download from primary URL. Trying alternate sizes..." -ForegroundColor Yellow
        
        # Fluent UI may not have exact sizes, try closest available
        $alternateSizes = @(16, 20, 24, 28, 32, 48)
        $closestSize = $alternateSizes | Sort-Object { [Math]::Abs($_ - $icon.Size) } | Select-Object -First 1
        $alternateUrl = "$baseUrl/ic_fluent_clock_$($closestSize)_regular.png"
        
        try {
            Invoke-WebRequest -Uri $alternateUrl -OutFile $outputPath -ErrorAction Stop
            Write-Host "  Downloaded size $closestSize as fallback to $outputPath" -ForegroundColor Green
        }
        catch {
            Write-Host "  Could not download icon. Please manually download from https://github.com/microsoft/fluentui-system-icons" -ForegroundColor Red
        }
    }
}

Write-Host "`nDone! If any icons failed, please download them manually from:"
Write-Host "https://github.com/microsoft/fluentui-system-icons/tree/main/assets/Clock/PNG" -ForegroundColor Cyan
