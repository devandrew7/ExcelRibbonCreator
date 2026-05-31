# PowerShell script to scan M365 local icons recursively and generate icons_map.js
$root = "D:\Antigravity\ExcelRibbonCreator\icons\2024-microsoft-365-content-icons"

function Get-DirMap($path) {
    $map = @{}
    $dirs = Get-ChildItem -Path $path -Directory
    foreach ($dir in $dirs) {
        $subdirs = Get-ChildItem -Path $dir.FullName -Directory
        if ($subdirs.Count -gt 0) {
            $map[$dir.Name] = Get-DirMap $dir.FullName
        } else {
            $files = Get-ChildItem -Path $dir.FullName -Filter *.svg
            $fileNames = @()
            foreach ($file in $files) {
                $fileNames += $file.Name
            }
            $map[$dir.Name] = $fileNames
        }
    }
    return $map
}

if (Test-Path $root) {
    Write-Host "스캐너를 가동합니다. 타겟 디렉토리: $root"
    $dirMap = Get-DirMap $root
    $json = ConvertTo-Json $dirMap -Depth 100
    
    $jsContent = "/**`n * Excel Custom Ribbon Creator - Automatically Generated Icons Map`n * Generated At: $(Get-Date)`n */`n`nconst ICONS_MAP = $json;`n`nif (typeof module !== 'undefined' && module.exports) {`n    module.exports = { ICONS_MAP };`n} else {`n    window.ICONS_MAP = ICONS_MAP;`n}"
    
    $outputPath = "D:\Antigravity\ExcelRibbonCreator\icons_map.js"
    Set-Content -Path $outputPath -Value $jsContent -Encoding utf8
    Write-Host "컴파일 완료! 생성 경로: $outputPath"
} else {
    Write-Warning "경로를 찾을 수 없습니다: $root"
}
