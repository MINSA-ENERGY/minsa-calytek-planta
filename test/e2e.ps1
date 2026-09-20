# Corre la E2E de los 3 roles en Edge headless. Desde PowerShell (no desde Bash: msedge no esta en el PATH).
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\test\e2e.ps1
# Copia del de minsa-proyectos-app (2026-09-13) con los roles de planta. ASCII puro a proposito (PS 5.1 lee sin BOM como ANSI).
param([string[]]$Roles = @('gerencia', 'trazabilidad', 'validador', 'lectura'))
$app = Split-Path -Parent $PSScriptRoot
$edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edge)) { Write-Host "No esta Edge en $edge"; exit 1 }
$srv = Start-Process -FilePath node -ArgumentList "servidor-local.js", "test/pruebas.html" -WorkingDirectory $app -PassThru -WindowStyle Hidden
Start-Sleep -Seconds 2
if ($srv.HasExited) { Write-Host 'PUERTO 8080 OCUPADO: el servidor murio al arrancar (otra sesion corriendo una E2E o capturas.mjs?). Lo que se midiera ahora seria de OTRA app.'; exit 2 }
$fallas = 0
try {
    foreach ($rol in $Roles) {
        $out = Join-Path $env:TEMP "planta-e2e-$rol.html"
        # Presupuesto de tiempo VIRTUAL de 120 s (antes 30): la corrida crece con cada version y al agotarse Edge vuelca el DOM a
        # medias (sin RESUMEN). Solo cuesta tiempo real si algo espera de verdad.
        & $edge --headless=new --disable-gpu --virtual-time-budget=120000 --dump-dom "http://localhost:8080/?rol=$rol" 2>$null | Out-File -Encoding utf8 $out
        Start-Sleep -Seconds 1
        $s = Get-Content $out -Raw -Encoding UTF8
        Write-Host "=== $rol"
        if ($s -match 'PRUEBAS TERMINADAS: ([^<]+)') { Write-Host ("  " + $Matches[1]); if ($Matches[1] -notmatch ' 0 falla') { $fallas++ } } else { Write-Host "  SIN RESUMEN (largo $($s.Length))"; $fallas++ }
        [regex]::Matches($s, '\[FALLA\][^\n]*') | ForEach-Object { Write-Host ("  " + $_.Value) }
    }
} finally { Stop-Process -Id $srv.Id -Force }
if ($fallas) { exit 1 }
exit 0
