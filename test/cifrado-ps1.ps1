# Arnes de test/cifrado.test.mjs (C-41, v0.43.0): extrae Material/Cifrar/Descifrar-Certificado de
# docs/exportar-planta.ps1 POR AST (no corre el exportador: no toca Graph ni el repo), cifra el documento de entrada,
# descifra el blob que produjo Node y prueba que un blob alterado se rechaza. Escribe el resultado en UTF-8 sin BOM.
# ASCII puro a proposito: PowerShell 5.1 lee un .ps1 sin BOM como ANSI.
param([Parameter(Mandatory = $true)][string]$Entrada, [Parameter(Mandatory = $true)][string]$Salida, [Parameter(Mandatory = $true)][string]$Exportador)
$ErrorActionPreference = 'Stop'
$errores = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Exportador, [ref]$null, [ref]$errores)
if ($errores -and $errores.Count) { [Console]::Error.WriteLine('el exportador no parsea: ' + $errores[0].Message); exit 3 }
$nombres = @('Material-Certificado', 'Cifrar-Certificado', 'Descifrar-Certificado')
$fns = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $nombres -contains $n.Name }, $true))
if ($fns.Count -ne 3) { [Console]::Error.WriteLine('se esperaban 3 funciones del esquema en el exportador, hay ' + $fns.Count); exit 4 }
foreach ($fn in $fns) { . ([scriptblock]::Create($fn.Extent.Text)) }

$utf8SinBom = New-Object System.Text.UTF8Encoding($false)
$in = [IO.File]::ReadAllText($Entrada, $utf8SinBom) | ConvertFrom-Json
$k = Material-Certificado ([string]$in.folio) ([string]$in.sufijo)
$alterado = 'aceptado'
try { Descifrar-Certificado ([string]$in.blobAlterado) $k | Out-Null } catch { $alterado = $_.Exception.Message }
$out = [ordered]@{
    nombre = $k.nombre
    blob = (Cifrar-Certificado $in.doc $k)
    descifrado = (Descifrar-Certificado ([string]$in.blobNode) $k)
    alterado = $alterado
}
[IO.File]::WriteAllText($Salida, ($out | ConvertTo-Json -Depth 6), $utf8SinBom)
