# Zet het webadres van de invoegtoepassing in de bestanden.
# Gebruik:  .\set-host.ps1 https://www.vermaat-it.nl/handtekening
param([Parameter(Mandatory = $true)][string]$HostUrl)

$HostUrl = $HostUrl.TrimEnd('/')
$uri = [Uri]$HostUrl
if ($uri.Scheme -ne 'https') { throw 'Het adres moet met https:// beginnen.' }
$origin = "$($uri.Scheme)://$($uri.Authority)"

foreach ($f in 'manifest.xml', 'signature.js', 'launchevent.js') {
    $p = Join-Path $PSScriptRoot $f
    $t = [IO.File]::ReadAllText($p)
    if ($t -notmatch '\{\{HOST\}\}|\{\{ORIGIN\}\}') { Write-Warning "$f is al ingesteld, overgeslagen."; continue }
    $t = $t.Replace('{{HOST}}', $HostUrl).Replace('{{ORIGIN}}', $origin)
    [IO.File]::WriteAllText($p, $t, (New-Object Text.UTF8Encoding $false))
    Write-Host "Bijgewerkt: $f"
}
Write-Host "Klaar. Upload nu alle bestanden naar $HostUrl"
