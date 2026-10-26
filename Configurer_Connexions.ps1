param([switch]$VerifierSeulement)
$ErrorActionPreference='Stop'
$attendu=Join-Path $PSScriptRoot 'Decisionnel_Logistique.xlsm'
$gestion=Join-Path $PSScriptRoot 'Gestion_Logistique.xlsm'
$exports=Join-Path $PSScriptRoot 'Exports'
$rapport=Join-Path $PSScriptRoot 'Projet15_Rapport_Auto.pbix'
$script=Join-Path $PSScriptRoot 'Ouvrir_Actualise.ps1'
$lien=Join-Path $PSScriptRoot 'Ouvrir_Rapport_Actualise.lnk'
foreach($p in @($attendu,$gestion,$rapport,$script)) {
    if(-not(Test-Path -LiteralPath $p -PathType Leaf)){throw "Fichier manquant : $p"}
}
try{$excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}
catch{throw 'Ouvrir le decisionnel de ce dossier dans Excel, fermer son menu (retour a la feuille de lancement), puis relancer.'}
$livres=@($excel.Workbooks | Where-Object {$_.FullName -eq $attendu})
if($livres.Count -ne 1){throw "Le decisionnel de ce dossier doit etre ouvert dans la session Excel active : $attendu. Fermer les autres copies du meme nom."}
$livre=$livres[0]
if($livre.ReadOnly){throw 'Le decisionnel est en lecture seule. Le fermer puis ouvrir la copie locale modifiable.'}
$ws=$livre.Worksheets.Item('Param_Dec')
$trace=[ordered]@{date=(Get-Date).ToString('o');classeur=$livre.FullName;enregistrement_effectue=$false;
    avant=[ordered]@{D01=$ws.Range('C2').Value2;D02=$ws.Range('C3').Value2;D03=$ws.Range('C4').Value2};
    apres=[ordered]@{D01=$gestion;D02=$exports;D03=$lien}}
if($VerifierSeulement){$trace | ConvertTo-Json -Depth 8; exit 0}
[void][IO.Directory]::CreateDirectory($exports)
[void][IO.Directory]::CreateDirectory((Join-Path $PSScriptRoot 'Rapports'))
$shell=New-Object -ComObject WScript.Shell
$shortcut=$shell.CreateShortcut($lien)
$shortcut.TargetPath=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$shortcut.Arguments='-NoProfile -WindowStyle Hidden -File "'+$script+'"'
$shortcut.WorkingDirectory=$PSScriptRoot
$shortcut.Description='Exporter les calculs Excel, actualiser et enregistrer Power BI'
$shortcut.IconLocation=(Join-Path $env:ProgramFiles 'Microsoft Power BI Desktop\bin\PBIDesktop.exe')+',0'
$shortcut.Save()
$ws.Range('C2').Value2=$gestion
$ws.Range('C3').Value2=$exports
$ws.Range('C4').Value2=$lien
$trace | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $PSScriptRoot ('Rapports\Configuration_Connexions_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.json')) -Encoding UTF8
Write-Host 'Les trois chemins de connexion ont ete adaptes dans la copie ouverte.'
Write-Host 'REVENIR DANS EXCEL ET ENREGISTRER LE DECISIONNEL (Ctrl+S).'
Write-Host 'Aucun code VBA ni aucune donnee metier ne sont modifies par ce script.'
Write-Host 'Le parametre DossierExports du rapport doit aussi etre regle comme explique dans le guide.'
Write-Host "Dossier a copier dans le modele Power BI : $exports"
