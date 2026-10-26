param([switch]$NonInteractif)
$ErrorActionPreference = 'Stop'
$verification = [ordered]@{ date=(Get-Date).ToString('o'); dossier=$PSScriptRoot; controles=@(); statut='a_verifier' }
function Ajouter-Controle([string]$nom,[bool]$ok,[string]$detail) {
    $verification.controles += [ordered]@{controle=$nom;ok=$ok;detail=$detail}
    $signe=if($ok){'OK'}else{'A CORRIGER'}
    Write-Host ("[{0}] {1} : {2}" -f $signe,$nom,$detail)
}
Ajouter-Controle 'Windows 64 bits' ([Environment]::Is64BitOperatingSystem) $env:OS
Ajouter-Controle 'Windows PowerShell 5.1' ($PSVersionTable.PSEdition -eq 'Desktop' -and $PSVersionTable.PSVersion -ge [version]'5.1') $PSVersionTable.PSVersion.ToString()
$typeExcel=[type]::GetTypeFromProgID('Excel.Application')
Ajouter-Controle 'Excel de bureau enregistre' ($null -ne $typeExcel) 'Excel Windows avec VBA est necessaire ; Excel en ligne ne convient pas.'
$bin=Join-Path $env:ProgramFiles 'Microsoft Power BI Desktop\bin'
foreach($nom in @('PBIDesktop.exe','Microsoft.PowerBI.AdomdClient.dll')) {
    $p=Join-Path $bin $nom
    Ajouter-Controle $nom (Test-Path -LiteralPath $p -PathType Leaf) $p
}
foreach($nom in @('Gestion_Logistique.xlsm','Decisionnel_Logistique.xlsm','Projet15_Modele_Auto.pbit','Projet15_Rapport_Auto.pbix','Ouvrir_Actualise.ps1','Configurer_Connexions.ps1','Exports\Referentiels_PowerBI.xlsx')) {
    Ajouter-Controle $nom (Test-Path -LiteralPath (Join-Path $PSScriptRoot $nom) -PathType Leaf) 'Fichier du dossier de demonstration.'
}
$exports=@(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'Exports') -Filter 'Resultats_PowerBI_*.xlsx' -File)
Ajouter-Controle 'Export natif initial' ($exports.Count -ge 1) 'Au moins un fichier de resultats date doit etre present.'
$zone=@()
foreach($nom in @('Gestion_Logistique.xlsm','Decisionnel_Logistique.xlsm','Ouvrir_Actualise.ps1','Configurer_Connexions.ps1')) {
    $p=Join-Path $PSScriptRoot $nom
    if(Test-Path -LiteralPath $p) {
        $marque=Get-Item -LiteralPath $p -Stream Zone.Identifier -ErrorAction SilentlyContinue
        if($marque){$zone+=$nom}
    }
}
Ajouter-Controle 'Fichiers transferes reconnus comme locaux' ($zone.Count -eq 0) $(if($zone.Count){'Consulter le guide : Proprietes > Debloquer, pour les seuls fichiers du projet verifies. '+($zone -join ', ')}else{'Aucune marque Internet detectee sur les fichiers controles.'})
$temoin=Join-Path $PSScriptRoot ('Rapports\test_ecriture_'+[guid]::NewGuid().ToString('N')+'.tmp')
$ecriture=$false
try { [void][IO.Directory]::CreateDirectory((Join-Path $PSScriptRoot 'Rapports')); [IO.File]::WriteAllText($temoin,'test'); Remove-Item -LiteralPath $temoin; $ecriture=$true } catch {}
Ajouter-Controle 'Dossier accessible en ecriture' $ecriture 'Les exports, les traces et le PBIX doivent pouvoir etre enregistres.'
$echecs=@($verification.controles | Where-Object {-not $_.ok})
$verification.statut=if($echecs.Count){'a_corriger'}else{'prerequis_detectes'}
$verification | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'Rapports\Verification_PC.json') -Encoding UTF8
Write-Host ''
Write-Host 'Ce controle detecte les composants. Le test du bouton et des chiffres dans le guide reste necessaire.'
if(-not $NonInteractif){Write-Host 'Lire 00_COMMENCER_ICI.html pour la suite.'}
if($echecs.Count){exit 1}
