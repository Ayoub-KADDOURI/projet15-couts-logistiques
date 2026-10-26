param([switch]$PreparerSourcesSeulement, [switch]$SansMessageErreur)
$ErrorActionPreference = 'Stop'
$racineAuto = $PSScriptRoot
$dossierAuto = Join-Path $racineAuto 'Exports'
$rapportAuto = Join-Path $racineAuto 'Projet15_Rapport_Auto.pbix'
$journalAuto = Join-Path $racineAuto 'Rapports\Derniere_Actualisation_Auto.json'
$journalEtat = [ordered]@{ debut=(Get-Date).ToString('o'); statut='demarrage'; etape=''; export=$null; erreur=$null }
$mutexAuto = $null; $verrouAuto = $false; $excelAuto = $null; $statusAvant = $null
function Etape-Auto([string]$nom) {
    $journalEtat.etape = $nom
    $journalEtat | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $journalAuto -Encoding UTF8
    Write-Host $nom
}
function Escape-Xml([object]$v) { [Security.SecurityElement]::Escape([string]$v) }
function Colonne-Excel([int]$n) {
    $s=''
    while ($n -gt 0) { $n--; $s=([char](65+($n%26))).ToString()+$s; $n=[math]::Floor($n/26) }
    return $s
}
function Ecrire-EntreeZip($zip,[string]$nom,[string]$contenu) {
    $entry=$zip.CreateEntry($nom)
    $stream=$entry.Open(); $writer=New-Object IO.StreamWriter($stream,([Text.UTF8Encoding]::new($false)))
    try { $writer.Write($contenu) } finally { $writer.Dispose(); $stream.Dispose() }
}
function Exporter-References($livre,[string]$destination) {
    Add-Type -AssemblyName System.IO.Compression
    $temporaire=Join-Path $dossierAuto ('References_'+[guid]::NewGuid().ToString('N')+'.tmp')
    $flux=[IO.File]::Create($temporaire)
    $zip=New-Object IO.Compression.ZipArchive($flux,[IO.Compression.ZipArchiveMode]::Create)
    try {
        $xmlBase='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        Ecrire-EntreeZip $zip '[Content_Types].xml' ($xmlBase+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        Ecrire-EntreeZip $zip '_rels/.rels' ($xmlBase+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        Ecrire-EntreeZip $zip 'xl/workbook.xml' ($xmlBase+'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/><sheet name="Parametres" sheetId="2" r:id="rId2"/></sheets></workbook>')
        Ecrire-EntreeZip $zip 'xl/_rels/workbook.xml.rels' ($xmlBase+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        Ecrire-EntreeZip $zip 'xl/styles.xml' ($xmlBase+'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
        $indexFeuille=0
        foreach ($nomFeuille in @('Budget','Parametres')) {
            $indexFeuille++
            $valeurs=$livre.Worksheets.Item($nomFeuille).UsedRange.Value2
            $xml=New-Object Text.StringBuilder
            [void]$xml.Append($xmlBase+'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
            for ($ligne=1;$ligne -le $valeurs.GetLength(0);$ligne++) {
                [void]$xml.Append('<row r="'+$ligne+'">')
                for ($colonne=1;$colonne -le $valeurs.GetLength(1);$colonne++) {
                    $v=$valeurs.GetValue($ligne,$colonne)
                    if ($null -eq $v) { continue }
                    $adresse=(Colonne-Excel $colonne)+$ligne
                    $entete=[string]$valeurs.GetValue(1,$colonne)
                    if ($v -is [string]) { [void]$xml.Append('<c r="'+$adresse+'" t="inlineStr"><is><t xml:space="preserve">'+(Escape-Xml $v)+'</t></is></c>') }
                    elseif ($v -is [bool]) { [void]$xml.Append('<c r="'+$adresse+'" t="b"><v>'+[int]$v+'</v></c>') }
                    else {
                        $style=''
                        if ($ligne -gt 1 -and ($entete -eq 'Mois' -or $entete.StartsWith('Date'))) { $style=' s="1"'; if ($livre.Date1904) { $v=[double]$v+1462 } }
                        $nombre=([double]$v).ToString('R',[Globalization.CultureInfo]::InvariantCulture)
                        [void]$xml.Append('<c r="'+$adresse+'"'+$style+'><v>'+$nombre+'</v></c>')
                    }
                }
                [void]$xml.Append('</row>')
            }
            [void]$xml.Append('</sheetData></worksheet>')
            Ecrire-EntreeZip $zip ('xl/worksheets/sheet'+$indexFeuille+'.xml') $xml.ToString()
        }
    } finally { $zip.Dispose(); $flux.Dispose() }
    Move-Item -LiteralPath $temporaire -Destination $destination -Force
}
function Verifier-ControlesExcel([Array]$tableau) {
    if ($null -eq $tableau -or $tableau.Rank -ne 2 -or $tableau.GetLength(1) -lt 6) {
        throw 'Les controles Excel sont absents ou incomplets. Relancer le calcul dans le decisionnel.'
    }
    $releves=@()
    for ($ligne=2;$ligne -le $tableau.GetLength(0);$ligne++) {
        $poste=[string]$tableau.GetValue($ligne,1)
        if (-not $poste) { continue }
        $statut=[string]$tableau.GetValue($ligne,6)
        $ecart=[double]$tableau.GetValue($ligne,4)
        if ($statut -ne 'OK') {
            throw ("Le controle '{0}' est marque '{1}' dans Excel (ecart {2:N2} MAD). Consulter les controles et relancer le calcul dans Excel." -f $poste,$statut,$ecart)
        }
        $releves += [pscustomobject][ordered]@{
            Poste=$poste; Grand_livre=[double]$tableau.GetValue($ligne,2)
            Calcule=[double]$tableau.GetValue($ligne,3); Ecart_MAD=$ecart
            Ecart_pct=[double]$tableau.GetValue($ligne,5); Etat=$statut
        }
    }
    if ($releves.Count -eq 0) { throw 'Aucun controle Excel disponible. Relancer le calcul dans le decisionnel.' }
    return $releves
}
function Lire-Etat($livre) {
    $data=$livre.Worksheets.Item('Etat').UsedRange.Value2; $result=@{}
    for ($i=2;$i -le $data.GetLength(0);$i++) { $result[[string]$data.GetValue($i,1)]=$data.GetValue($i,3) }
    return $result
}
function Trouver-Controle($fenetre,[string]$id) {
    $condition=New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty,$id)),
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::IsOffscreenProperty,$false)))
    $fenetre.FindFirst([System.Windows.Automation.TreeScope]::Descendants,$condition)
}
function Cliquer-Controle($element) {
    if ($null -eq $element) { throw 'Le bouton Power BI attendu est introuvable.' }
    $pattern=$element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    $pattern.Invoke()
}
function Connecter-Modele([int]$pidPbi) {
    $serveur=Get-CimInstance Win32_Process -Filter "name='msmdsrv.exe'" | Where-Object ParentProcessId -eq $pidPbi
    if (-not $serveur -or $serveur.CommandLine -notmatch '-s\s+"([^"]+)"') { return $null }
    $portFichier=Join-Path $Matches[1] 'msmdsrv.port.txt'
    if (-not (Test-Path -LiteralPath $portFichier)) { return $null }
    $portLu=[int](Get-Content -LiteralPath $portFichier -Encoding Unicode -Raw).Trim([char]0).Trim()
    $connexion=New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection("Data Source=localhost:$portLu")
    $connexion.Open(); return $connexion
}
function Lire-Modele($connexion) {
    $cmd=$connexion.CreateCommand(); $cmd.CommandTimeout=8
    $cmd.CommandText='EVALUATE ROW("Fichier",MAX(Perimetre[Fichier]),"Actualise",MAX(Perimetre[Actualise]),"Global",[Cout global],"CA",[CA HT],"Commandes",[Nb commandes],"Budget",[Budget total])'
    $reader=$cmd.ExecuteReader()
    try {
        if (-not $reader.Read()) { throw 'Le modèle ne renvoie aucune donnée.' }
        $result=@{}
        for ($i=0;$i -lt $reader.FieldCount;$i++) { $result[$reader.GetName($i).Trim('[',']')]=$reader.GetValue($i) }
        return $result
    } finally { $reader.Dispose() }
}
try {
    [void][IO.Directory]::CreateDirectory($dossierAuto)
    [void][IO.Directory]::CreateDirectory((Join-Path $racineAuto 'Rapports'))
    $hashMutex=[BitConverter]::ToString(([Security.Cryptography.SHA256]::Create()).ComputeHash([Text.Encoding]::UTF8.GetBytes($racineAuto))).Replace('-','').Substring(0,20)
    $mutexAuto=New-Object Threading.Mutex($false,('Local\Projet15PBI_'+$hashMutex))
    $verrouAuto=$mutexAuto.WaitOne(0)
    if (-not $verrouAuto) { exit 0 }
    Etape-Auto 'Lecture du décisionnel ouvert'
    $excelAuto=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    $statusAvant=$excelAuto.StatusBar
    $cheminDecisionnelAuto=Join-Path $racineAuto 'Decisionnel_Logistique.xlsm'
    $livres=@($excelAuto.Workbooks | Where-Object { $_.FullName -eq $cheminDecisionnelAuto })
    if ($livres.Count -ne 1) { throw 'Ouvrir le Decisionnel_Logistique.xlsm de ce dossier dans la session Excel active, puis relancer. Fermer les autres copies du meme nom.' }
    $livreAuto=$livres[0]; $etatAuto=Lire-Etat $livreAuto
    $excelAuto.StatusBar='Préparation du rapport Power BI actualisé...'
    $periodeData=$livreAuto.Worksheets.Item('Periodes').UsedRange.Value2; $periodesMap=@{}
    for ($i=2;$i -le $periodeData.GetLength(0);$i++) { $periodesMap[[string]$periodeData.GetValue($i,1)]=@($periodeData.GetValue($i,3),$periodeData.GetValue($i,4),$periodeData.GetValue($i,5)) }
    $importsData=$livreAuto.Worksheets.Item('Import_Periodes').UsedRange.Value2; $debuts=@(); $fins=@()
    for ($i=2;$i -le $importsData.GetLength(0);$i++) {
        $codePeriode=[string]$importsData.GetValue($i,1)
        if (-not $codePeriode) { continue }
        if (-not $periodesMap.ContainsKey($codePeriode) -or $periodesMap[$codePeriode][2] -ne 'Clôturée') { throw "La période $codePeriode n'est pas clôturée. La valider dans l'application avant la restitution." }
        $debuts+=[double]$periodesMap[$codePeriode][0]; $fins+=[double]$periodesMap[$codePeriode][1]
    }
    if ($debuts.Count -eq 0) { throw 'Aucune période importée. Actualiser les données dans le décisionnel.' }
    if (-not $etatAuto['E-03'] -or [double]$etatAuto['E-02'] -gt [double]$etatAuto['E-03']) {
        Etape-Auto 'Recalcul natif des données importées'
        $chrono=[Diagnostics.Stopwatch]::StartNew()
        $debutCalcul=[DateTime]::FromOADate(($debuts | Measure-Object -Minimum).Minimum)
        $finCalcul=[DateTime]::FromOADate(($fins | Measure-Object -Maximum).Maximum)
        $calculOk=$excelAuto.Run(("'"+$livreAuto.Name+"'!modCalcul.CalculerTout"),$debutCalcul,$finCalcul)
        if (-not $calculOk) { throw 'Le calcul natif a échoué. Consulter les contrôles dans Excel.' }
        $chrono.Stop()
        $null=$excelAuto.Run(("'"+$livreAuto.Name+"'!modEtat.EcrireEtat"),'E-03',(Get-Date))
        $null=$excelAuto.Run(("'"+$livreAuto.Name+"'!modEtat.EcrireEtat"),'E-04',([math]::Round($chrono.Elapsed.TotalSeconds,1)))
        $journalEtat.recalcul_natif=$true; $etatAuto=Lire-Etat $livreAuto
    } else { $journalEtat.recalcul_natif=$false }
    $controles=$livreAuto.Worksheets.Item('Calc_Controles').UsedRange.Value2
    $journalEtat.controles_excel=@(Verifier-ControlesExcel $controles)
    $journalEtat.regle_controles='Validation native Excel ; ecarts conserves et traces. Les alertes de rapprochement restent bloquantes.'
    $alertes=$livreAuto.Worksheets.Item('Calc_Alertes').UsedRange.Value2
    for ($i=2;$i -le $alertes.GetLength(0);$i++) {
        if ($alertes.GetValue($i,3) -eq 'Rapprochement' -and $alertes.GetValue($i,9) -eq 'À examiner') { throw 'Une alerte de rapprochement bloque la restitution. La traiter dans Excel.' }
    }
    Etape-Auto 'Export natif et copie des références en mémoire'
    $fichierMinute=Join-Path $dossierAuto ('Resultats_PowerBI_'+(Get-Date -Format 'yyyy-MM-dd_HHmm')+'.xlsx')
    if (Test-Path -LiteralPath $fichierMinute) {
        try { $verifFlux=[IO.File]::Open($fichierMinute,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None); $verifFlux.Dispose() }
        catch { throw 'Fermer dans Excel le fichier de résultats de cette minute avant de relancer le rapport.' }
    }
    $livreAuto.Worksheets.Item('Param_Dec').Range('C3').Value2=$dossierAuto
    $exportOk=$excelAuto.Run(("'"+$livreAuto.Name+"'!modCalcul.ExporterPowerBI"),'')
    if (-not $exportOk) { throw 'La macro native a refusé l''export.' }
    $etatAuto=Lire-Etat $livreAuto; $exportAuto=[string]$etatAuto['E-12']
    if (-not (Test-Path -LiteralPath $exportAuto -PathType Leaf)) { throw 'Le fichier exporté est introuvable.' }
    Exporter-References $livreAuto (Join-Path $dossierAuto 'Referentiels_PowerBI.xlsx')
    $journalEtat.export=[IO.Path]::GetFileName($exportAuto)
    $attendu=@{Global=[double]$etatAuto['E-05'];CA=[double]$etatAuto['E-06'];Commandes=[double]$etatAuto['E-07'];Budget=[double]$etatAuto['E-08']}
    $journalEtat.valeurs_excel=$attendu
    if ($PreparerSourcesSeulement) { $journalEtat.statut='sources_pretes'; Etape-Auto 'Sources préparées'; exit 0 }
    if (-not (Test-Path -LiteralPath $rapportAuto -PathType Leaf)) { throw 'Le rapport Projet15_Rapport_Auto.pbix est introuvable.' }
    $binPbi=Join-Path $env:ProgramFiles 'Microsoft Power BI Desktop\bin'
    $exePbi=Join-Path $binPbi 'PBIDesktop.exe'
    if (-not (Test-Path -LiteralPath $exePbi)) { throw 'Power BI Desktop est introuvable dans Program Files.' }
    [void][Reflection.Assembly]::LoadFrom((Join-Path $binPbi 'Microsoft.PowerBI.AdomdClient.dll'))
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class Projet15Fenetre { [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h,int c); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); }'
    $processusRapport=@(Get-CimInstance Win32_Process -Filter "name='PBIDesktop.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($rapportAuto,[StringComparison]::OrdinalIgnoreCase) -ge 0 })
    if ($processusRapport.Count -gt 1) { throw 'Ce rapport est ouvert plusieurs fois. Fermer ses fenêtres en double puis relancer le bouton.' }
    if ($processusRapport.Count -eq 1) { $pidRapportAuto=[int]$processusRapport[0].ProcessId; $journalEtat.fenetre_reutilisee=$true }
    else {
        $pbiAuto=Start-Process -FilePath $exePbi -ArgumentList ('"'+$rapportAuto+'"') -PassThru
        $pidRapportAuto=$pbiAuto.Id; $journalEtat.fenetre_reutilisee=$false
    }
    $journalEtat.pid_powerbi=$pidRapportAuto
    Etape-Auto 'Ouverture de Power BI'
    $limiteOuverture=(Get-Date).AddSeconds(150); $fenetrePbi=$null; $connexionAuto=$null
    do {
        Start-Sleep -Milliseconds 500
        try {
            $pbiActif=Get-Process -Id $pidRapportAuto
            if ($pbiActif.MainWindowHandle -eq 0) { continue }
            $fenetrePbi=[System.Windows.Automation.AutomationElement]::FromHandle($pbiActif.MainWindowHandle)
            $connexionAuto=Connecter-Modele $pidRapportAuto
            if ($null -eq $connexionAuto) { continue }
            $null=Lire-Modele $connexionAuto
            break
        } catch { if ($null -ne $connexionAuto) { $connexionAuto.Dispose(); $connexionAuto=$null } }
    } while ((Get-Date) -lt $limiteOuverture)
    if ($null -eq $connexionAuto) { throw 'Power BI n''a pas terminé son ouverture dans le délai prévu.' }
    try {
        $limiteInterface=(Get-Date).AddSeconds(50); $boutonActualiser=$null
        do {
            try {
                $pbiActif=Get-Process -Id $pidRapportAuto
                [void][Projet15Fenetre]::ShowWindow($pbiActif.MainWindowHandle,9)
                $fenetrePbi=[System.Windows.Automation.AutomationElement]::FromHandle($pbiActif.MainWindowHandle)
                $ongletAccueil=Trouver-Controle $fenetrePbi 'home'
                if ($null -ne $ongletAccueil) { $ongletAccueil.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() }
                $boutonActualiser=Trouver-Controle $fenetrePbi 'refreshQueries'
            } catch { $boutonActualiser=$null }
            if ($null -eq $boutonActualiser) { Start-Sleep -Milliseconds 500 }
        } while ($null -eq $boutonActualiser -and (Get-Date) -lt $limiteInterface)
        if ($null -eq $boutonActualiser) { throw 'Power BI est ouvert mais son interface ne permet pas encore de lancer l''actualisation.' }
        [void][Projet15Fenetre]::SetForegroundWindow($pbiActif.MainWindowHandle)
        $debutActualisation=Get-Date
        Etape-Auto 'Actualisation des données dans Power BI'
        Cliquer-Controle $boutonActualiser
        $limiteActualisation=(Get-Date).AddSeconds(180); $valide=$false; $valeursModele=$null
        do {
            Start-Sleep -Milliseconds 600
            try {
                $valeursModele=Lire-Modele $connexionAuto
                $valide=([string]$valeursModele.Fichier -eq $journalEtat.export -and [DateTime]$valeursModele.Actualise -ge $debutActualisation.AddSeconds(-1))
                foreach ($k in @('Global','CA','Commandes','Budget')) { if ([math]::Abs([double]$valeursModele[$k]-$attendu[$k]) -gt 0.01) { $valide=$false } }
            } catch { $valide=$false }
        } while (-not $valide -and (Get-Date) -lt $limiteActualisation)
        if (-not $valide) { throw 'L''actualisation Power BI n''a pas été confirmée. Vérifier les erreurs de chargement du rapport.' }
        $journalEtat.valeurs_powerbi=$valeursModele
        Etape-Auto 'Sauvegarde du rapport actualisé'
        $limiteBoutonSauvegarde=(Get-Date).AddSeconds(45); $boutonSauvegarde=$null
        do {
            try {
                $pbiActif=Get-Process -Id $pidRapportAuto
                $fenetrePbi=[System.Windows.Automation.AutomationElement]::FromHandle($pbiActif.MainWindowHandle)
                $boutonSauvegarde=Trouver-Controle $fenetrePbi 'save'
                if ($null -ne $boutonSauvegarde -and -not $boutonSauvegarde.Current.IsEnabled) { $boutonSauvegarde=$null }
            } catch { $boutonSauvegarde=$null }
            if ($null -eq $boutonSauvegarde) { Start-Sleep -Milliseconds 500 }
        } while ($null -eq $boutonSauvegarde -and (Get-Date) -lt $limiteBoutonSauvegarde)
        if ($null -eq $boutonSauvegarde) { throw 'Power BI ne permet pas de sauvegarder le rapport. Vérifier ses éventuelles erreurs de chargement.' }
        $dateAvant=(Get-Item -LiteralPath $rapportAuto).LastWriteTimeUtc
        Cliquer-Controle $boutonSauvegarde
        $limiteSauvegarde=(Get-Date).AddSeconds(35)
        do { Start-Sleep -Milliseconds 500; $fichierPbi=Get-Item -LiteralPath $rapportAuto } while ($fichierPbi.LastWriteTimeUtc -le $dateAvant -and (Get-Date) -lt $limiteSauvegarde)
        if ($fichierPbi.LastWriteTimeUtc -le $dateAvant) { throw 'Les données sont actualisées mais la sauvegarde du PBIX n''a pas été confirmée.' }
        Start-Sleep -Milliseconds 800
        $journalEtat.statut='succes'; $journalEtat.fin=(Get-Date).ToString('o')
        Etape-Auto 'Rapport actualisé et enregistré'
    } finally { $connexionAuto.Dispose() }
} catch {
    $journalEtat.statut='erreur'; $journalEtat.erreur=$_.Exception.Message; $journalEtat.fin=(Get-Date).ToString('o')
    $journalEtat | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $journalAuto -Encoding UTF8
    if (-not $SansMessageErreur) { Add-Type -AssemblyName System.Windows.Forms; [void][Windows.Forms.MessageBox]::Show($journalEtat.erreur,'Actualisation Power BI',[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Error) }
    Write-Error $journalEtat.erreur
    exit 1
} finally {
    if ($null -ne $excelAuto) { try { $excelAuto.StatusBar=$statusAvant } catch {} }
    if ($verrouAuto) { $mutexAuto.ReleaseMutex() }
    if ($null -ne $mutexAuto) { $mutexAuto.Dispose() }
}
