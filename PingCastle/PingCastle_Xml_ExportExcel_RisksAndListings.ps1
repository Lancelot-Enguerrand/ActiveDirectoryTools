Param([string]$BaseDir,[string]$Domain)
#--- Déclarations Variables
#- Date
$date = Get-Date -Format yyyyMMdd
#- Répertoire de travail
if(!($BaseDir)){$BaseDir = "."}
$WorkDir = "$Basedir\Rapports\$date\"
if(Test-Path $WorkDir){Write-Host "Le répertoire d'aujourd'hui existe déjà : $WorkDir"}else{md $WorkDir}
$XMLPath = "$WorkDir\ad_hc_$domain.xml"
#- Template Risques
$riskFileName = "Risques"
$templateExcel = "$Basedir\$riskFileName" + "_Template.xlsx"
$excelDocumentedRisks = "$Basedir\$riskFileName" + ".xlsx"
#- Filtres Listing
$rulesDirectory = "$Basedir\ExportRules"
$rulesFilesList = Get-ChildItem $rulesDirectory

#--- Pingcastle
#- Paramètres - Emplacement et version
$pcVersion = "3.1.0.1"
$PingCastleDir = "$Basedir\Pingcastle_$pcVersion"
$PingCastleExecutable = "$PingCastleDir\PingCastle.exe"
$ArgumentList = "--healthcheck","--server $domain","--level Full","--no-enum-limit"#,"--skip-dc-rpc","--datefile"
#- Lancement de PingCastle
if(Test-Path $XMLPath){Write-Host "Rapport déjà existant : $XMLPath"}else{
    Start-Process -FilePath $PingCastleExecutable -ArgumentList $ArgumentList -WorkingDirectory $WorkDir -Wait
}
#- Vérification du fichier de sortie Pingcastle
if(!(Test-Path $XMLPath)){Write-Host "Rapport non trouvé : $XMLPath" ;Exit}
$PingCastleXML = [XML](Get-Content $XMLPath)

#--- Paramètres Excel
$OutputExcelRiskFile = "$WorkDir\$riskFileName" + "_$date.xlsx"
$OutputExcelImportantFile = "$WorkDir\Important_$date.xlsx"
$ongletRisque = 'Rapport'
$ongletImportant = 'Important'
#$NomTableauRisque = 'Risques'
$NomTableauImportant = 'Important'
$DefaultTableStyle = "Medium2"
#$styleTableauRisque = 'Medium21'
$styleTableauImportant = 'Medium22'

#--- Export Excel
if(!(Test-Path $OutputExcelRiskFile)){Copy-Item $templateExcel $OutputExcelRiskFile}
$PingCastleXML = [XML](Get-Content $XMLPath)
$RiskList = $PingCastleXML | select -ExpandProperty HealthCheckdata | select -ExpandProperty RiskRules |  select -ExpandProperty HealthcheckRiskRule | Select Points,Category,Model,RiskId,Rationale
#- Ajout Propriétés supplémentaires
$RiskList | Add-Member -MemberType NoteProperty -Name "Responsable" -Value ""
$RiskList | Add-Member -MemberType NoteProperty -Name "Vérification" -Value ""
$RiskList | Add-Member -MemberType NoteProperty -Name "Méthode" -Value ""
$RiskList | Add-Member -MemberType NoteProperty -Name "Résolution" -Value ""
$RiskList | Add-Member -MemberType NoteProperty -Name "Statut" -Value "A faire"
#- Ajout Notes
$rapportInitialExcel = Import-Excel $excelDocumentedRisks -WorksheetName $ongletRisque
Foreach($risk in $rapportInitialExcel){
    if($risk.Vérification){
        Foreach($newrisk in $RiskList){
            if($newrisk.RiskId -like $risk.RiskId){
                $newrisk.Responsable = $risk.Responsable
                $newrisk.Vérification = $risk.Vérification
                if($risk.Statut -like "En attente"){$newrisk.Statut = "En attente"}
            }
        }
    }
}
#- Insertion Résultat dans le Fichier Excel
$RiskList | Export-Excel -Path $OutputExcelRiskFile  `
    -WorksheetName $ongletRisque `
    -AutoSize

#--- Export Infos importantes
$ImportantValues = "KrbtgtLastChangeDate","KrbtgtLastVersion","AdminAccountName","AdminLastLoginDate","SchemaLastChanged","LastADBackup"
$ImportantList = $PingCastleXML | select -ExpandProperty HealthCheckdata | select $ImportantValues
#-Insertion Résultat dans le Fichier Excel
if(!(Test-Path $OutputExcelImportantFile)){
$ImportantList | Export-Excel -Path $OutputExcelImportantFile `
    -WorksheetName $ongletImportant `
    -TableName $NomTableauImportant `
    -TableStyle $styleTableauImportant `
    -AutoSize
}

#--- Export Listes Diverses
foreach($rulesfile in $rulesFilesList){
    $Rules = Get-Content "$rulesDirectory\$rulesFile" | Select-String '^[^#]' | ConvertFrom-Csv
    Foreach($rule in $Rules){
        $Resultat = $PingCastleXML | select -ExpandProperty HealthCheckdata | select -ExpandProperty $rule.Balise1 |  select -ExpandProperty $rule.Balise2
        $ExportFilePath = $WorkDir + $rulesfile.BaseName + "_" + $date + ".xlsx"
        if($rule.Style){$TableStyle = $rule.Style}else{$TableStyle = $DefaultTableStyle}
        if(!(Test-Path $ExportFilePath) -and ($Resultat)){
            $Resultat | Export-Excel -Path "$ExportFilePath" `
                -WorksheetName $rule.Name `
                -TableName $rule.Name `
                -TableStyle $TableStyle `
                -AutoSize
        }
    }
}