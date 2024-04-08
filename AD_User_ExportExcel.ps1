Param([string]$User,[string]$RulesFile,[string]$OutputPath,[string]$FileName,[bool]$OpenFile)
#--- Modules Verification ---#
if(!(Get-Module -ListAvailable -Name ActiveDirectory) -or !(Get-Module -ListAvailable -Name ImportExcel)){
    if(!(Get-Module -ListAvailable -Name ActiveDirectory)){
        Write-Host "Module Manquant : ActiveDirectory"
        Write-Host "Installation : Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online"}
    if(!(Get-Module -ListAvailable -Name ImportExcel)){
        Write-Host "Module Manquant : ImportExcel"
        Write-Host "Installation : https://github.com/dfinke/ImportExcel"}
    Exit}
#--- File Opening after Extract ---#
if(!($OpenFile)){$OpenFile = $False}
#--- Rules File ---#
if(!($RulesFile)){$RulesFile = Read-Host "Indiquer le fichier de règles "}
$rulesFileDir = ".\"
$rulesFileName = "ADUserExportExcelParameter.csv"
$rulesFilePath = "$rulesFileDir\$rulesFileName"
if(!($RulesFile)){$RulesFile = $rulesFilePath}
if(!(Test-Path $rulesFile)){$RulesFile = Read-Host "Indiquer le fichier de règles "}
if(!(Test-Path $rulesFile)){Write-Host "Rules files not found" ;break}
$rulesFileContent = Get-Content $RulesFile | Select-String '^[^#]' | ConvertFrom-Csv -Delimiter ";" | Sort Priority
#--- Ask For Username ---#
if(!($User)){$User = Read-Host "Entrer un nom d'utilisateur "}
#--- AD Request ---#
try{$GroupList = Get-ADUser $user -Properties Memberof | select -ExpandProperty Memberof | Foreach{Get-ADGroup $_ -Properties Description,Memberof}}catch{Write-Host "Utilisateur introuvable"; Exit}
$UserDN = Get-ADUser $User | select -ExpandProperty DistinguishedName
$recursiveGroupList = Get-ADGroup -LDAPFilter ("(member:1.2.840.113556.1.4.1941:={0})" -f $UserDN) -Properties Name,Description,DistinguishedName,GroupCategory | Select Name,Description,DistinguishedName,GroupCategory
$insertGroupList = $recursiveGroupList

#--- Output File ---#
$date = Get-Date -Format yyyyMMdd
$extensionFile = ".xlsx"
if(!($FileName)){$filename = "$userADM" + "_" + "$date"}
if(($filename.Substring($filename.Length -$extensionFile.Length, $extensionFile.Length)) -like $extensionFile){$filename = $filename.Substring($FileName.Length - $extensionFile.Length)}
if(!($OutputPath)){$OutputPath = ".\"}
$outputExcelFile = $OutputPath + $filename + $extensionFile
For(($i = 1);Test-Path $outputExcelFile;$i++){
    Write-Host("Le fichier existe déjà : $outputExcelFile")
    $outputExcelFile = $OutputPath + $filename + '_' + $i.ToString() + $extensionFile
}
#--- Paramètres Excel ---#
#--- Row Column Parameters ---#
$currentRow = 1
$currentColumn = 1
$stepRow = 1
$missingColumn = 0
#-- Personnalisation Tableau --#
$nomFeuille = "Bilan"
$nomTableauOther = "Autres"
$styleTableauOther = "Medium4"
$nomTableauBuiltin = "Builtin"
$styleTableauBuiltin = "Medium4"

#--- Création Tableau de Liste ---#
$listLists = @{}
foreach($rule in $rulesFileContent){$listName = $rule.Name;$listLists[$listName]=@($Rule.Name)}

#--- Clean CSV Content ---#
Foreach($rule in $rulesFileContent){while($rule.Priority.Length -lt 3){$rule.Priority = '0' + $rule.Priority};if(!($rule.Style)){$rule.Style = "Medium2"}}
#--- Filter and Insert to Lists ---#
$sortedpriorityrules = $rulesFileContent| Sort Priority
Foreach($group in $insertGroupList)
{
    $inserted = $false
    Foreach($rule in $sortedpriorityrules){
        if(($inserted -like $false) -and ($group.($rule.Property) -like $rule.Filter)){
            if(($group.($rule.Display)) -and ($rule.Display -notlike "Name")){
            $listLists[$rule.Name]+=($group.($rule.Display)).substring($rule.Descriptioncrop)}
            else{$listLists[$rule.Name]+=($group.Name).Substring($rule.NameCrop)}
            $inserted = $True
        }
    }
}

#--- Export Excel ---#
$sorteddisplayrules = $rulesFileContent | Sort Column,Row
Foreach($rule in $sorteddisplayrules)
{
    if(($rule.Column -gt 0) -and ($listLists[$rule.Name].count -gt 1)){
        if(($currentColumn + $missingColumn) -lt $rule.Column){$missingColumn += ($rule.Column - ($currentColumn + 1 + $missingColumn));$currentRow = 1}
        $CurrentColumn = $rule.Column - $missingColumn
        $listLists[$rule.Name] | Export-Excel -Path "$outputExcelFile" `
            -WorksheetName $nomFeuille `
            -TableStyle $rule.Style `
            -TableName $rule.Name `
            -StartRow $currentRow `
            -StartColumn $CurrentColumn `
            -AutoSize
        $currentRow += $listLists[$rule.Name].Count + $stepRow
    }
}
#--- Liste Brute
if($GroupList){$GroupList | Select Name,Description,DistinguishedName | sort Name | Export-Excel -Path "$outputExcelFile" `
    -WorksheetName "MemberOf" `
    -TableStyle $styleTableauOther `
    -TableName "MemberOfBase" `
    -AutoSize}
#--- Liste Complete
if($recursiveGroupList){$recursiveGroupList | Select Name,Description,DistinguishedName | sort Name | Export-Excel -Path "$outputExcelFile" `
    -WorksheetName "ALL" `
    -TableStyle $styleTableauOther `
    -TableName "MemberOfRecursif" `
    -AutoSize}

if($OpenFile){& "$outputExcelFile"}