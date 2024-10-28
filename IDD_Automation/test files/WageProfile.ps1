

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

$excel = New-Object -ComObject "Excel.Application"
$excel.visible = $False
$excel.displayalerts = $False
$workbook = $excel.workbooks.Add()
$XLSDocName = "\WTKConfigurationTemplate_WP_.xlsx"
$XLSDoc = $ScriptDir + $XLSDocName 
$SourceDocName = "\WSAWageProfile.xml"
$XMLDoc = $ScriptDir + $SourceDocName
$RuleWorkSheet = $workbook.worksheets.add()
$RuleWorkSheet.Name = "AdjustmentRule"


$SourceSelectionEntry = Read-Host -Prompt 'Please select source: Enter (A) for live API Call, (F) for XML File'

$SourceSelection = $SourceSelectionEntry.ToUpper

Write-Host $SourceSelectionEntry

if ($SourceSelectionEntry -eq "A") {


$eTIMEInstance = Read-Host -Prompt 'Please enter eTIME Instance Name'
$DataCenterEntry = Read-Host -Prompt 'Please enter Data Center (DC2 or DC4)'
$UserName = Read-Host -Prompt 'Please enter your MOTIF Account Name'
$SAS70Password = Read-Host -Prompt 'Please enter your MOTIF Account Password'

$XLSDocName = "_WTKConfigurationTemplate_WP_.xlsx"
$XLSDocName = "\" + $eTIMEInstance + $XLSDocName
$XLSDoc = $ScriptDir + $XLSDocName 

$DataCenter = $DataCenterEntry.ToUpper()



   if ($DataCenter -eq "DC2") {
                         $eTIMEHost = "https://adpeet2.adphc.com/" + $eTIMEInstance + "/XmlService"
                              }
   elseif ($DataCenter -eq "DC4")
                              {
                         $eTIMEHost = "https://sdmdc4.esi.adp.com/" + $eTIMEInstance + "/XmlService"
                              }  
   else {
            Do {
                $DataCenterEntry = Read-Host -Prompt 'You have not entered a valid Data Center, please re-enter'
                $DataCenter = $DataCenterEntry.ToUpper()
                }
            While (($DataCenter -ne "DC4") -and ($DataCenter -ne "DC2"))
   
   
        }

<# 
XmlService Request and API Calls
#>

$XMLHeader = "<?xml version='1.0'?><Kronos_WFC version='1.0'><Request Object='System' Action='Logon' Username='" + $UserName + "' Password='" + $SAS70Password + "'/>"

$RuleParams = $XMLHeader +
"<Request Action='RetrieveAllforUpdate'><WSAWageProfile/></Request></Kronos_WFC>"

try {
$Response = Invoke-WebRequest -UseBasicParsing $eTIMEHost -ContentType "text/xml" -Method POST -Body $RuleParams -ErrorAction Stop
$StatusCode = $Response.StatusCode
    }

catch
{
    $StatusCode = $_.Exception.Response.StatusCode.value__
    $ErrorMessage = "Unable to connect to instance.  HTTP Error " + $StatusCode + ". Please validate connection and try again."
    $Finish = Read-Host -Prompt $ErrorMessage
    exit
}


<# 
Parsing XML with Select-Xml and XPath
#>

$XMLPath = "/Kronos_WFC/Response/WSAWageProfile"
$RuleDetail = Select-Xml -Content $Response -XPath $XMLPath | Select-Object -ExpandProperty Node

}
else 
{
[XML]$XMLContent = Get-Content $XMLDoc
$RuleDetail = $XMLContent.root.Kronos_WFC.Response.WSAWageProfile
}

$RuleWorkSheet.Rows.Item(1).font.bold = $true
$RuleWorkSheet.Rows.Item(1).HorizontalAlignment = -4108
$RuleWorkSheet.Rows.Item(2).font.bold = $true

$RuleWorkSheet.Range("A1") = "Adjustment Rule"
$RuleWorkSheet.Range("E1") = "Indicator For"
$RuleWorkSheet.Range("F1") = "Trigger"
$RuleWorkSheet.Range("L1") = "Wage Allocation"
$RuleWorkSheet.Range("P1") = "Bonus Allocation"

$MergedCells = $RuleWorkSheet.Range("A1:D1")
$MergedCells.Select()
$MergedCells.MergeCells = $true

$MergedCells = $RuleWorkSheet.Range("F1:J1")
$MergedCells.Select()
$MergedCells.MergeCells = $true
$MergedCells.Interior.ColorIndex = 10

$MergedCells = $RuleWorkSheet.Range("L1:O1")
$MergedCells.Select()
$MergedCells.MergeCells = $true 
$MergedCells.Interior.ColorIndex = 3

$MergedCells = $RuleWorkSheet.Range("P1:Z1")
$MergedCells.Select()
$MergedCells.MergeCells = $true 
$MergedCells.Interior.ColorIndex = 6


$RuleWorkSheet.Range("A2") = "Name"
$RuleWorkSheet.Range("A2").ColumnWidth=33
$RuleWorkSheet.Range("B2") = "Description"
$RuleWorkSheet.Range("B2").ColumnWidth=33
$RuleWorkSheet.Range("C2") = "EffectiveDate"
$RuleWorkSheet.Range("D2") = "ExpirationDate"
$RuleWorkSheet.Range("E2") = "New Trigger"
$RuleWorkSheet.Range("F2") = "JobOrLocation"
$RuleWorkSheet.Range("F2").ColumnWidth=40
$RuleWorkSheet.Range("G2") = "LCEntries"
$RuleWorkSheet.Range("H2") = "CostCenter"
$RuleWorkSheet.Range("I2") = "PayCodes"
$RuleWorkSheet.Range("I2").ColumnWidth=30
$RuleWorkSheet.Range("J2") = "MatchAnyWhere"
$RuleWorkSheet.Range("K2") = "AllocationType"
$RuleWorkSheet.Range("L2") = "Amount"
$RuleWorkSheet.Range("M2") = "Type"
$RuleWorkSheet.Range("N2") = "useHighestWage"
$RuleWorkSheet.Range("O2") = "overrideIfPrimaryJob"
$RuleWorkSheet.Range("P2") = "bonusRateAmount"
$RuleWorkSheet.Range("Q2") = "bonusRateHourlyRate"
$RuleWorkSheet.Range("R2") = "JobCodeType"
$RuleWorkSheet.Range("S2") = "JobOrLocation"
$RuleWorkSheet.Range("T2") = "Paycode"
$RuleWorkSheet.Range("U2") = "timeAmountMinimumTime"
$RuleWorkSheet.Range("V2") = "timeAmountMaximumTime"
$RuleWorkSheet.Range("W2") = "timeAmountMaximumAmount"
$RuleWorkSheet.Range("X2") = "TimePeriod"
$RuleWorkSheet.Range("Y2") = "weekStart"
$RuleWorkSheet.Range("Z2") = "oncePerDay"

$RowNumber = 3

foreach ($Rule in $RuleDetail) {

$RuleName = $Rule.Name

Write-Host $RuleName

    $RuleWorkSheet.Columns.Item(1).Rows.Item($RowNumber) = $RuleName
    $RuleWorkSheet.Columns.Item(3).Rows.Item($RowNumber) = "1753-01-01"
    $RuleWorkSheet.Columns.Item(3).Rows.Item($RowNumber).NumberFormat = "YYYY-MM-DD"
    $RuleWorkSheet.Columns.Item(4).Rows.Item($RowNumber) = "3000-01-01"
    $RuleWorkSheet.Columns.Item(4).Rows.Item($RowNumber).NumberFormat = "YYYY-MM-DD"

        $PayCodes = $Rule.WgpPayCodes.WSAWageProfilePayCode
        $PayCodeTotal = $PayCodes.Count

        $LLDef = $Rule.WageProfileLLDef.WSAWageProfileLaborLevelLinkage

        foreach ($LL in $LLDef) {
        $LLName = $LL.LaborLevelName
    
        $Triggers = $LL.Adjustments.WSAWageAdjustment
        $TriggerTotal = $Triggers.Count
        $TriggerCounter = 1

        foreach ($Trigger in $Triggers) {
        
        $LabLevelEntry = $Trigger.LaborLevelEntryName

        $LabAccount = $LLName + ": " + $LabLevelEntry
        
        $AdjType = "Wage"

        $TypeCode = $Trigger.Type
            switch ($TypeCode) {"0" {$Type = "FlatRate"}
                               "1" {$Type = "Addition"}
                               "2" {$Type = "Multiplier NOT SUPPORTED"}
                               }
        
        $Amount = $Trigger.Amount
        $MatchAnywhere = "TRUE"
        $UseHighestWage = $Rule.UseHighestWageSwitch
        $OverridePrimaryJob = $Rule.OverrideIfHomeSwitch
        
        
        $RuleWorkSheet.Columns.Item(5).Rows.Item($RowNumber) = "New"
        $RuleWorkSheet.Columns.Item(6).Rows.Item($RowNumber) = $LabAccount
        $RuleWorkSheet.Columns.Item(7).Rows.Item($RowNumber) = ",,,,"
        $RuleWorkSheet.Columns.Item(10).Rows.Item($RowNumber) = $MatchAnywhere <#What is it#>
        $RuleWorkSheet.Columns.Item(11).Rows.Item($RowNumber) = $AdjType
        $RuleWorkSheet.Columns.Item(12).Rows.Item($RowNumber) = $Amount
        $RuleWorkSheet.Columns.Item(13).Rows.Item($RowNumber) = $Type
        $RuleWorkSheet.Columns.Item(14).Rows.Item($RowNumber) = $UseHighestWage
        $RuleWorkSheet.Columns.Item(15).Rows.Item($RowNumber) = $OverridePrimaryJob
        $RuleWorkSheet.Columns.Item(16).Rows.Item($RowNumber) = $BonusRate
        $RuleWorkSheet.Columns.Item(17).Rows.Item($RowNumber) = $BonusHourlyRate
        $RuleWorkSheet.Columns.Item(18).Rows.Item($RowNumber) = $LaborAccountType
        $RuleWorkSheet.Columns.Item(19).Rows.Item($RowNumber) = $JoborLocation  <#What is it#>
        $RuleWorkSheet.Columns.Item(20).Rows.Item($RowNumber) = $BonusPayCode
        $RuleWorkSheet.Columns.Item(21).Rows.Item($RowNumber) = $TimeAmountMinimumTime
        $RuleWorkSheet.Columns.Item(22).Rows.Item($RowNumber) = $TimeAmountMaximumTime
        $RuleWorkSheet.Columns.Item(23).Rows.Item($RowNumber) = $TimeAmountMaximumAmount
        $RuleWorkSheet.Columns.Item(24).Rows.Item($RowNumber) = $TimePeriod
        $RuleWorkSheet.Columns.Item(25).Rows.Item($RowNumber) = $WeekStart
        $RuleWorkSheet.Columns.Item(26).Rows.Item($RowNumber) = $OncePerDay

            if ($PayCodes -ne "") {

            $PayCodeCounter = 1

            foreach ($PayCode in $PayCodes) {
            $PayCode = $PayCode.Name
            $RuleWorkSheet.Columns.Item(9).Rows.Item($RowNumber) = $PayCode
            $PayCodeCounter++
            if ($PayCodeCounter -le $PayCodeTotal){$RowNumber++}

            }
            }
            
            if ($TriggerCounter -le $TriggerTotal){$RowNumber++}
            $TriggerCounter++

        }
        }
$RowNumber++


}


$workbook.SaveAs($XLSDoc)
$workbook.close($false)
$excel.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
