#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
#%%%%%%%%%%%%%%%%%%%%%%%%%  Global Variables  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

$crlf = "`r`n"
$workingDir = Convert-Path .
$dbPath = $workingDir + "\db\itag.accdb"
$tempCsv = $workingDir + "\tmp\tmp.csv"
$jsonDir = $workingDir + "\exports\json\"
$csvDir = $workingDir + "\exports\csv\"

$adOpenStatic = 3
$adLockOptimistic = 3

$connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
$connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= $dbPath"
$command = $connection.CreateCommand()

#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
#%%%%%%%%%%%%%%%%%%%%%%%%%%%%  Functions  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function retConvertDate([string]$dateVal) {
    $newDate = $dateVal.split(" ")
    return $newDate[0]
}

function retNormTaxonomy([string]$oldStr) {
    $newStr = $oldStr -replace '\d\d\d\;\#', ''
    $newStr = $newStr -replace '\d\d\;\#', ''
    $newStr = $newStr -replace '\d\;\#', ''
    $newStr = $newStr -replace '\#', ''
    $newStr = $newStr -replace '\?', '&'
    $newStr = $newStr -replace "\:", " :: "
    $newStr = $newStr -replace '[^\x00-\x7f]', '&'
    return $newStr
}

function retHighLevelCapability([string]$oldStr) {
    $newStr = $oldStr -replace " \:\: ", "%"
    $newStrArray = $newStr.split("%")
    if ($newStrArray.Length -gt 2) {
        $category = $newStrArray[0]
        $class = $newStrArray[1]
        $newStr = $category + " :: " + $class
    } Else {
        $newStr = $oldStr
    }
    return $newStr
}

function importData {
    write-host "%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%  Refreshing Access Databases from ITAG Team Site Content  %%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%"

    $accessApp = New-Object -com Access.Application
    $accessApp.OpenCurrentDatabase($dbPath)
    $accessApp.visible = $True
    $accessApp.DoCmd.RunMacro("refreshItagData")
    $accessApp.CloseCurrentDatabase()
    $accessApp.quit()
}

function exportDecisions {
    write-host "%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%  Exporting ITAG Decisions Information to CSV and JSON  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%"

    $csv = $csvDir + "decisions.csv"
    $json = $jsonDir +  "decisions.json"
    $qryString = "select * from DecisionRegistry"

    $command.CommandText = $qryString
    $adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
    $dataset = New-Object -TypeName System.Data.DataSet
    $adapter.Fill($dataset)
    $dataset.Tables[0] | export-csv $tempCsv -NoTypeInformation

    Import-Csv $tempCsv | select @{Name = "nSubmissionID"; Expression = {$_."Submission ID"}},
        @{Name = "sItagID"; Expression = {$_."ITAG ID"}},
        @{Name = "dDateReviewed"; Expression = {retConvertDate($_."Date Reviewed")}},
        @{Name = "sItagGovernanceType"; Expression = {$_."ITAG Governance Type"}},
        @{Name = "sDecision"; Expression = {$_."Decision"}},
        @{Name = "bConstraints"; Expression = {$_."Any Applied Conditions or Constraints?"}},
        @{Name = "bActionItems"; Expression = {$_."Any Assigned Action Items?"}} | Export-csv $csv -NoTypeInformation

    import-csv $csv | ConvertTo-Json | New-Item -path $jsonDir -Name "decisions.json" -ItemType file -force
}

function exportDevonContacts {
    write-host "%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%  Exporting ITAG Sharepoint Site Contact Information to CSV and JSON  %%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%"

    $csv = $csvDir + "contacts.csv"
    $json = $jsonDir + "contacts.json"
    $qryString = "select * from UserInfo"

    $command.CommandText = $qryString
    $adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
    $dataset = New-Object -TypeName System.Data.DataSet
    $adapter.Fill($dataset)
    $dataset.Tables[0] | export-csv $tempCsv -NoTypeInformation

    Import-Csv $tempCsv | select @{Name = "nID"; Expression = {$_."ID"}},
        @{Name = "sRequesters"; Expression = {$_."Name"}},
        @{Name = "sRequesterDepartment"; Expression = {$_."Department"}},
        @{Name = "sTitle"; Expression = {$_."Title"}} | Export-csv $csv -NoTypeInformation

    import-csv $csv | ConvertTo-Json | New-Item -path $jsonDir -Name "contacts.json" -ItemType file -force
}

function exportRequests {
    write-host "%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%  Exporting ITAG Submissions Information to CSV and JSON  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%"

    $csv = $csvDir + "requests.csv"
    $json = $jsonDir + "requests.json"

    $qryString = "Select * from [RequestRegistry]"
    $command.CommandText = $qryString
    $adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
    $dataset = New-Object -TypeName System.Data.DataSet
    $adapter.Fill($dataset)
    $dataset.Tables[0] | export-csv $tempCsv -NoTypeInformation

    Import-Csv $tempCsv | select @{Name = "nID"; Expression = {$_."ID"}},
        @{Name = "sContentType"; Expression = {$_."Content Type"}},
        @{Name = "sProductManufacturer"; Expression = {$_."Product Manufacturer"}},
        @{Name = "sProductName"; Expression = {$_."Product Name"}},
        @{Name = "dDateRequested"; Expression = {retConvertDate($_."Date Requested")}},
        @{Name = "sTechnologyHostingModel"; Expression = {$_."Technology Hosting Model"}},
        @{Name = "sApplicationHostingModel"; Expression = {$_."Application Hosting Model"}},
        @{Name = "sDeploymentType"; Expression = {$_."Deployment Type"}},
        @{Name = "sPurposeForEngagement"; Expression = {$_."Purpose for Engaging the Governance Board?"}},
        @{Name = "bNewITCapability"; Expression = {$_."New IT Capability to Devon?"}},
        @{Name = "sITCapability"; Expression = {retNormTaxonomy($_."What is the Primary IT Capability that this Technology Provides?")}},
        @{Name = "sBusinessCapability"; Expression = {retNormTaxonomy($_."What are the Business Capabilities that this Application Support")}},
        @{Name = "bCurrentAppWithCapabilities"; Expression = {$_."Does Devon Currently License Any Applications that Have Similar "}},
        @{Name = "sNumberOfUsers"; Expression = {$_."How Many Users will be Using this Application?"}},
        @{Name = "bIntegrationRequired"; Expression = {$_."Will the Application Either Consume Data From, or Provide Data t"}},
        @{Name = "sIntegrationMethod"; Expression = {$_."Integration Method"}},
        @{Name = "bExternalDataSources"; Expression = {$_."Is the Data in Either the Source System or Target System owned b"}},
        @{Name = "sDataObjectsInvolvedInIntegration"; Expression = {retNormTaxonomy($_."What Types of Data will be Provided or Consumed as Part of the I")}},
        @{Name = "sSourceSystem"; Expression = {$_."Proposed Source Application or System"}},
        @{Name = "sTargetSystem"; Expression = {$_."Proposed Target Application or System"}},
        @{Name = "sProductNameWithManufacture"; Expression = {$_."Product Name with Product Manufacturer"}},
        @{Name = "bOpenSource"; Expression = {$_."Is This Product Open Source?"}},
        @{Name = "nEstGaCost"; Expression = {$_."Estimated G&A Cost"}},
        @{Name = "nEstCapitalCost"; Expression = {$_."Estimated Capital Cost"}} | Export-csv $csv -NoTypeInformation

    import-csv $csv | ConvertTo-Json | New-Item -path $jsonDir -Name "requests.json" -ItemType file -force
}

function exportConsolidated {
    write-host "%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%  Merging Decisions and Submissions Information, then Exporting Information to CSV and JSON  %%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%"

    $csv = $csvDir + "consolidated.csv"
    $json = $jsonDir + "consolidated.json"

    $qryString = "Select * from [DecisionRegistry]"
    $oDecDbConn = New-Object -comobject ADODB.Connection
    $oDecDbConn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + $dbPath)
    $oRsDec= New-Object -comobject ADODB.Recordset
    $oRsDec.Open($qryString, $oDecDbConn,$adOpenStatic,$adLockOptimistic)
    $oRsDec.MoveFirst()

    $csvContent = [char]34 + "nSubmissionID" + [char]34 + "," + [char]34 + "sItagID" + [char]34 + "," + [char]34 + "dDateReviewed" + [char]34 + "," + [char]34 + "sItagGovernanceType" + [char]34 + "," + [char]34 + "sDecision" + [char]34 + "," + [char]34 + "bConstraints" + [char]34 + "," + [char]34 + "bActionItems" + [char]34 + "," + [char]34 + "sContentType" + [char]34 + "," + [char]34 + "sProductManufacturer" + [char]34 + "," + [char]34 + "sProductName" + [char]34 + "," + [char]34 + "dDateRequested" + [char]34 + "," + [char]34 + "sTechnologyHostingModel" + [char]34 + "," + [char]34 + "sApplicationHostingModel" + [char]34 + "," + [char]34 + "sDeploymentType" + [char]34 + "," + [char]34 + "sPurposeForEngagement" + [char]34 + "," + [char]34 + "bNewITCapability" + [char]34 + "," + [char]34 + "sITCapability" + [char]34 + "," + [char]34 + "sBusinessCapability" + [char]34 + "," + [char]34 + "bCurrentLicenses" + [char]34 + "," + [char]34 + "sNumberOfUsers" + [char]34 + "," + [char]34 + "bIntegrationRequired" + [char]34 + "," + [char]34 + "sIntegrationMethod" + [char]34 + "," + [char]34 + "bExternalDataSources" + [char]34 + "," + [char]34 + "sDataObjectsInvolvedInIntegration" + [char]34 + "," + [char]34 + "sSourceSystem" + [char]34 + "," + [char]34 + "sTargetSystem" + [char]34 + "," + [char]34 + "sProductNameWithManufacture" + [char]34 + "," + [char]34 + "bOpenSource" + [char]34 + "," + [char]34 + "nEstGaCost" + [char]34 + "," + [char]34 + "nEstCapitalCost" + [char]34 + "," + [char]34 + "sITClass" + [char]34 + $crlf

    do{
        $nSubmissionID = $oRsDec.Fields.Item("Submission ID").Value ;
		$sItagID = $oRsDec.Fields.Item("ITAG ID").Value ;
		$dDateReviewed = retConvertDate($oRsDec.Fields.Item("Date Reviewed").Value) ;
		$sItagGovernanceType = $oRsDec.Fields.Item("ITAG Governance Type").Value ;
		$sDecision = $oRsDec.Fields.Item("Decision").Value ;
		$bConstraints = $oRsDec.Fields.Item("Any Applied Conditions or Constraints?").Value ;
		$bActionItems = $oRsDec.Fields.Item("Any Assigned Action Items?").Value ;

            $qryStrRequest = "Select * from [RequestRegistry] Where ID = " + $nSubmissionID
            $oReqDbConn = New-Object -comobject ADODB.Connection
            $oReqDbConn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + $dbPath)
            $oRsReq = New-Object -comobject ADODB.Recordset
            $oRsReq.Open($qryStrRequest, $oReqDbConn,$adOpenStatic,$adLockOptimistic)

            $nID = $oRsReq.Fields.Item("ID").Value ;
            $sContentType = $oRsReq.Fields.Item("Content Type").Value ;
            $sProductManufacturer = $oRsReq.Fields.Item("Product Manufacturer").Value ;
            $sProductName = $oRsReq.Fields.Item("Product Name").Value ;
            $dDateRequested = retConvertDate($oRsReq.Fields.Item("Date Requested").Value) ;
            $sTechnologyHostingModel = $oRsReq.Fields.Item("Technology Hosting Model").Value ;
            $sApplicationHostingModel = $oRsReq.Fields.Item("Application Hosting Model").Value ;
            $sDeploymentType = $oRsReq.Fields.Item("Deployment Type").Value ;
            $sNumberOfUsers = $oRsReq.Fields.Item("How Many Users will be Using this Application?").Value ;
            $bIntegrationRequired = $oRsReq.Fields.Item("Will the Application Either Consume Data From, or Provide Data t").Value ;
            $sIntegrationMethod = $oRsReq.Fields.Item("Integration Method").Value ;
            $bExternalDataSources = $oRsReq.Fields.Item("Is the Data in Either the Source System or Target System owned b").Value ;
            $sDataObjectsInvolvedInIntegration = retNormTaxonomy($oRsReq.Fields.Item("What Types of Data will be Provided or Consumed as Part of the I").Value) ;
            $sSourceSystem = $oRsReq.Fields.Item("Proposed Source Application or System").Value ;
            $sPurposeForEngagement = $oRsReq.Fields.Item("Purpose for Engaging the Governance Board?").Value ;
            $bNewITCapability = $oRsReq.Fields.Item("New IT Capability to Devon?").Value ;
            $sITCapability = retNormTaxonomy([string]$oRsReq.Fields.Item("What is the Primary IT Capability that this Technology Provides?").Value) ;
            $sBusinessCapability = retNormTaxonomy($oRsReq.Fields.Item("What are the Business Capabilities that this Application Support").Value) ;

            if ($sITCapability -eq "" -and $sBusinessCapability -eq "") {$sITClass = "Integration"}
                ElseIf ($sITCapability -ne "") {$sITClass = retHighLevelCapability($sITCapability)}
                Else {$sITClass = retHighLevelCapability($sBusinessCapability)}

            $bCurrentAppWithCapabilities = $oRsReq.Fields.Item("Does Devon Currently License Any Applications that Have Similar ").Value ;
            $sTargetSystem = $oRsReq.Fields.Item("Proposed Target Application or System").Value ;
            $sProductNameWithManufacture = $oRsReq.Fields.Item("Product Name with Product Manufacturer").Value ;
            $bOpenSource = $oRsReq.Fields.Item("Is This Product Open Source?").Value ;
            $nEstGaCost = $oRsReq.Fields.Item("Estimated G&A Cost").Value ;
            $nEstCapitalCost = $oRsReq.Fields.Item("Estimated Capital Cost").Value ;

            $csvContent = $csvContent + [char]34 + $nSubmissionID + [char]34 + "," + [char]34 + $sItagID + [char]34 + "," + [char]34 + $dDateReviewed + [char]34 + "," + [char]34 + $sItagGovernanceType + [char]34 + "," + [char]34 + $sDecision + [char]34 + "," + [char]34 + $bConstraints + [char]34 + "," + [char]34 + $bActionItems + [char]34 + "," + [char]34 + $sContentType + [char]34 + "," + [char]34 + $sProductManufacturer + [char]34 + "," + [char]34 + $sProductName + [char]34 + "," + [char]34 + $dDateRequested + [char]34 + "," + [char]34 + $sTechnologyHostingModel + [char]34 + "," + [char]34 + $sApplicationHostingModel + [char]34 + "," + [char]34 + $sDeploymentType + [char]34 + "," + [char]34 + $sPurposeForEngagement + [char]34 + "," + [char]34 + $bNewITCapability + [char]34 + "," + [char]34 + $sITCapability + [char]34 + "," + [char]34 + $sBusinessCapability + [char]34 + "," + [char]34 + $bCurrentLicenses + [char]34 + "," + [char]34 + $sNumberOfUsers + [char]34 + "," + [char]34 + $bIntegrationRequired + [char]34 + "," + [char]34 + $sIntegrationMethod + [char]34 + "," + [char]34 + $bExternalDataSources + [char]34 + "," + [char]34 + $sDataObjectsInvolvedInIntegration + [char]34 + "," + [char]34 + $sSourceSystem + [char]34 + "," + [char]34 + $sTargetSystem + [char]34 + "," + [char]34 + $sProductNameWithManufacture + [char]34 + "," + [char]34 + $bOpenSource + [char]34 + "," + [char]34 + $nEstGaCost + [char]34 + "," + [char]34 + $nEstCapitalCost + [char]34 + "," + [char]34 + $sITClass + [char]34 + $crlf

            $oRsReq.Close()
            $oReqDbConn.Close()
    $oRsDec.MoveNext()
    }
    until ($oRsDec.EOF -eq $True)

    New-Item -path $csvDir -Name "consolidated.csv" -Value $csvContent -ItemType file -force
    import-csv $csv | ConvertTo-Json | New-Item -path $jsonDir -Name "consolidated.json" -ItemType file -force
}

function exportUserMapping {
    write-host "%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%  Merging User Information with Submissions, then Exporting Information to CSV and JSON  %%%%%%%%%%%%%%"
    write-host "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    write-host "%"

    $csv = $csvDir + "userInfo.csv"
    $json = $jsonDir + "userInfo.json"

     $csvContent = [char]34 + "nID" + [char]34 + "," + [char]34 + "sRequesters" + [char]34 + "," + [char]34 + "sRequesterDepartment" + [char]34 + "," + [char]34 + "sContentType" + [char]34 + "," + [char]34 + "sProductManufacturer" + [char]34 + "," + [char]34 + "sProductName" + [char]34 + "," + [char]34 + "dDateRequested" + [char]34 + "," + [char]34 + "sTechnologyHostingModel" + [char]34 + "," + [char]34 + "sApplicationHostingModel" + [char]34 + "," + [char]34 + "sDeploymentType" + [char]34 + "," + [char]34 + "bNewITCapability" + [char]34 + "," + [char]34 + "sITCapability" + [char]34 + "," + [char]34 + "sBusinessCapability" + [char]34 + $crlf

    $qryStrRequest = "Select * from [RequestRegistry]"
    $oReqDbConn = New-Object -comobject ADODB.Connection
    $oReqDbConn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + $dbPath)
    $oRsReq = New-Object -comobject ADODB.Recordset
    $oRsReq.Open($qryStrRequest, $oReqDbConn,$adOpenStatic,$adLockOptimistic)
    $oRsReq.MoveFirst()

    do{
        $nID = $oRsReq.Fields.Item("ID").Value ;
        $sRequesters = $oRsReq.Fields.Item("Requesters").Value ;
        $sContentType = $oRsReq.Fields.Item("Content Type").Value ;
        $sProductManufacturer = $oRsReq.Fields.Item("Product Manufacturer").Value ;
        $sProductName = $oRsReq.Fields.Item("Product Name").Value ;
        $dDateRequested = retConvertDate($oRsReq.Fields.Item("Date Requested").Value) ;
        $sTechnologyHostingModel = $oRsReq.Fields.Item("Technology Hosting Model").Value ;
        $sApplicationHostingModel = $oRsReq.Fields.Item("Application Hosting Model").Value ;
        $sDeploymentType = $oRsReq.Fields.Item("Deployment Type").Value ;
        $bNewITCapability = $oRsReq.Fields.Item("New IT Capability to Devon?").Value ;
        $sITCapability = retNormTaxonomy($oRsReq.Fields.Item("What is the Primary IT Capability that this Technology Provides?").Value) ;
        $sBusinessCapability = retNormTaxonomy($oRsReq.Fields.Item("What are the Business Capabilities that this Application Support").Value) ;
        if ($sRequesters -ne ""){
            $submitters = $sRequesters.split(";")
            foreach($person in $submitters){
                    $qryString = "Select * from [UserInfo] Where ID = " + $person
                    $oUserDbConn = New-Object -comobject ADODB.Connection
                    $oUserDbConn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + $dbPath)
                    $oRsUser = New-Object -comobject ADODB.Recordset
                    $oRsUser.Open($qryString, $oUserDbConn,$adOpenStatic,$adLockOptimistic)
                    If ($oRsUser.EOF -ne $True){
                        $submitter = $oRsUser.Fields.Item("Name").Value ;
                        $submitterDepartment = $oRsUser.Fields.Item("Department").Value ;

                        $csvContent = $csvContent + [char]34 + $nID + [char]34 + "," + [char]34 + $submitter + [char]34 + "," + [char]34 + $submitterDepartment + [char]34 + "," + [char]34 + $sContentType + [char]34 + "," + [char]34 + $sProductManufacturer + [char]34 + "," + [char]34 + $sProductName + [char]34 + "," + [char]34 + $dDateRequested + [char]34 + "," + [char]34 + $sTechnologyHostingModel + [char]34 + "," + [char]34 + $sApplicationHostingModel + [char]34 + "," + [char]34 + $sDeploymentType + [char]34 + "," + [char]34 + $bNewITCapability + [char]34 + "," + [char]34 + $sITCapability + [char]34 + "," + [char]34 + $sBusinessCapability + [char]34 + $crlf
                    }
                $oRsUser.Close()
                $oUserDbConn.Close()
            }
        }
        $oRsReq.MoveNext()
    }
    until ($oRsReq.EOF -eq $True)

    New-Item -path $csvDir -Name "userInfo.csv" -Value $csvContent -ItemType file -force
    import-csv $csv | ConvertTo-Json | New-Item -path $jsonDir -Name "userInfo.json" -ItemType file -force
}

#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
#%%%%%%%%%%%%%%%%%%%%%%%%%%  Script Execution %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

importData
exportDecisions
exportDevonContacts
exportRequests
exportConsolidated
exportUserMapping
