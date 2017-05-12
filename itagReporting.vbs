'Build base filesytem object for needed for simple filesystem operations
Set oFs = CreateObject("Scripting.FileSystemObject")

'Root directory for the script working area. The Microsoft Access database and output content is found in directories beneath the root.
workingDir = oFs.GetAbsolutePathName(".")
accessdbLoc = workingDir & "\db\itag.accdb"
csvOutputLoc =  workingDir & "\exports\csv\"
jsonOutputLoc = workingDir & "\exports\json\"

consolidatedCsv = csvOutputLoc & "consolidated.csv"
decisionsCsv = csvOutputLoc & "decisions.csv"
requestCsv= csvOutputLoc & "requests.csv"
userInfoCsv = csvOutputLoc & "userInfo.csv"

consolidatedJson = jsonOutputLoc & "consolidated.json"
decisionsJson = jsonOutputLoc & "decisions.json"
requestJson= jsonOutputLoc & "requests.json"
userInfoJson = jsonOutputLoc & "userInfo.json"

'Source Microsoft Access table names
DecisionsSrc = "DecisionRegistry"
RequestsSrc = "RequestRegistry"
UserInfoSrc = "UserInfo"

'Build base querys for data retrieval'
qryDecisions = "SELECT * FROM " & DecisionsSrc
qryRequests = "SELECT * FROM " & RequestsSrc
qryUser = "SELECT * FROM " & UserInfoSrc

'Establish database connection
Set oDbConn = CreateObject("ADODB.Connection")
sDbConn = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & accessdbLoc
oDbConn.open sDbConn

RefreshData
'ArchiveOldDocuments
BuildDecisionsDocument
BuildRequestsDocument
BuildUsersDocument
BuildConsolidatedRecordsDocument

'Kill Global Objects
oDbConn.Close
Set oDbConn = Nothing
Set oFs = Nothing
wscript.echo "All Done.... Completed Data Refresh and Export Processing."

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%% Functions and Subroutines %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Function fRegReplace(sString, sExpression, sReplace)
	Set  oRegEx = CreateObject("VBScript.RegExp")
	oRegEx.Global = True
	oRegEx.IgnoreCase = True
	oRegEx.Pattern = sExpression
	fRegReplace = oRegEx.Replace(sString, sReplace)
	Set oRegEx = Nothing
end Function

Function fNormalizeTaxonomy(sString)
	tempVal = fRegReplace(sString, "\d\d\d\;\#", "")
	tempVal = fRegReplace(tempVal, "\d\d\;\#", "")
	tempVal = fRegReplace(tempVal, "\d\;\#", "")
	tempVal = fRegReplace(tempVal, "\#","")
	fNormalizeTaxonomy = tempVal
end Function

Sub RefreshData
	wscript.echo "Refreshing Access Databases."
	Set accessApp = createObject("Access.Application")
	accessApp.visible = true
	accessApp.OpenCurrentDataBase(accessdbLoc)
	accessApp.DoCmd.RunMacro "refreshItagData"
	Set accessApp = Nothing
End Sub

Sub ArchiveOldDocuments
	wscript.echo "Archiving older CSV Documents."
	sDate = cstr(date())
	sDate = replace(sDate, "/", "-")
	sTime = cstr(formatDateTime(time(),3))
	sTime = replace(sTime, ":", "")
	sTime = replace(sTime, " AM", "")
	sTime = replace(sTime, " PM", "")
	bkpFolderName = sDate & "_" & sTime
	bkpLoc = csvOutputLoc & "hist\" & bkpFolderName & "\"

	If oFs.FileExists(decisionsCsv) Then
		if Not oFs.FolderExists(bkpLoc) Then
			oFs.CreateFolder bkpLoc
		End If
		oFs.MoveFile decisionsCsv, bkpLoc
	End If

	If oFs.FileExists(requestCsv) Then
		if Not oFs.FolderExists(bkpLoc) Then
			oFs.CreateFolder bkpLoc
		End If
		oFs.MoveFile requestCsv, bkpLoc
	End If

	If oFs.FileExists(userInfoCsv) Then
		if Not oFs.FolderExists(bkpLoc) Then
			oFs.CreateFolder bkpLoc
		End If
		oFs.MoveFile userInfoCsv, bkpLoc
	End If

	If oFs.FileExists(consolidatedCsv) Then
		if Not oFs.FolderExists(bkpLoc) Then
			oFs.CreateFolder bkpLoc
		End If
		oFs.MoveFile consolidatedCsv, bkpLoc
	End If
End Sub

Sub BuildDecisionsDocument
	wscript.echo "Exporting DecisionRegistry Table to CSV and Json."
	Set rs = oDbConn.execute(qryDecisions)
	Set oCsv = oFS.CreateTextFile(decisionsCsv,True)
	Set oJson = oFS.CreateTextFile(decisionsJson,True)

	oCsv.WriteLine """nSubmissionID""" & "," & """sItagID""" & "," & """dDateReviewed""" & "," & """sItagGovernanceType""" & "," & """sDecision""" & "," & """bConstraints""" & "," & """bActionItems"""
	oJson.WriteLine "["

	DO WHILE NOT rs.EOF
		nSubmissionID = rs("Submission ID")
		sItagID = rs("ITAG ID")
		dDateReviewed = rs("Date Reviewed")
		sItagGovernanceType = rs("ITAG Governance Type")
		sDecision = rs("Decision")
		bConstraints = rs("Any Applied Conditions or Constraints?")
		bActionItems = rs("Any Assigned Action Items?")

		oJson.WriteLine vbtab & "{"
		oJson.WriteLine vbtab & """nSubmissionID""" & ": " & nSubmissionID & ","
		oJson.WriteLine vbtab & """sItagID""" & ": " & chr(34) & sItagID & chr(34) & ","
		oJson.WriteLine vbtab & """dDateReviewed""" & ": " & chr(34) & dDateReviewed & chr(34) & ","
		oJson.WriteLine vbtab & """sItagGovernanceType""" & ": " & chr(34) & sItagGovernanceType & chr(34) & ","
		oJson.WriteLine vbtab & """sDecision""" & ": " & chr(34) & sDecision & chr(34) & ","
		oJson.WriteLine vbtab & """bConstraints""" & ": " & chr(34) & bConstraints & chr(34) & ","
		oJson.WriteLine vbtab & """bActionItems""" & ": " & chr(34) & bActionItems & chr(34)
		oCsv.WriteLine nSubmissionID & "," & chr(34) & sItagID & chr(34) & "," & dDateReviewed & "," & chr(34) & sItagGovernanceType & chr(34) & "," & chr(34) & sDecision & chr(34) & "," & bConstraints & "," & bActionItems

		rs.MoveNext

		If NOT rs.EOF Then
			oJson.WriteLine vbtab & "},"
		Else
			oJson.WriteLine vbtab & "}"
		End If
	Loop
	oJson.Writeline "]"

	'Kill objects used to create the ITAG DECISIONS document
	Set rs = Nothing
	Set oCsv = Nothing
	Set oJson = Nothing
End Sub

Sub BuildRequestsDocument
	wscript.echo "Exporting RequestRegistry Table to CSV and Json."
	Set rs = oDbConn.execute(qryRequests)
	Set oCsv = oFS.CreateTextFile(requestCsv,True)
	Set oJson = oFS.CreateTextFile(requestJson,True)

	oCsv.WriteLine """nID""" & "," & """sContentType""" & "," & """sProductManufacturer""" & "," & """sProductName""" & "," & """sRequesters""" & "," & """dDateRequested""" & "," & """sTechnologyHostingModel""" & "," & """sApplicationHostingModel""" & "," & """sDeploymentType""" & "," & """sPurposeForEngagement""" & "," & """bNewITCapability""" & "," & """sITCapability""" & "," & """sBusinessCapability""" & "," & """bCurrentLicenses""" & "," & """sNumberOfUsers""" & "," & """bIntegrationRequired""" & "," & """sIntegrationMethod""" & "," & """bExternalDataSources""" & "," & """sDataObjectsInvolvedInIntegration""" & "," & """sSourceSystem""" & "," & """sTargetSystem""" & "," & """sProductNameWithManufacture""" & "," & """bOpenSource""" & "," & """nEstGaCost""" & "," & """nEstCapitalCost"""
	oJson.WriteLine "["

	DO WHILE NOT rs.EOF
		nID = rs("ID")
		sContentType = rs("Content Type")
		sProductManufacturer = rs("Product Manufacturer")
		sProductName = rs("Product Name")
		sRequesters = rs("Requesters")
		dDateRequested = rs("Date Requested")
		sTechnologyHostingModel = rs("Technology Hosting Model")
		sApplicationHostingModel = rs("Application Hosting Model")
		sDeploymentType = rs("Deployment Type")
		sPurposeForEngagement = rs("Purpose for Engaging the Governance Board?")
		bNewITCapability = rs("New IT Capability to Devon?")

		If IsNull(rs("What is the Primary IT Capability that this Technology Provides?")) Then
			sITCapability = rs("What is the Primary IT Capability that this Technology Provides?")
		Else
			sITCapability = fNormalizeTaxonomy(rs("What is the Primary IT Capability that this Technology Provides?"))
		End If

		If IsNull(rs("What are the Business Capabilities that this Application Support")) Then
			sBusinessCapability = rs("What are the Business Capabilities that this Application Support")
		Else
			sBusinessCapability = fNormalizeTaxonomy(rs("What are the Business Capabilities that this Application Support"))
		End If

		bCurrentAppWithCapabilities = rs("Does Devon Currently License Any Applications that Have Similar ")
		sNumberOfUsers = rs("How Many Users will be Using this Application?")
		bIntegrationRequired = rs("Will the Application Either Consume Data From, or Provide Data t")
		sIntegrationMethod = rs("Integration Method")
		bExternalDataSources = rs("Is the Data in Either the Source System or Target System owned b")

		If IsNull(rs("What Types of Data will be Provided or Consumed as Part of the I")) Then
			sDataObjectsInvolvedInIntegration = rs("What Types of Data will be Provided or Consumed as Part of the I")
		Else
			sDataObjectsInvolvedInIntegration = fNormalizeTaxonomy(rs("What Types of Data will be Provided or Consumed as Part of the I"))
		End If

		sSourceSystem = rs("Proposed Source Application or System")
		sTargetSystem = rs("Proposed Target Application or System")
		sProductNameWithManufacture = rs("Product Name with Product Manufacturer")
		bOpenSource = rs("Is This Product Open Source?")
		nEstGaCost = rs("Estimated G&A Cost")
		nEstCapitalCost = rs("Estimated Capital Cost")

		oJson.WriteLine vbtab & "{"
		oJson.WriteLine vbtab & """nID""" & ": " & nID & ","
		oJson.WriteLine vbtab & """sContentType""" & ": " & chr(34) & sContentType & chr(34) & ","
		oJson.WriteLine vbtab & """sProductManufacturer""" & ": " & chr(34) & sProductManufacturer & chr(34) & ","
		oJson.WriteLine vbtab & """sProductName""" & ": " & chr(34) & sProductName & chr(34) & ","
		oJson.WriteLine vbtab & """sRequesters""" & ": " & chr(34) & sRequesters & chr(34) & ","
		oJson.WriteLine vbtab & """dDateRequested""" & ": " & chr(34) & dDateRequested & chr(34) & ","
		oJson.WriteLine vbtab & """sTechnologyHostingModel""" & ": " & chr(34) & sTechnologyHostingModel & chr(34) & ","
		oJson.WriteLine vbtab & """sApplicationHostingModel""" & ": " & chr(34) & sApplicationHostingModel & chr(34) & ","
		oJson.WriteLine vbtab & """sDeploymentType""" & ": " & chr(34) & sDeploymentType & chr(34) & ","
		oJson.WriteLine vbtab & """sPurposeForEngagement""" & ": " & chr(34) & sPurposeForEngagement & chr(34) & ","
		oJson.WriteLine vbtab & """bNewITCapability""" & ": " & chr(34) & bNewITCapability & chr(34) & ","
		oJson.WriteLine vbtab & """sITCapability""" & ": " & chr(34) & sITCapability & chr(34) & ","
		oJson.WriteLine vbtab & """sBusinessCapability""" & ": " & chr(34) & sBusinessCapability & chr(34) & ","
		oJson.WriteLine vbtab & """bCurrentLicenses""" & ": " & chr(34) & bCurrentAppWithCapabilities & chr(34) & ","
		oJson.WriteLine vbtab & """sNumberOfUsers""" & ": " & chr(34) & sNumberOfUsers & chr(34) & ","
		oJson.WriteLine vbtab & """bIntegrationRequired""" & ": " & chr(34) & bIntegrationRequired & chr(34) & ","
		oJson.WriteLine vbtab & """sIntegrationMethod""" & ": " & chr(34) & sIntegrationMethod & chr(34) & ","
		oJson.WriteLine vbtab & """bExternalDataSources""" & ": " & chr(34) & bExternalDataSources & chr(34) & ","
		oJson.WriteLine vbtab & """sDataObjectsInvolvedInIntegration""" & ": " & chr(34) & sDataObjectsInvolvedInIntegration & chr(34) & ","
		oJson.WriteLine vbtab & """sSourceSystem""" & ": " & chr(34) & sSourceSystem & chr(34) & ","
		oJson.WriteLine vbtab & """sTargetSystem""" & ": " & chr(34) & sTargetSystem & chr(34) & ","
		oJson.WriteLine vbtab & """sProductNameWithManufacture""" & ": " & chr(34) & sProductNameWithManufacture & chr(34) & ","
		oJson.WriteLine vbtab & """bOpenSource""" & ": " & chr(34) & bOpenSource & chr(34) & ","
		oJson.WriteLine vbtab & """nEstGaCost""" & ": " & nEstGaCost & ","
		oJson.WriteLine vbtab & """nEstCapitalCost""" & ": " & nEstCapitalCost

		oCsv.WriteLine nID & "," & chr(34) & sContentType & chr(34) & "," & chr(34) & sProductManufacturer & chr(34) & "," & chr(34) & sProductName & chr(34) & "," & chr(34) & sRequesters & chr(34) & "," & dDateRequested & "," & chr(34) & sTechnologyHostingModel & chr(34) & "," & chr(34) & sApplicationHostingModel & chr(34) & "," & chr(34) & sDeploymentType & chr(34) & "," & chr(34) & sPurposeForEngagement & chr(34) & "," & bNewITCapability & "," & chr(34) & sITCapability & chr(34) & "," & chr(34) & sBusinessCapability & chr(34) & "," & bCurrentAppWithCapabilities & "," & chr(34) & sNumberOfUsers & chr(34) & ","& bIntegrationRequired & "," & chr(34) & sIntegrationMethod& chr(34) & ","& bExternalDataSources & "," & chr(34) & sDataObjectsInvolvedInIntegration & chr(34) & "," & chr(34) & sSourceSystem & chr(34) & "," & chr(34) & sTargetSystem & chr(34) & "," & chr(34) & sProductNameWithManufacture & chr(34) & "," & bOpenSource & "," & nEstGaCost & "," & nEstCapitalCost
		rs.MoveNext

		If NOT rs.EOF Then
			oJson.WriteLine vbtab & "},"
		Else
			oJson.WriteLine vbtab & "}"
		End If
	Loop
	oJson.Writeline "]"

	'Kill objects used to create the ITAG SUBMISSIONS REQUEST document
	Set rs = Nothing
	Set oCsv = Nothing
	Set oJson = Nothing
End Sub

Sub BuildUsersDocument
	wscript.echo "Exporting UserInfo Table to CSV and Json."
	Set rsReq = oDbConn.execute(qryRequests)
	Set oCsv = oFS.CreateTextFile(userInfoCsv, True)
	Set oJson = oFS.CreateTextFile(userInfoJson,True)

	oCsv.WriteLine """nID""" & "," & """sRequesters""" & "," & """sRequesterDepartment""" & "," & """sContentType""" & "," & """sProductManufacturer""" & "," & """sProductName""" & "," & """dDateRequested""" & "," & """sTechnologyHostingModel""" & "," & """sApplicationHostingModel""" & "," & """sDeploymentType""" & "," & """bNewITCapability""" & "," & """sITCapability""" & "," & """sBusinessCapability"""
	oJson.WriteLine "["

	Do While NOT rsReq.EOF
		recordCount = recordCount + 1
		rsReq.MoveNext
	Loop

	rsReq.MoveFirst
	DO WHILE NOT rsReq.EOF
		recordCheck = recordCheck + 1

		nID = rsReq("ID")
		sRequesters = rsReq("Requesters")
		sContentType = rsReq("Content Type")
		sProductManufacturer = rsReq("Product Manufacturer")
		sProductName = rsReq("Product Name")
		dDateRequested = rsReq("Date Requested")
		sTechnologyHostingModel = rsReq("Technology Hosting Model")
		sApplicationHostingModel = rsReq("Application Hosting Model")
		sDeploymentType = rsReq("Deployment Type")
		bNewITCapability = rsReq("New IT Capability to Devon?")

		If IsNull(rsReq("What is the Primary IT Capability that this Technology Provides?")) Then
			sITCapability = rsReq("What is the Primary IT Capability that this Technology Provides?")
		Else
			sITCapability = fNormalizeTaxonomy(rsReq("What is the Primary IT Capability that this Technology Provides?"))
		End If

		If IsNull(rsReq("What are the Business Capabilities that this Application Support")) Then
			sBusinessCapability = rsReq("What are the Business Capabilities that this Application Support")
		Else
			sBusinessCapability = fNormalizeTaxonomy(rsReq("What are the Business Capabilities that this Application Support"))
		End If

		if sRequesters <> "" Then
			aRequesters = Split(sRequesters,";")
			For i = lbound(aRequesters) to ubound(aRequesters)
				Set rsUserInfo = oDbConn.execute(qryUser & " Where ID = " & aRequesters(i))
				If Not rsUserInfo.EOF then
					oJson.WriteLine vbtab & "{"
					oJson.WriteLine vbtab & """nID""" & ": " & nID & ","
					oJson.WriteLine vbtab & """sRequesters""" & ": " & chr(34) & rsUserInfo("Name") & chr(34) & ","
					oJson.WriteLine vbtab & """sRequesterDepartment""" & ": " & chr(34) & rsUserInfo("Department") & chr(34) & ","
					oJson.WriteLine vbtab & """sContentType""" & ": " & chr(34) & sContentType & chr(34) & ","
					oJson.WriteLine vbtab & """sProductManufacturer""" & ": " & chr(34) & sProductManufacturer & chr(34) & ","
					oJson.WriteLine vbtab & """sProductName""" & ": " & chr(34) & sProductName & chr(34) & ","
					oJson.WriteLine vbtab & """dDateRequested""" & ": " & chr(34) & dDateRequested & chr(34) & ","
					oJson.WriteLine vbtab & """sTechnologyHostingModel""" & ": " & chr(34) & sTechnologyHostingModel & chr(34) & ","
					oJson.WriteLine vbtab & """sApplicationHostingModel""" & ": "& chr(34) & sApplicationHostingModel & chr(34) & ","
					oJson.WriteLine vbtab & """sDeploymentType""" & ": " & chr(34) & sDeploymentType & chr(34) & ","
					oJson.WriteLine vbtab & """bNewITCapability""" & ": " & chr(34) & bNewITCapability & chr(34) & ","
					oJson.WriteLine vbtab & """sITCapability""" & ": " & chr(34) & sITCapability & chr(34) & ","
					oJson.WriteLine vbtab & """sBusinessCapability""" & ": " & chr(34) & sBusinessCapability & chr(34)

					oCsv.WriteLine nID & "," & chr(34) & rsUserInfo("Name") & chr(34) & "," & chr(34) & rsUserInfo("Department") & chr(34) & "," & chr(34) & sContentType & chr(34) & "," & chr(34) & sProductManufacturer & chr(34) & "," & chr(34) & sProductName & chr(34) & "," & dDateRequested & "," & chr(34) & sTechnologyHostingModel & chr(34) & "," & chr(34) & sApplicationHostingModel & chr(34) & "," & chr(34) & sDeploymentType & chr(34) & "," & bNewITCapability & "," & chr(34) & sITCapability & chr(34) & "," & chr(34) & sBusinessCapability & chr(34)
					If (recordCheck = recordCount) AND (i = ubound(aRequesters)) Then
						oJson.WriteLine vbtab & "}"
					Else
						oJson.WriteLine vbtab & "},"
					End If
				End If
			Next
		End If
		rsReq.MoveNext
	Loop
	oJson.Writeline "]"

	'Kill objects used to create the USERINFORMATION document
	Set rsReq = Nothing
	Set rsUserInfo = Nothing
	Set oCsv = Nothing
	Set oJson = Nothing
End Sub

Sub BuildConsolidatedRecordsDocument
	wscript.echo "Building Consolidated Records Datasets in CSV and Json."
	Set rsDec = oDbConn.execute(qryDecisions)
	Set oCsv = oFS.CreateTextFile(consolidatedCsv, True)
	Set oJson = oFS.CreateTextFile(consolidatedJson,True)

	oCsv.WriteLine """nSubmissionID""" & "," & """sItagID""" & "," & """dDateReviewed""" & "," & """sItagGovernanceType""" & "," & """sDecision""" & "," & """bConstraints""" & "," & """bActionItems""" & "," & """sContentType""" & "," & """sProductManufacturer""" & "," & """sProductName""" & "," & """dDateRequested""" & "," & """sTechnologyHostingModel""" & "," & """sApplicationHostingModel""" & "," & """sDeploymentType""" & "," & """sPurposeForEngagement""" & "," & """bNewITCapability""" & "," & """sITCapability""" & "," & """sBusinessCapability""" & "," & """bCurrentLicenses""" & "," & """sNumberOfUsers""" & "," & """bIntegrationRequired""" & "," & """sIntegrationMethod""" & "," & """bExternalDataSources""" & "," & """sDataObjectsInvolvedInIntegration""" & "," & """sSourceSystem""" & "," & """sTargetSystem""" & "," & """sProductNameWithManufacture""" & "," & """bOpenSource""" & "," & """nEstGaCost""" & "," & """nEstCapitalCost"""
	oJson.WriteLine "["

	DO WHILE NOT rsDec.EOF
		nSubmissionID = rsDec("Submission ID")
		sItagID = rsDec("ITAG ID")
		dDateReviewed = rsDec("Date Reviewed")
		sItagGovernanceType = rsDec("ITAG Governance Type")
		sDecision = rsDec("Decision")
		bConstraints = rsDec("Any Applied Conditions or Constraints?")
		bActionItems = rsDec("Any Assigned Action Items?")

		Set rsReq = oDbConn.execute(qryRequests & " Where ID = " & nSubmissionID)
			nID = rsReq("ID")
			sContentType = rsReq("Content Type")
			sProductManufacturer = rsReq("Product Manufacturer")
			sProductName = rsReq("Product Name")
			dDateRequested = rsReq("Date Requested")
			sTechnologyHostingModel = rsReq("Technology Hosting Model")
			sApplicationHostingModel = rsReq("Application Hosting Model")
			sDeploymentType = rsReq("Deployment Type")
			sNumberOfUsers = rsReq("How Many Users will be Using this Application?")
			bIntegrationRequired = rsReq("Will the Application Either Consume Data From, or Provide Data t")
			sIntegrationMethod = rsReq("Integration Method")
			bExternalDataSources = rsReq("Is the Data in Either the Source System or Target System owned b")

			If IsNull(rsReq("What Types of Data will be Provided or Consumed as Part of the I")) Then
				sDataObjectsInvolvedInIntegration = rsReq("What Types of Data will be Provided or Consumed as Part of the I")
			Else
				sDataObjectsInvolvedInIntegration = fNormalizeTaxonomy(rsReq("What Types of Data will be Provided or Consumed as Part of the I"))
			End If

			sSourceSystem = rsReq("Proposed Source Application or System")
			sPurposeForEngagement = rsReq("Purpose for Engaging the Governance Board?")
			bNewITCapability = rsReq("New IT Capability to Devon?")

			If IsNull(rsReq("What is the Primary IT Capability that this Technology Provides?")) Then
				sITCapability = rsReq("What is the Primary IT Capability that this Technology Provides?")
			Else
				sITCapability = fNormalizeTaxonomy(rsReq("What is the Primary IT Capability that this Technology Provides?"))
			End If

			If IsNull(rsReq("What are the Business Capabilities that this Application Support")) Then
				sBusinessCapability = rsReq("What are the Business Capabilities that this Application Support")
			Else
				sBusinessCapability = fNormalizeTaxonomy(rsReq("What are the Business Capabilities that this Application Support"))
			End If

			bCurrentAppWithCapabilities = rsReq("Does Devon Currently License Any Applications that Have Similar ")
			sTargetSystem = rsReq("Proposed Target Application or System")
			sProductNameWithManufacture = rsReq("Product Name with Product Manufacturer")
			bOpenSource = rsReq("Is This Product Open Source?")
			nEstGaCost = rsReq("Estimated G&A Cost")
			nEstCapitalCost = rsReq("Estimated Capital Cost")

			oJson.WriteLine vbtab & "{"
			oJson.WriteLine vbtab & """nSubmissionID""" & ": " & nSubmissionID & ","
			oJson.WriteLine vbtab & """sItagID""" & ": " & chr(34) & sItagID & chr(34) & ","
			oJson.WriteLine vbtab & """dDateReviewed""" & ": " & chr(34) & dDateReviewed & chr(34) & ","
			oJson.WriteLine vbtab & """sItagGovernanceType""" & ": " & chr(34) & sItagGovernanceType & chr(34) & ","
			oJson.WriteLine vbtab & """sDecision""" & ": " & chr(34) & sDecision & chr(34) & ","
			oJson.WriteLine vbtab & """bConstraints""" & ": " & chr(34) & bConstraints & chr(34) & ","
			oJson.WriteLine vbtab & """bActionItems""" & ": " & chr(34) & bActionItems & chr(34) & ","
			oJson.WriteLine vbtab & """sContentType""" & ": " & chr(34) & sContentType & chr(34) & ","
			oJson.WriteLine vbtab & """sProductManufacturer""" & ": " & chr(34) & sProductManufacturer & chr(34) & ","
			oJson.WriteLine vbtab & """sProductName""" & ": " & chr(34) & sProductName & chr(34) & ","
			oJson.WriteLine vbtab & """dDateRequested""" & ": " & chr(34) & dDateRequested & chr(34) & ","
			oJson.WriteLine vbtab & """sTechnologyHostingModel""" & ": " & chr(34) & sTechnologyHostingModel & chr(34) & ","
			oJson.WriteLine vbtab & """sApplicationHostingModel""" & ": " & chr(34) & sApplicationHostingModel & chr(34) & ","
			oJson.WriteLine vbtab & """sDeploymentType""" & ": " & chr(34) & sDeploymentType & chr(34) & ","
			oJson.WriteLine vbtab & """sPurposeForEngagement""" & ": " & chr(34) & sPurposeForEngagement & chr(34) & ","
			oJson.WriteLine vbtab & """bNewITCapability""" & ": " & chr(34) & bNewITCapability & chr(34) & ","
			oJson.WriteLine vbtab & """sITCapability""" & ": " & chr(34) & sITCapability & chr(34) & ","
			oJson.WriteLine vbtab & """sBusinessCapability""" & ": " & chr(34) & sBusinessCapability & chr(34) & ","
			oJson.WriteLine vbtab & """bCurrentLicenses""" & ": " & chr(34) & bCurrentAppWithCapabilities & chr(34) & ","
			oJson.WriteLine vbtab & """sNumberOfUsers""" & ": " & chr(34) & sNumberOfUsers & chr(34) & ","
			oJson.WriteLine vbtab & """bIntegrationRequired""" & ": " & chr(34) & bIntegrationRequired & chr(34) & ","
			oJson.WriteLine vbtab & """sIntegrationMethod""" & ": " & chr(34) & sIntegrationMethod & chr(34) & ","
			oJson.WriteLine vbtab & """bExternalDataSources""" & ": " & chr(34) & bExternalDataSources & chr(34) & ","
			oJson.WriteLine vbtab & """sDataObjectsInvolvedInIntegration""" & ": " & chr(34) & sDataObjectsInvolvedInIntegration & chr(34) & ","
			oJson.WriteLine vbtab & """sSourceSystem""" & ": " & chr(34) & sSourceSystem & chr(34) & ","
			oJson.WriteLine vbtab & """sTargetSystem""" & ": " & chr(34) & sTargetSystem & chr(34) & ","
			oJson.WriteLine vbtab & """sProductNameWithManufacture""" & ": " & chr(34) & sProductNameWithManufacture & chr(34) & ","
			oJson.WriteLine vbtab & """bOpenSource""" & ": " & chr(34) & bOpenSource & chr(34) & ","
			oJson.WriteLine vbtab & """nEstGaCost""" & ": " & nEstGaCost & ","
			oJson.WriteLine vbtab & """nEstCapitalCost""" & ": " & nEstCapitalCost

			oCsv.WriteLine nSubmissionID & "," & chr(34) & sItagID & chr(34) & "," & dDateReviewed & "," & chr(34) & sItagGovernanceType & chr(34) & "," & chr(34) & sDecision & chr(34) & "," & bConstraints & "," & bActionItems & "," & chr(34) & sContentType & chr(34) & "," & chr(34) & sProductManufacturer & chr(34) & "," & chr(34) & sProductName & chr(34) & "," & dDateRequested & "," & chr(34) & sTechnologyHostingModel & chr(34) & "," & chr(34) & sApplicationHostingModel & chr(34) & "," & chr(34) & sDeploymentType & chr(34) & "," & chr(34) & sPurposeForEngagement & chr(34) & "," & bNewITCapability & "," & chr(34) & sITCapability & chr(34) & "," & chr(34) & sBusinessCapability & chr(34) & "," & bCurrentAppWithCapabilities & "," & chr(34) & sNumberOfUsers & chr(34) & ","& bIntegrationRequired & "," & chr(34) & sIntegrationMethod& chr(34) & ","& bExternalDataSources & "," & chr(34) & sDataObjectsInvolvedInIntegration & chr(34) & "," & chr(34) & sSourceSystem & chr(34) & "," & chr(34) & sTargetSystem & chr(34) & "," & chr(34) & sProductNameWithManufacture & chr(34) & "," & bOpenSource & "," & nEstGaCost & "," & nEstCapitalCost

		rsDec.MoveNext

		If NOT rsDec.EOF Then
			oJson.WriteLine vbtab & "},"
		Else
			oJson.WriteLine vbtab & "}"
		End If
	Loop
	oJson.Writeline "]"

	'Kill objects used to create the ITAG DECISIONS document
	Set rsReq = Nothing
	Set rsDec = Nothing
	Set oCsv = Nothing
	Set oJson = Nothing
End Sub
