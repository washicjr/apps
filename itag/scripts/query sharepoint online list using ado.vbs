Set oDbConn = CreateObject("ADODB.Connection")
Set oFs = CreateObject("Scripting.FileSystemObject")

workingDir = oFs.GetAbsolutePathName(".")
outputLoc =  workingDir & "\datasets\json\"
decisionsOutput = outputLoc & "decisions.json"

BuildDecisionsJson

wscript.echo "All Done"

Sub BuildDecisionsJson
	Set oJson = oFS.CreateTextFile(decisionsOutput ,True)
	oDbConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes;DATABASE=https://dvn.sharepoint.com/teams/aigb;LIST={9eca56e4-0d17-4e6d-b1c9-99028c5b129f};"
	Set rs = oDbConn.Execute("Select * from [ITAG Decision Registry]")

	oJson.WriteLine "["
	DO WHILE NOT rs.EOF
		oJson.WriteLine vbtab & "{"
		oJson.WriteLine vbtab & """nSubmissionID""" & ": " & rs("Submission ID") & ","
		oJson.WriteLine vbtab & """sItagID""" & ": " & chr(34) & rs("ITAG ID") & chr(34) & ","
		oJson.WriteLine vbtab & """dDateReviewed""" & ": " & rs("Date Reviewed") & ","
		oJson.WriteLine vbtab & """sItagGovernanceType""" & ": " & chr(34) & rs("ITAG Governance Type") & chr(34) & ","
		oJson.WriteLine vbtab & """sDecision""" & ": " & chr(34) & rs("Decision") & chr(34) & ","
		oJson.WriteLine vbtab & """bConstraints""" & ": " & rs("Any Applied Conditions or Constraints?") & ","
		oJson.WriteLine vbtab & """bActionItems""" & ": " & rs("Any Assigned Action Items?") & ","
		rs.MoveNext
		If NOT rs.EOF Then
			oJson.WriteLine vbtab & "},"
		Else
			oJson.WriteLine vbtab & "}"
		End If
	Loop
	oJson.Writeline "]"
End Sub
