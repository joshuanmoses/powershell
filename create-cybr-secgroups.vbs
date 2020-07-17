'====================================

on error resume next

Const ForReading = 1
Const ForWrite = 2
Const ForAppending = 8

const ADS_GROUP_TYPE_GLOBAL_GROUP        = &h00000002
const ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP  = &h00000004
const ADS_GROUP_TYPE_LOCAL_GROUP         = &h00000004
const ADS_GROUP_TYPE_UNIVERSAL_GROUP     = &h00000008
const ADS_GROUP_TYPE_SECURITY_ENABLED    = &h80000000

Set fso = WScript.CreateObject("Scripting.FileSystemObject")

arrGroups = array("ChangePassword", "ViewPassword", "ConnectAs", "Approvers")

set oArgs=wscript.arguments

if oArgs.Count > 0 then
	strSafeName = oArgs(0)
else
	MsgBox "No safe name specified." & vbcrlf & "Script exiting"
	wscript.quit
end if

set oArgs = Nothing

blnYesNo = MsgBox ("Create groups for safe name " & vbcrlf & chr(34) & strSafeName &chr(34) & " ?", vbYesNo)
if blnYesNo = vbNo then
	wscript.quit
end if

set objOU = getobject("LDAP://ou=CyberArk Groups,ou=Domain Groups,dc=corp,dc=yourdomain,dc=com")

for x = 0 to UBound(arrGroups)
		strGroupName = "CA_" & strSafeName & "_" & arrGroups(x)
		strSAMAccountName = strGroupName

		err.clear

		set objgroup = objOU.create("Group", "cn=" & strGroupName)

		objGroup.put "samaccountname", strSAMAccountName
		objGroup.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP or ADS_GROUP_TYPE_SECURITY_ENABLED

		objGroup.SetInfo
		if err.number = 0 then
			MsgBox strGroupname & " created."
		else
			MsgBox "Error creating " & vbcrlf & chr(34) & strGroupname & chr(34) & vbcrlf & err.number & " " & err.description
		end if

		set Group = Nothing

Next

MsgBox "Script Finished."
