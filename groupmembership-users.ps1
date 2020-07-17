import-module activedirectory
$date = get-date -format yyyyMMddHHmmss
$Users = Get-ADUser -Filter *
Foreach ($User in $Users) {
	$User_All = Get-ADUser $User -Properties samaccountname,memberof
	$User_SAM = $User_All.samaccountname
	$User_MemberOfGroups = $User_All.memberof
	Foreach ($User_MemberOf in $User_MemberOfGroups) {
		write "$User_SAM;$User_MemberOf" | out-file "CSV Files\Group Membership - Users-$date.csv" -Append
	}
}
