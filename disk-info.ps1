#
#	Title:		disk info.ps1
#
#	Purpose:	disk info for all volumes including mount points of servers in
#
#	v1:			Joshua Moses
#				initial script
#
#	Servers:	NDCMMGTP04
#
#	Location:	D:\Program Files\PIMM Scripts\disk info
#
#	Date:		4/3/13
#
#	Output:		
#
#	Input:		servers.txt

$TotalGB = @{Name="Capacity(GB)";expression={[math]::round(($_.Capacity/ 1073741824),2)}}
$FreeGB = @{Name="FreeSpace(GB)";expression={[math]::round(($_.FreeSpace / 1073741824),2)}}
$FreePerc = @{Name="Free(%)";expression={[math]::round(((($_.FreeSpace / 1073741824)/($_.Capacity / 1073741824)) * 100),0)}}

function get-mountpoints {
$volumes = Get-WmiObject -computer $server win32_volume
$volumes | Select SystemName, Label, $TotalGB, $FreeGB, $FreePerc | Format-Table -AutoSize
}

$servers = (Get-Content .\servers.txt)

foreach ($server in $servers){
get-mountpoints
}
