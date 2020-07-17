import-module activedirectory
$date = get-date -format yyyyMMddHHmmss
get-aduser -filter * -properties homedirectory | select-object samaccountname,homedirectory | export-csv "CSV Files\HomeDirectory-$date.csv" -notype
