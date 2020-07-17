get-qaduser -IncludeAllProperties | select-object samAccountName,FirstName,LastName,DisplayName,ProfilePath | where-object {$_.ProfilePath -like "ro" } | export-csv D:\test.csv
