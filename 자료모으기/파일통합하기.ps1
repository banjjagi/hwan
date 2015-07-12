cd D:\db\자료\작업중\백업
get-childItem *.csv -Recurse | Rename-Item -NewName { $_.Name -replace ".csv", '.csev' }
type *.csev >all.csv
