.\Get-FileShareInventory-WithLastAccess.ps1 -MappingCsv ".\Config\FileShareMapping.csv" -IncludeFileDetail
.\Get-FileSharePermissions.ps1 -MappingCsv ".\Config\FileShareMapping.csv" -ThrottleLimit 1
.\Get-BlockedExtensions.ps1 -MappingCsv ".\Config\FileShareMapping.csv"
.\Get-DuplicateFiles.ps1 -MappingCsv ".\Config\FileShareMapping.csv"
