.\Get-FileShareInventory-WithLastAccess.ps1 -MappingCsv ".\Config\FileShareMappingByOU_Resume.csv" -IncludeFileDetail
.\Get-FileSharePermissions.ps1 -MappingCsv ".\Config\FileShareMappingByOU_Resume.csv" -ThrottleLimit 1
.\Get-BlockedExtensions.ps1 -MappingCsv ".\Config\FileShareMappingByOU_Resume.csv"
.\Get-DuplicateFiles.ps1 -MappingCsv ".\Config\FileShareMappingByOU_Resume.csv"
.\Get-InvalidCharacters.ps1 -MappingCsv ".\Config\FileShareMappingByOU_Resume.csv"
.\Get-PathTooLong.ps1 -MappingCsv ".\Config\FileShareMappingByOU_Resume.csv" -LongueurMax 400