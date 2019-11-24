param(
[Parameter(Mandatory=$true)][string]$Username
)

Get-ADUser -identity $Username -Properties Name,Title,Company,StreetAddress,City,State,PostalCode,OfficePhone,Mobile,Fax | 
Select-Object Name,Title,Company,StreetAddress,City,State,PostalCode,OfficePhone,Mobile,Fax |
Export-Csv -Path "$Username.csv" -NoTypeInformation