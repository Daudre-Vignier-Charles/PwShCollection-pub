Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Windows.Forms

$path = "{0}\PasswordMustChange.txt" -f $env:TEMP

$validityPeriod = 90 #days
$searchBase = "OU=etc,OU=AnotherOne,OU=AnOU,DC=EXEMPLE,DC=NET"
$noAccessGroup = "CN=noAccess,CN=Users,DC=EXEMPLE,DC=NET"

$ALL=Get-ADUser  -SearchBase $searchBase `
            -Properties memberof,displayName,PasswordLastSet `
            -Filter {
                Memberof -ne $noAccessGroup -and
                PasswordExpired -eq $false -and
                PasswordNeverExpires -eq $false -and
                Enabled -eq $true
            } |
ForEach-Object {
    if ( ! $_.LockedOUt -and $_.PasswordLastSet -ne $null ) {
        $delay=((($_.PasswordLastSet).AddDays($validityPeriod)).Subtract((Get-Date))).Days
        if ( $delay -le 5 -and $delay -gt 0 ) {
            $data = [PSCustomObject]@{
                joursRestants = [int]$delay
                nom = [string]$_.displayName
            }
            return $data
        }
    }
} | Sort-Object joursRestants | Format-Table -Expand Both -AutoSize @{L='Nom';E={$_.nom}},@{L='Jours restants';E={$_.joursRestants};Alignment="left"} | Out-String

$stream =  [System.IO.StreamWriter] $path
$stream.Write($ALL)
$stream.Close()
Start-Process "notepad.exe" -Wait -ArgumentList $path
Remove-Item -Path $path
