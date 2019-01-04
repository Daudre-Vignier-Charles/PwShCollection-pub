Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Windows.Forms

$path = "{0}\PswdExpire_{1}.txt" -f $env:TEMP, (Get-Date -f yyyy-MM-dd-hh-mm)

#region CONFIGURATION
$validityPeriod = 90 #days
$searchBase     = ""
$noAccessGroup  = ""
#endregion

#region PROCESSING
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
        $delay=((($_.PasswordLastSet).AddDays($validityPeriod)).Subtract((Get-Date)))
        if ( $delay.Days -ge 0 -and $delay.Days -le 4 ) {
            $data = [PSCustomObject]@{
                jours   = [int]$delay.Days
                heures  = [int]$delay.Hours
                minutes = [int]$delay.Minutes
                nom     = [string]$_.displayName}
            return $data
        }
        elseif ($delay.Days -lt 0) {
        $data = [PSCustomObject]@{
                jours   = [string]"--"
                heures  = [string]"--"
                minutes = [string]"--"
                nom     = [string]$_.displayName}
            return $data
        }
    }
} |
Sort-Object jours,heures,minutes |
Format-Table -Expand Both `
             -AutoSize @{L='Nom';E={$_.nom}},
                       @{L='Jours';E={$_.jours};Alignment="left"},
                       @{L='Heures';E={$_.heures};Alignment="left"},
                       @{L='Minutes';E={$_.minutes};Alignment="left"} |
Out-String
#endregion

#region RAPPORT
try
{
    $stream =  [System.IO.StreamWriter] $path
    $stream.Write($ALL)
    $stream.Close()
    Start-Process "notepad.exe" -Wait -ArgumentList $path
}
finally
{
    Remove-Item -Path $path
}
#endregion
