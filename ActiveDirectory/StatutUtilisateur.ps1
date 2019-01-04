Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)

function UserSelector(){
    Begin {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "User selector"
        $form.AutoSize = $true
        $form.AutoSizeMode = "GrowAndShrink" # OR GrowOnly
        $form.StartPosition = "CenterScreen"

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.AutoSize = $true
        $listBox.Font = New-Object System.Drawing.Font("Tahoma", 10)

        [void]$script:result

        [void] $listBox.add_MouseDoubleClick({
            $script:result = $listBox.SelectedIndex
            $form.DialogResult = "OK" 
            $form.Close()
            return
        })
    }
    Process{
        [void] $listBox.Items.Add($_.displayName + " (" + $_.sAMAccountName + ")")
    }
    End{
        if ($listBox.Items.Count -eq 1) {
            return 0
        }
        [void] $form.Controls.Add($listBox)
        $ret = $form.ShowDialog()
        if ( $ret -eq "OK" ) {
            return $script:result
        } elseif ($ret -eq "Cancel" ) {
            return -1
        }
    }
}

$title = 'Statut utilisateur'
$msg   = "Identifiant de l'utilisateur:"
$noAccessGroup = ""


[string] $user = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
if ($user -eq "") {
    exit
}
[string] $likeUser = "*{0}*" -f $user
$UserExist = Get-ADUser -Filter {sAMAccountName -eq $user}
If ($UserExist -eq $Null){
    $users = Get-ADUser -Properties "memberof","sAMAccountName","displayName" -Filter { (sAMAccountName -like $likeUser) -or
                                                                                        (displayName -like $likeuser) -and
                                                                                        (Memberof -ne $noAccessGroup)} | Select-Object "sAMAccountName","displayName"
    if ($users.Length -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("L'utilisateur {0} n'existe pas." -f $user,'Status utilisateur : {0}' -f $user,'OK')
        exit
    } else {
        [int]$uindex = $users | UserSelector
        if ( $uindex -eq -1 ) {
            exit
        }
        $user = $users[$uindex].sAMAccountName
    }
}
# LockOut Status for "user" :
$result = Get-ADUser $user -Properties * | Select-Object -Property CN, Enabled, LockedOut,PasswordExpired, BadPwdCount, LastBadPasswordAttempt, PasswordLastSet, Deleted
if ($result.LockedOut -eq $True){
    $result.LockedOut = "oui"
} else {
    $result.LockedOut = "non"
}
if ($result.PasswordExpired -eq $True){
    $result.PasswordExpired = "oui"
} else {
    $result.PasswordExpired = "non"
}
if ($result.Enabled -eq $True){
    $result.Enabled = "non"
} else {
    $result.Enabled = "oui"
}
[string] $end_result = @"
Compte désactivé`t`t: {0}
Compte verrouillé`t`t: {1}
Mot de passe expiré`t: {2}
Nombre de mauvais mdp`t: {3}
Dernier échec`t`t: {4}
Date de MAJ du MDP`t: {5}
"@ -f $result.Enabled, $result.LockedOut.ToString(), $result.PasswordExpired.ToString(),
    $result.BadPwdCount.ToString(), $result.LastBadPasswordAttempt.Datetime,
    $result.PasswordLastSet.Datetime
[System.Windows.Forms.MessageBox]::Show($end_result,'Status utilisateur : {0}' -f $user,'OK')