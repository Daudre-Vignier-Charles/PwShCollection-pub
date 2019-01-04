Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

Hide-Console

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$adServer = "FIRST.EXEMPLE.NET"

$title = 'OU Object Counter'
$msg   = "LDAP Path"

[string] $ldap_path = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

$objnum=Get-ADObject -LDAPFilter:"(objectClass=organizationalPerson)" -Properties:userPrincipalName -SearchBase:$ldap_path -SearchScope:"OneLevel" -Server:$adServer | measure

[System.Windows.Forms.MessageBox]::Show($objnum.Count.ToString(),"Nombre d'objets :" -f $user,'OK')