Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)

Add-Type -AssemblyName System.IO

$expired = Search-ADAccount -AccountExpired|Select-Object -Property Name
$disabled = Search-ADAccount -AccountDisabled|Select-Object -Property Name
$locked = Search-ADAccount -LockedOut|Select-Object -Property Name
$pexpired = Search-ADAccount -PasswordExpired|Select-Object -Property Name

$path = "{0}\ADUsrStatus_{1}.txt" -f $env:TEMP,(Get-Date -f yyyy-MM-dd-hh-mm)

$stream =  [System.IO.StreamWriter] $path
$stream.WriteLine("Expired account :")
$stream.WriteLine("---------------")
$expired | ForEach-Object {
    $stream.WriteLine($_.Name)
}
$stream.WriteLine("`r`nDisabled account :")
$stream.WriteLine("----------------")
$disabled | ForEach-Object {
    $stream.WriteLine($_.Name)
}
$stream.WriteLine("`r`nLocked account :")
$stream.WriteLine("--------------")
$locked | ForEach-Object {
    $stream.WriteLine($_.Name)
}
$stream.WriteLine("`r`nExpired password account :")
$stream.WriteLine("------------------------")
$pexpired | ForEach-Object {
    $stream.WriteLine($_.Name)
}
$stream.Close()
Start-Process "notepad.exe" -Wait -ArgumentList $path
Remove-Item -Path $path