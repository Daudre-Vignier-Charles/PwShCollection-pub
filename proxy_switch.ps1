function Main {
    function Show-Notif
    {

      [CmdletBinding(SupportsShouldProcess = $true)]
      param
      (
        [Parameter(Mandatory=$true)]
        $Text,

        [Parameter(Mandatory=$true)]
        $Title,

        [ValidateSet('None', 'Info', 'Warning', 'Error')]
        $Icon = 'Info',
        $Timeout = 10000
      )

      Add-Type -AssemblyName System.Windows.Forms
	  
	  $proxyServer = "http://proxy.exemple.net:8080"

      if ($script:notif -eq $null)
      {
        $script:notif = New-Object System.Windows.Forms.NotifyIcon
      }

      $path                    = Get-Process -id $pid | Select-Object -ExpandProperty Path
      $notif.Icon            = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
      $notif.NotifIcon  = $Icon
      $notif.NotifText  = $Text
      $notif.NotifTitle = $Title
      $notif.Visible         = $true

      $notif.ShowNotif($Timeout)
    }
    $ProxyStatus = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\"Internet Settings" -Name ProxyEnable | Select-Object -ExpandProperty ProxyEnable
    $ProxyURL    = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\"Internet Settings" -Name ProxyServer | Select-Object -ExpandProperty ProxyServer
    $notifText = ""
    if ( $ProxyStatus -eq "1" ) {
        Set-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\"Internet Settings" -Name ProxyEnable -Value 0
        $notifText="Proxy désactivé"
    } else {
        Set-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\"Internet Settings" -Name ProxyEnable -Value 1
        $notifText="Proxy activé"
    }
    if ( $ProxyURL -ne $proxyServer ) {
        Set-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\"Internet Settings" -Name ProxyServer  -Value $proxyServer
        $notifText+=", URL du proxy rétablie"
    }
    Show-Notif -Title Système -Text $notifText -Timeout 1200
    Start-Sleep 5
    $Script:notif.Dispose()
}
main
