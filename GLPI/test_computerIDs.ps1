# Copyright 2018-2019 Charles Daudré-Vignier <charles@daudre-vignier.fr>
# All rights reserved.
<#
  The author provide the program 'as is' without warranty of any kind,
  expressed or implied, including, but not limited to,
  the implied warranties of merchantability and fitness for a particular purpose.
#>
<#
  The author grants unlimited rights of use, execution,
  modification and copying to DSOgroup and any of its affiliates
  provided that the original copyright and legal notices are included.
#>


# Avertissement : le fichier doit être encodé en windows1252

<# MESSAGE d'erreur:
powershell error
   explication de l'erreur + "erreur retournée par la commande ping de cmd"

Hôte inconnu :
    Pas d'enregistrement DNS

Une erreur irrécupérable s’est produite lors d’une recherche sur la base de données :
    Enregistrement DNS mais "impossible de joindre l’hôte de destination", généralement un VPN.

Problème avec une partie du filterspec ou avec l’ensemble du tampon providerspecific :
    Enregistrement DSN "mais durée de vie TTL expirée lors du transit", généralement un VPN.

Erreur due à des ressources insuffisantes :
    L'ordinateur ne répond pas aux ping, "délai d’attente de la demande dépassé".
    Soit l'utilisateur a récemment éteint ou mis en veille son ordinateur.
    Soit son ordinateur ou un pare-feu intermédiaire peut bloquer les pings.
    Son adresse IP est alors demandée directement au DNS et imprimée. #>

$ErrorGlpiBadStatut = "Attention, l'ordinateur {0} est déclaré en {1} !"
$ErrorGlpiBadUser = "Attention, l'utilisateur réel de l'ordinateur {0} n'est pas celui enregistré dans GLPI !"
$ErrorGlpiNoUser = "Attention, l'ordinateur {0} n'a pas d'utilisateur enregistré dans GLPI !"
$ErrorGlpiNoSerial = " |
 | Attention, un ordinateur a été trouvé sans numéro de série !
 | Cet ordinateur ne sera pas traité sans numéro de série.
 |"
$ErrorNoFile = "Erreur fatale, aucun fichier sélectionné."

$Separator = "--------------------------------------------------------------------------------------------------------------------------------"

$PopulateCSV = @(
)

$StatutsUnused = @(
    "En spare",
    "En configuration",
    "En recherche",
    "En sortie d'immobilisation",
    "En sortie de parc",
    "Mis au rebut"
)

$StatutsUsed = @(
    "Domicile",
    "En fonctionnement"
)

$Mask = "255.255.255.0"

$IPSubnets = @{
    "169.254.0.0" = "demo Network"
}

# Function ask with GUI for a CSV file, get this one from GLPI.
Function Get-FileName($initialDirectory, $type) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null ;
    if ( $type -eq "open" ) {
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog ;
    }
    if ( $type -eq "save" ) {
        $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog ;
    }
    $OpenFileDialog.initialDirectory = $initialDirectory ;
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv" ;
    $OpenFileDialog.ShowDialog() | Out-Null ;
    $OpenFileDialog.filename ;
}

# Function return subnet from ip and mask
Function Get-IPSubnet($IP, $Mask) {
    [Net.IPAddress] $IP = $IP;
    [Net.IPAddress] $Mask = $Mask ;
    [Net.IPAddress] $Subnet = $IP.Address -band $Mask.Address ;
    $Subnet.IPAddressToString ;
}

Clear-Host ;

$CSVFilePath = Get-FileName "C:\" "open" ;
if ( -Not $CSVFilePath ) {
    Write-Host $ErrorNoFile -ForegroundColor Black -BackgroundColor Red ;
    Read-Host -Prompt "Appuyez sur entrer pour quitter " ;
    return
}
Write-Host ( "Ouverture de " + $CSVFilePath + "`n`n" ) ;
$CSVFile = Import-Csv -Delimiter ";" -Encoding UTF8 $CSVFilePath ;

Write-Host ( "Test des IDs :`n`n" ) ;
# Ping each of IDs and ask DSN for IP if Error
foreach ( $Computer in $CSVFile ) {
    if ( $Computer.'Numéro de série' ) {
        $NewRow = [pscustomobject][ordered] @{
            'Numéro de série' = $Computer.'Numéro de série'
            'Statut GLPI' = $Computer.Statut
            'Utilisé' = ""
            'Utilisateur GLPI' = $Computer.Utilisateur
            'Utilisateur1' = ""
            'u1' = ""
            'Utilisateur2' = ""
            'u2' = ""
            'Utilisateur3' = ""
            'u3' = ""
            'IP' = ""
            'Réseau' = ""
            'Immobilisation' = $Computer.'Informations financières et administratives - Numéro d''immobilisation'
        }
        $ID = $Computer.'Numéro de série' ;
        $ADComputer = Get-ADComputer -Properties * -Filter {CN -eq "$($Computer.'Numéro de série')"} ;
        if ( $ADComputer.IPv4Address ) {
            $NewRow.'Utilisé' = "OUI" ;
            $UserSubnet = Get-IPSubnet $ADComputer.IPv4Address $Mask ;
            $NewRow.IP = $ADComputer.IPv4Address ;
            $NewRow.'Réseau' = $IPSubnets[$UserSubnet] ;
            Write-Host ( "{0} est connecté ou a récemment été connecté avec l'adresse IP {1}" -f $Computer.'Numéro de série', $ADComputer.IPv4Address) ;
            Write-Host ( "Son adresse IP appartient au réseau {0}" -f $IPSubnets[$UserSubnet] ) ;
            Write-Host ( "    Tentative de ping de {0} :" -f $Computer.'Numéro de série' ) ;
            $Ret = Test-Connection -Count 1 $Computer.'Numéro de série' -ErrorVariable PingError -ErrorAction Continue 2>$null ;
            if ( $Ret ) {
                Write-Host ( "    " ) -NoNewline ;
                Write-Host ( "{0} est actuellement connecté." -f $Computer.'Numéro de série' ) -ForegroundColor black -BackgroundColor cyan ;
                [array] $UserList = Get-ChildItem \\$ADComputer.IPv4Address\c$\Users 2>$NULL | Sort-Object -Descending -Property LastWriteTime ;
                if ( $Computer.Statut | Where-Object { $_ -in $StatutsUnused } ) {
                    Write-Host ( $ErrorGlpiBadStatut -f $Computer.'Numéro de série', $Computer.Statut ) -ForegroundColor red -BackgroundColor black ;
                }
                if ( $Computer.Utilisateur ) {
                    $Utilisateur = $Computer.Utilisateur -split " " ;
                    $Utilisateur = $Utilisateur[1][0] + $Utilisateur[0] ;
                    if ( $UserList -AND $Utilisateur -ne $UserList[0].Name ) {
                        Write-Host ( $ErrorGlpiBadUser -f $Computer.'Numéro de série' ) -ForegroundColor red -BackgroundColor black ;
                        Write-Host ( "    Utilisateur réel      : {0}" -f $UserList[0] ) ;
                        Write-Host ( "    Utilisateur dans GLPI : {0}" -f $Utilisateur ) ;
                    }
                }
                else {
                    Write-Host ( $ErrorGlpiNoUser -f $Computer.'Numéro de série' ) -ForegroundColor red -BackgroundColor black ;
                }
                if ( $UserList ) {
                    $NewRow.Utilisateur1 = $UserList[0] ;
                    $NewRow.u1 = $UserList[0].LastAccessTime ;
                    $NewRow.Utilisateur2 = $UserList[1] ;
                    $NewRow.u2 = $UserList[0].LastAccessTime ;
                    $NewRow.Utilisateur3 = $UserList[2] ;
                    $NewRow.u3 = $UserList[0].LastAccessTime ;
                    Write-Host ( "    Les trois dernier utilisateurs de l'ordinateur {0} sont :" -f $Computer.'Numéro de série') ;
                    Write-Host ( "        {0} {1}" -f $UserList[0].LastAccessTime, $UserList[0] ) ;
                    Write-Host ( "        {0} {1}" -f $UserList[1].LastAccessTime, $UserList[1] ) ;
                    Write-Host ( "        {0} {1}" -f $UserList[2].LastAccessTime, $UserList[2] ) ;
                }
            }
            if ($PingError[0]) {
                $PingError = $PingError[0].ToString() ;
                Write-Host ( "    " ) -NoNewline ;
                Write-Host ( $PingError ) -ForegroundColor black -BackgroundColor cyan ;
            }
        }
        else {
            $NewRow.Utilisé = "NON" ;
            Write-Host ( "{0} n'est pas enregistré dans les DNS et n'est pas connecté." -f $Computer.'Numéro de série') ;
        }
    }
    else {
        Write-Host $ErrorGlpiNoSerial -ForegroundColor red -BackgroundColor black ;
    }
    Write-Host ( $Separator ) ;
    $PopulateCSV += $NewRow ;
}

Write-Host ( "Sélectionnez un fichier pour sauvergarder les résultats." ) ;
Read-Host -Prompt "Appuyer sur entrer pour continuer " ;
$AskForSave = $true ;
while ( $AskForSave -eq $true ) {
    $CSVExportPath = Get-FileName "C:\" "save" ;
    if ( -Not $CSVExportPath ) {
        Write-Host $ErrorNoFile -ForegroundColor Black -BackgroundColor Red ;
        Write-Host "Voulez vous enregistrer le fichier de résultat ? Y/N" ;
        $key = $Host.UI.RawUI.ReadKey()
        if ( $key.Character -eq "y") {
            $AskForSave = $true ;
        }
        elseif ( $key.Character -eq "n") {
            $AskForSave = $false ;
        }
    }
    else {
        $AskForSave = $false ;
    }
}
if ( $CSVExportPath ) {
    $PopulateCSV | Export-Csv $CSVExportPath -NoTypeInformation -Delimiter ";" -Encoding UTF8 ;
}
