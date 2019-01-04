# Copyright 2018-2019 Charles Daudr�-Vignier <charles@daudre-vignier.fr>
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


# Avertissement : le fichier doit �tre encod� en windows1252

<# MESSAGE d'erreur:
powershell error
   explication de l'erreur + "erreur retourn�e par la commande ping de cmd"

H�te inconnu :
    Pas d'enregistrement DNS

Une erreur irr�cup�rable s�est produite lors d�une recherche sur la base de donn�es :
    Enregistrement DNS mais "impossible de joindre l�h�te de destination", g�n�ralement un VPN.

Probl�me avec une partie du filterspec ou avec l�ensemble du tampon providerspecific :
    Enregistrement DSN "mais dur�e de vie TTL expir�e lors du transit", g�n�ralement un VPN.

Erreur due � des ressources insuffisantes :
    L'ordinateur ne r�pond pas aux ping, "d�lai d�attente de la demande d�pass�".
    Soit l'utilisateur a r�cemment �teint ou mis en veille son ordinateur.
    Soit son ordinateur ou un pare-feu interm�diaire peut bloquer les pings.
    Son adresse IP est alors demand�e directement au DNS et imprim�e. #>

$ErrorGlpiBadStatut = "Attention, l'ordinateur {0} est d�clar� en {1} !"
$ErrorGlpiBadUser = "Attention, l'utilisateur r�el de l'ordinateur {0} n'est pas celui enregistr� dans GLPI !"
$ErrorGlpiNoUser = "Attention, l'ordinateur {0} n'a pas d'utilisateur enregistr� dans GLPI !"
$ErrorGlpiNoSerial = " |
 | Attention, un ordinateur a �t� trouv� sans num�ro de s�rie !
 | Cet ordinateur ne sera pas trait� sans num�ro de s�rie.
 |"
$ErrorNoFile = "Erreur fatale, aucun fichier s�lectionn�."

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
    if ( $Computer.'Num�ro de s�rie' ) {
        $NewRow = [pscustomobject][ordered] @{
            'Num�ro de s�rie' = $Computer.'Num�ro de s�rie'
            'Statut GLPI' = $Computer.Statut
            'Utilis�' = ""
            'Utilisateur GLPI' = $Computer.Utilisateur
            'Utilisateur1' = ""
            'u1' = ""
            'Utilisateur2' = ""
            'u2' = ""
            'Utilisateur3' = ""
            'u3' = ""
            'IP' = ""
            'R�seau' = ""
            'Immobilisation' = $Computer.'Informations financi�res et administratives - Num�ro d''immobilisation'
        }
        $ID = $Computer.'Num�ro de s�rie' ;
        $ADComputer = Get-ADComputer -Properties * -Filter {CN -eq "$($Computer.'Num�ro de s�rie')"} ;
        if ( $ADComputer.IPv4Address ) {
            $NewRow.'Utilis�' = "OUI" ;
            $UserSubnet = Get-IPSubnet $ADComputer.IPv4Address $Mask ;
            $NewRow.IP = $ADComputer.IPv4Address ;
            $NewRow.'R�seau' = $IPSubnets[$UserSubnet] ;
            Write-Host ( "{0} est connect� ou a r�cemment �t� connect� avec l'adresse IP {1}" -f $Computer.'Num�ro de s�rie', $ADComputer.IPv4Address) ;
            Write-Host ( "Son adresse IP appartient au r�seau {0}" -f $IPSubnets[$UserSubnet] ) ;
            Write-Host ( "    Tentative de ping de {0} :" -f $Computer.'Num�ro de s�rie' ) ;
            $Ret = Test-Connection -Count 1 $Computer.'Num�ro de s�rie' -ErrorVariable PingError -ErrorAction Continue 2>$null ;
            if ( $Ret ) {
                Write-Host ( "    " ) -NoNewline ;
                Write-Host ( "{0} est actuellement connect�." -f $Computer.'Num�ro de s�rie' ) -ForegroundColor black -BackgroundColor cyan ;
                [array] $UserList = Get-ChildItem \\$ADComputer.IPv4Address\c$\Users 2>$NULL | Sort-Object -Descending -Property LastWriteTime ;
                if ( $Computer.Statut | Where-Object { $_ -in $StatutsUnused } ) {
                    Write-Host ( $ErrorGlpiBadStatut -f $Computer.'Num�ro de s�rie', $Computer.Statut ) -ForegroundColor red -BackgroundColor black ;
                }
                if ( $Computer.Utilisateur ) {
                    $Utilisateur = $Computer.Utilisateur -split " " ;
                    $Utilisateur = $Utilisateur[1][0] + $Utilisateur[0] ;
                    if ( $UserList -AND $Utilisateur -ne $UserList[0].Name ) {
                        Write-Host ( $ErrorGlpiBadUser -f $Computer.'Num�ro de s�rie' ) -ForegroundColor red -BackgroundColor black ;
                        Write-Host ( "    Utilisateur r�el      : {0}" -f $UserList[0] ) ;
                        Write-Host ( "    Utilisateur dans GLPI : {0}" -f $Utilisateur ) ;
                    }
                }
                else {
                    Write-Host ( $ErrorGlpiNoUser -f $Computer.'Num�ro de s�rie' ) -ForegroundColor red -BackgroundColor black ;
                }
                if ( $UserList ) {
                    $NewRow.Utilisateur1 = $UserList[0] ;
                    $NewRow.u1 = $UserList[0].LastAccessTime ;
                    $NewRow.Utilisateur2 = $UserList[1] ;
                    $NewRow.u2 = $UserList[0].LastAccessTime ;
                    $NewRow.Utilisateur3 = $UserList[2] ;
                    $NewRow.u3 = $UserList[0].LastAccessTime ;
                    Write-Host ( "    Les trois dernier utilisateurs de l'ordinateur {0} sont :" -f $Computer.'Num�ro de s�rie') ;
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
            $NewRow.Utilis� = "NON" ;
            Write-Host ( "{0} n'est pas enregistr� dans les DNS et n'est pas connect�." -f $Computer.'Num�ro de s�rie') ;
        }
    }
    else {
        Write-Host $ErrorGlpiNoSerial -ForegroundColor red -BackgroundColor black ;
    }
    Write-Host ( $Separator ) ;
    $PopulateCSV += $NewRow ;
}

Write-Host ( "S�lectionnez un fichier pour sauvergarder les r�sultats." ) ;
Read-Host -Prompt "Appuyer sur entrer pour continuer " ;
$AskForSave = $true ;
while ( $AskForSave -eq $true ) {
    $CSVExportPath = Get-FileName "C:\" "save" ;
    if ( -Not $CSVExportPath ) {
        Write-Host $ErrorNoFile -ForegroundColor Black -BackgroundColor Red ;
        Write-Host "Voulez vous enregistrer le fichier de r�sultat ? Y/N" ;
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
