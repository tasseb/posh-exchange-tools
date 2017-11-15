<#
.Synopsis
   Cette fonction permet d'obtenir des informations sur le ou les utilisateurs passés en paramètres
.DESCRIPTION
   Cette fonction permet d'obtenir des informations sur le ou les utilisateurs passés en paramètres
   Les informations retournées sont les suivantes:
        * Objet Utilisateur  = Nom, Prénom, Display Name, OU de l'utilisateur AD, extension téléphonique
        * Objet AdresseEmail = Liste des adresses emails et alias de l'utilisateur
        * Objet EmailDelegation = Liste des adresses emails sur lesquelle sl'utilisateur possède une délégation
        * Objet Ordinateur = Liste des ordinateurs dont la description de l'Objet AD contient le nom d'utilisateur
        * Objet Smartphone = Liste des smartphones et tablettes de l'utilisateur connectés et synchronisés à Exchange

    L'utilisation de l'option -Filtre permet de sélectionner une partie des informations au lieu d'afficher tout.
    Les différentes valeurs acceptées sont User | Group | Email | SharedBal | Computer | Mobile
        * Filtre User affiche les informations du compte AD de l'utilisateur
        * Filtre Group affiche les groupes Active Directory auquel l'utilisateur appartient
        * Filtre Email affiche l'adresse email principale ainsi que les "alias" de l'utilisateur
        * Filtre SharedBal affiche les adresses emails que l'utilisateur a en délégation
        * Filtre Computer affiche les informations des ordinateurs liés aux comptes utilisateurs (dont l'objet AD contient le nom d'utilisateur)
        * Filtre Mobile affiche les informations des appareils mobiles liés au compte Exchange de l'utilisateur

    L'utilisation de l'option -Brut retourne un résultat non formaté, sinon le résultat est retourné formaté

.PARAMETER User
   Le paramètre User est indispensable et peut contenir plusieurs valeurs.
   Les valeurs acceptés sont soit le nom de l'utilisateur (Display Name), ou l'alias AD.
.PARAMETER Filtre
    L'utilisation de l'option -Filtre permet de sélectionner une partie des informations au lieu d'afficher tout.
    Les différentes valeurs acceptées sont User | Group | Email | SharedBal | Computer | Mobile
.PARAMETER Brut
    Lorsqu'il est utilisé, affiche les informations de manière brut (sans formatage). Peut-être utile pour un traitement des informations par une autre commande.
.EXAMPLE
   Get-UserInfo a.nonyme
   Retourne les informations de l'utilisateur dont l'alias est a.nonyme
.EXAMPLE
   Get-UserInfo "Albert Nonyme"
   Retourne les informations de l'utilisateur dont le Display Name est "Albert Nonyme"
.EXAMPLE
   Get-UserInfo a.nonyme,n.ainconnu
   Retourne les informations des utilisateurs dont les alias sont a.nonyme et n.ainconnu
.EXAMPLE
   Get-UserInfo a.nonyme,n,ainconnu -Filter User,Group,Mobile
   Retourne les informations Utilisateurs, appartenance aux groupes AD, information sur les appareils mobiles des utilisateurs dont les alias sont a.nonyme et n.ainconnu
#>
function Get-UserInfo
{
    [CmdletBinding()]
    Param
    (
        # User - supporte les valeurs multiples (plusieurs utilisateurs passés en paramètre)
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String[]]$User,
        # Filtre - permet de filtrer le rapport par défaut en sélectionnant un ou plusieurs critères (User,Group,Email,SharedBal,Computer,Mobile)
        [String[]]$Filtre,
        # Brut - switch permettant de formater la sortie de manière brut (objet non formaté), sinon le résultat est affiché comme un rapport complet des informations récupérées sur l'utilisateur
        [Switch]$Brut
    )

    Begin{
            # La variable $textcolor permet de changer la couleur des différents titres des valeurs retournées (purement "esthétique"). Vert par défaut.
            $textcolor = "Green"
    }
    Process
    {
        #Boucle foreach pour chacun des utilisateurs passés en paramètre
        foreach ($chaine in $User) {
            #Récupération des propriétés Name et Alias de l'utilisateur pour les passer en paramètre aux futures commandes (permet de prendre en paramètre en DisplayName ou un Alias)
            $u = Get-Mailbox $chaine | Select-Object Name,Alias
            
            $GetUser           = Get-User -Identity $u.Alias | Select-Object -Property SamAccountName,Sid,OrganizationalUnit,Displayname,FirstName,LastName,Phone
            $GetGroupe         = Get-ADPrincipalGroupMembership $u.Alias | Sort-Object Name | Select-Object -ExpandProperty Name
            $GetEmailAddresses = Get-Mailbox -Identity $u.Alias | Select-Object -ExpandProperty EmailAddresses
            $GetSharedBal      = Get-Mailbox | Get-MailboxPermission -User $u.Alias | Select-Object Identity
            $GetCompInfo       = Get-ADComputer -Filter * -Properties CanonicalName,Description,IPV4Address,LastLogonDate,OperatingSystem | Where-Object {$_.Description -like ('*{0}*' -f $u.Alias)} | Select-Object -Property Name,CanonicalName,Description,IPV4Address,LastLogonDate,OperatingSystem
            $GetMobile         = Get-MobileDevice | Where-Object {$_.UserDisplayName -like $u.Name} | Select-Object -Property FriendlyName,DeviceOS,DeviceType

            #Construction d'un objet PSCustomObject contenant toutes les propriétés précédemment récupérées
            $UserInfo = [ordered]@{
                'Utilisateur' = $GetUser;
                'ADGroupe' = $GetGroupe;
                'AdresseEmail' = $GetEmailAddresses;
                'EmailDelegation' = Split-Path $GetSharedBal.Identity -Leaf;
                'Ordinateur' = $GetCompInfo;
                'Smartphone' = $GetMobile
            }
            $objUserInfo = New-Object -TypeName PSCustomObject -Property $UserInfo

            If ($Brut) {
                Write-Output $objUserInfo
                
            }
            Else {
                If ($Filtre) {
                    foreach ($valeur in $Filtre) {
                        Switch ($valeur) {
                            "User" {
                                   Write-Host ""
                                   Write-Host -ForegroundColor $textcolor "-------------------- Informations utilisateur $chaine ----------------------------------"
                                   $objUserInfo.Utilisateur
                                   }
                            "Group" {
                                    Write-Host ""
                                    Write-Host -ForegroundColor $textcolor "-------------------- L'utilisateur $chaine est membre des groupes AD suivants ----------"
                                    $objUserInfo.ADGroupe
                                    }  
                            "Email" {
                                    Write-Host ""
                                    Write-Host -ForegroundColor $textcolor "-------------------- Alias emails de $chaine -------------------------------------------"
                                    $objUserInfo.AdresseEmail
                                    }
                            "SharedBal" {
                                        Write-Host ""
                                        Write-Host -ForegroundColor $textcolor "-------------------- Boîtes emails en délégation pour $chaine --------------------------"
                                        $objUserInfo.EmailDelegation
                                        }
                            "Computer" {
                                       Write-Host ""
                                       Write-Host -ForegroundColor $textcolor "-------------------- Informations ordinateur(s) affectés à $chaine ---------------------"
                                       $objUserInfo.Ordinateur 
                                       }
                            "Mobile" {
                                        Write-Host ""
                                        Write-Host -ForegroundColor $textcolor "-------------------- Informations appareil(s) mobile(s) de $chaine ---------------------"
                                        $objUserInfo.Smartphone
                                     }
                            default {Write-Warning "Ce paramètre n'est pas pris en compte. Seuls User | Group | Email | SharedBal | Computer | Mobile sont acceptés"}
                        }
                    }
                }
                Else {
                    
                    Write-Host ""
                    Write-Host -ForegroundColor $textcolor "-------------------- Informations utilisateur $chaine ----------------------------------"
                    $objUserInfo.Utilisateur
                    Write-Host ""
                    Write-Host -ForegroundColor $textcolor "-------------------- L'utilisateur $chaine est membre des groupes AD suivants ----------"
                    $objUserInfo.ADGroupe
                    Write-Host ""
                    Write-Host -ForegroundColor $textcolor "-------------------- Alias emails de $chaine -------------------------------------------"
                    $objUserInfo.AdresseEmail
                    Write-Host ""
                    Write-Host -ForegroundColor $textcolor "-------------------- Boîtes emails en délégation pour $chaine --------------------------"
                    $objUserInfo.EmailDelegation
                    Write-Host ""
                    Write-Host -ForegroundColor $textcolor "-------------------- Informations ordinateur(s) affectés à $chaine ---------------------"
                    $objUserInfo.Ordinateur
                    Write-Host ""
                    Write-Host -ForegroundColor $textcolor "-------------------- Informations appareil(s) mobile(s) de $chaine ---------------------"
                    $objUserInfo.Smartphone                
                }
                
            }

        }

    }
    End{}
}