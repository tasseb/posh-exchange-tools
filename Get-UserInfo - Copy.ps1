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
    L'utilisation de le l'option -Brut retourne un résultat non formaté, sinon le résultat est retourné formaté
.PARAMETER User
   Le paramètre User est indispensable et peut contenir plusieurs valeurs.
   Les valeurs acceptés sont soit le nom de l'utilisateur (Display Name), ou l'alias AD.
.EXAMPLE
   Get-UserInfo a.nonyme
   Retourne les informations de l'utilisateur dont l'alias est a.nonyme
.EXAMPLE
   Get-UserInfo "Albert Nonyme"
   Retourne les informations de l'utilisateur dont le Display Name est "Albert Nonyme"
.EXAMPLE
   Get-UserInfo a.nonyme,n.ainconnu
   Retourne les informations des utilisateurs dont les alias sont a.nonyme et n.ainconnu
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

        # Brut - switch permettant de formater la sortie de manière brut (objet non formaté), sinon le résultat est affiché comme un rapport complet des informations récupérées sur l'utilisateur
        [Switch]$Brut
    )

    Begin{}
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

            If (!$Brut) {
                $textcolor = "Green"
                Write-Host ""
                Write-Host -ForegroundColor $textcolor "-------------------- Informations utilisateur $chaine ----------------------------------"
                $objUserInfo.Utilisateur
                Write-Host ""
                Write-Host -ForegroundColor $textcolor "-------------------- L'utilisateur $chaine est membre des groupes AD suivants ----------------------------------"
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
            Else {
                Write-Output $objUserInfo
            }

        }

    }
    End{}
}