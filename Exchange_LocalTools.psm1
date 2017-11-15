#region Import-PSSession sur serveur Exchange
#Préalable à l'utilisation des commandes contenues dans ce module = Importer le lot de commande Exchange en provenance du serveur de messagerie
    #Récupération du nom de serveur
    $Serveur = Read-Host -Prompt "Entrer le nom du serveur Exchange"
    #Création de la PSSession Microsoft.Exchange
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Serveur/PowerShell/" -Authentication Kerberos -WarningAction SilentlyContinue
    #Importation des commandes
    Import-PSSession $Session -WarningAction SilentlyContinue
#endregion

<#
.Synopsis
   Cette fonction permet d'ajouter des délégations sur des boîtes email Exchange
.DESCRIPTION
   Cette fonction permet d'ajouter des délégations sur des boîtes email Exchange avec un accès complet à la boîte, et la possibilité d'envoyer des emails depuis cette boîte.
   ATTENTION: l'utilisation des délégations est utile pour des boîtes emails génériques, mais n'est pas recommandée sur des boîtes d'utilisateur (pour des raisons légales notamment).
.PARAMETER Target 
   Ce paramètre est obligatoire et correspond à l'alias ou le nom de la boîte email (générique) que l'on souhaite déléguer à un utilisateur
.PARAMETER User
   Ce paramètre est obligatoire et correspond à l'alias ou le nom de l'utilisateur à qui l'on ajoute une délégation
.EXAMPLE
   Add-Delegation -Target contact -User jean.phumune
   Ajout d'une délégation sur la boîte contact pour l'utilisateur dont l'alias est jean.phumune
.EXAMPLE
   Add-Delegation -Target "Contact Mailbox" -User "Jean Phumune"
   Ajout d'une délégation sur la boîte contact pour l'utilisateur "Jean Phumune"
#>
function Add-Delegation
{
    [CmdletBinding()]
    Param
    (
        #Alias ou nom de la boîte email (générique) que nous souhaitons déléguer
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$Target,

        #Alias ou nom de la boîte email (utilisateur) à qui nous allons ajouter une délégation
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string]$User
    )

    Begin
    {
        Write-Verbose "Récupération des propriétés Name des valeurs retournées"
        $TargetName = Get-Mailbox $Target
        $UserName = Get-Mailbox $User
        Write-Verbose "Ajout d'une délégation de la boîte $Target pour l'utilisateur $User"
    }
    Process
    {
        #Full Access
        Write-Verbose "Ajout du full access sur la boîte $Target pour l'utilisateur $User"
        Add-MailboxPermission $TargetName.Name -User $UserName.Name -AccessRights FullAccess -InheritanceType All -AutoMapping $true

        #Send As
        Write-Verbose "Ajout du send as sur la boîte $Target pour l'utilisateur $User"
        Add-ADPermission -Identity $TargetName.Name -User $UserName.Name -AccessRights ExtendedRight -ExtendedRights "Send As"
    }
    End
    {
        Write-Information "-------------------------------------------------------------------"

        #Nettoyage des variables paramètres de la fonction
        $Target = $null
        $User = $null
    }
}
<#
.Synopsis
   Cette fonction permet d'ajouter des délégations sur des boîtes email Exchange à l'aide d'un fichier CSV passé en paramètre.
.DESCRIPTION
   Cette fonction permet d'ajouter des délégations sur des boîtes email Exchange avec un accès complet à la boîte, et la possibilité d'envoyer des emails depuis cette boîte.
   Un fichier CSV est passé en paramètre sous la forme "Email à déléguer;Utilisateur bénéficiaire;".
   ATTENTION: l'utilisation des délégations est utile pour des boîtes emails génériques, mais n'est pas recommandée sur des boîtes d'utilisateur (pour des raisons légales notamment).
.PARAMETER Path
   Le paramètre Path correspond au chemin complet vers le fichier CSV à importer. S'il n'est pas précisé, une boîte de dialogue sera ouverte pour la sélection du fichier.
.EXAMPLE
   Add-DelegationCSV -Path "C:\fichier.csv"
   
   Le nom du fichier CSV n'a pas d'importance, son emplacement non plus tant que le chemin est atteignable par l'ordinateur qui lance la commande
   Le format du fichier CSV est sans en-tête, sous le format suivant: "Nom ou alias BAL générique;Nom ou alias du bénéficiare"
   Exemple de délégation de la boîte dont l'alias est contact aux utilisateurs Jean Phumune et Aurore Mahler:
   contact;Jean Phumune;
   contact;Aurore Mahler;

#>
function Add-DelegationCSV
{
    [CmdletBinding()]
    Param
    (
        #Path - chemin du fichier csv contenant la liste des boîtes à déléguer avec leur bénéficiaire
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$Path

    )

    Begin
    {
        #Si pas de fichier CSV précisé dans la commande, ouverture de la fenêtre de sélection
        if (!$Path) {
            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.Filter = "CSV (*csv)| *.csv"
            $OpenFileDialog.ShowDialog() | Out-Null
            $Path = $OpenFileDialog.FileName
        }
    }
    Process
    {
        #Import du fichier CSV et ajout des délégations pour chaque ligne
        Import-CSV $Path -Header MailboxName,UserName -Delimiter ";" | Foreach-Object{
            Write-Host "BAL:" $_.MailboxName "| UserName:" $_.UserName
            Write-Host""
            Add-Delegation -Target $_.MailboxName -User $_.UserName        
            Start-Sleep -Milliseconds 500
            Write-Host -ForegroundColor Cyan "---------------------------------------------------------------------------------------------------------"
        }
    }
    End
    {
    }
}
<#
.Synopsis
   Cette fonction permet de lister les délégations d'une boîte exchange.
.DESCRIPTION
   Cette fonction permet de lister l'ensemble des utilisateurs ayant un droit "Full Access" sur la boîte email donnée en paramètre.
   La fonction sélectionne uniquement les comptes utilisateurs correspondant à des boîtes emails, et filtre les utilisateurs génériques ayant automatqieuemtn des droits sur les boîtes emails.
.PARAMETER Mailbox
   Paramètre obligatoire correspondant au nom de la boîte sur laquelle nous souhaitons lister les délégations.
   L'Alias de la boîte comme sa propriété Name peut-être renseignée dans ce paramètre.
.EXAMPLE
   Get-DelegationList -Mailbox test
   Liste les délégations "Full Access" sur la boîte dont l'alias est "test"
.EXAMPLE
   Get-DelegationList -Mailbox "Boîte Test de la Société X"
   Liste les délégations "Full Access" sur la boîte dont le nom est "Boîte Test de la Société X"
#>
function Get-DelegationList
{
    [CmdletBinding()]
    Param
    (
        # Mailbox - nom ou alias de la boîte sur laquelle nous voulons effectuer cette commande
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Mailbox

    )

    Begin
    {
        Write-Verbose "Récupération de l'objet Mailbox passé en paramètre"
        $Target = Get-Mailbox $Mailbox
    }
    Process
    {
        Write-Verbose "Récupération de la liste des permissions sur la boîte $Target"
        $PermissionList = Get-MailboxPermission $Target.Alias | Where-Object {$_.user.tostring() -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false} | Select-Object Identity,User,AccessRights

        $data = foreach($Permission in $PermissionList) {
            [PSCustomObject][ordered]@{
                'Shared Mailbox' = Split-Path $Permission.Identity -Leaf;
                'User'           = Split-Path $Permission.User -Leaf;
                'Delegation'     = Split-Path $Permission.AccessRights -Leaf
            }
        }
        $data
    }
    End
    {
    }
}
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
<#
.Synopsis
   Cette fonction permet de lister le contenu d'un groupe Active Directory passé en paramètre.
.DESCRIPTION
   Cette fonction permet de lister le contenu d'un groupe Active Directory passé en paramètre.
   Si ce groupe contient des sous groupes, leur contenu est également listé (premier niveau seulement)
.PARAMETER ADGroup
    Paramètre obligatoire, il s'agit du groupe Active Directory dont nous souhaitons avoir le contenu
    
.EXAMPLE
   List-GroupMember NomGroup-GS
   Liste le contenu du group NomGroup-GS, ainsi que le contenu des sous-groupes qu'il contient.

#>
function List-GroupMember
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # ADGroup paramètre obligatoire correspondant au nom du groupe Active Directory
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String]$ADGroup
    )

    $UserGroup = @()
    $objRoot = Get-ADGroupMember "$ADGroup" | Sort-Object objectClass
    foreach ($obj in $objRoot) {
        
        if ($obj.objectClass -eq "group") {
            Write-Host -ForegroundColor DarkYellow "________________________________________"
            Write-Host -ForegroundColor Yellow "Groupe" $obj.Name
            Get-ADGroupMember $obj | ft Name,objectClass
        }
        else {
             $Utilisateur = [ordered]@{
                            'Name' = $obj.name;
                            'objectClass' = $obj.objectClass
                            }
            

           $objUtilisateur = New-Object -TypeName PSObject -Property $Utilisateur
           $UserGroup += $objUtilisateur
        }

    }
    if ($UserGroup) { 
        Write-Host -ForegroundColor DarkYellow "________________________________________"
        Write-Host -ForegroundColor Yellow "Racine du Groupe" $ADGroup
        Write-Output $UserGroup
    }
}
<#
.Synopsis
   Cette fonction permet la création d'une ou plusieurs boîtes génériques de type Shared-Mailbox
.DESCRIPTION
   Cette fonction permet la création d'une ou plusieurs boîtes génériques de type Shared-Mailbox à l'aide d'un fichier CSV dans lequel sont passés les différents paramètres nécessaires à la création des boîtes.
   Le fichier CSV doit être écrit sous la forme suivante, le header compris (obligatoire), exemple avec les boîtes "comptabilite" et "contact"
   
   Alias;Prenom;Nom;Description;Password;
   comptabilite;Service;Comptabilité;Boîte générique du service de la comptabilité;MotDePasseUtiliséPourCetteBoîte;
   contact;Service;Contact;Boîte générique contact;MotDePasseDeLaBoîteEnQuestion;

   ATTENTION: ce script a été créé pour une utilisation sur le domaine Active Directory de la société Swiss Risk & Care et nécessite des modifications pour une utilisation en dehors de ce domaine.
.PARAMETER Path
   Le paramètre Path correspond au chemin complet vers le fichier CSV à importer. S'il n'est pas précisé, une boîte de dialogue sera ouverte pour la sélection du fichier.
.EXAMPLE
   New-SharedBAL "D:\temp\dossier\SharedBAL.csv"
   Création des boîtes listées dans le fichier SharedBAL.csv en type Shared-Mailbox
.EXAMPLE
   New-SharedBAL
   Ouverture d'une boîte de dialogue pour sélectionner un fichier CSV contenant une liste de Shared-Mailbox à créer.
#>
function New-SharedBAL
{
    [CmdletBinding()]
    Param
    (
        # PathCSV
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $PathCSV
    )

    Begin
    {
        
        #Si pas de fichier CSV précisé dans la commande, ouverture de la fenêtre de sélection
        if (!$PathCSV) {
            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.Filter = "CSV (*csv)| *.csv"
            $OpenFileDialog.ShowDialog() | Out-Null
            $PathCSV = $OpenFileDialog.FileName
        }
        #Récupération du fichier CSV
        $FichierCSV = Import-Csv -Path $PathCSV -Delimiter ";"
    }
    Process
    {
        $i = 0
        $FichierCSV | ForEach-Object {
            <# ------------------------------------------------ Initialisation des Variables --------------------------------- #>
            #Chemin OU création User AD
            $Chemin_OU = "OU=Collaborative Mailboxes,OU=Users,OU=SRNC,DC=srnc,DC=lan"

            # $Alias_Utilisateur = login AD et alias principal de l'email
            $Alias_Utilisateur = $FichierCSV[$i].Alias

            # $Prenom = Nom Client
            $Prenom = $FichierCSV[$i].Prenom

            # $Nom = healthcare / absences / payroll
            $Nom = $FichierCSV[$i].Nom

            # $Description = "SRNC - <Nom CareDesk> - BAL Client : <Type> <Nom Client>"
            $Description = $FichierCSV[$i].Description

            # $Mot_De_Passe = A générer et enregistrer dans Keepass Swiss Risk & Care > Collaboratives Mailboxes
            $Mot_De_Passe = (ConvertTo-SecureString $FichierCSV[$i].Password -AsPlainText -force)
            $DisplayName = "$Prenom $Nom"
            $AdresseEmail = $Alias_Utilisateur+"@swissriskcare.ch"
            $AD_Utilisateur = $Alias_Utilisateur+"@srnc.lan"
            <# -------------------------------------------- FIn Initialitation des variables --------------------------------- #>
            
            <# ------------------------------------------------------------ Tests pour Debug --------------------------------- #>
            <#
            Write-Host ""
            Write-Host -ForegroundColor Cyan "----------------------- DEBUT DU TEST n° $i -----------"
            Write-Host "Alias: $Alias_Utilisateur"
            Write-Host "Prenom: $Prenom Nom: $Nom"
            Write-Host "Description: $Description"
            Write-Host "Mot de Passe: $Mot_De_Passe"
            Write-Host "OU: $Chemin_OU"
            Write-Host "Display Name: $DisplayName / Email: $AdresseEmail / AD Login: $AD_Utilisateur"
            Write-Host -ForegroundColor Cyan "----------------------- FIN DU TEST n° $i -------------"
            #>
            <# ------------------------------------------------------- Fin tests pour Debug --------------------------------- #>
           
            #Creation User AD dans OU dans le chemin $chemin_OU
            New-ADUser -Server "DC01.srnc.lan" -Name $DisplayName -SamAccountName $Alias_Utilisateur -DisplayName $DisplayName -Surname $Nom -GivenName $Prenom -UserPrincipalName $AD_Utilisateur -Description $Description -Office $Description -AccountPassword $Mot_De_Passe -PasswordNeverExpires $true -CannotChangePassword $true -Path $Chemin_OU -Enabled $false
            Start-Sleep -Seconds 10

            #Création boîte mail Exchange
            Enable-Mailbox -Shared -Identity $AD_Utilisateur -Database "DB02"

           
            $i ++
        }
    }
    End
    {
    }
}
<#
.Synopsis
   Cette fonction permet de supprimer des délégations sur des boîtes email Exchange
.DESCRIPTION
   Cette fonction permet de supprimer des délégations sur des boîtes email Exchange, suppression de l'accès complet à la boîte, et suppression de le possibilité d'envoyer des emails depuis cette boîte.
.PARAMETER Target 
   Ce paramètre est obligatoire et correspond à l'alias ou le nom de la boîte email (générique) sur laquelle nous supprimons une délégation
.PARAMETER User
   Ce paramètre est obligatoire et correspond à l'alias ou le nom de l'utilisateur à qui l'on supprime une délégation
.EXAMPLE
   Remove-Delegation -Target contact -User jean.phumune
   Suppression d'une délégation sur la boîte contact pour l'utilisateur dont l'alias est jean.phumune
.EXAMPLE
   Remove-Delegation -Target "Contact Mailbox" -User "Jean Phumune"
   Suppression d'une délégation sur la boîte contact pour l'utilisateur "Jean Phumune"
#>
function Remove-Delegation
{
    [CmdletBinding()]
    Param
    (
        #Alias ou nom de la boîte email (générique) sur laquelle nous supprimons une délégation
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$Target,

        #Alias ou nom de la boîte email (utilisateur) à qui nous allons supprimer une délégation
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string]$User
    )

    Begin
    {
        Write-Verbose "Récupération des propriétés Name des valeurs retournées"
        $TargetName = Get-Mailbox $Target
        $UserName = Get-Mailbox $User
        Write-Verbose "Suppression d'une délégation de la boîte $Target pour l'utilisateur $User"
    }
    Process
    {
        #Full Access
        Write-Verbose "Suppression du full access sur la boîte $Target pour l'utilisateur $User"
        Remove-MailboxPermission $TargetName.Name -User $UserName.Name -AccessRights FullAccess -InheritanceType All -Confirm:$false

        #Send As
        Write-Verbose "Suppression du send as sur la boîte $Target pour l'utilisateur $User"
        Remove-ADPermission -Identity $TargetName.Name -User $UserName.Name -AccessRights ExtendedRight -ExtendedRights "Send As" -Confirm:$false
    }
    End
    {
        Write-Information "--------------------------------------------------------------------------"
    }
}
<#
.Synopsis
   Cette fonction permet de supprimer des délégations sur des boîtes email Exchange à l'aide d'un fichier CSV passé en paramètre
.DESCRIPTION
   Cette fonction permet de supprimer des délégations sur des boîtes email Exchange, suppression de l'accès complet à la boîte, et suppression de le possibilité d'envoyer des emails depuis cette boîte.
   Un fichier CSV est passé en paramètre sous la forme "Email à dédéléguer;Utilisateur bénéficiaire;"
.PARAMETER Path
   Le paramètre Path correspond au chemin complet vers le fichier CSV à importer. S'il n'est pas précisé, une boîte de dialogue sera ouverte pour la sélection du fichier.
.EXAMPLE
   Remove-DelegationCSV -Path "C:\fichier.csv"
   
   Le nom du fichier CSV n'a pas d'importance, son emplacement non plus tant que le chemin est atteignable par l'ordinateur qui lance la commande
   Le format du fichier CSV est sans en-tête, sous le format suivant: "Nom ou alias BAL générique;Nom ou alias du bénéficiare"
   Exemple de suppression de délégation de la boîte dont l'alias est contact aux utilisateurs Jean Phumune et Aurore Mahler:
   contact;Jean Phumune;
   contact;Aurore Mahler;

#>
function Remove-DelegationCSV
{
    [CmdletBinding()]
    Param
    (
        #Path - chemin du fichier csv contenant la liste des boîtes à supprimer avec leur bénéficiaire
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$Path

    )

    Begin
    {
        #Si pas de fichier CSV précisé dans la commande, ouverture de la fenêtre de sélection
        if (!$Path) {
            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.Filter = "CSV (*csv)| *.csv"
            $OpenFileDialog.ShowDialog() | Out-Null
            $Path = $OpenFileDialog.FileName
        }
    }
    Process
    {
        #Import du fichier CSV et suppression des délégations pour chaque ligne
        Import-CSV $Path -Header MailboxName,UserName -Delimiter ";" | Foreach-Object{
            Write-Host "BAL:" $_.MailboxName "| UserName:" $_.UserName
            Write-Host ""
            Remove-Delegation -Target $_.MailboxName -User $_.UserName        
            Start-Sleep -Milliseconds 500
            Write-Host -ForegroundColor Cyan "---------------------------------------------------------------------------------------------------------"
        }
    }
    End
    {
    }
}