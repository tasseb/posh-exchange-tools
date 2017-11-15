<#
.Synopsis
   Cette fonction permet d'importer les commandes exchange
.DESCRIPTION
   Cette fonction permet d'importer les commandes exchange depuis le serveur exchange
   Certaines opérations retourneront une erreur car elles ne peuvent être jouées que sur la console EMC du serveur Exchange.
.PARAMETER Serveur
   Ce paramètre est obligatoire et doit contenir le nom réseau du serveur (ex: "serveur-exchange.contoso.com")
.PARAMETER Utilisateur
   Ce paramètre permet de renseigner le nom d'utilisateur utilisé pour se connecter au serveur Exchange.
   ATTENTION: l'utilisateur doit disposer de droits d'administration sur le serveur Exchange.
   Exemple "admin-user@contoso.com".
.EXAMPLE
   Import-ExchangeSession -Serveur serveur-exchange.contoso.com -Utilisateur admin-user@contoso.com
   Se connecter au serveur Exchange serveur-exchange.contoso.com avec l'utilisateur admin-user@contoso.com
.EXAMPLE
   Import-ExchangeSession -Serveur serveur-exchange.contoso.com
   Se connecter au serveur Exchange serveur-exchange.contoso.com sans préciser le nom d'utilisateur (demandé ensuite)
#>
function Import-ExchangeSession
{
    [CmdletBinding()]
    Param
    (
        #Serveur Exchange (forme dns)
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$Serveur,

        #Utilisateur (avec droit sur le serveur Exchange)
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Utilisateur
    )

        Write-Verbose "Récupération login et mot de passe"
        if ($Utilisateur) {
            $UserCredential = Get-Credential -Credential $Utilisateur
        }
        else {
            $UserCredential = Get-Credential
        }

        Write-Verbose "Création de la PSSession sur le serveur Exchange $Serveur"
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Serveur/PowerShell/" -Authentication Kerberos -Credential $UserCredential

        Write-Verbose "Import de la session et du jeu de commande EMC"
        Import-PSSession $Session

}