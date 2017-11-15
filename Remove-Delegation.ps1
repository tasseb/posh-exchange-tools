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