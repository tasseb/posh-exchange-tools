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