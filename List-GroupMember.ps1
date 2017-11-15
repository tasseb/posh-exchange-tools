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