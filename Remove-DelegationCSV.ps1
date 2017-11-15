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