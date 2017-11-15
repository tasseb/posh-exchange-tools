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