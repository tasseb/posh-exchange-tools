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