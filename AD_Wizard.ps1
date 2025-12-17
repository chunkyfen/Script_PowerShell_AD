Add-Type -AssemblyName System.Windows.Forms   # Charge la librairie Windows Forms pour créer une interface graphique
Add-Type -AssemblyName System.Drawing         # Charge la librairie pour gérer les polices et couleurs
Import-Module ActiveDirectory                 # Charge le module Active Directory pour utiliser les cmdlets AD

$form = New-Object Windows.Forms.Form         # Crée une nouvelle fenêtre
$form.Text = "nouvel employé"                 # Définit le titre de la fenêtre
$form.Size = '400,450'                        # Définit la taille de la fenêtre

# --- OU principale ---
$lblOU = New-Object Windows.Forms.Label -Property @{Text="OU";Location='10,20'}   # Crée un label "OU"
$form.Controls.Add($lblOU)                                                        # Ajoute le label à la fenêtre

$cbOU = New-Object Windows.Forms.ComboBox -Property @{Location='120,20';Width=200} # Crée une liste déroulante pour les OU
(Get-ADOrganizationalUnit -Filter *).DistinguishedName | ForEach-Object { $cbOU.Items.Add($_) } # Remplit la liste avec les OU existantes
$form.Controls.Add($cbOU)                                                         # Ajoute la liste à la fenêtre

$btnAddOU = New-Object Windows.Forms.Button -Property @{Text="Ajouter OU";Location='330,20'} # Crée un bouton "Ajouter OU"
$btnAddOU.Add_Click({                                                             # Action quand on clique sur le bouton
    $newOU = [System.Windows.Forms.InputBox]::Show("Nom de la nouvelle OU","Créer OU") # Demande le nom de la nouvelle OU
    if ($newOU) { New-ADOrganizationalUnit -Name $newOU -Path "DC=script,DC=local"    # Crée la nouvelle OU dans AD
        $cbOU.Items.Add("OU=$newOU,DC=script,DC=local") }                             # Ajoute la nouvelle OU à la liste
})
$form.Controls.Add($btnAddOU)                                                     # Ajoute le bouton à la fenêtre

# --- Groupe ---
$lblGrp = New-Object Windows.Forms.Label -Property @{Text="Groupe";Location='10,60'} # Crée un label "Groupe"
$form.Controls.Add($lblGrp)                                                         # Ajoute le label à la fenêtre

$cbGrp = New-Object Windows.Forms.ComboBox -Property @{Location='120,60';Width=200} # Crée une liste déroulante pour les groupes
(Get-ADGroup -Filter *).Name | ForEach-Object { $cbGrp.Items.Add($_) }              # Remplit la liste avec les groupes existants
$form.Controls.Add($cbGrp)                                                          # Ajoute la liste à la fenêtre

$btnAddGrp = New-Object Windows.Forms.Button -Property @{Text="Ajouter Groupe";Location='330,60'} # Crée un bouton "Ajouter Groupe"
$btnAddGrp.Add_Click({                                                              # Action quand on clique sur le bouton
    $newGrp = [System.Windows.Forms.InputBox]::Show("Nom du nouveau groupe","Créer Groupe") # Demande le nom du nouveau groupe
    if ($newGrp) { New-ADGroup -Name $newGrp -GroupScope Global -Path "DC=script,DC=local" # Crée le groupe dans AD
        $cbGrp.Items.Add($newGrp) }                                                 # Ajoute le groupe à la liste
})
$form.Controls.Add($btnAddGrp)                                                     # Ajoute le bouton à la fenêtre

# --- Champs utilisateur ---
$txtNom = New-Object Windows.Forms.TextBox -Property @{Location='120,100'}         # Zone de texte pour le Nom
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Nom";Location='10,100'})) # Label "Nom"
$form.Controls.Add($txtNom)                                                        # Ajoute la zone de texte

$txtPrenom = New-Object Windows.Forms.TextBox -Property @{Location='120,140'}      # Zone de texte pour le Prénom
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Prénom";Location='10,140'})) # Label "Prénom"
$form.Controls.Add($txtPrenom)                                                     # Ajoute la zone de texte

$cbJour = New-Object Windows.Forms.ComboBox -Property @{Location='120,180'}        # Liste déroulante pour le Jour
1..31 | ForEach-Object { $cbJour.Items.Add($_.ToString("00")) }                    # Remplit avec 01 à 31
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Jour";Location='10,180'})) # Label "Jour"
$form.Controls.Add($cbJour)                                                        # Ajoute la liste

$cbMois = New-Object Windows.Forms.ComboBox -Property @{Location='120,220'}        # Liste déroulante pour le Mois
1..12 | ForEach-Object { $cbMois.Items.Add($_.ToString("00")) }                    # Remplit avec 01 à 12
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Mois";Location='10,220'})) # Label "Mois"
$form.Controls.Add($cbMois)                                                        # Ajoute la liste

$cbAnnee = New-Object Windows.Forms.ComboBox -Property @{Location='120,260'}       # Liste déroulante pour l’Année
1980..2005 | ForEach-Object { $cbAnnee.Items.Add($_) }                             # Remplit avec 1980 à 2005
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Année";Location='10,260'})) # Label "Année"
$form.Controls.Add($cbAnnee)                                                       # Ajoute la liste

# --- Résultat ---
$lblResult = New-Object Windows.Forms.Label -Property @{                           # Label pour afficher UserName et Password
    ForeColor='Red'; Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Bold));
    Location='10,300'; Size='350,40'
}
$form.Controls.Add($lblResult)                                                     # Ajoute le label à la fenêtre

# --- Boutons ---
$btnGen = New-Object Windows.Forms.Button -Property @{Text="Générer";Location='10,360'} # Bouton "Générer"
$btnGen.Add_Click({                                                                # Action quand on clique
    if ($txtNom.Text -and $txtPrenom.Text -and $cbAnnee.SelectedItem) {            # Vérifie que Nom, Prénom, Année sont remplis
        $u = ($txtPrenom.Text.Substring(0,1).ToLower() + $txtNom.Text.ToLower())   # Génère UserName = 1ère lettre prénom + nom
        $p = ($txtNom.Text.ToLower() + $cbAnnee.SelectedItem + "?")                # Génère Password = nom + année + ?
        $lblResult.Text = "UserName: $u `nPassword: $p"                            # Affiche le résultat en rouge/gras
    }
})
$form.Controls.Add($btnGen)                                                        # Ajoute le bouton

$btnSave = New-Object Windows.Forms.Button -Property @{Text="Enregistrer";Location='120,360'} # Bouton "Enregistrer"
$btnSave.Add_Click({                                                               # Action quand on clique
    $nom = $txtNom.Text; $prenom = $txtPrenom.Text; $annee = $cbAnnee.SelectedItem # Récupère les valeurs saisies
    $username = ($prenom.Substring(0,1).ToLower() + $nom.ToLower())                # Génère UserName
    $password = ($nom.ToLower() + $annee + "?")                                    # Génère Password
    $securePass = ConvertTo-SecureString $password -AsPlainText -Force             # Convertit le mot de passe en SecureString

    New-ADUser -Name "$prenom $nom" `                                              # Crée l’utilisateur AD
               -SamAccountName $username `
               -UserPrincipalName "$username@script.local" `
               -AccountPassword $securePass `
               -Enabled $true `
               -Path $cbOU.SelectedItem

    Add-ADGroupMember -Identity $cbGrp.SelectedItem -Members $username             # Ajoute l’utilisateur au groupe choisi

    $txtNom.Clear(); $txtPrenom.Clear()                                            # Réinitialise les champs
    $cbJour.SelectedIndex=$cbMois.SelectedIndex=$cbAnnee.SelectedIndex=$cbGrp.SelectedIndex=$cbOU.SelectedIndex=-1
    $lblResult.Text="Utilisateur $username créé"                                   # Message de confirmation
})
$form.Controls.Add($btnSave)                                                       # Ajoute le bouton

$btnCancel = New-Object Windows.Forms.Button -Property @{Text="Annuler";Location='230,360'} # Bouton "Annuler"
$btnCancel.Add_Click({ $form.Close() })                                            # Ferme la fenêtre quand on clique
$form.Controls.Add($btnCancel)                                                     # Ajoute le bouton

$form.ShowDialog()                                                                 # Affiche la fenêtre et attend l’interaction
