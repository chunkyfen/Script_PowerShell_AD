Add-Type -AssemblyName System.Windows.Forms   # Charge la librairie Windows Forms (nécessaire pour créer une interface graphique)
Add-Type -AssemblyName System.Drawing         # Charge la librairie pour gérer les polices, couleurs et styles
Import-Module ActiveDirectory                 # Charge le module Active Directory pour utiliser les cmdlets AD (New-ADUser, Add-ADGroupMember)

$form = New-Object Windows.Forms.Form         # Crée une nouvelle fenêtre graphique
$form.Text = "nouvel employé"                 # Définit le titre affiché en haut de la fenêtre
$form.Size = '400,450'                        # Définit la taille de la fenêtre (largeur x hauteur en pixels)

# --- OU principale ---
$lblOU = New-Object Windows.Forms.Label -Property @{Text="OU";Location='10,20'}   # Crée un label "OU" placé à la position (10,20)
$form.Controls.Add($lblOU)                                                        # Ajoute ce label à la fenêtre

$cbOU = New-Object Windows.Forms.ComboBox -Property @{Location='120,20';Width=250} # Crée une liste déroulante (ComboBox) pour choisir une OU
(Get-ADOrganizationalUnit -Filter *).DistinguishedName | ForEach-Object { $cbOU.Items.Add($_) } # Remplit la liste avec toutes les OU existantes dans AD
$form.Controls.Add($cbOU)                                                         # Ajoute la liste déroulante à la fenêtre

# --- Groupe ---
$lblGrp = New-Object Windows.Forms.Label -Property @{Text="Groupe";Location='10,60'} # Crée un label "Groupe"
$form.Controls.Add($lblGrp)                                                         # Ajoute ce label à la fenêtre

$cbGrp = New-Object Windows.Forms.ComboBox -Property @{Location='120,60';Width=250} # Crée une liste déroulante pour choisir un groupe
(Get-ADGroup -Filter *).Name | ForEach-Object { $cbGrp.Items.Add($_) }              # Remplit la liste avec tous les groupes existants dans AD
$form.Controls.Add($cbGrp)                                                          # Ajoute la liste déroulante à la fenêtre

# --- Champs utilisateur ---
$txtNom = New-Object Windows.Forms.TextBox -Property @{Location='120,100'}         # Zone de texte pour saisir le Nom
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Nom";Location='10,100'})) # Label "Nom"
$form.Controls.Add($txtNom)                                                        # Ajoute la zone de texte à la fenêtre

$txtPrenom = New-Object Windows.Forms.TextBox -Property @{Location='120,140'}      # Zone de texte pour saisir le Prénom
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Prénom";Location='10,140'})) # Label "Prénom"
$form.Controls.Add($txtPrenom)                                                     # Ajoute la zone de texte

$cbJour = New-Object Windows.Forms.ComboBox -Property @{Location='120,180'}        # Liste déroulante pour choisir le Jour
1..31 | ForEach-Object { $cbJour.Items.Add($_.ToString("00")) }                    # Remplit la liste avec les jours de 01 à 31
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Jour";Location='10,180'})) # Label "Jour"
$form.Controls.Add($cbJour)                                                        # Ajoute la liste déroulante

$cbMois = New-Object Windows.Forms.ComboBox -Property @{Location='120,220'}        # Liste déroulante pour choisir le Mois
1..12 | ForEach-Object { $cbMois.Items.Add($_.ToString("00")) }                    # Remplit la liste avec les mois de 01 à 12
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Mois";Location='10,220'})) # Label "Mois"
$form.Controls.Add($cbMois)                                                        # Ajoute la liste déroulante

$cbAnnee = New-Object Windows.Forms.ComboBox -Property @{Location='120,260'}       # Liste déroulante pour choisir l’Année
1980..2005 | ForEach-Object { $cbAnnee.Items.Add($_) }                             # Remplit la liste avec les années de 1980 à 2005
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Année";Location='10,260'})) # Label "Année"
$form.Controls.Add($cbAnnee)                                                       # Ajoute la liste déroulante

# --- Résultat ---
$lblResult = New-Object Windows.Forms.Label -Property @{                           # Crée un label pour afficher UserName et Password
    ForeColor='Red'; Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Bold)); # Texte rouge et gras
    Location='10,300'; Size='350,40'                                               # Position et taille du label
}
$form.Controls.Add($lblResult)                                                     # Ajoute le label à la fenêtre

# --- Boutons ---
$btnGen = New-Object Windows.Forms.Button -Property @{Text="Générer";Location='10,360'} # Bouton "Générer"
$btnGen.Add_Click({                                                                # Action exécutée quand on clique sur le bouton
    if ($txtNom.Text -and $txtPrenom.Text -and $cbAnnee.SelectedItem) {            # Vérifie que Nom, Prénom et Année sont remplis
        $u = ($txtPrenom.Text.Substring(0,1).ToLower() + $txtNom.Text.ToLower())   # Génère UserName = 1ère lettre du prénom + nom
        $p = ($txtNom.Text.ToLower() + $cbAnnee.SelectedItem + "?")                # Génère Password = nom + année + ?
        $lblResult.Text = "UserName: $u `nPassword: $p"                            # Affiche UserName et Password dans le label résultat
    }
})
$form.Controls.Add($btnGen)                                                        # Ajoute le bouton "Générer" à la fenêtre

$btnSave = New-Object Windows.Forms.Button -Property @{Text="Enregistrer";Location='120,360'} # Bouton "Enregistrer"
$btnSave.Add_Click({                                                               # Action exécutée quand on clique sur le bouton
    $nom = $txtNom.Text; $prenom = $txtPrenom.Text; $annee = $cbAnnee.SelectedItem # Récupère les valeurs saisies
    $username = ($prenom.Substring(0,1).ToLower() + $nom.ToLower())                # Génère UserName
    $password = ($nom.ToLower() + $annee + "?")                                    # Génère Password
    $securePass = ConvertTo-SecureString $password -AsPlainText -Force             # Convertit le mot de passe en SecureString (obligatoire pour AD)

    # Crée l’utilisateur dans Active Directory
    New-ADUser -Name "$prenom $nom" -SamAccountName $username -UserPrincipalName "$username@script.local" -AccountPassword $securePass -Enabled $true -Path $cbOU.SelectedItem


    Add-ADGroupMember -Identity $cbGrp.SelectedItem -Members $username             # Ajoute l’utilisateur au groupe sélectionné

    $txtNom.Clear(); $txtPrenom.Clear()                                            # Vide les champs Nom et Prénom
    $cbJour.SelectedIndex=$cbMois.SelectedIndex=$cbAnnee.SelectedIndex=$cbGrp.SelectedIndex=$cbOU.SelectedIndex=-1 # Réinitialise les listes déroulantes
    $lblResult.Text="Utilisateur $username créé"                                   # Affiche un message de confirmation
})
$form.Controls.Add($btnSave)                                                       # Ajoute le bouton "Enregistrer" à la fenêtre

$btnCancel = New-Object Windows.Forms.Button -Property @{Text="Annuler";Location='230,360'} # Bouton "Annuler"
$btnCancel.Add_Click({ $form.Close() })                                            # Ferme la fenêtre quand on clique sur "Annuler"
$form.Controls.Add($btnCancel)                                                     # Ajoute le bouton "Annuler" à la fenêtre

$form.ShowDialog()                                                                 # Affiche la fenêtre et attend l’interaction de l’utilisateur
