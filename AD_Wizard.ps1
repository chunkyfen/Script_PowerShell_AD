Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

$form = New-Object Windows.Forms.Form
$form.Text = "nouvel employe"
$form.Size = '450,580'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# --- Section: Information personnelle ---
$grpInfo = New-Object Windows.Forms.GroupBox
$grpInfo.Text = "Information personnelle"
$grpInfo.Location = '15,10'
$grpInfo.Size = '410,180'
$form.Controls.Add($grpInfo)

# --- Nom ---
$lblNom = New-Object Windows.Forms.Label -Property @{Text="Nom :";Location='20,30';Size='100,20'}
$grpInfo.Controls.Add($lblNom)
$txtNom = New-Object Windows.Forms.TextBox -Property @{Location='130,28';Width=250}
$grpInfo.Controls.Add($txtNom)

# --- Prenom ---
$lblPrenom = New-Object Windows.Forms.Label -Property @{Text="Prénom :";Location='20,65';Size='100,20'}
$grpInfo.Controls.Add($lblPrenom)
$txtPrenom = New-Object Windows.Forms.TextBox -Property @{Location='130,63';Width=250}
$grpInfo.Controls.Add($txtPrenom)

# --- Groupe ---
$lblGrp = New-Object Windows.Forms.Label -Property @{Text="Groupe :";Location='20,100';Size='100,20'}
$grpInfo.Controls.Add($lblGrp)
$cbGrp = New-Object Windows.Forms.ComboBox -Property @{Location='130,98';Width=250;DropDownStyle='DropDownList'}
$cbGrp.Items.Add("Informatique") | Out-Null
$cbGrp.Items.Add("comptabilite") | Out-Null
$cbGrp.Items.Add("RH") | Out-Null
$grpInfo.Controls.Add($cbGrp)

# --- Date de naissance ---
$lblDate = New-Object Windows.Forms.Label -Property @{Text="Date de naissance :";Location='20,135';Size='100,20'}
$grpInfo.Controls.Add($lblDate)

$cbJour = New-Object Windows.Forms.ComboBox -Property @{Location='130,133';Width=60;DropDownStyle='DropDownList'}
1..31 | ForEach-Object { $cbJour.Items.Add($_.ToString("00")) | Out-Null }
$grpInfo.Controls.Add($cbJour)

$lblSlash1 = New-Object Windows.Forms.Label -Property @{Text="/";Location='195,135';Size='10,20'}
$grpInfo.Controls.Add($lblSlash1)

$cbMois = New-Object Windows.Forms.ComboBox -Property @{Location='205,133';Width=60;DropDownStyle='DropDownList'}
1..12 | ForEach-Object { $cbMois.Items.Add($_.ToString("00")) | Out-Null }
$grpInfo.Controls.Add($cbMois)

$lblSlash2 = New-Object Windows.Forms.Label -Property @{Text="/";Location='270,135';Size='10,20'}
$grpInfo.Controls.Add($lblSlash2)

$cbAnnee = New-Object Windows.Forms.ComboBox -Property @{Location='280,133';Width=100;DropDownStyle='DropDownList'}
1980..2005 | ForEach-Object { $cbAnnee.Items.Add($_) | Out-Null }
$grpInfo.Controls.Add($cbAnnee)

# --- Bouton Générer ---
$btnGenerate = New-Object Windows.Forms.Button -Property @{
    Text="Generer le User Name && Password"
    Location='80,210'
    Width=280
    Height=35
    Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))
}
$form.Controls.Add($btnGenerate)

# --- Section: Résultat ---
$grpResult = New-Object Windows.Forms.GroupBox
$grpResult.Text = "Resultat"
$grpResult.Location = '15,260'
$grpResult.Size = '410,150'
$form.Controls.Add($grpResult)

# --- User Name ---
$lblUserNameTitle = New-Object Windows.Forms.Label -Property @{
    Text="User Name :"
    Location='20,30'
    Size='100,20'
}
$grpResult.Controls.Add($lblUserNameTitle)

$lblUserName = New-Object Windows.Forms.Label -Property @{
    Text=""
    ForeColor='Red'
    Font=(New-Object Drawing.Font("Arial",11,[Drawing.FontStyle]::Bold))
    Location='130,28'
    Size='250,25'
    BorderStyle='FixedSingle'
    BackColor='White'
    TextAlign='MiddleLeft'
}
$grpResult.Controls.Add($lblUserName)

# --- Password ---
$lblPasswordTitle = New-Object Windows.Forms.Label -Property @{
    Text="Password :"
    Location='20,70'
    Size='100,20'
}
$grpResult.Controls.Add($lblPasswordTitle)

$lblPassword = New-Object Windows.Forms.Label -Property @{
    Text=""
    ForeColor='Red'
    Font=(New-Object Drawing.Font("Arial",11,[Drawing.FontStyle]::Bold))
    Location='130,68'
    Size='250,25'
    BorderStyle='FixedSingle'
    BackColor='White'
    TextAlign='MiddleLeft'
}
$grpResult.Controls.Add($lblPassword)

# --- Message de statut ---
$lblStatus = New-Object Windows.Forms.Label -Property @{
    Text=""
    ForeColor='Green'
    Font=(New-Object Drawing.Font("Arial",9,[Drawing.FontStyle]::Bold))
    Location='20,105'
    Size='370,30'
    TextAlign='MiddleCenter'
}
$grpResult.Controls.Add($lblStatus)

# --- Boutons Enregistrer et Annuler ---
$btnSave = New-Object Windows.Forms.Button -Property @{
    Text="Enregistrer"
    Location='100,430'
    Width=120
    Height=35
    Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))
}
$form.Controls.Add($btnSave)

$btnCancel = New-Object Windows.Forms.Button -Property @{
    Text="Annuler"
    Location='230,430'
    Width=120
    Height=35
    Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))
}
$form.Controls.Add($btnCancel)

# --- Événement du bouton Générer ---
$btnGenerate.Add_Click({
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $annee = $cbAnnee.SelectedItem
    
    # Validation
    if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
        [Windows.Forms.MessageBox]::Show("Nom et prenom requis!", "Erreur", 'OK', 'Error')
        return
    }
    
    if ($null -eq $annee) {
        [Windows.Forms.MessageBox]::Show("Annee de naissance requise!", "Erreur", 'OK', 'Error')
        return
    }
    
    # Génération du User Name: première lettre du prénom + nom
    $username = ($prenom.Substring(0,1).ToLower() + $nom.ToLower())
    $username = $username -replace '[^a-z0-9]', ''
    
    # Génération du Password: nom + année + ?
    $password = ($nom.ToLower() + $annee + "?")
    
    # Affichage
    $lblUserName.Text = $username
    $lblPassword.Text = $password
    $lblStatus.Text = ""
})

# --- Événement du bouton Enregistrer ---
$btnSave.Add_Click({
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $jour = $cbJour.SelectedItem
    $mois = $cbMois.SelectedItem
    $annee = $cbAnnee.SelectedItem
    $groupName = $cbGrp.SelectedItem
    $username = $lblUserName.Text
    $password = $lblPassword.Text
    
    # Validation complète
    if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
        [Windows.Forms.MessageBox]::Show("Nom et prenom requis!", "Erreur", 'OK', 'Error')
        return
    }
    
    if ($null -eq $groupName) {
        [Windows.Forms.MessageBox]::Show("Groupe requis!", "Erreur", 'OK', 'Error')
        return
    }
    
    if ($null -eq $jour -or $null -eq $mois -or $null -eq $annee) {
        [Windows.Forms.MessageBox]::Show("Date de naissance complete requise!", "Erreur", 'OK', 'Error')
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($password)) {
        [Windows.Forms.MessageBox]::Show("Veuillez d'abord generer le User Name et Password!", "Erreur", 'OK', 'Error')
        return
    }
    
    # Vérifier si l'utilisateur existe déjà
    try {
        $userExists = Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue
        if ($userExists) {
            $counter = 2
            $originalUsername = $username
            while (Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue) {
                $username = "$originalUsername$counter"
                $counter++
            }
            $lblUserName.Text = $username
        }
    } catch {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = "Erreur verification utilisateur"
        return
    }
    
    $securePass = ConvertTo-SecureString $password -AsPlainText -Force
    $dateNaissance = "$annee-$mois-$jour"
    
    try {
        # Obtenir le domaine par défaut
        $ouPath = "CN=Users," + (Get-ADDomain).DistinguishedName
        
        # Création utilisateur AD
        $newUserParams = @{
            Name = "$prenom $nom"
            GivenName = $prenom
            Surname = $nom
            SamAccountName = $username
            UserPrincipalName = "$username@script.local"
            AccountPassword = $securePass
            Enabled = $true
            Path = $ouPath
            Description = "Date de naissance: $dateNaissance - Groupe: $groupName"
            ChangePasswordAtLogon = $false
        }
        
        New-ADUser @newUserParams -ErrorAction Stop
        Start-Sleep -Seconds 1
        
        # Ajout au groupe
        try {
            Add-ADGroupMember -Identity $groupName -Members $username -ErrorAction Stop
        } catch {
            # Si le groupe n'existe pas, on continue quand même
            Write-Host "Avertissement: Groupe $groupName non trouve" -ForegroundColor Yellow
        }
        
        # RESET: Remettre tous les champs à leur état initial
        $txtNom.Clear()
        $txtPrenom.Clear()
        $cbJour.SelectedIndex = -1
        $cbMois.SelectedIndex = -1
        $cbAnnee.SelectedIndex = -1
        $cbGrp.SelectedIndex = -1
        $lblUserName.Text = ""
        $lblPassword.Text = ""
        
        $lblStatus.ForeColor = 'Green'
        $lblStatus.Text = "Utilisateur cree avec succes!"
        
        [Windows.Forms.MessageBox]::Show("Utilisateur $username cree avec succes dans le groupe $groupName!", "Succes", 'OK', 'Information')
    }
    catch {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = "Erreur lors de la creation"
        [Windows.Forms.MessageBox]::Show("Erreur: $($_.Exception.Message)", "Erreur", 'OK', 'Error')
    }
})

# --- Événement du bouton Annuler ---
$btnCancel.Add_Click({ $form.Close() })

# Afficher le formulaire
$form.ShowDialog()
