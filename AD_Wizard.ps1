# ============================
# Chargement des librairies nécessaires
# ============================
Add-Type -AssemblyName System.Windows.Forms   # Interface graphique Windows Forms
Add-Type -AssemblyName System.Drawing         # Gestion des polices et couleurs
Import-Module ActiveDirectory                 # Cmdlets Active Directory (New-ADUser, Add-ADGroupMember)

# ============================
# Création de la fenêtre principale
# ============================
$form = New-Object Windows.Forms.Form
$form.Text = "New Employee"
$form.Size = '450,580'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# ============================
# Bloc d'informations personnelles
# ============================
$grpInfo = New-Object Windows.Forms.GroupBox
$grpInfo.Text = "Personal Information"
$grpInfo.Location = '15,10'
$grpInfo.Size = '410,180'
$form.Controls.Add($grpInfo)

# Champs Nom et Prénom
$lblNom = New-Object Windows.Forms.Label -Property @{Text="Last Name :";Location='20,30';Size='100,20'}
$grpInfo.Controls.Add($lblNom)
$txtNom = New-Object Windows.Forms.TextBox -Property @{Location='130,28';Width=250}
$grpInfo.Controls.Add($txtNom)

$lblPrenom = New-Object Windows.Forms.Label -Property @{Text="First Name :";Location='20,65';Size='100,20'}
$grpInfo.Controls.Add($lblPrenom)
$txtPrenom = New-Object Windows.Forms.TextBox -Property @{Location='130,63';Width=250}
$grpInfo.Controls.Add($txtPrenom)

# Sélection du groupe
$lblGrp = New-Object Windows.Forms.Label -Property @{Text="Group :";Location='20,100';Size='100,20'}
$grpInfo.Controls.Add($lblGrp)
$cbGrp = New-Object Windows.Forms.ComboBox -Property @{Location='130,98';Width=250;DropDownStyle='DropDownList'}
$cbGrp.Items.Add("Informatique") | Out-Null
$cbGrp.Items.Add("comptabilite") | Out-Null
$cbGrp.Items.Add("RH") | Out-Null
$grpInfo.Controls.Add($cbGrp)

# Date de naissance
$lblDate = New-Object Windows.Forms.Label -Property @{Text="Date of Birth :";Location='20,135';Size='100,20'}
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

# ============================
# Bouton pour générer login et mot de passe
# ============================
$btnGenerate = New-Object Windows.Forms.Button -Property @{
    Text="Generate User Name && Password"
    Location='80,210'
    Width=280
    Height=35
    Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))
}
$form.Controls.Add($btnGenerate)

# ============================
# Bloc de résultats (affichage login/mot de passe)
# ============================
$grpResult = New-Object Windows.Forms.GroupBox
$grpResult.Text = "Result"
$grpResult.Location = '15,260'
$grpResult.Size = '410,150'
$form.Controls.Add($grpResult)

# Affichage du login
$lblUserNameTitle = New-Object Windows.Forms.Label -Property @{Text="User Name :";Location='20,30';Size='100,20'}
$grpResult.Controls.Add($lblUserNameTitle)

$lblUserName = New-Object Windows.Forms.TextBox -Property @{
    Text=""
    ForeColor='Red'
    Font=(New-Object Drawing.Font("Arial",11,[Drawing.FontStyle]::Bold))
    Location='130,28'
    Width=250
    BackColor='White'
}
$grpResult.Controls.Add($lblUserName)

# Affichage du mot de passe
$lblPasswordTitle = New-Object Windows.Forms.Label -Property @{Text="Password :";Location='20,70';Size='100,20'}
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

# Label de statut (succès/erreur)
$lblStatus = New-Object Windows.Forms.Label -Property @{
    Text=""
    ForeColor='Green'
    Font=(New-Object Drawing.Font("Arial",9,[Drawing.FontStyle]::Bold))
    Location='20,105'
    Size='370,30'
    TextAlign='MiddleCenter'
}
$grpResult.Controls.Add($lblStatus)

# ============================
# Boutons Save et Cancel
# ============================
$btnSave = New-Object Windows.Forms.Button -Property @{Text="Save";Location='100,430';Width=120;Height=35;Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))}
$form.Controls.Add($btnSave)

$btnCancel = New-Object Windows.Forms.Button -Property @{Text="Cancel";Location='230,430';Width=120;Height=35;Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))}
$form.Controls.Add($btnCancel)

# ============================
# Logique du bouton Generate
# ============================
$btnGenerate.Add_Click({
    # Nettoyage des accents et génération login/mot de passe
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $annee = $cbAnnee.SelectedItem
    
    if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
        [Windows.Forms.MessageBox]::Show("Last name and first name are required!", "Error", 'OK', 'Error')
        return
    }
    if ($null -eq $annee) {
        [Windows.Forms.MessageBox]::Show("Year of birth is required!", "Error", 'OK', 'Error')
        return
    }
    
    $prenomClean = $prenom -replace '[àâäã]','a' -replace '[éèêë]','e' -replace '[îï]','i' -replace '[ôö]','o' -replace '[ùûü]','u' -replace '[ç]','c'
    $nomClean = $nom -replace '[àâäã]','a' -replace '[éèêë]','e' -replace '[îï]','i' -replace '[ôö]','o' -replace '[ùûü]','u' -replace '[ç]','c'
    
    $username = ($prenomClean.Substring(0,1).ToLower() + $nomClean.ToLower()) -replace '[^a-z0-9]', ''
    $password = ($nomClean.ToLower() + $annee + "?")
    
    $lblUserName.Text = $username
    $lblPassword.Text = $password
    $lblStatus.Text = ""
})

# ============================
# Logique du bouton Save (création AD + ajout groupe)
# ============================
$btnSave.Add_Click({
    # Vérification des champs
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $jour = $cbJour.SelectedItem
    $mois = $cbMois.SelectedItem
    $annee = $cbAnnee.SelectedItem
    $groupName = $cbGrp.SelectedItem
    $username = $lblUserName.Text
    $password = $lblPassword.Text
    
    if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
        [Windows.Forms.MessageBox]::Show("Last name and first name are required!", "Error", 'OK', 'Error')
