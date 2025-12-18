Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

$form = New-Object Windows.Forms.Form
$form.Text = "Nouvel employe"
$form.Size = '400,480'

# --- OU principale ---
$lblOU = New-Object Windows.Forms.Label -Property @{Text="OU";Location='10,20'}
$form.Controls.Add($lblOU)
$cbOU = New-Object Windows.Forms.ComboBox -Property @{Location='120,20';Width=250}
(Get-ADOrganizationalUnit -Filter *).DistinguishedName | ForEach-Object { $cbOU.Items.Add($_) }
$form.Controls.Add($cbOU)

# --- Groupe ---
$lblGrp = New-Object Windows.Forms.Label -Property @{Text="Groupe";Location='10,60'}
$form.Controls.Add($lblGrp)
$cbGrp = New-Object Windows.Forms.ComboBox -Property @{Location='120,60';Width=250}
(Get-ADGroup -Filter *).Name | ForEach-Object { $cbGrp.Items.Add($_) }
$form.Controls.Add($cbGrp)

# --- Champs utilisateur ---
$txtNom = New-Object Windows.Forms.TextBox -Property @{Location='120,100'}
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Nom";Location='10,100'}))
$form.Controls.Add($txtNom)

$txtPrenom = New-Object Windows.Forms.TextBox -Property @{Location='120,140'}
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Prenom";Location='10,140'}))
$form.Controls.Add($txtPrenom)

# --- Bloc Date de naissance ---
$lblDate = New-Object Windows.Forms.Label -Property @{Text="Date de naissance";Location='10,180'}
$form.Controls.Add($lblDate)

$cbJour = New-Object Windows.Forms.ComboBox -Property @{Location='120,200'}
1..31 | ForEach-Object { $cbJour.Items.Add($_.ToString("00")) }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Jour";Location='10,200'}))
$form.Controls.Add($cbJour)

$cbMois = New-Object Windows.Forms.ComboBox -Property @{Location='120,240'}
1..12 | ForEach-Object { $cbMois.Items.Add($_.ToString("00")) }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Mois";Location='10,240'}))
$form.Controls.Add($cbMois)

$cbAnnee = New-Object Windows.Forms.ComboBox -Property @{Location='120,280'}
1910..2010 | ForEach-Object { $cbAnnee.Items.Add($_) }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Annee";Location='10,280'}))
$form.Controls.Add($cbAnnee)

# --- Resultat ---
$lblResult = New-Object Windows.Forms.Label -Property @{
    ForeColor='Red'; Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Bold));
    Location='10,320'; Size='350,40'
}
$form.Controls.Add($lblResult)

# --- Boutons ---
$btnSave = New-Object Windows.Forms.Button -Property @{Text="Enregistrer";Location='120,370'}
$btnSave.Add_Click({
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $jour = $cbJour.SelectedItem
    $mois = $cbMois.SelectedItem
    $annee = $cbAnnee.SelectedItem
    
    # Validation des champs requis
    if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
        $lblResult.Text = "Nom et prenom requis!"
        return
    }
    
    if ($null -eq $cbOU.SelectedItem -or $null -eq $cbGrp.SelectedItem) {
        $lblResult.Text = "OU et Groupe requis!"
        return
    }
    
    if ($null -eq $jour -or $null -eq $mois -or $null -eq $annee) {
        $lblResult.Text = "Date de naissance complete requise!"
        return
    }
    
    # Génération du nom d'utilisateur
    $username = ($prenom.Substring(0,1).ToLower() + $nom.ToLower())
    
    # Vérifier si l'utilisateur existe déjà
    $userExists = Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue
    if ($userExists) {
        # Ajouter un numéro si l'utilisateur existe
        $counter = 2
        $originalUsername = $username
        while (Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue) {
            $username = "$originalUsername$counter"
            $counter++
        }
    }
    
    # Génération du mot de passe
    $password = ($nom.ToLower() + $annee + "?")
    $securePass = ConvertTo-SecureString $password -AsPlainText -Force
    
    # Date de naissance formatée
    $dateNaissance = "$annee-$mois-$jour"
    
    try {
        # Création utilisateur AD avec login et mot de passe
        New-ADUser -Name "$prenom $nom" `
                   -GivenName $prenom `
                   -Surname $nom `
                   -SamAccountName $username `
                   -UserPrincipalName "$username@script.local" `
                   -AccountPassword $securePass `
                   -Enabled $true `
                   -Path $cbOU.SelectedItem `
                   -Description "Date de naissance: $dateNaissance"
        
        # Ajout au groupe
        Add-ADGroupMember -Identity $cbGrp.SelectedItem -Members $username
        
        # Reset
        $txtNom.Clear()
        $txtPrenom.Clear()
        $cbJour.SelectedIndex = -1
        $cbMois.SelectedIndex = -1
        $cbAnnee.SelectedIndex = -1
        $cbGrp.SelectedIndex = -1
        $cbOU.SelectedIndex = -1
        
        $lblResult.ForeColor = 'Green'
        $lblResult.Text = "Utilisateur $username cree avec mot de passe $password"
    }
    catch {
        $lblResult.ForeColor = 'Red'
        $lblResult.Text = "Erreur: $($_.Exception.Message)"
    }
})
$form.Controls.Add($btnSave)

$btnCancel = New-Object Windows.Forms.Button -Property @{Text="Annuler";Location='230,370'}
$btnCancel.Add_Click({ $form.Close() })
$form.Controls.Add($btnCancel)

$form.ShowDialog()
