Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

$form = New-Object Windows.Forms.Form
$form.Text = "Nouvel employe"
$form.Size = '400,520'
$form.StartPosition = 'CenterScreen'

# --- OU principale ---
$lblOU = New-Object Windows.Forms.Label -Property @{Text="OU";Location='10,20';AutoSize=$true}
$form.Controls.Add($lblOU)
$cbOU = New-Object Windows.Forms.ComboBox -Property @{Location='120,20';Width=250;DropDownStyle='DropDownList'}
try {
    (Get-ADOrganizationalUnit -Filter *).DistinguishedName | ForEach-Object { $cbOU.Items.Add($_) | Out-Null }
} catch {
    [Windows.Forms.MessageBox]::Show("Erreur chargement OUs: $($_.Exception.Message)", "Erreur")
}
$form.Controls.Add($cbOU)

# --- Groupe ---
$lblGrp = New-Object Windows.Forms.Label -Property @{Text="Groupe";Location='10,60';AutoSize=$true}
$form.Controls.Add($lblGrp)
$cbGrp = New-Object Windows.Forms.ComboBox -Property @{Location='120,60';Width=250;DropDownStyle='DropDownList'}
try {
    (Get-ADGroup -Filter *).Name | ForEach-Object { $cbGrp.Items.Add($_) | Out-Null }
} catch {
    [Windows.Forms.MessageBox]::Show("Erreur chargement groupes: $($_.Exception.Message)", "Erreur")
}
$form.Controls.Add($cbGrp)

# --- Champs utilisateur ---
$txtNom = New-Object Windows.Forms.TextBox -Property @{Location='120,100';Width=250}
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Nom";Location='10,100';AutoSize=$true}))
$form.Controls.Add($txtNom)

$txtPrenom = New-Object Windows.Forms.TextBox -Property @{Location='120,140';Width=250}
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Prenom";Location='10,140';AutoSize=$true}))
$form.Controls.Add($txtPrenom)

# --- Bloc Date de naissance ---
$lblDate = New-Object Windows.Forms.Label -Property @{Text="Date de naissance";Location='10,180';AutoSize=$true}
$form.Controls.Add($lblDate)

$cbJour = New-Object Windows.Forms.ComboBox -Property @{Location='120,200';Width=60;DropDownStyle='DropDownList'}
1..31 | ForEach-Object { $cbJour.Items.Add($_.ToString("00")) | Out-Null }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Jour";Location='10,200';AutoSize=$true}))
$form.Controls.Add($cbJour)

$cbMois = New-Object Windows.Forms.ComboBox -Property @{Location='120,240';Width=60;DropDownStyle='DropDownList'}
1..12 | ForEach-Object { $cbMois.Items.Add($_.ToString("00")) | Out-Null }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Mois";Location='10,240';AutoSize=$true}))
$form.Controls.Add($cbMois)

$cbAnnee = New-Object Windows.Forms.ComboBox -Property @{Location='120,280';Width=80;DropDownStyle='DropDownList'}
1910..2010 | ForEach-Object { $cbAnnee.Items.Add($_) | Out-Null }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Annee";Location='10,280';AutoSize=$true}))
$form.Controls.Add($cbAnnee)

# --- Resultat ---
$lblResult = New-Object Windows.Forms.Label -Property @{
    ForeColor='Red'; Font=(New-Object Drawing.Font("Arial",9,[Drawing.FontStyle]::Bold));
    Location='10,320'; Size='370,60'; BorderStyle='FixedSingle'
}
$form.Controls.Add($lblResult)

# --- Boutons ---
$btnSave = New-Object Windows.Forms.Button -Property @{Text="Enregistrer";Location='120,400';Width=100}
$btnSave.Add_Click({
    $lblResult.Text = "Traitement en cours..."
    $lblResult.ForeColor = 'Blue'
    $form.Refresh()
    
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $jour = $cbJour.SelectedItem
    $mois = $cbMois.SelectedItem
    $annee = $cbAnnee.SelectedItem
    $ouPath = $cbOU.SelectedItem
    $groupName = $cbGrp.SelectedItem
    
    # Validation des champs requis
    if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
        $lblResult.ForeColor = 'Red'
        $lblResult.Text = "ERREUR: Nom et prenom requis!"
        return
    }
    
    if ($null -eq $ouPath) {
        $lblResult.ForeColor = 'Red'
        $lblResult.Text = "ERREUR: OU requise!"
        return
    }
    
    if ($null -eq $groupName) {
        $lblResult.ForeColor = 'Red'
        $lblResult.Text = "ERREUR: Groupe requis!"
        return
    }
    
    if ($null -eq $jour -or $null -eq $mois -or $null -eq $annee) {
        $lblResult.ForeColor = 'Red'
        $lblResult.Text = "ERREUR: Date de naissance complete requise!"
        return
    }
    
    # Génération du nom d'utilisateur
    $username = ($prenom.Substring(0,1).ToLower() + $nom.ToLower())
    $username = $username -replace '[^a-z0-9]', ''  # Supprimer caractères spéciaux
    
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
        }
    } catch {
        $lblResult.ForeColor = 'Red'
        $lblResult.Text = "ERREUR verification utilisateur: $($_.Exception.Message)"
        return
    }
    
    # Génération du mot de passe
    $password = ($nom.Substring(0,1).ToUpper() + $nom.Substring(1).ToLower() + $annee + "!")
    $securePass = ConvertTo-SecureString $password -AsPlainText -Force
    
    # Date de naissance formatée
    $dateNaissance = "$annee-$mois-$jour"
    
    try {
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
            Description = "Date de naissance: $dateNaissance"
            ChangePasswordAtLogon = $false
        }
        
        New-ADUser @newUserParams -ErrorAction Stop
        
        # Petite pause pour s'assurer que l'utilisateur est créé
        Start-Sleep -Seconds 2
        
        # Vérifier que l'utilisateur a bien été créé
        $createdUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction Stop
        if (-not $createdUser) {
            throw "Utilisateur non trouve apres creation"
        }
        
        # Ajout au groupe
        try {
            Add-ADGroupMember -Identity $groupName -Members $username -ErrorAction Stop
            $groupStatus = " et ajoute au groupe $groupName"
        } catch {
            $groupStatus = " MAIS erreur ajout groupe: $($_.Exception.Message)"
        }
        
        # Reset
        $txtNom.Clear()
        $txtPrenom.Clear()
        $cbJour.SelectedIndex = -1
        $cbMois.SelectedIndex = -1
        $cbAnnee.SelectedIndex = -1
        $cbGrp.SelectedIndex = -1
        $cbOU.SelectedIndex = -1
        
        $lblResult.ForeColor = 'Green'
        $lblResult.Text = "SUCCES: Utilisateur $username cree$groupStatus`nMot de passe: $password"
    }
    catch {
        $lblResult.ForeColor = 'Red'
        $errorMsg = $_.Exception.Message
        $lblResult.Text = "ERREUR creation: $errorMsg"
        
        # Log détaillé dans la console PowerShell
        Write-Host "=== ERREUR DETAILLEE ===" -ForegroundColor Red
        Write-Host "Message: $errorMsg" -ForegroundColor Red
        Write-Host "Username: $username" -ForegroundColor Yellow
        Write-Host "OU: $ouPath" -ForegroundColor Yellow
        Write-Host "Groupe: $groupName" -ForegroundColor Yellow
        Write-Host $_.Exception.GetType().FullName -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
    }
})
$form.Controls.Add($btnSave)

$btnCancel = New-Object Windows.Forms.Button -Property @{Text="Annuler";Location='230,400';Width=100}
$btnCancel.Add_Click({ $form.Close() })
$form.Controls.Add($btnCancel)

# Vérifier les permissions au démarrage
try {
    $domain = Get-ADDomain
    $lblResult.ForeColor = 'Green'
    $lblResult.Text = "Connecte a: $($domain.DNSRoot)"
} catch {
    $lblResult.ForeColor = 'Red'
    $lblResult.Text = "ATTENTION: Probleme connexion AD"
}

$form.ShowDialog()
