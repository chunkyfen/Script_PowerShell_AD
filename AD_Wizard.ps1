Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Import-Module ActiveDirectory

# Helper function to create controls
function New-Control($type, $props) {
    $ctrl = New-Object "System.Windows.Forms.$type"
    $props.GetEnumerator() | ForEach-Object { $ctrl.$($_.Key) = $_.Value }
    return $ctrl
}

# Main Form
$form = New-Control Form @{
    Text = "New Employee"; Size = '450,580'; StartPosition = 'CenterScreen'
    FormBorderStyle = 'FixedDialog'; MaximizeBox = $false
}

# Personal Information Group
$grpInfo = New-Control GroupBox @{Text = "Personal Information"; Location = '15,10'; Size = '410,180'}
$form.Controls.Add($grpInfo)

# Input Controls
$txtNom = New-Control TextBox @{Location = '130,28'; Width = 250}
$txtPrenom = New-Control TextBox @{Location = '130,63'; Width = 250}
$cbGrp = New-Control ComboBox @{Location = '130,98'; Width = 250; DropDownStyle = 'DropDownList'}
@("Informatique", "comptabilite", "RH") | ForEach-Object { $cbGrp.Items.Add($_) | Out-Null }

$cbJour = New-Control ComboBox @{Location = '130,133'; Width = 60; DropDownStyle = 'DropDownList'}
$cbMois = New-Control ComboBox @{Location = '205,133'; Width = 60; DropDownStyle = 'DropDownList'}
$cbAnnee = New-Control ComboBox @{Location = '280,133'; Width = 100; DropDownStyle = 'DropDownList'}

1..31 | ForEach-Object { $cbJour.Items.Add($_.ToString("00")) | Out-Null }
1..12 | ForEach-Object { $cbMois.Items.Add($_.ToString("00")) | Out-Null }
1980..2005 | ForEach-Object { $cbAnnee.Items.Add($_) | Out-Null }

# Add labels and controls to group
$grpInfo.Controls.AddRange(@(
    (New-Control Label @{Text = "Last Name :"; Location = '20,30'; Size = '100,20'}),
    $txtNom,
    (New-Control Label @{Text = "First Name :"; Location = '20,65'; Size = '100,20'}),
    $txtPrenom,
    (New-Control Label @{Text = "Group :"; Location = '20,100'; Size = '100,20'}),
    $cbGrp,
    (New-Control Label @{Text = "Date of Birth :"; Location = '20,135'; Size = '100,20'}),
    $cbJour,
    (New-Control Label @{Text = "/"; Location = '195,135'; Size = '10,20'}),
    $cbMois,
    (New-Control Label @{Text = "/"; Location = '270,135'; Size = '10,20'}),
    $cbAnnee
))

# Generate Button
$btnGenerate = New-Control Button @{
    Text = "Generate User Name && Password"; Location = '80,210'; Size = '280,35'
    Font = New-Object Drawing.Font("Arial", 10)
}
$form.Controls.Add($btnGenerate)

# Result Group
$grpResult = New-Control GroupBox @{Text = "Result"; Location = '15,260'; Size = '410,150'}
$form.Controls.Add($grpResult)

$lblUserName = New-Control TextBox @{
    Location = '130,28'; Width = 250; BackColor = 'White'
    ForeColor = 'Red'; Font = New-Object Drawing.Font("Arial", 11, [Drawing.FontStyle]::Bold)
}
$lblPassword = New-Control Label @{
    Location = '130,68'; Size = '250,25'; BackColor = 'White'; BorderStyle = 'FixedSingle'
    ForeColor = 'Red'; Font = New-Object Drawing.Font("Arial", 11, [Drawing.FontStyle]::Bold)
    TextAlign = 'MiddleLeft'
}
$lblStatus = New-Control Label @{
    Location = '20,105'; Size = '370,30'; TextAlign = 'MiddleCenter'
    ForeColor = 'Green'; Font = New-Object Drawing.Font("Arial", 9, [Drawing.FontStyle]::Bold)
}

$grpResult.Controls.AddRange(@(
    (New-Control Label @{Text = "User Name :"; Location = '20,30'; Size = '100,20'}),
    $lblUserName,
    (New-Control Label @{Text = "Password :"; Location = '20,70'; Size = '100,20'}),
    $lblPassword,
    $lblStatus
))

# Action Buttons
$btnSave = New-Control Button @{
    Text = "Save"; Location = '100,430'; Size = '120,35'
    Font = New-Object Drawing.Font("Arial", 10)
}
$btnCancel = New-Control Button @{
    Text = "Cancel"; Location = '230,430'; Size = '120,35'
    Font = New-Object Drawing.Font("Arial", 10)
}
$form.Controls.AddRange(@($btnSave, $btnCancel))

# Helper function to clean accents
function Remove-Accents($text) {
    $text -replace '[àâäã]','a' -replace '[éèêë]','e' -replace '[îï]','i' -replace '[ôö]','o' -replace '[ùûü]','u' -replace '[ç]','c'
}

# Generate Button Click
$btnGenerate.Add_Click({
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
    
    $prenomClean = Remove-Accents $prenom
    $nomClean = Remove-Accents $nom
    
    $lblUserName.Text = ($prenomClean[0].ToString().ToLower() + $nomClean.ToLower()) -replace '[^a-z0-9]', ''
    $lblPassword.Text = $nomClean.ToLower() + $annee + "?"
    $lblStatus.Text = ""
})

# Save Button Click
$btnSave.Add_Click({
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $jour = $cbJour.SelectedItem
    $mois = $cbMois.SelectedItem
    $annee = $cbAnnee.SelectedItem
    $groupName = $cbGrp.SelectedItem
    $username = $lblUserName.Text.Trim()
    $password = $lblPassword.Text
    
    # Validation
    if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
        [Windows.Forms.MessageBox]::Show("Last name and first name are required!", "Error", 'OK', 'Error')
        return
    }
    if ($null -eq $groupName) {
        [Windows.Forms.MessageBox]::Show("Group is required!", "Error", 'OK', 'Error')
        return
    }
    if ($null -eq $jour -or $null -eq $mois -or $null -eq $annee) {
        [Windows.Forms.MessageBox]::Show("Complete date of birth is required!", "Error", 'OK', 'Error')
        return
    }
    if ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($password)) {
        [Windows.Forms.MessageBox]::Show("Please generate the User Name and Password first!", "Error", 'OK', 'Error')
        return
    }
    
    # Check if user exists
    try {
        if (Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue) {
            [Windows.Forms.MessageBox]::Show("Username '$username' already exists! Please use a different name.", "Error", 'OK', 'Error')
            return
        }
    } catch {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = "Error verifying user"
        return
    }
    
    # Create user
    try {
        $newUserParams = @{
            Name = "$prenom $nom"
            GivenName = $prenom
            Surname = $nom
            SamAccountName = $username
            UserPrincipalName = "$username@script.local"
            AccountPassword = (ConvertTo-SecureString $password -AsPlainText -Force)
            Enabled = $true
            Path = "OU=Structure,DC=script,DC=local"
            Description = "Date of birth: $annee-$mois-$jour - Group: $groupName"
            ChangePasswordAtLogon = $false
        }
        
        New-ADUser @newUserParams -ErrorAction Stop
        Start-Sleep -Seconds 2
        
        $groupStatus = "and added to group $groupName"
        try {
            Add-ADGroupMember -Identity $groupName -Members $username -ErrorAction Stop
        } catch {
            $groupStatus = "but failed to add to group"
        }
        
        # Clear form
        $txtNom.Clear(); $txtPrenom.Clear()
        $cbJour.SelectedIndex = $cbMois.SelectedIndex = $cbAnnee.SelectedIndex = $cbGrp.SelectedIndex = -1
        $lblUserName.Text = $lblPassword.Text = ""
        
        $lblStatus.ForeColor = 'Green'
        $lblStatus.Text = "User created successfully!"
        [Windows.Forms.MessageBox]::Show("User $username created successfully $groupStatus!", "Success", 'OK', 'Information')
    }
    catch {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = "Error during creation"
        [Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message, "Error", 'OK', 'Error')
    }
})

$btnCancel.Add_Click({ $form.Close() })

$form.ShowDialog() | Out-Null
