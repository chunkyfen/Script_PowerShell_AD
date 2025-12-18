Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

$form = New-Object Windows.Forms.Form
$form.Text = "New Employee"
$form.Size = '450,580'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# --- Section: Personal Information ---
$grpInfo = New-Object Windows.Forms.GroupBox
$grpInfo.Text = "Personal Information"
$grpInfo.Location = '15,10'
$grpInfo.Size = '410,180'
$form.Controls.Add($grpInfo)

# --- Last Name ---
$lblNom = New-Object Windows.Forms.Label -Property @{Text="Last Name :";Location='20,30';Size='100,20'}
$grpInfo.Controls.Add($lblNom)
$txtNom = New-Object Windows.Forms.TextBox -Property @{Location='130,28';Width=250}
$grpInfo.Controls.Add($txtNom)

# --- First Name ---
$lblPrenom = New-Object Windows.Forms.Label -Property @{Text="First Name :";Location='20,65';Size='100,20'}
$grpInfo.Controls.Add($lblPrenom)
$txtPrenom = New-Object Windows.Forms.TextBox -Property @{Location='130,63';Width=250}
$grpInfo.Controls.Add($txtPrenom)

# --- Group ---
$lblGrp = New-Object Windows.Forms.Label -Property @{Text="Group :";Location='20,100';Size='100,20'}
$grpInfo.Controls.Add($lblGrp)
$cbGrp = New-Object Windows.Forms.ComboBox -Property @{Location='130,98';Width=250;DropDownStyle='DropDownList'}
$cbGrp.Items.Add("Informatique") | Out-Null
$cbGrp.Items.Add("comptabilite") | Out-Null
$cbGrp.Items.Add("RH") | Out-Null
$grpInfo.Controls.Add($cbGrp)

# --- Date of Birth ---
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

# --- Generate Button ---
$btnGenerate = New-Object Windows.Forms.Button -Property @{
    Text="Generate User Name && Password"
    Location='80,210'
    Width=280
    Height=35
    Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))
}
$form.Controls.Add($btnGenerate)

# --- Section: Result ---
$grpResult = New-Object Windows.Forms.GroupBox
$grpResult.Text = "Result"
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

# --- Status Message ---
$lblStatus = New-Object Windows.Forms.Label -Property @{
    Text=""
    ForeColor='Green'
    Font=(New-Object Drawing.Font("Arial",9,[Drawing.FontStyle]::Bold))
    Location='20,105'
    Size='370,30'
    TextAlign='MiddleCenter'
}
$grpResult.Controls.Add($lblStatus)

# --- Save and Cancel Buttons ---
$btnSave = New-Object Windows.Forms.Button -Property @{
    Text="Save"
    Location='100,430'
    Width=120
    Height=35
    Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))
}
$form.Controls.Add($btnSave)

$btnCancel = New-Object Windows.Forms.Button -Property @{
    Text="Cancel"
    Location='230,430'
    Width=120
    Height=35
    Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Regular))
}
$form.Controls.Add($btnCancel)

# --- Generate Button Event ---
$btnGenerate.Add_Click({
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $annee = $cbAnnee.SelectedItem
    
    # Validation
    if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
        [Windows.Forms.MessageBox]::Show("Last name and first name are required!", "Error", 'OK', 'Error')
        return
    }
    
    if ($null -eq $annee) {
        [Windows.Forms.MessageBox]::Show("Year of birth is required!", "Error", 'OK', 'Error')
        return
    }
    
    # Generate User Name: first letter of first name + last name
    # Remove accents for username
    $prenomClean = $prenom
    $prenomClean = $prenomClean -replace '[àâäã]','a' -replace '[éèêë]','e' -replace '[îï]','i' -replace '[ôö]','o' -replace '[ùûü]','u' -replace '[ç]','c'
    $nomClean = $nom
    $nomClean = $nomClean -replace '[àâäã]','a' -replace '[éèêë]','e' -replace '[îï]','i' -replace '[ôö]','o' -replace '[ùûü]','u' -replace '[ç]','c'
    
    $username = ($prenomClean.Substring(0,1).ToLower() + $nomClean.ToLower())
    $username = $username -replace '[^a-z0-9]', ''
    
    # Generate Password: lastname + year + ?
    $password = ($nomClean.ToLower() + $annee + "?")
    
    # Display
    $lblUserName.Text = $username
    $lblPassword.Text = $password
    $lblStatus.Text = ""
})

# --- Save Button Event ---
$btnSave.Add_Click({
    $nom = $txtNom.Text.Trim()
    $prenom = $txtPrenom.Text.Trim()
    $jour = $cbJour.SelectedItem
    $mois = $cbMois.SelectedItem
    $annee = $cbAnnee.SelectedItem
    $groupName = $cbGrp.SelectedItem
    $username = $lblUserName.Text
    $password = $lblPassword.Text
    
    # Complete validation
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
    
    # Check if user already exists
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
        $lblStatus.Text = "Error verifying user"
        return
    }
    
    $securePass = ConvertTo-SecureString $password -AsPlainText -Force
    $dateNaissance = "$annee-$mois-$jour"
    
    try {
        # Use the Structure OU path
        $ouPath = "OU=Structure,DC=script,DC=local"
        
        # Create AD user
        $newUserParams = @{
            Name = "$prenom $nom"
            GivenName = $prenom
            Surname = $nom
            SamAccountName = $username
            UserPrincipalName = "$username@script.local"
            AccountPassword = $securePass
            Enabled = $true
            Path = $ouPath
            Description = "Date of birth: $dateNaissance - Group: $groupName"
            ChangePasswordAtLogon = $false
        }
        
        New-ADUser @newUserParams -ErrorAction Stop
        
        # Wait a moment for AD replication
        Start-Sleep -Seconds 2
        
        # Add to group
        try {
            Add-ADGroupMember -Identity $groupName -Members $username -ErrorAction Stop
            $groupStatus = "and added to group $groupName"
        } catch {
            $groupStatus = "but failed to add to group: $($_.Exception.Message)"
            Write-Host "Warning: $groupStatus" -ForegroundColor Yellow
        }
        
        # RESET: Reset all fields to initial state
        $txtNom.Clear()
        $txtPrenom.Clear()
        $cbJour.SelectedIndex = -1
        $cbMois.SelectedIndex = -1
        $cbAnnee.SelectedIndex = -1
        $cbGrp.SelectedIndex = -1
        $lblUserName.Text = ""
        $lblPassword.Text = ""
        
        $lblStatus.ForeColor = 'Green'
        $lblStatus.Text = "User created successfully!"
        
        [Windows.Forms.MessageBox]::Show("User $username created successfully $groupStatus!", "Success", 'OK', 'Information')
    }
    catch {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = "Error during creation"
        [Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", 'OK', 'Error')
    }
})

# --- Cancel Button Event ---
$btnCancel.Add_Click({ $form.Close() })

# Display the form
$form.ShowDialog()
