Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

$form = New-Object Windows.Forms.Form
$form.Text = "nouvel employé"
$form.Size = '400,450'

# --- OU principale ---
$lblOU = New-Object Windows.Forms.Label -Property @{Text="OU";Location='10,20'}
$form.Controls.Add($lblOU)

$cbOU = New-Object Windows.Forms.ComboBox -Property @{Location='120,20';Width=200}
(Get-ADOrganizationalUnit -Filter *).DistinguishedName | ForEach-Object { $cbOU.Items.Add($_) }
$form.Controls.Add($cbOU)

$btnAddOU = New-Object Windows.Forms.Button -Property @{Text="Ajouter OU";Location='330,20'}
$btnAddOU.Add_Click({
    $newOU = [System.Windows.Forms.InputBox]::Show("Nom de la nouvelle OU","Créer OU")
    if ($newOU) { New-ADOrganizationalUnit -Name $newOU -Path "DC=script,DC=local"
        $cbOU.Items.Add("OU=$newOU,DC=script,DC=local") }
})
$form.Controls.Add($btnAddOU)

# --- Groupe ---
$lblGrp = New-Object Windows.Forms.Label -Property @{Text="Groupe";Location='10,60'}
$form.Controls.Add($lblGrp)

$cbGrp = New-Object Windows.Forms.ComboBox -Property @{Location='120,60';Width=200}
(Get-ADGroup -Filter *).Name | ForEach-Object { $cbGrp.Items.Add($_) }
$form.Controls.Add($cbGrp)

$btnAddGrp = New-Object Windows.Forms.Button -Property @{Text="Ajouter Groupe";Location='330,60'}
$btnAddGrp.Add_Click({
    $newGrp = [System.Windows.Forms.InputBox]::Show("Nom du nouveau groupe","Créer Groupe")
    if ($newGrp) { New-ADGroup -Name $newGrp -GroupScope Global -Path "DC=script,DC=local"
        $cbGrp.Items.Add($newGrp) }
})
$form.Controls.Add($btnAddGrp)

# --- Champs utilisateur ---
$txtNom = New-Object Windows.Forms.TextBox -Property @{Location='120,100'}
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Nom";Location='10,100'}))
$form.Controls.Add($txtNom)

$txtPrenom = New-Object Windows.Forms.TextBox -Property @{Location='120,140'}
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Prénom";Location='10,140'}))
$form.Controls.Add($txtPrenom)

$cbJour = New-Object Windows.Forms.ComboBox -Property @{Location='120,180'}
1..31 | ForEach-Object { $cbJour.Items.Add($_.ToString("00")) }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Jour";Location='10,180'}))
$form.Controls.Add($cbJour)

$cbMois = New-Object Windows.Forms.ComboBox -Property @{Location='120,220'}
1..12 | ForEach-Object { $cbMois.Items.Add($_.ToString("00")) }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Mois";Location='10,220'}))
$form.Controls.Add($cbMois)

$cbAnnee = New-Object Windows.Forms.ComboBox -Property @{Location='120,260'}
1980..2005 | ForEach-Object { $cbAnnee.Items.Add($_) }
$form.Controls.Add((New-Object Windows.Forms.Label -Property @{Text="Année";Location='10,260'}))
$form.Controls.Add($cbAnnee)

# --- Résultat ---
$lblResult = New-Object Windows.Forms.Label -Property @{
    ForeColor='Red'; Font=(New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Bold));
    Location='10,300'; Size='350,40'
}
$form.Controls.Add($lblResult)

# --- Boutons ---
$btnGen = New-Object Windows.Forms.Button -Property @{Text="Générer";Location='10,360'}
$btnGen.Add_Click({
    if ($txtNom.Text -and $txtPrenom.Text -and $cbAnnee.SelectedItem) {
        $u = ($txtPrenom.Text.Substring(0,1).ToLower() + $txtNom.Text.ToLower())
        $p = ($txtNom.Text.ToLower() + $cbAnnee.SelectedItem + "?")
        $lblResult.Text = "UserName: $u `nPassword: $p"
    }
})
$form.Controls.Add($btnGen)

$btnSave = New-Object Windows.Forms.Button -Property @{Text="Enregistrer";Location='120,360'}
$btnSave.Add_Click({
    $nom = $txtNom.Text; $prenom = $txtPrenom.Text; $annee = $cbAnnee.SelectedItem
    $username = ($prenom.Substring(0,1).ToLower() + $nom.ToLower())
    $password = ($nom.ToLower() + $annee + "?")
    $securePass = ConvertTo-SecureString $password -AsPlainText -Force

    New-ADUser -Name "$prenom $nom" `
               -SamAccountName $username `
               -UserPrincipalName "$username@script.local" `
               -AccountPassword $securePass `
               -Enabled $true `
               -Path $cbOU.SelectedItem

    Add-ADGroupMember -Identity $cbGrp.SelectedItem -Members $username

    $txtNom.Clear(); $txtPrenom.Clear()
    $cbJour.SelectedIndex=$cbMois.SelectedIndex=$cbAnnee.SelectedIndex=$cbGrp.SelectedIndex=$cbOU.SelectedIndex=-1
    $lblResult.Text="Utilisateur $username créé"
})
$form.Controls.Add($btnSave)

$btnCancel = New-Object Windows.Forms.Button -Property @{Text="Annuler";Location='230,360'}
$btnCancel.Add_Click({ $form.Close() })
$form.Controls.Add($btnCancel)

$form.ShowDialog()

