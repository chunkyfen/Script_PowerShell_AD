Add-Type -AssemblyName System.Windows.Forms   # Charge la librairie Windows Forms
Add-Type -AssemblyName System.Drawing         # Charge la librairie pour polices et couleurs
Import-Module ActiveDirectory                 # Charge le module Active Directory

$form = New-Object Windows.Forms.Form
$form.Text = "Nouvel employe"                 # Titre de la fenetre
$form.Size = '400,480'                        # Taille de la fenetre

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
$btnGen = New-Object Windows.Forms.Button -Property @{Text="Generer";Location='10,370'}
$btnGen.Add_Click({
    if ($txtNom.Text -and $txtPrenom.Text -and $cbAnnee.SelectedItem) {
        $u = ($txtPrenom.Text.Substring(0,1).ToLower() + $txtNom.Text.ToLower())
        $p = ($txtNom.Text.ToLower() + $cbAnnee.SelectedItem + "?")
        $lblResult.Text = "UserName: $u `nPassword: $p"
    }
})
$form.Controls.Add($btnGen)

$btnSave = New-Object Windows.Forms.Button -Property @{Text="Enregistrer";Location='120,370'}
$btnSave.Add_Click({
    $nom = $txtNom.Text
    $prenom = $txtPrenom.Text
    $annee = $cbAnnee.SelectedItem
    $username = ($prenom.Substring(0,1).ToLower() + $nom.ToLower())
    $password = ($nom.ToLower() + $annee + "?")
    $securePass = ConvertTo-SecureString $password -AsPlainText -Force

    # Creation utilisateur AD (une seule ligne)
    New-ADUser -Name "$prenom $nom" -SamAccountName $username -UserPrincipalName "$username@script.local" -AccountPassword $securePass -Enabled $true -Path $cbOU.SelectedItem

    # Ajout au groupe
    Add-ADGroupMember -Identity $cbGrp.SelectedItem -Members $username

    # Reset
    $txtNom.Clear(); $txtPrenom.Clear()
    $cbJour.SelectedIndex=$cbMois.SelectedIndex=$cbAnnee.SelectedIndex=$cbGrp.SelectedIndex=$cbOU.SelectedIndex=-1
    $lblResult.Text="Utilisateur $username cree"
})
$form.Controls.Add($btnSave)

$btnCancel = New-Object Windows.Forms.Button -Property @{Text="Annuler";Location='230,370'}
$btnCancel.Add_Click({ $form.Close() })
$form.Controls.Add($btnCancel)

$form.ShowDialog()
