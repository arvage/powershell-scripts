
## Convert tool for regular mailbox to shared or other way"
## This PS will gives you the option to convert a regular email to a shared mailbox and give give full access to another account for send/receive.
## it also comes with email forwarding in case you need to convert and forward incoming emails.
## 

## Make sure to update below parameters

##### Main Configuration #####
##############################
$domain = "your_domain_name"
$domainController = "your_DC_IP_OR_DNS_NAME"
$DCpath = "PATH_TO_OU"       ## e.g. "OU=SBS Users,OU=Users,OU=MyBusiness,DC=google,DC=com"
$exchangeServer = "DNSname_or_IP_of_your_exchange" 
##############################

# Graphical Form
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$MainForm = New-Object System.Windows.Forms.Form 
$MainForm.Text = "Email type convertion"
$MainForm.Size = New-Object System.Drawing.Size(320,320) 
$MainForm.Location.X = 250
$MainForm.Location.Y = 250
$MainForm.KeyPreview = $True

## form labels, listbox and textboxes
$ConvertLabel = New-Object System.Windows.Forms.Label
$ConvertLabel.Location = New-Object System.Drawing.Size(10,20) 
$ConvertLabel.Size = New-Object System.Drawing.Size(280,20) 
$ConvertLabel.Text = "Convert to:"
$ConvertLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)

$ConvertListBox = New-Object System.Windows.Forms.ListBox 
$ConvertListBox.Location = New-Object System.Drawing.Size(10,40) 
$ConvertListBox.Size = New-Object System.Drawing.Size(280,20) 
$ConvertListBox.Height = 40
[void] $ConvertListBox.Items.Add("Shared")
[void] $ConvertListBox.Items.Add("Regular")

$CredUserLabel = New-Object System.Windows.Forms.Label
$CredUserLabel.Location = New-Object System.Drawing.Size(10,85) 
$CredUserLabel.Size = New-Object System.Drawing.Size(280,13) 
$CredUserLabel.Text = "Your Admin Username: (without kagan\)"

$UserText = New-Object System.Windows.Forms.TextBox
$UserText.Location = New-Object System.Drawing.Size(10,100) 
$UserText.Size = New-Object System.Drawing.Size(280,20)

$MailboxLabel = New-Object System.Windows.Forms.Label
$MailboxLabel.Location = New-Object System.Drawing.Size(10,125) 
$MailboxLabel.Size = New-Object System.Drawing.Size(280,13) 
$MailboxLabel.Text = "Mailbox to Convert: (without @kaganonline.com)"

$MailboxText = New-Object System.Windows.Forms.TextBox
$MailboxText.Location = New-Object System.Drawing.Size(10,140) 
$MailboxText.Size = New-Object System.Drawing.Size(280,20)

$DelegationLabel = New-Object System.Windows.Forms.Label
$DelegationLabel.Location = New-Object System.Drawing.Size(10,165)
$DelegationLabel.Size = New-Object System.Drawing.Size(280,13)
$DelegationLabel.Text = "(Optional) Full Delegation Access to Username:"

$DelegationText = New-Object System.Windows.Forms.TextBox
$DelegationText.Location = New-Object System.Drawing.Size(10,180)
$DelegationText.Size = New-Object System.Drawing.Size(280,13)

$ForwardLabel = New-Object System.Windows.Forms.Label
$ForwardLabel.Location = New-Object System.Drawing.Size(10,205)
$ForwardLabel.Size = New-Object System.Drawing.Size(280,13)
$ForwardLabel.Text = "(Optional) Email to forward to:"

$ForwardText = New-Object System.Windows.Forms.TextBox
$ForwardText.Location = New-Object System.Drawing.Size(10,220)
$ForwardText.Size = New-Object System.Drawing.Size(280,13)

## convert button design
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(60,255)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "Convert"

## convert button action
$OKButton.Add_Click({
    if ($UserText.Text -eq '')
    {
    [System.Windows.MessageBox]::Show("Please provide your admin account","Error","Ok","Warning");
    }
    else
    {$userempty = 1}

    if ($MailboxText.Text -eq '')
    {
    [System.Windows.MessageBox]::Show("Please provide account for conversion","Error","Ok","Warning");
    }
    else
    {$emailempty=1}
    if ($emailempty -eq 1 -and $userempty -eq 1)
    {
    $credentials = "$domain\" + $UserText.Text;
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeServer/PowerShell/ -Authentication Kerberos -Credential $credentials;
    Import-PSSession $Session;
    Set-Mailbox -Identity $MailboxText.Text -Type $ConvertListBox.SelectedItem.ToString() | Out-Null;
    if ($DelegationText.Text -ne '')
    {
    Add-MailboxPermission -Identity $MailboxText.Text -User $DelegationText.Text -AccessRights FullAccess | Out-Null
    }
    if ($ForwardText.Text -ne '')
    {
    Set-Mailbox $MailboxText.Text -ForwardingAddress $ForwardText.Text -DeliverToMailboxAndForward $False -Force
    }


    ## Regular Account Selected
    if ($ConvertListBox.SelectedItem -ne 'Shared')
    {
        Enable-ADAccount -Identity $MailboxText.Text -Server $domainController -Confirm:$false;
        Get-ADUser $MailboxText.Text | Move-ADObject -Server $domainController -TargetPath $DCpath;
        $fulluser = Get-MailboxPermission -Identity $MailboxText.Text | Where-Object {($_.Isinherited -ne 'True' ) ` -and ($_.User -notlike '*SELF')} | select User;
        $user = $fulluser -creplace '^[^]]*\\', '' -creplace '.{1}$'
        if ([string]::IsNullOrEmpty($user))
            {Write-Host 'No Delegation Found'}
        else
            {Remove-MailboxPermission -Identity $MailboxText.Text -User $user -AccessRights FullAccess -Confirm:$false}
    }
    else
    {
        Disable-ADAccount -Identity $MailboxText.Text -Server $domainController -Confirm:$false;
        Get-ADUser $MailboxText.Text | Move-ADObject -Server $domainController -TargetPath $DCpath
    }
    Remove-PSSession $Session;
    $MainForm.Close();
    }

})

## if enter key pressed
$MainForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {
    if ($UserText.Text -eq '')
    {
    [System.Windows.MessageBox]::Show("Please provide your admin account","Error","Ok","Warning");
    }
    else
    {$userempty = 1}

    if ($MailboxText.Text -eq '')
    {
    [System.Windows.MessageBox]::Show("Please provide account for conversion","Error","Ok","Warning");
    }
    else
    {$emailempty=1}
    if ($emailempty -eq 1 -and $userempty -eq 1)
    {
        $credentials = "$domain\" + $UserText.Text;
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeServer/PowerShell/ -Authentication Kerberos -Credential $credentials;
        Import-PSSession $Session;
        Set-Mailbox -Identity $MailboxText.Text -Type $ConvertListBox.SelectedItem.ToString() | Out-Null;
    if ($DelegationText.Text -ne '')
        {
        Add-MailboxPermission -Identity $MailboxText.Text -User $DelegationText.Text -AccessRights FullAccess | Out-Null
        }
    if ($ForwardText.Text -ne '')
    {
    Set-Mailbox $MailboxText.Text -ForwardingAddress $ForwardText.Text -DeliverToMailboxAndForward $False -Force
    }


    ## Regular Account Selected
    if ($ConvertListBox.SelectedItem -ne 'Shared')
        {
        Enable-ADAccount -Identity $MailboxText.Text -Server $domainController -Confirm:$false;
        Get-ADUser $MailboxText.Text | Move-ADObject -Server $domainController -TargetPath $DCpath;
        $fulluser = Get-MailboxPermission -Identity $MailboxText.Text | Where-Object {($_.Isinherited -ne 'True' ) ` -and ($_.User -notlike '*SELF')} | select User;
        $user = $fulluser -creplace '^[^]]*\\', '' -creplace '.{1}$'
        if ([string]::IsNullOrEmpty($user))
            {Write-Host 'No Delegation Found'}
        else
            {Remove-MailboxPermission -Identity $MailboxText.Text -User $user -AccessRights FullAccess -Confirm:$false}
        }
    else
    {
        Disable-ADAccount -Identity $MailboxText.Text -Server $domainController -Confirm:$false;
        Get-ADUser $MailboxText.Text | Move-ADObject -Server $domainController -TargetPath $DCpath
    }
    Remove-PSSession $Session;
    $MainForm.Close();
    }
    }
})

## if esc key pressed
$MainForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$MainForm.Close()}})


## cancel button design
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(170,255)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$MainForm.Close()})
## adds created elements into form
$MainForm.Topmost = $True
$MainForm.Add_Shown({$MainForm.Activate()})
$MainForm.Controls.Add($ConvertLabel) 
$MainForm.Controls.Add($ConvertListBox) 
$MainForm.Controls.Add($CredUserLabel)
$MainForm.Controls.Add($UserText)
$MainForm.Controls.Add($MailboxLabel)
$MainForm.Controls.Add($MailboxText)
$MainForm.Controls.Add($DelegationLabel)
$MainForm.Controls.Add($DelegationText)
$MainForm.Controls.Add($ForwardLabel)
$MainForm.Controls.Add($ForwardText)
$MainForm.Controls.Add($OKButton)
$MainForm.Controls.Add($CancelButton)
## showing mainform 
$MainForm.ShowDialog()
