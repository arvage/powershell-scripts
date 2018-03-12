## Automatic Reply Setup Tool"

##### Main Configuration #####
##############################
$domain = "your_domain_name"
$domainController = "your_DC_IP_OR_DNS_NAME"
$DCpath = "PATH_TO_OU"       ## e.g. "OU=SBS Users,OU=Users,OU=MyBusiness,DC=google,DC=com"
$exchangeServer = "XNG13CAS01" 
##############################


# Connect to Exchange Server
$UserCredential = Get-Credential -Credential $null
$Session = New-PSSession -Verbose -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeServer/PowerShell/ -Authentication Kerberos -Credential $UserCredential 
Import-PSSession -Verbose -DisableNameChecking $Session | Out-Null;


# Graphical Form
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$MainForm = New-Object System.Windows.Forms.Form 
$MainForm.Text = "Automatic Reply Config"
$MainForm.Size = New-Object System.Drawing.Size(320,250) 
$MainForm.Location.X = 250
$MainForm.Location.Y = 250
$MainForm.KeyPreview = $True

$UserEmailLabel = New-Object System.Windows.Forms.Label
$UserEmailLabel.Location = New-Object System.Drawing.Size(10,10) 
$UserEmailLabel.Size = New-Object System.Drawing.Size(280,13) 
$UserEmailLabel.Text = "Email:"

$UserEmail = New-Object System.Windows.Forms.TextBox
$UserEmail.Location = New-Object System.Drawing.Size(10,30) 
$UserEmail.Size = New-Object System.Drawing.Size(280,20)

$AutoReplyLabel = New-Object System.Windows.Forms.Label
$AutoReplyLabel.Location = New-Object System.Drawing.Size(10,60) 
$AutoReplyLabel.Size = New-Object System.Drawing.Size(280,13) 
$AutoReplyLabel.Text = "Auto Reply Message:   (Need <br> at end of each line)"

$AutoReplyText = New-Object System.Windows.Forms.TextBox
$AutoReplyText.Location = New-Object System.Drawing.Size(10,80) 
$AutoReplyText.Size = New-Object System.Drawing.Size(280,60)
$AutoReplyText.Multiline = $True

$EnabledLabel = New-Object System.Windows.Forms.Label
$EnabledLabel.Location = New-Object System.Drawing.Size(10,150) 
$EnabledLabel.Size = New-Object System.Drawing.Size(50,13) 
$EnabledLabel.Text = "Enabled:"

$EnabledStatus = New-Object System.Windows.Forms.CheckBox
$EnabledStatus.Location = New-Object System.Drawing.Size(60,150) 
$EnabledStatus.Size = New-Object System.Drawing.Size(280,13)
$EnabledStatus.Checked = $True

$EnabledStatus.Add_CheckStateChanged({
if ($EnabledStatus.Checked)
    { 
     $AutoReplyText.Enabled = $true 
    }
else
    {
     $AutoReplyText.Enabled = $false
    }
})

$MainForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
        {
        $AutoReplyText.Text += "<br>" 
        $AutoReplyText.SelectionStart = $AutoReplyText.Text.Length
        $AutoReplyText.SelectionLength = 0
        }
})
## apply button design
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(60,180)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "Apply"

## cancel button design
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(170,180)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$MainForm.Close()})


## convert button action
$OKButton.Add_Click({
  if ($EnabledStatus.Checked)
  {
    if ($UserEmail.Text -eq "")
    {
    [System.Windows.Forms.MessageBox]::Show("Email Empty!","Error","Ok","Warning");
    }
    if ($AutoReplyText.Text -eq "")
    {
    [System.Windows.Forms.MessageBox]::Show("Auto Reply Message Empty!","Error","Ok","Warning");
    }
    if ($UserEmail.Text -ne "" -and $AutoReplyText.Text -ne "")
    {
    $answer = [System.Windows.Forms.MessageBox]::Show("Activating " + $UserEmail.Text + " Auto Reply?","Confirm" ,"OkCancel","Warning");
        if ($answer -eq 'Ok')
        {
        
        Set-MailboxAutoReplyConfiguration $UserEmail.Text -AutoReplyState Enabled –ExternalMessage $AutoReplyText.Text –InternalMessage $AutoReplyText.Text
        Remove-PSSession $Session;
        [System.Windows.Forms.MessageBox]::Show("Activated!","Success","Ok","Information");
        $MainForm.Close();
        }
    }
  }
  if (!$EnabledStatus.Checked)
  {
  if ($UserEmail.Text -eq "")
    {
    [System.Windows.Forms.MessageBox]::Show("Email Empty!","Error","Ok","Warning");
    }
    if ($UserEmail.Text -ne "")
    {
    $answer = [System.Windows.Forms.MessageBox]::Show("Disabling " + $UserEmail.Text + " Auto Reply?","Confirm" ,"OkCancel","Warning");
        if ($answer -eq 'Ok')
        {
        Set-MailboxAutoReplyConfiguration $UserEmail.Text -AutoReplyState Disabled –ExternalMessage $null –InternalMessage $null
        Remove-PSSession $Session;
        [System.Windows.Forms.MessageBox]::Show("Deactivated!","Success","Ok","Information");
        $MainForm.Close();
        }
    }
  }
})

## if esc key pressed
$MainForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$MainForm.Close()}})

## adds created elements into form
$MainForm.Topmost = $True
$MainForm.Add_Shown({$MainForm.Activate()})
$MainForm.Controls.Add($UserEmailLabel)
$MainForm.Controls.Add($UserEmail)
$MainForm.Controls.Add($AutoReplyLabel)
$MainForm.Controls.Add($AutoReplyText)
$MainForm.Controls.Add($EnabledLabel)
$MainForm.Controls.Add($EnabledStatus)
$MainForm.Controls.Add($OKButton)
$MainForm.Controls.Add($CancelButton)

## showing mainform 
$MainForm.ShowDialog()


