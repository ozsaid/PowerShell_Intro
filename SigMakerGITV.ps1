using namespace System.Windows.Forms
using namespace System.Drawing
 #Global Varibles
# Define the Outlook version
$outlookVersion = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' | Select-Object -ExpandProperty VersionToReport).split(".")[0]+".0"

# Combine the registry path
$registryPath = "HKCU:\Software\Microsoft\Office\$outlookVersion\Outlook\"

# Get the default Outlook profile name from the registry
$defaultProfileName = (Get-ItemProperty -Path $registryPath -Name "DefaultProfile").DefaultProfile
$registryPath = "HKCU:\Software\Microsoft\Office\$outlookVersion\Outlook\Profiles\"+$defaultProfileName+"\9375CFF0413111d3B88A00104B2A6676"
$TEST=(Get-ChildItem $registryPath).Name.Replace("HKEY_CURRENT_USER","HKCU:")
$emailRegex = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
foreach ($PT in $TEST) {
                        $CHK=Get-ItemPropertyValue -path $PT -PSProperty "Account Name" 
                        if ($CHK -match $emailRegex) {$GLOBAL:RK=$PT
                                                      $global:UN=$chk
                                                      break
                                                      }
                       } 
#Connect to Graph and CheckVaribles

$tenantId = "TENANT ID "
$client_id = "APP ID"
$client_secret = "APP SECRET"
$resource = "https://graph.microsoft.com"

# Construct token request
$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/token"
$body = @{
    grant_type    = "client_credentials"
    client_id     = $client_id
    client_secret = $client_secret
    resource      = $resource
}

# Get access token
$response = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body

# Extract access token from the response
$accessToken = $response.access_token


# Define your Graph API endpoint and user ID
$graphApiEndpoint = "https://graph.microsoft.com/beta/users/$UN"

# Define your access token (replace with your actual access token)

# Make the API request
try {
$userData = Invoke-RestMethod -Uri $graphApiEndpoint -Headers @{
    Authorization = "Bearer $accessToken"
} -Method Get
    }
    catch {Write-host "Cannot get Data"}



#Generated Form Function
function GenerateForm {

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$EMAIL = New-Object System.Windows.Forms.TextBox
$office = New-Object System.Windows.Forms.TextBox
$mobile = New-Object System.Windows.Forms.TextBox
$title = New-Object System.Windows.Forms.TextBox
$LastName = New-Object System.Windows.Forms.TextBox
$FirstName = New-Object System.Windows.Forms.TextBox
$label6 = New-Object System.Windows.Forms.Label
$label5 = New-Object System.Windows.Forms.Label
$label4 = New-Object System.Windows.Forms.Label
$label3 = New-Object System.Windows.Forms.Label
$label2 = New-Object System.Windows.Forms.Label
$label1 = New-Object System.Windows.Forms.Label
$SEND = New-Object System.Windows.Forms.Button
$OUTLOOK = New-Object System.Windows.Forms.Button
$CANCEL = New-Object System.Windows.Forms.Button
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$Reply = New-Object System.Windows.Forms.CheckBox
$Send1 = New-Object System.Windows.Forms.CheckBox
$form1.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$CANCEL_OnClick= 
{
#TODO: Place custom script here
$form1.Close()
}

$handler_SEND_Click= 
{
$HTM='<p><strong><span style="font-size: medium;">%%FirstName%% %%LastName%%</span></strong><br /><span style="font-size: small;">%%Title%%<br /><span style="color: #0099ff;">Office:</span>%%Phone%% &nbsp; <span style="color: #99cc00;"><strong>|</strong> </span><span style="color: #0099ff;">Mobile:</span> %%MobilePhone%%<br /><span style="text-decoration: underline; color: #0000ff; font-size: smaller;"><span>%%WindowsEmailAddress%%</span></span></span></p>
<p><a href="https://www.Mrcloud.com/"><img style="float: left;" src="data:image/jpeg;base64,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" /></a></p>'

#TODO: Place custom script here
#TODO: Place custom script here
$global:FN=$FirstName.Text
$global:LN=$LastName.Text
$global:OF=$office.Text
$global:MB=$mobile.text
$global:TL=$title.Text
$global:EM=$EMAIL.Text
$HTM=$HTM.Replace('%%FirstName%%',$FirstName.text)
$HTM=$HTM.Replace('%%LastName%%',$LastName.Text)
$HTM=$HTM.Replace('%%Title%%',$title.Text)
$HTM=$HTM.Replace('%%MobilePhone%%',$mobile.Text)
$HTM=$HTM.Replace("%%Phone%%",$office.Text)
$HTM=$HTM.Replace("%%WindowsEmailAddress%%",$EMAIL.text)
if ($office.Text -eq $null -or $office.Text-eq "") {$HTM=$HTM.Replace("Mobile:","") 
                                                    $HTM=$HTM.Replace("Office:","Mobile:") 
                                                    $HTM=$HTM.Replace("%%PHONE%%",$mobile.Text)
                                                    $HTM=$HTM.Replace("%%MobilePhone%%","") 
                                                    }
if ($mobile.Text -eq $null -or $mobile.Text-eq "") {$HTM=$HTM.Replace("Mobile:","") 
                                                    $HTM=$HTM.Replace("%%MobilePhone%%","") 
                                                    }
 $global:HTM=$HTM                                                    

# PSv5+:
# Import namespaces so that types can be referred by
# their mere name (e.g., `Form` rather than `System.Windows.Forms.Form`)
#
'<table id="signWrapper" border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;" width="100%"><tbody><tr><td align="left" valign="top"><table align="left" border="0" cellpadding="0" cellspacing="0" style="max-width:600px; border-collapse:collapse;" width="600"><tbody><tr><td align="left" style="padding: 10px; background-color: rgb(255, 255, 255);" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" style="max-width:600px; width:100%; border-collapse:collapse;" width="600"><tbody><tr><td align="left" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;" width="100%"><tbody><tr class="signRow"><td><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;" width="100%"><tbody><tr><td style="font-family: arial; font-size: 12px; line-height: 18px; padding: 0px; vertical-align: middle; text-align: left;" width="50%"><span style="font-size: 14px; color: rgb(0, 0, 0); font-family: Arial, Helvetica, sans-serif;"><strong>%%FIRSTNAME%%  %%LASTNAME%%</strong></span><br><span style="font-size: 12px; color: rgb(0, 0, 0); font-family: Arial, Helvetica, sans-serif;">%%TITLE%%</span><span style="font-size: 12px;"><br></span></td><td style="font-family: arial; font-size: 12px; line-height: 18px; padding: 0px; text-align: right; vertical-align: top;" width="50%"><br></td></tr></tbody></table></td></tr>
<tr class="signRow"><td height="10" style="font-size:10px; line-height:10px;"> </td></tr>

<tr class="signRow"><td height="10" style="font-size: 10px; line-height: 10px; vertical-align: bottom;"><strong> </strong><span style="font-size: 12px; color: rgb(84, 172, 210);"><strong>Office</strong></span><strong>: </strong><strong><span style="color: rgb(0, 0, 0);"><span style="font-size: 12px;">%%Phone%%</span>  </span></strong><span style="color: rgb(97, 189, 109);"> |   </span><span style="color: rgb(84, 172, 210);"><strong><span style="font-size: 12px;">Mobile </span></strong></span><span style="color: rgb(0, 0, 0);"><strong><span style="font-size: 12px;">:%%MobilePhone%%</span></strong></span><br><span style="color: rgb(0, 0, 0);"><strong><span style="font-size: 12px;">​</span></strong></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 12px;"><strong> %%WindowsEmailAddress%%</strong></span></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 12px;"><strong>​</strong></span></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 14px; line-height: 1.5;"><strong><img src="https://a.bybrand.io/img-658422c280966.jpg" style="display: inline-block; vertical-align: middle; margin: 0px; float: none; max-width: 600px; text-align: center; padding: 0px; border: none; width: 113px; height: 40px;" width="113" height="40">   </strong></span></span><span style="color: rgb(124, 112, 107);"><span style="font-size: 14px; line-height: 1.5;"><a href="https://www.Mrcloud.co.il/" target="_blank" rel="noopener noreferrer" style="text-decoration: none solid transparent; color: inherit;"><strong>www.Mrcloud.co.il</strong></a></span></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 14px;"><strong>​</strong></span></span><span style="font-size: 14px;"><br></span><span style="color: rgb(44, 130, 201);"><span style="font-size: 12px;"><strong>​</strong></span></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 12px;"><strong> </strong><strong><br></strong></span></span></td></tr><tr class="signRow"><td height="10" style="font-size:10px; line-height:10px;"><strong> </strong></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table>'

# Load the WinForms assembly.
Add-Type -AssemblyName System.Windows.Forms

# Create a form.
$form = [Form] @{
    ClientSize      = [Point]::new(400, 400)
    Text            = "WebBrowser-Control Demo"
}

# Create a web-browser control, make it as large as the inside of the form,
# and assign the HTML text.
# PSv5+:
# Import namespaces so that types can be referred by
# their mere name (e.g., `Form` rather than `System.Windows.Forms.Form`)
#
'<table id="signWrapper" border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;" width="100%"><tbody><tr><td align="left" valign="top"><table align="left" border="0" cellpadding="0" cellspacing="0" style="max-width:600px; border-collapse:collapse;" width="600"><tbody><tr><td align="left" style="padding: 10px; background-color: rgb(255, 255, 255);" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" style="max-width:600px; width:100%; border-collapse:collapse;" width="600"><tbody><tr><td align="left" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;" width="100%"><tbody><tr class="signRow"><td><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;" width="100%"><tbody><tr><td style="font-family: arial; font-size: 12px; line-height: 18px; padding: 0px; vertical-align: middle; text-align: left;" width="50%"><span style="font-size: 14px; color: rgb(0, 0, 0); font-family: Arial, Helvetica, sans-serif;"><strong>%%FIRSTNAME%%  %%LASTNAME%%</strong></span><br><span style="font-size: 12px; color: rgb(0, 0, 0); font-family: Arial, Helvetica, sans-serif;">%%TITLE%%</span><span style="font-size: 12px;"><br></span></td><td style="font-family: arial; font-size: 12px; line-height: 18px; padding: 0px; text-align: right; vertical-align: top;" width="50%"><br></td></tr></tbody></table></td></tr>
<tr class="signRow"><td height="10" style="font-size:10px; line-height:10px;"> </td></tr>

<tr class="signRow"><td height="10" style="font-size: 10px; line-height: 10px; vertical-align: bottom;"><strong> </strong><span style="font-size: 12px; color: rgb(84, 172, 210);"><strong>Office</strong></span><strong>: </strong><strong><span style="color: rgb(0, 0, 0);"><span style="font-size: 12px;">%%Phone%%</span>  </span></strong><span style="color: rgb(97, 189, 109);"> |   </span><span style="color: rgb(84, 172, 210);"><strong><span style="font-size: 12px;">Mobile </span></strong></span><span style="color: rgb(0, 0, 0);"><strong><span style="font-size: 12px;">:%%MobilePhone%%</span></strong></span><br><span style="color: rgb(0, 0, 0);"><strong><span style="font-size: 12px;">​</span></strong></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 12px;"><strong> %%WindowsEmailAddress%%</strong></span></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 12px;"><strong>​</strong></span></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 14px; line-height: 1.5;"><strong><img src="https://a.bybrand.io/img-658422c280966.jpg" style="display: inline-block; vertical-align: middle; margin: 0px; float: none; max-width: 600px; text-align: center; padding: 0px; border: none; width: 113px; height: 40px;" width="113" height="40">   </strong></span></span><span style="color: rgb(124, 112, 107);"><span style="font-size: 14px; line-height: 1.5;"><a href="https://www.Mrcloud.co.il/" target="_blank" rel="noopener noreferrer" style="text-decoration: none solid transparent; color: inherit;"><strong>www.Mrcloud.co.il</strong></a></span></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 14px;"><strong>​</strong></span></span><span style="font-size: 14px;"><br></span><span style="color: rgb(44, 130, 201);"><span style="font-size: 12px;"><strong>​</strong></span></span><br><span style="color: rgb(44, 130, 201);"><span style="font-size: 12px;"><strong> </strong><strong><br></strong></span></span></td></tr><tr class="signRow"><td height="10" style="font-size:10px; line-height:10px;"><strong> </strong></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table>'

# Load the WinForms assembly.
Add-Type -AssemblyName System.Windows.Forms

# Create a form.
$form = [Form] @{
    ClientSize      = [Point]::new(400, 400)
    Text            = "WebBrowser-Control Demo"
}

# Create a web-browser control, make it as large as the inside of the form,
# and assign the HTML text.
$sb = [WebBrowser] @{
  ClientSize = $form.ClientSize
  DocumentText = $HTM
}



# Add the web-browser control to the form...
$form.Controls.Add($sb)

# ... and display the form as a dialog (synchronously).
$form.ShowDialog()

# Clean up.
$form.Dispose()
}

$handler_OUTLOOK_Click= 
{
$global:R=$Reply.Checked
$global:S=$Send1.Checked
$global:FN=$FirstName.Text
$global:LN=$LastName.Text
$global:OF=$office.Text
$global:MB=$mobile.text
$global:TL=$title.Text
$global:EM=$EMAIL.Text
$HTM=$HTM.Replace('%%FirstName%%',$FirstName.text)
$HTM=$HTM.Replace('%%LastName%%',$LastName.Text)
$HTM=$HTM.Replace('%%Title%%',$title.Text)
$HTM=$HTM.Replace('%%MobilePhone%%',$mobile.Text)
$HTM=$HTM.Replace("%%Phone%%",$office.Text)
$HTM=$HTM.Replace("%%WindowsEmailAddress%%",$EMAIL.text)
if ($office.Text -eq $null -or $office.Text-eq "") {$HTM=$HTM.Replace("Mobile:","") 
                                                    $HTM=$HTM.Replace("Office:","Mobile:") 
                                                    $HTM=$HTM.Replace("%%PHONE%%",$mobile.Text)
                                                    $HTM=$HTM.Replace("%%MobilePhone%%","") 
                                                    }
if ($mobile.Text -eq $null -or $mobile.Text-eq "") {$HTM=$HTM.Replace("Mobile:","") 
                                                    $HTM=$HTM.Replace("%%MobilePhone%%","") 
                                                    }
 $global:HTM=$HTM                                                    

if ($global:r)
    { $propertyName = 'Reply-Forward Signature'
      $propertyValue = 'Mrcloud'
      New-ItemProperty -Path $global:rk -Name $propertyName -Value $propertyValue -PropertyType String -Force
      }
if ($global:S)
       { $propertyName = 'New Signature'
         $propertyValue = 'Mrcloud'
       New-ItemProperty -Path $global:rk -Name $propertyName -Value $propertyValue -PropertyType String -Force
      }


$HTM | out-file $env:appdata\Microsoft\Signatures\MrCloud.htm

[System.Windows.Forms.MessageBox]::Show('Signiture Was Save to Outlook ')
return
}


#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 572
$System_Drawing_Size.Width = 480
$form1.ClientSize = $System_Drawing_Size
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.Name = "form1"
$form1.Text = "MyCloudSecSigniture"

$EMAIL.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 125
$System_Drawing_Point.Y = 254
$EMAIL.Location = $System_Drawing_Point
$EMAIL.Name = "EMAIL"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 203
$EMAIL.Size = $System_Drawing_Size
$EMAIL.TabIndex = 13
$EMAIL.Text=$userData.mail
$email.Add_TextChanged({
    if (-not [regex]::IsMatch($email.Text, '^[a-zA-Z0-9-+&@_.!()\s]+$')) {
        $email.BackColor = 'lightpink'  # Set the background color to indicate validation error
        [System.Windows.Forms.MessageBox]::Show("Only English and Numbers Allowed", "שגיאת הזנה", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $email.Clear()
        $email.BackColor = 'White'
    } else {
        $email.BackColor = 'White'      # Set the background color back to normal
    }
})

$form1.Controls.Add($EMAIL)

$office.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 125
$System_Drawing_Point.Y = 212
$office.Location = $System_Drawing_Point
$office.Name = "office"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 204
$office.Size = $System_Drawing_Size
$office.TabIndex = 12
$office.Text=$userData.businessPhones
$office.Add_TextChanged({
    if (-not [regex]::IsMatch($office.Text, '^[0-9\-+]+$') -or $office.Text -eq $null) {
        $office.BackColor = 'LightPink'  # Set the background color to indicate validation error
        [System.Windows.Forms.MessageBox]::Show("+-,0-9 Allowed", "שגיאת הזנה", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $office.Clear()
        $office.BackColor = 'White'
    } else {
        $office.BackColor = 'White'      # Set the background color back to normal
    }
    })

$form1.Controls.Add($office)

$mobile.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 125
$System_Drawing_Point.Y = 167
$mobile.Location = $System_Drawing_Point
$mobile.Name = "mobile"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 204
$mobile.Size = $System_Drawing_Size
$mobile.TabIndex = 11
$mobile.Text=$userData.mobilePhone
$mobile.Add_TextChanged({
    if (-not [regex]::IsMatch($mobile.Text, '^[0-9\-+]+$') -or $mobile.Text -eq $null) {
        $mobile.BackColor = 'LightPink'  # Set the background color to indicate validation error
        [System.Windows.Forms.MessageBox]::Show("+-,0-9 Allowed", "שגיאת הזנה", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $mobile.Clear()
        $mobile.BackColor = 'White'
    } else {
        $mobile.BackColor = 'White'      # Set the background color back to normal
    }
    })

$form1.Controls.Add($mobile)

$title.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 127
$System_Drawing_Point.Y = 121
$title.Location = $System_Drawing_Point
$title.Name = "title"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 202
$title.Size = $System_Drawing_Size
$title.TabIndex = 10
$title.Text=$userData.jobTitle
$title.Add_TextChanged({
    if (-not [regex]::IsMatch($title.Text, '^[a-zA-Z0-9-+&\s]+$')) {
        $title.BackColor = 'lightpink'  # Set the background color to indicate validation error
        [System.Windows.Forms.MessageBox]::Show("Only English and Numbers Allowed", "שגיאת הזנה", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $title.Clear()
        $title.BackColor = 'White'
    } else {
        $title.BackColor = 'White'      # Set the background color back to normal
    }
})

$form1.Controls.Add($title)

$LastName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 127
$System_Drawing_Point.Y = 77
$LastName.Location = $System_Drawing_Point
$LastName.Name = "LastName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 202
$LastName.Size = $System_Drawing_Size
$LastName.TabIndex = 9
$LastName.Text=$userData.surname
$lastname.Add_TextChanged({
    if (-not [regex]::IsMatch($lastname.Text, '^[a-zA-Z0-9-+&\s]+$')) {
        $lastname.BackColor = 'lightpink'  # Set the background color to indicate validation error
        [System.Windows.Forms.MessageBox]::Show("Only English and Numbers Allowed", "שגיאת הזנה", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $lastname.Clear()
        $lastname.BackColor = 'White'
    } else {
        $lastname.BackColor = 'White'      # Set the background color back to normal
    }
})

$form1.Controls.Add($LastName)

$FirstName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 129
$System_Drawing_Point.Y = 34
$FirstName.Location = $System_Drawing_Point
$FirstName.Name = "FirstName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$FirstName.Size = $System_Drawing_Size
$FirstName.TabIndex = 8
$FirstName.Text=$userData.givenName
$firstname.Add_TextChanged({
    if (-not [regex]::IsMatch($firstname.Text, '^[a-zA-Z0-9-+&\s]+$')) {
        $firstname.BackColor = 'lightpink'  # Set the background color to indicate validation error
        [System.Windows.Forms.MessageBox]::Show("Only English and Numbers Allowed", "שגיאת הזנה", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $firstname.Clear()
        $firstname.BackColor = 'White'
    } else {
        $firstname.BackColor = 'White'      # Set the background color back to normal
    }
})

$form1.Controls.Add($FirstName)

$label6.DataBindings.DefaultDataSourceUpdateMode = 0
$label6.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 26
$System_Drawing_Point.Y = 257
$label6.Location = $System_Drawing_Point
$label6.Name = "label6"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 26
$System_Drawing_Size.Width = 122
$label6.Size = $System_Drawing_Size
$label6.TabIndex = 7
$label6.Text = "E-mail:"

$form1.Controls.Add($label6)

$label5.DataBindings.DefaultDataSourceUpdateMode = 0
$label5.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 26
$System_Drawing_Point.Y = 215
$label5.Location = $System_Drawing_Point
$label5.Name = "label5"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$label5.Size = $System_Drawing_Size
$label5.TabIndex = 6
$label5.Text = "Office Phone:"

$form1.Controls.Add($label5)

$label4.DataBindings.DefaultDataSourceUpdateMode = 0
$label4.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 26
$System_Drawing_Point.Y = 165
$label4.Location = $System_Drawing_Point
$label4.Name = "label4"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 27
$System_Drawing_Size.Width = 86
$label4.Size = $System_Drawing_Size
$label4.TabIndex = 5
$label4.Text = "Mobile:"

$form1.Controls.Add($label4)

$label3.DataBindings.DefaultDataSourceUpdateMode = 0
$label3.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 26
$System_Drawing_Point.Y = 121
$label3.Location = $System_Drawing_Point
$label3.Name = "label3"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 28
$System_Drawing_Size.Width = 92
$label3.Size = $System_Drawing_Size
$label3.TabIndex = 4
$label3.Text = "Title:"

$form1.Controls.Add($label3)

$label2.DataBindings.DefaultDataSourceUpdateMode = 0
$label2.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 26
$System_Drawing_Point.Y = 77
$label2.Location = $System_Drawing_Point
$label2.Name = "label2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$label2.Size = $System_Drawing_Size
$label2.TabIndex = 3
$label2.Text = "Last Name:"

$form1.Controls.Add($label2)

$label1.DataBindings.DefaultDataSourceUpdateMode = 0
$label1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 32
$label1.Location = $System_Drawing_Point
$label1.Name = "label1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 32
$System_Drawing_Size.Width = 137
$label1.Size = $System_Drawing_Size
$label1.TabIndex = 2
$label1.Text = "First Name:"

$form1.Controls.Add($label1)

$SEND.BackColor = [System.Drawing.Color]::FromArgb(255,0,255,0)

$SEND.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 170
$System_Drawing_Point.Y = 470
$SEND.Location = $System_Drawing_Point
$SEND.Name = "SEND"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 50
$System_Drawing_Size.Width = 134
$SEND.Size = $System_Drawing_Size
$SEND.TabIndex = 1
$SEND.Text = "PREVIEW"
$SEND.UseVisualStyleBackColor = $False
$SEND.add_Click($handler_SEND_Click)

$form1.Controls.Add($SEND)

$OUTLOOK.BackColor = [System.Drawing.Color]::FromArgb(175,0,175,0)

$OUTLOOK.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 310
$System_Drawing_Point.Y = 470
$OUTLOOK.Location = $System_Drawing_Point
$OUTLOOK.Name = "OUTLOOK"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 50
$System_Drawing_Size.Width = 134
$OUTLOOK.Size = $System_Drawing_Size
$OUTLOOK.TabIndex = 1
$OUTLOOK.Text = "APPLY"
$OUTLOOK.UseVisualStyleBackColor = $False
$OUTLOOK.add_Click($handler_OUTLOOK_Click)

$form1.Controls.Add($OUTLOOK)

$CANCEL.BackColor = [System.Drawing.Color]::FromArgb(255,184,134,11)

$CANCEL.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 20
$System_Drawing_Point.Y = 470
$CANCEL.Location = $System_Drawing_Point
$CANCEL.Name = "CANCEL"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 50
$System_Drawing_Size.Width = 134
$CANCEL.Size = $System_Drawing_Size
$CANCEL.TabIndex = 0
$CANCEL.Text = "CANCEL"
$CANCEL.UseVisualStyleBackColor = $False
$CANCEL.add_Click($CANCEL_OnClick)

$form1.Controls.Add($CANCEL)
$Reply.DataBindings.DefaultDataSourceUpdateMode = 0
$Reply.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9.75,1,3,0)
$Reply.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 313
$Reply.Location = $System_Drawing_Point
$Reply.Name = "Reply"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 32
$System_Drawing_Size.Width = 124
$Reply.Size = $System_Drawing_Size
$Reply.TabIndex = 1
$Reply.Text = "Default Reply"
$Reply.UseVisualStyleBackColor = $True

$form1.Controls.Add($Reply)

$Send1.DataBindings.DefaultDataSourceUpdateMode = 0
$Send1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",10,1,3,0)
$Send1.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 360
$Send1.Location = $System_Drawing_Point
$Send1.Name = "Send1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 35
$System_Drawing_Size.Width = 126
$Send1.Size = $System_Drawing_Size
$Send1.TabIndex = 0
$Send1.Text = "Default Send"
$Send1.UseVisualStyleBackColor = $True

$form1.Controls.Add($Send1)

#add an image to the form using Base64
$base64ImageString = "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"
$imageBytes = [Convert]::FromBase64String($base64ImageString)
$ms = New-Object IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$ms.Write($imageBytes, 0, $imageBytes.Length);
$alkanelogo = [System.Drawing.Image]::FromStream($ms, $true)

$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $alkanelogo.Size.Width;
$pictureBox.Height =  $alkanelogo.Size.Height; 
$pictureBox.Location = New-Object System.Drawing.Size(145,300) 
$pictureBox.Image = $alkanelogo;
$form1.Controls.Add($pictureBox)

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null

} #End Function

#Call the Function
GenerateForm
