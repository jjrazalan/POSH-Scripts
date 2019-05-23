<# 
.SYNOPSIS
GUI for ADUser creation, modification, and remote terminal page.

.DESCRIPTION
Create ADUsers tab prompts for necessary fields to create users and then will call the createUser functions to set variables and call the actual create users function for AD.
Update Users prompts for username or has filters to search for by surname and first name. It then will pull data from AD where you can update the values. 
Allows admins to easily enter a computer name and access it's remote terminal.

.AUTHOR
    Josh Razalan
#>
function createUser {
    Import-Module ActiveDirectory

    #ensure username is consistent
    $Username = $Usernamef.text.tolower()
    $Desc = ''

    #Read user data from each field form and assign the data to a variable as below
    $Firstname = $Firstnamef.text
    $MI = $MIf.text
    $Lastname = $Lastnamef.text
    $Password = $Passwordf.text
    $Office = $Officef.text
    $Department = $Departmentf.text
    $JobTitle = $Jobtitlef.text
    $O365 = $O365f.Checked
    $License = $Licensef.SelectedItem
    
    #Set desciption by user intials
    $tempDesc = $Firstname + ' ' + $MI + ' ' + $Lastname	
    $tempDesc.split(' ') | ForEach-Object {$Desc += $_[0]}
    $Desc = $Desc.ToUpper()  

    #Set correct syntax for Display Names.
    if (!$MI) {
        $Name = $Firstname + ' ' + $Lastname
    }
    else { 
        $Name = $Firstname + ' ' + $MI + '. ' + $Lastname
    }
    
    #Connect to Microsoft Online
    Get-MsolDomain -ErrorAction SilentlyContinue
    if ($?) {
    
    }
    else {
        Connect-MsolService
    }
    
    #Check if user exists then call methods to create user in AD
    if (-Not (Get-ADUser -Filter 'SamAccountName -eq $Username')) {
        try{
            . Functions.ps1
            createUser $Firstname $MI $Lastname $Username $Password $Office $Department $JobTitle $Desc $Name 
            . AddO365Users.ps1
            CreateO365User $Firstname $Lastname $Username $Password $Name $O365 $License $Office $Department $JobTitle
            $Label10.text = "$Username sucessfully created"
        }
        catch{
            $Label10.text = "$Username creation unsucessful"
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("$Username already exists", 'Error', 'Ok', 'Error')
        $Label10.text = "$Username already exists"
    }

    #Clears out form boxes for multiple entries
    $Firstnamef.text = ''
    $MIf.text = ''
    $Lastnamef.text = ''
    $Usernamef.text = ''
    $Passwordf.text = ''
    $Officef.SelectedItem = $null
    $Departmentf.text = ''
    $JobTitlef.text = ''
    $O365f.Checked = $false
}

function searchUser {
    Import-Module ActiveDirectory

    $User = ""
    $Search = $Searchf.text
    $Filter = $Filterf.text

    #Switch determines what type of filter it will use
    switch ($Filter) {
        "name" {
            $Search = $Search + "*"
            $User = Get-ADUser -Filter 'name -like $Search' -Properties *
        }
        "surname" {
            $Search = $Search + "*"
            $User = Get-ADUser -Filter 'SurName -like $Search' -Properties *
        }
        "username" {
            $User = Get-ADUser -Filter 'SamAccountName -eq $Search' -Properties *
        }
    }
    
    if ($User -eq "") {
        $namel.text = "User not found."
    }
    #Fill in form textboxes with user information
    else {
        $namel.text = $User.Name
        $Phonel.text = "Phone:"
        $Phonef.text = $User.OfficePhone
        $Phonef.Visible = $true
        $Jobtitlel.text = "Job Title:"
        $Jobtitlef1.text = $User.Title
        $Jobtitlef1.Visible = $true
        $Officel.text = "Office:"
        $Officef1.text = $User.Office
        $Officef1.Visible = $true
        $Departmentl.text = "Department:"
        $Departmentf1.text = $User.Department
        $Departmentf1.Visible = $true
        $Descl.text = "Description:"
        $Descf.text = $User.Description
        $Descf.Visible = $true
        $DNl.text = $User.DistinguishedName
        $Update.Visible = $true

    }
}

function updateUser {
    if ($Jobtitlef1.text -eq "" -or $Officef1.text -eq "" -or $Descf.text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Values are missing!", 'Error', 'Ok', 'Error')
    }
    else {
        try {
            if ($Phonef.text -eq "") {
                $Phonef.text = " "
            }
            if ($Departmentf1.text -eq "") {
                $Departmentf1.text = " "
            }
            Set-ADUser -Identity $DNl.text `
                -OfficePhone $Phonef.text `
                -Title $Jobtitlef1.text `
                -Office $Officef1.text `
                -Department $Departmentf1.text

            $namel.text = "Update successful"
            $Phonel.text = ""
            $Phonef.text = " "
            $Phonef.Visible = $false
            $Jobtitlel.text = ""
            $Jobtitlef1.text = " "
            $Jobtitlef1.Visible = $false
            $Officel.text = ""
            $Officef1.text = " "
            $Officef1.Visible = $false
            $Departmentl.text = ""
            $Departmentf1.text = " "
            $Departmentf1.Visible = $false
            $Descl.text = ""
            $Descf.text = " "
            $Descf.Visible = $false
            $Update.Visible = $false
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Could not set User's values", 'Error', 'Ok', 'Error')
        }
    } 
}

function generateForm {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    #Initialize tab controls
    $tabControl = New-Object System.Windows.Forms.TabControl
    $CreateUserPage = New-Object System.Windows.Forms.TabPage
    $GetUserPage = New-Object System.Windows.Forms.TabPage
    $RemoteTerminalPage = New-Object System.Windows.Forms.TabPage
    $tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
    $tabControl.Location = New-Object System.Drawing.Point(5, 0)
    $tabControl.Name = "tabControl"

    #begin GUI{ 

    $CreateADUsers = New-Object system.Windows.Forms.Form
    $CreateADUsers.ClientSize = '785,540'
    $CreateADUsers.text = "IT AD tools"
    $CreateADUsers.TopMost = $false

    #Uncomment and add file path for icon
    #$Icon = New-Object system.drawing.icon (".ico")
    #$CreateADUsers.Icon = $Icon

    #Create tabs
    $tabControl.Size = '775,535'
    $tabControl.Font = 'Microsoft Sans Serif,12'
    $CreateADUsers.Controls.Add($tabControl)



    #Tab for AD Users Creation
    $CreateUserPage.Text = "Create AD Users"
    $tabControl.Controls.Add($CreateUserPage)

    $Label1 = New-Object system.Windows.Forms.Label
    $Label1.text = "AD Users Creation:"
    $Label1.AutoSize = $true
    $Label1.width = 30
    $Label1.height = 10
    $Label1.location = New-Object System.Drawing.Point(6, 5)
    $Label1.Font = 'Microsoft Sans Serif,18,style=Bold'


    $Label2 = New-Object system.Windows.Forms.Label
    $Label2.text = "First Name:"
    $Label2.AutoSize = $true
    $Label2.width = 30
    $Label2.height = 10
    $Label2.location = New-Object System.Drawing.Point(6, 75)
    $Label2.Font = 'Microsoft Sans Serif,14'

    $Firstnamef = New-Object system.Windows.Forms.TextBox
    $Firstnamef.multiline = $false
    $Firstnamef.width = 185
    $Firstnamef.height = 20
    $Firstnamef.location = New-Object System.Drawing.Point(146, 75)
    $Firstnamef.Font = 'Microsoft Sans Serif,14'

    $Label4 = New-Object system.Windows.Forms.Label
    $Label4.text = "Last Name:"
    $Label4.AutoSize = $true
    $Label4.width = 30
    $Label4.height = 10
    $Label4.location = New-Object System.Drawing.Point(440, 75)
    $Label4.Font = 'Microsoft Sans Serif,14'

    $Label3 = New-Object system.Windows.Forms.Label
    $Label3.text = "MI:"
    $Label3.AutoSize = $true
    $Label3.width = 20
    $Label3.height = 10
    $Label3.location = New-Object System.Drawing.Point(340, 75)
    $Label3.Font = 'Microsoft Sans Serif,14'

    $MIf = New-Object system.Windows.Forms.TextBox
    $MIf.multiline = $false
    $MIf.width = 35
    $MIf.height = 20
    $MIf.location = New-Object System.Drawing.Point(395, 75)
    $MIf.Font = 'Microsoft Sans Serif,14'
    $MIf.MaxLength = 1

    $Lastnamef = New-Object system.Windows.Forms.TextBox
    $Lastnamef.multiline = $false
    $Lastnamef.width = 185
    $Lastnamef.height = 20
    $Lastnamef.location = New-Object System.Drawing.Point(581, 75)
    $Lastnamef.Font = 'Microsoft Sans Serif,14'

    $Label5 = New-Object system.Windows.Forms.Label
    $Label5.text = "Username:"
    $Label5.AutoSize = $true
    $Label5.width = 25
    $Label5.height = 10
    $Label5.location = New-Object System.Drawing.Point(6, 150)
    $Label5.Font = 'Microsoft Sans Serif,14'

    $Usernamef = New-Object system.Windows.Forms.TextBox
    $Usernamef.multiline = $false
    $Usernamef.Visible = $true
    $Usernamef.width = 185
    $Usernamef.height = 20
    $Usernamef.location = New-Object System.Drawing.Point(146, 150)
    $Usernamef.Font = 'Microsoft Sans Serif,14'

    $Label6 = New-Object system.Windows.Forms.Label
    $Label6.text = "Password:"
    $Label6.AutoSize = $true
    $Label6.width = 25
    $Label6.height = 10
    $Label6.location = New-Object System.Drawing.Point(340, 150)
    $Label6.Font = 'Microsoft Sans Serif,14'

    $Passwordf = New-Object system.Windows.Forms.TextBox
    $Passwordf.multiline = $false
    $Passwordf.width = 185
    $Passwordf.height = 20
    $Passwordf.location = New-Object System.Drawing.Point(500, 150)
    $Passwordf.Font = 'Microsoft Sans Serif,14'

    $Label7 = New-Object system.Windows.Forms.Label
    $Label7.text = "Office:"
    $Label7.AutoSize = $true
    $Label7.width = 25
    $Label7.height = 10
    $Label7.location = New-Object System.Drawing.Point(6, 225)
    $Label7.Font = 'Microsoft Sans Serif,14'

    $Officef = New-Object system.Windows.Forms.ComboBox
    $Officef.width = 185
    $Officef.height = 20
    $Officef.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    <#$Officef.AutoCompleteSource = 'ListItems'
    $Officef.AutoCompleteMode = 'Append'#>
    $Officef.Items.AddRange(@('Colorado'))
    $Officef.location = New-Object System.Drawing.Point(146, 225)
    $Officef.Font = 'Microsoft Sans Serif,14'

    $Label8 = New-Object system.Windows.Forms.Label
    $Label8.text = "Department:"
    $Label8.AutoSize = $true
    $Label8.width = 25
    $Label8.height = 10
    $Label8.location = New-Object System.Drawing.Point(340, 225)
    $Label8.Font = 'Microsoft Sans Serif,14'

    $Departmentf = New-Object system.Windows.Forms.TextBox
    $Departmentf.multiline = $false
    $Departmentf.width = 185
    $Departmentf.height = 20
    $Departmentf.location = New-Object System.Drawing.Point(500, 225)
    $Departmentf.Font = 'Microsoft Sans Serif,14'

    $Label9 = New-Object system.Windows.Forms.Label
    $Label9.text = "Job Title:"
    $Label9.AutoSize = $true
    $Label9.width = 25
    $Label9.height = 10
    $Label9.location = New-Object System.Drawing.Point(6, 300)
    $Label9.Font = 'Microsoft Sans Serif,14'

    $Jobtitlef = New-Object system.Windows.Forms.TextBox
    $Jobtitlef.multiline = $false
    $Jobtitlef.width = 185
    $Jobtitlef.height = 20
    $Jobtitlef.location = New-Object System.Drawing.Point(146, 300)
    $Jobtitlef.Font = 'Microsoft Sans Serif,14'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.text = "Submit"
    $Submit.width = 125
    $Submit.height = 55
    $Submit.location = New-Object System.Drawing.Point(6, 375)
    $Submit.Font = 'Microsoft Sans Serif,14'

    $O365f = New-Object system.Windows.Forms.CheckBox
    $O365f.text = "Assign O365 License"
    $O365f.AutoSize = $false
    $O365f.width = 220
    $O365f.height = 28
    $O365f.location = New-Object System.Drawing.Point(340, 303)
    $O365f.Font = 'Microsoft Sans Serif,14'

    $Licensef = New-Object system.Windows.Forms.ComboBox
    $Licensef.width = 185
    $Licensef.height = 20
    $Licensef.Visible = $false
    $Licensef.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $Licensef.Items.AddRange(@('E3', 'ProPlus', 'E1'))
    $Licensef.SelectedItem = $Licensef.Items[0]
    $Licensef.location = New-Object System.Drawing.Point(581, 300)
    $Licensef.Font = 'Microsoft Sans Serif,14'

    $Open = New-Object system.Windows.Forms.Button
    $Open.text = "Open CSV"
    $Open.width = 125
    $Open.height = 55
    $Open.location = New-Object System.Drawing.Point(150, 375)
    $Open.Font = 'Microsoft Sans Serif,14'

    $Create = New-Object system.Windows.Forms.Button
    $Create.text = "Create with CSV"
    $Create.width = 125
    $Create.height = 55
    $Create.location = New-Object System.Drawing.Point(300, 375)
    $Create.Font = 'Microsoft Sans Serif,14'

    $Label10 = New-Object system.Windows.Forms.Label
    $Label10.text = ""
    $Label10.AutoSize = $true
    $Label10.width = 25
    $Label10.height = 10
    $Label10.location = New-Object System.Drawing.Point(6, 450)
    $Label10.Font = 'Microsoft Sans Serif,12'

    $Label11 = New-Object system.Windows.Forms.Label
    $Label11.text = ""
    $Label11.AutoSize = $true
    $Label11.width = 25
    $Label11.height = 10
    $Label11.location = New-Object System.Drawing.Point(6, 475)
    $Label11.Font = 'Microsoft Sans Serif,12'

    $CreateUserPage.controls.AddRange(@($Label1, $Label2, $Firstnamef, $Label3, $Label4, $MIf, $Lastnamef, $Label5, $Usernamef,
            $Label6, $Passwordf, $Label7, $Officef, $Label8, $Departmentf, $Label9, $Jobtitlef, $O365f, $Licensef, $Submit, $Open, $Create, $Label10, $Label11))


    #Get AD User page
    $GetUserPage.Text = "Get AD Users"
    $tabControl.Controls.Add($GetUserPage)

    $Label1 = New-Object system.Windows.Forms.Label
    $Label1.text = "Search:"
    $Label1.AutoSize = $true
    $Label1.width = 30
    $Label1.height = 10
    $Label1.location = New-Object System.Drawing.Point(6, 15)
    $Label1.Font = 'Microsoft Sans Serif,14'

    $filterf = New-Object system.Windows.Forms.ComboBox
    $filterf.width = 125
    $filterf.height = 20
    $filterf.text = 'username'
    $filterf.AutoCompleteSource = 'ListItems'
    $filterf.AutoCompleteMode = 'Append'
    $filterf.Items.AddRange(@('username', 'surname', 'name'))
    $filterf.location = New-Object System.Drawing.Point(146, 15)
    $filterf.Font = 'Microsoft Sans Serif,14'

    $Searchf = New-Object system.Windows.Forms.TextBox
    $Searchf.multiline = $false
    $Searchf.width = 300
    $Searchf.height = 20
    $Searchf.location = New-Object System.Drawing.Point(270, 15)
    $Searchf.Font = 'Microsoft Sans Serif,16'

    $Search = New-Object system.Windows.Forms.Button
    $Search.text = "Search"
    $Search.width = 125
    $Search.height = 34
    $Search.location = New-Object System.Drawing.Point(610, 14)
    $Search.Font = 'Microsoft Sans Serif,14'

    $namel = New-Object system.Windows.Forms.Label
    $namel.text = ""
    $namel.AutoSize = $true
    $namel.width = 30
    $namel.height = 10
    $namel.location = New-Object System.Drawing.Point(6, 75)
    $namel.Font = 'Microsoft Sans Serif,14'

    $Phonel = New-Object system.Windows.Forms.Label
    $Phonel.text = ""
    $Phonel.AutoSize = $true
    $Phonel.width = 25
    $Phonel.height = 10
    $Phonel.location = New-Object System.Drawing.Point(6, 150)
    $Phonel.Font = 'Microsoft Sans Serif,14'

    $Phonef = New-Object system.Windows.Forms.TextBox
    $Phonef.multiline = $false
    $Phonef.Visible = $false
    $Phonef.text = " "
    $Phonef.width = 185
    $Phonef.height = 20
    $Phonef.location = New-Object System.Drawing.Point(146, 150)
    $Phonef.Font = 'Microsoft Sans Serif,14'

    $Jobtitlel = New-Object system.Windows.Forms.Label
    $Jobtitlel.text = ""
    $Jobtitlel.AutoSize = $true
    $Jobtitlel.width = 25
    $Jobtitlel.height = 10
    $Jobtitlel.location = New-Object System.Drawing.Point(340, 150)
    $Jobtitlel.Font = 'Microsoft Sans Serif,14'

    $Jobtitlef1 = New-Object system.Windows.Forms.TextBox
    $Jobtitlef1.multiline = $false
    $Jobtitlef1.Visible = $false
    $Jobtitlef1.width = 185
    $Jobtitlef1.height = 20
    $Jobtitlef1.location = New-Object System.Drawing.Point(500, 150)
    $Jobtitlef1.Font = 'Microsoft Sans Serif,14'

    $Officel = New-Object system.Windows.Forms.Label
    $Officel.text = ""
    $Officel.AutoSize = $true
    $Officel.width = 25
    $Officel.height = 10
    $Officel.location = New-Object System.Drawing.Point(6, 225)
    $Officel.Font = 'Microsoft Sans Serif,14'

    $Officef1 = New-Object system.Windows.Forms.ComboBox
    $Officef1.width = 185
    $Officef1.height = 20
    $Officef1.Visible = $false
    $Officef1.AutoCompleteSource = 'ListItems'
    $Officef1.AutoCompleteMode = 'Append'
    $Officef1.Items.AddRange(@('Colorado'))
    $Officef1.location = New-Object System.Drawing.Point(146, 225)
    $Officef1.Font = 'Microsoft Sans Serif,14'

    $Departmentl = New-Object system.Windows.Forms.Label
    $Departmentl.text = ""
    $Departmentl.AutoSize = $true
    $Departmentl.width = 25
    $Departmentl.height = 10
    $Departmentl.location = New-Object System.Drawing.Point(340, 225)
    $Departmentl.Font = 'Microsoft Sans Serif,14'

    $Departmentf1 = New-Object system.Windows.Forms.TextBox
    $Departmentf1.multiline = $false
    $Departmentf1.Visible = $false
    $Departmentf1.text = " "
    $Departmentf1.width = 185
    $Departmentf1.height = 20
    $Departmentf1.location = New-Object System.Drawing.Point(500, 225)
    $Departmentf1.Font = 'Microsoft Sans Serif,14'

    $Descl = New-Object system.Windows.Forms.Label
    $Descl.text = ""
    $Descl.AutoSize = $true
    $Descl.width = 25
    $Descl.height = 10
    $Descl.location = New-Object System.Drawing.Point(6, 300)
    $Descl.Font = 'Microsoft Sans Serif,14'

    $Descf = New-Object system.Windows.Forms.TextBox
    $Descf.multiline = $false
    $Descf.Visible = $false
    $Descf.width = 185
    $Descf.height = 20
    $Descf.location = New-Object System.Drawing.Point(146, 300)
    $Descf.Font = 'Microsoft Sans Serif,14'

    $DNl = New-Object system.Windows.Forms.Label
    $DNl.text = ""
    $DNl.Visible = $false
    $DNl.width = 25
    $DNl.height = 10
    $DNl.location = New-Object System.Drawing.Point(340, 225)
    $DNl.Font = 'Microsoft Sans Serif,14'

    $Update = New-Object system.Windows.Forms.Button
    $Update.text = "Update"
    $Update.Visible = $false
    $Update.width = 125
    $Update.height = 55
    $Update.location = New-Object System.Drawing.Point(6, 375)
    $Update.Font = 'Microsoft Sans Serif,14'


    $GetUserPage.controls.AddRange(@($Label1, $filterf, $Searchf, $Search, $namel, $Phonel, $Phonef,
            $Jobtitlel, $Jobtitlef1, $Officel, $Officef1, $Departmentl, $Departmentf1, $Descl, $Descf, $DNl, $Update))

    #Connect to remote computer terminal page
    $RemoteTerminalPage.Text = "Remote Terminal"
    $tabControl.Controls.add($RemoteTerminalPage)
    $RemoteTerminalPage.Add_Click({
        & "P:\Global\IT\POSH\CreateUsers\Remote Computer.lnk"
    })

    #region gui events {
    #Submit form to create user
    $Submit.Add_Click( {createUser})

    $Jobtitlef1.Add_KeyDown( {
            if ($_.KeyCode -eq "Enter") {
                createUser
            }
        })

    $Open.Add_Click( {Invoke-Item "ADUsers.csv"})

    $Create.Add_Click( {& AddADUsers.ps1})

    $O365f.Add_Click( {
            if ($O365f.Checked) {
                $Licensef.Visible = $true
            }
            else {
                $Licensef.Visible = $false
            }
        })

    $Search.Add_Click( {searchUser})

    $Searchf.Add_KeyDown( {
            if ($_.KeyCode -eq "Enter") {
                searchUser
            }
        })

    $Update.Add_Click( {updateUser})

    [void]$CreateADUsers.ShowDialog()

}
generateForm