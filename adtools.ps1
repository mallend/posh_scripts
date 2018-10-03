## AD Tools allows admin to Add and Delete AD accounts from a single form from the Add User and Delete User tabs, respectively
# Future addon will include adding and deleting Computer objects

Try { Import-Module Actividnumrectory } # Try to import Actividnumrectory
Catch { [System.Windows.Forms.MessageBox]::Show("Error importing Actividnumrectory Module. Please try again.") } # If fails retry

# Global variables Add User
$global:lastName = $null
$global:firstName = $null
$global:initial = $null
$global:email = $null
$global:Idnumber = $null
$global:symbol = $null
$global:bldg = $null
$global:title = $null
$global:city = $null
$global:state = $null
$global:zip = $null
$global:phone = $null
$global:org = $null
$global:company = $null
$global:affl = $null
$global:rank = $null
$global:country = $null
$global:password = $null
$global:kerberos = $null
$global:OU = $null

#Global variables Remove User
$global:REMOVEemail = $null
$global:REMOVEIdnumber = $null
$global:REMOVEDN = $null
$global:REMOVidnumsplayName = $null


Function CreateForm #Creates the GUI form
{
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    #Setup Form
    $Form = New-Object System.Windows.Forms.Form
    $TabControl = New-Object System.Windows.Forms.Tabcontrol
    $AddUsersPage = New-Object System.Windows.Forms.TabPage
    $RemoveUsersPage = New-Object System.Windows.Forms.TabPage
    $RemoveComputersPage = New-Object System.Windows.Forms.TabPage

    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    #Form
    $Form.Text = "Actividnumrectory Tools"
    $Form.Name = "AD Tools"
    $Form.DataBindings.DefaultDataSourceUpdateMode = 0
    $Form.StartPosition = "CenterScreen"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 725
    $System_Drawing_Size.Height = 450
    $Form.ClientSize = $System_Drawing_Size

    #Tabs
    $TabControl.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 0
    $System_Drawing_Point.Y = 0
    $TabControl.Location = $System_Drawing_Point
    $TabControl.Name = "TabControl"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 450
    $System_Drawing_Size.Width = 725
    $TabControl.Size = $System_Drawing_Size
    $Form.Controls.Add($TabControl)

    #AddUsersPage
    $AddUsersPage.DataBindings.DefaultDataSourceUpdateMode = 0
    $AddUsersPage.UseVisualStyleBackColor = $True
    $AddUsersPage.Name = "AddUserPage"
    $AddUsersPage.Text = "Add User"
    $TabControl.Controls.Add($AddUsersPage)
#BUTTONS
    #CheckButton
    $AddCheckButton = New-Object System.Windows.Forms.Button
    $AddCheckButton.Location = New-Object System.Drawing.Size(165,375)
    $AddCheckButton.Size = New-Object System.Drawing.Size(85,30)
    $AddCheckButton.Text = "Check"
    $AddUsersPage.Controls.Add($AddCheckButton)
    $AddCheckButton.Add_Click({CheckUser})
    #CreateButton
    $AddCreateButton = New-Object System.Windows.Forms.Button
    $AddCreateButton.Location = New-Object System.Drawing.Size(265,375)
    $AddCreateButton.Size = New-Object System.Drawing.Size(85,30)
    $AddCreateButton.Text = "Create User"
    $AddCreateButton.ForeColor = 'Red'
    $AddUsersPage.Controls.Add($AddCreateButton)
    $AddCreateButton.Add_Click({CreateUser})
    #ClearButton
    $AddClearButton = New-Object System.Windows.Forms.Button
    $AddClearButton.Location = New-Object System.Drawing.Size(365,375)
    $AddClearButton.Size = New-Object System.Drawing.Size(85,30)
    $AddClearButton.Text = "Clear"
    $AddUsersPage.Controls.Add($AddClearButton)
    $AddClearButton.Add_Click({AddUserClear})
    #ExitButton
    $AddExitButton = New-Object System.Windows.Forms.Button
    $AddExitButton.Location = New-Object System.Drawing.Size(465,375)
    $AddExitButton.Size = New-Object System.Drawing.Size(85,30)
    $AddExitButton.Text = "Exit"
    $AddUsersPage.Controls.Add($AddExitButton)
    $AddExitButton.Add_Click({$Form.Close()})
#TextBoxes and Labels
    #Instructions
    $AddInstructionsLabel = New-Object System.Windows.Forms.Label
    $AddInstructionsLabel.Location = New-Object System.Drawing.Size(105,5)
    $AddInstructionsLabel.Size = New-Object System.Drawing.Size (600, 15)
    $AddInstructionsLabel.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
    $AddInstructionsLabel.Text = "Use this form to create a new user in Actividnumrectory. Click the 'Check' Button before clicking 'Create'."
    $AddUsersPage.Controls.Add($AddInstructionsLabel)
    #Last name
    $AddLastNameTextbox = New-Object System.Windows.Forms.TextBox
    $AddLastNameTextbox.Location = New-Object System.Drawing.Size(115, 35)
    $AddLastNameTextbox.Size = New-Object System.Drawing.Size(100, 20)        
    $AddLastNameLabel = New-Object System.Windows.Forms.Label
    $AddLastNameLabel.Text = "Last Name:"
    $AddLastNameLabel.Location = New-Object System.Drawing.Size (10, 35)
    $AddLastNameLabel.Size = New-Object System.Drawing.Size(85, 20)
    $AddLastNameLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddLastNameLabel)
    $AddUsersPage.Controls.Add($AddLastNameTextbox)
    #First name
    $AddFirstNameTextbox = New-Object System.Windows.Forms.TextBox
    $AddFirstNameTextbox.Location = New-Object System.Drawing.Size(115, 65)
    $AddFirstNameTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $AddFirstNameLabel = New-Object System.Windows.Forms.Label
    $AddFirstNameLabel.Text = "First Name:"
    $AddFirstNameLabel.Location = New-Object System.Drawing.Size (10, 65)
    $AddFirstNameLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddFirstNameLabel)
    $AddUsersPage.Controls.Add($AddFirstNameTextbox)
    #Initial
    $AddInitialTextbox = New-Object System.Windows.Forms.TextBox
    $AddInitialTextbox.Location = New-Object System.Drawing.Size(115, 95)
    $AddInitialTextbox.Size = New-Object System.Drawing.Size(22, 20)    
    $AddInitialLabel = New-Object System.Windows.Forms.Label
    $AddInitialLabel.Text = "Initial:"
    $AddInitialLabel.Location = New-Object System.Drawing.Size (10, 95)
    $AddInitialLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddInitialLabel)
    $AddUsersPage.Controls.Add($AddInitialTextbox)
    #Email
    $AddEmailTextbox = New-Object System.Windows.Forms.TextBox
    $AddEmailTextbox.Location = New-Object System.Drawing.Size(115, 125)
    $AddEmailTextbox.Size = New-Object System.Drawing.Size(150, 20)    
    $AddEmailLabel = New-Object System.Windows.Forms.Label
    $AddEmailLabel.Text = "Email:"
    $AddEmailLabel.Location = New-Object System.Drawing.Size (10, 125)
    $AddEmailLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddEmailLabel)
    $AddUsersPage.Controls.Add($AddEmailTextbox)
    #Idnumber
    $AddIdnumberTextbox = New-Object System.Windows.Forms.TextBox
    $AddIdnumberTextbox.Location = New-Object System.Drawing.Size(115, 155)
    $AddIdnumberTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $AddIdnumberLabel = New-Object System.Windows.Forms.Label
    $AddIdnumberLabel.Text = "Id Number:"
    $AddIdnumberLabel.Location = New-Object System.Drawing.Size (10, 155)
    $AddIdnumberLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddIdnumberLabel)
    $AddUsersPage.Controls.Add($AddIdnumberTextbox)
    #OU Listbox
    $UserOUListbox = New-Object System.Windows.Forms.ListBox
    $UserOUListbox.Location = New-Object System.Drawing.Point(315,35)
    $UserOUListbox.Size = New-Object System.Drawing.Size(20,20)
    $UserOUListbox.Height = 80
    $UserOUListbox.Width = 120
    [void] $UserOUListbox.Items.Add('Item1')
    [void] $UserOUListbox.Items.Add('Item2')
    [void] $UserOUListbox.Items.Add('Item3')
    [void] $UserOUListbox.Items.Add('Item4')
    [void] $UserOUListbox.Items.Add('Item5')
    [void] $UserOUListbox.Items.Add('Item6')
    [void] $UserOUListbox.Items.Add('Item7')
    $UserOULabel = New-Object System.Windows.Forms.Label
    $UserOULabel.Text = "OU:"
    $UserOULabel.Location = New-Object System.Drawing.Size(275, 35)
    $UserOULabel.Size = New-Object System.Drawing.Size(35,20)
    $UserOULabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($UserOULabel)
    $AddUsersPage.Controls.Add($UserOUListbox)
    #Office
    $AddOfficeTextbox = New-Object System.Windows.Forms.TextBox
    $AddOfficeTextbox.Location = New-Object System.Drawing.Size(585, 35)
    $AddOfficeTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $AddOfficeLabel = New-Object System.Windows.Forms.Label
    $AddOfficeLabel.Text = "Office Symbol:"
    $AddOfficeLabel.Location = New-Object System.Drawing.Size (460, 35)
    $AddOfficeLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddOfficeLabel.Size = New-Object System.Drawing.Size(120, 20)
    $AddUsersPage.Controls.Add($AddOfficeTextbox)
    $AddUsersPage.Controls.Add($AddOfficeLabel)
    #JobTitle
    $AddJobTextbox = New-Object System.Windows.Forms.TextBox
    $AddJobTextbox.Location = New-Object System.Drawing.Size(585, 65)
    $AddJobTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $AddJobLabel = New-Object System.Windows.Forms.Label
    $AddJobLabel.Text = "Job Title:"
    $AddJobLabel.Location = New-Object System.Drawing.Size (460, 65)
    $AddJobLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddJobLabel)
    $AddUsersPage.Controls.Add($AddJobTextbox)
    #Phone
    $AddPhoneTextbox = New-Object System.Windows.Forms.TextBox
    $AddPhoneTextbox.Location = New-Object System.Drawing.Size(585, 95)
    $AddPhoneTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $AddPhoneLabel = New-Object System.Windows.Forms.Label
    $AddPhoneLabel.Text = "Phone Number:"
    $AddPhoneLabel.Location = New-Object System.Drawing.Size (460, 95)
    $AddPhoneLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddPhoneLabel.Size = New-Object System.Drawing.Size(120, 20)
    $AddUsersPage.Controls.Add($AddPhoneLabel)
    $AddUsersPage.Controls.Add($AddPhoneTextbox)
    #Description
    $AddDescriptionTextbox = New-Object System.Windows.Forms.TextBox
    $AddDescriptionTextbox.Location = New-Object System.Drawing.Size(585, 125)
    $AddDescriptionTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $AddDescriptionLabel = New-Object System.Windows.Forms.Label
    $AddDescriptionLabel.Text = "Description:"
    $AddDescriptionLabel.Location = New-Object System.Drawing.Size (460, 125)
    $AddDescriptionLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddDescriptionLabel)
    $AddUsersPage.Controls.Add($AddDescriptionTextbox)
    #Company
    $AddCompanyTextbox = New-Object System.Windows.Forms.TextBox
    $AddCompanyTextbox.Location = New-Object System.Drawing.Size(585, 155)
    $AddCompanyTextbox.Size = New-Object System.Drawing.Size(100, 20) 
    $AddCompanyTextbox.Text = "Company"   
    $AddCompanyLabel = New-Object System.Windows.Forms.Label
    $AddCompanyLabel.Text = "Company:"
    $AddCompanyLabel.Location = New-Object System.Drawing.Size (460, 155)
    $AddCompanyLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddCompanyLabel)
    $AddUsersPage.Controls.Add($AddCompanyTextbox)
    #Bldg
    $AddBldgTextbox = New-Object System.Windows.Forms.TextBox
    $AddBldgTextbox.Location = New-Object System.Drawing.Size(585, 185)
    $AddBldgTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $AddBldgLabel = New-Object System.Windows.Forms.Label
    $AddBldgLabel.Text = "Building #:"
    $AddBldgLabel.Location = New-Object System.Drawing.Size (460, 185)
    $AddBldgLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddBldgLabel)
    $AddUsersPage.Controls.Add($AddBldgTextbox)
    #City
    $AddCityTextbox = New-Object System.Windows.Forms.TextBox
    $AddCityTextbox.Location = New-Object System.Drawing.Size(585, 215)
    $AddCityTextbox.Size = New-Object System.Drawing.Size(100, 20)
    $AddCityTextbox.Text = "City"    
    $AddCityLabel = New-Object System.Windows.Forms.Label
    $AddCityLabel.Text = "City:"
    $AddCityLabel.Location = New-Object System.Drawing.Size (460, 215)
    $AddCityLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddCityLabel)
    $AddUsersPage.Controls.Add($AddCityTextbox)
    #State
    $AddStateTextbox = New-Object System.Windows.Forms.TextBox
    $AddStateTextbox.Location = New-Object System.Drawing.Size(585, 245)
    $AddStateTextbox.Size = New-Object System.Drawing.Size(25, 20)
    $AddStateTextbox.Text = "ST"    
    $AddStateLabel = New-Object System.Windows.Forms.Label
    $AddStateLabel.Text = "State:"
    $AddStateLabel.Location = New-Object System.Drawing.Size (460, 245)
    $AddStateLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddStateLabel)
    $AddUsersPage.Controls.Add($AddStateTextbox)
    #ZipCode
    $AddZipTextbox = New-Object System.Windows.Forms.TextBox
    $AddZipTextbox.Location = New-Object System.Drawing.Size(585, 275)
    $AddZipTextbox.Size = New-Object System.Drawing.Size(75, 20)
    $AddZipTextbox.Text = "Zip"    
    $AddZipLabel = New-Object System.Windows.Forms.Label
    $AddZipLabel.Text = "Zip Code:"
    $AddZipLabel.Location = New-Object System.Drawing.Size (460, 275)
    $AddZipLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddZipLabel)
    $AddUsersPage.Controls.Add($AddZipTextbox)
    #Affiliation
    $AfflTextbox = New-Object System.Windows.Forms.ListBox
    $AfflTextbox.Location = New-Object System.Drawing.Point(115,185)
    $AfflTextbox.Size = New-Object System.Drawing.Size(85,20)
    $AfflTextbox.Height = 50
    [void] $AfflTextbox.Items.Add('MIL')
    [void] $AfflTextbox.Items.Add('CIV')
    [void] $AfflTextbox.Items.Add('CTR')
    [void] $AfflTextbox.Items.Add('FM')
    $AddUsersPage.Controls.Add($AfflTextbox)
    $AfflLabel = New-Object System.Windows.Forms.Label
    $AfflLabel.Text = "Affiliation:"
    $AfflLabel.Location = New-Object System.Drawing.Size(10, 185)
    $AfflLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AfflLabel)
    #Rank
    $AddRankTextbox = New-Object System.Windows.Forms.TextBox
    $AddRankTextbox.Location = New-Object System.Drawing.Size(115, 245)
    $AddRankTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $AddRankLabel = New-Object System.Windows.Forms.Label
    $AddRankLabel.Text = "Rank:"
    $AddRankLabel.Location = New-Object System.Drawing.Size (10, 245)
    $AddRankLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $AddUsersPage.Controls.Add($AddRankLabel)
    $AddUsersPage.Controls.Add($AddRankTextbox)

    #RemoveUsersPage
    $RemoveUsersPage.DataBindings.DefaultDataSourceUpdateMode = 0
    $RemoveUsersPage.UseVisualStyleBackColor = $True
    $RemoveUsersPage.Name = "RemoveUsersPage"
    $RemoveUsersPage.Text = "Remove User"
    $TabControl.Controls.Add($RemoveUsersPage)
#Buttons
    #FindButton
    $RemoveFindButton = New-Object System.Windows.Forms.Button
    $RemoveFindButton.Location = New-Object System.Drawing.Size(165,375)
    $RemoveFindButton.Size = New-Object System.Drawing.Size(85,30)
    $RemoveFindButton.Text = "Find"
    $RemoveUsersPage.Controls.Add($RemoveFindButton)
    $RemoveFindButton.Add_Click({FindUser})
    #DeleteButton
    $RemoveDeleteButton = New-Object System.Windows.Forms.Button
    $RemoveDeleteButton.Location = New-Object System.Drawing.Size(265,375)
    $RemoveDeleteButton.Size = New-Object System.Drawing.Size(85,30)
    $RemoveDeleteButton.Text = "Delete"
    $RemoveDeleteButton.ForeColor = 'Red'
    $RemoveUsersPage.Controls.Add($RemoveDeleteButton)
    $RemoveDeleteButton.Add_Click({DeleteUser})
    #ClearButton
    $RemoveClearButton = New-Object System.Windows.Forms.Button
    $RemoveClearButton.Location = New-Object System.Drawing.Size(365,375)
    $RemoveClearButton.Size = New-Object System.Drawing.Size(85,30)
    $RemoveClearButton.Text = "Clear"
    $RemoveUsersPage.Controls.Add($RemoveClearButton)
    $RemoveClearButton.Add_Click({RemoveUserClear})
    #ExitButton
    $RemoveExitButton = New-Object System.Windows.Forms.Button
    $RemoveExitButton.Location = New-Object System.Drawing.Size(465,375)
    $RemoveExitButton.Size = New-Object System.Drawing.Size(85,30)
    $RemoveExitButton.Text = "Exit"
    $RemoveUsersPage.Controls.Add($RemoveExitButton)
    $RemoveExitButton.Add_Click({$Form.Close()})
#Textboxes and Labels
    #Instructions
    $RemInstructionsLabel = New-Object System.Windows.Forms.Label
    $RemInstructionsLabel.Location = New-Object System.Drawing.Size(105,5)
    $RemInstructionsLabel.Size = New-Object System.Drawing.Size (600, 15)
    $RemInstructionsLabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Regular)
    $RemInstructionsLabel.Text = "Use this form to delete a user from Actividnumrectory. Click the 'Find' Button to verify user."
    $RemoveUsersPage.Controls.Add($RemInstructionsLabel)
    #Email
    $RemEmailTextbox = New-Object System.Windows.Forms.TextBox
    $RemEmailTextbox.Location = New-Object System.Drawing.Size(225, 55)
    $RemEmailTextbox.Size = New-Object System.Drawing.Size(150, 20)    
    $RemEmailLabel = New-Object System.Windows.Forms.Label
    $RemEmailLabel.Text = "Email Address:"
    $RemEmailLabel.Location = New-Object System.Drawing.Size (105, 55)
    $RemEmailLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $RemEmailLabel.Size = New-Object System.Drawing.Size(125, 20)
    $RemoveUsersPage.Controls.Add($RemEmailTextbox)
    $RemoveUsersPage.Controls.Add($RemEmailLabel)
    #Idnumber
    $RemIdnumberTextbox = New-Object System.Windows.Forms.TextBox
    $RemIdnumberTextbox.Location = New-Object System.Drawing.Size(450, 55)
    $RemIdnumberTextbox.Size = New-Object System.Drawing.Size(100, 20)    
    $RemIdnumberLabel = New-Object System.Windows.Forms.Label
    $RemIdnumberLabel.Text = "Idnumber:"
    $RemIdnumberLabel.Location = New-Object System.Drawing.Size (400, 55)
    $RemIdnumberLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $RemoveUsersPage.Controls.Add($RemIdnumberTextbox)
    $RemoveUsersPage.Controls.Add($RemIdnumberLabel)
    #Groupbox
    $global:group = New-Object System.Windows.Forms.GroupBox
    $global:group.Location = New-Object System.Drawing.Size(195,135)
    $global:group.Size = New-Object System.Drawing.Size (325, 150)
    $global:group.Text = "User Found:"
    $RemoveUsersPage.Controls.Add($global:group)

    #RemoveComputersPage
    $RemoveComputersPage.DataBindings.DefaultDataSourceUpdateMode = 0
    $RemoveComputersPage.UseVisualStyleBackColor = $True
    $RemoveComputersPage.Name = "RemoveComputersPage"
    $RemoveComputersPage.Text = "Remove Computers"
    $TabControl.Controls.Add($RemoveComputersPage)
#Buttons
    #Find Button
    $RmPCFindButton = New-Object System.Windows.Forms.Button
    $RmPCFindButton.Location = New-Object System.Drawing.Size(165,375)
    $RmPCFindButton.Size = New-Object System.Drawing.Size(85,30)
    $RmPCFindButton.Text = "Find"
    $RemoveComputersPage.Controls.Add($RmPCFindButton)
    $RmPCFindButton.Add_Click({FindADComputer})
    #DeleteButton
    $RmPCDeleteButton = New-Object System.Windows.Forms.Button
    $RmPCDeleteButton.Location = New-Object System.Drawing.Size(265,375)
    $RmPCDeleteButton.Size = New-Object System.Drawing.Size(85,30)
    $RmPCDeleteButton.Text = "Delete"
    $RmPCDeleteButton.ForeColor = 'Red'
    $RemoveComputersPage.Controls.Add($RmPCDeleteButton)
    $RmPCDeleteButton.Add_Click({DeleteComputer})
    #ClearButton
    $RmPCClearButton = New-Object System.Windows.Forms.Button
    $RmPCClearButton.Location = New-Object System.Drawing.Size(365,375)
    $RmPCClearButton.Size = New-Object System.Drawing.Size(85,30)
    $RmPCClearButton.Text = "Clear"
    $RemoveComputersPage.Controls.Add($RmPCClearButton)
    $RmPCClearButton.Add_Click({RemoveComputerClear})
    #ExitButton
    $RmPCExitButton = New-Object System.Windows.Forms.Button
    $RmPCExitButton.Location = New-Object System.Drawing.Size(465,375)
    $RmPCExitButton.Size = New-Object System.Drawing.Size(85,30)
    $RmPCExitButton.Text = "Exit"
    $RemoveComputersPage.Controls.Add($RmPCExitButton)
    $RmPCExitButton.Add_Click({$Form.Close()})
#Textboxes and Labels
    #Instructions
    $RmPCInstructionsLabel = New-Object System.Windows.Forms.Label
    $RmPCInstructionsLabel.Location = New-Object System.Drawing.Size(95,5)
    $RmPCInstructionsLabel.Size = New-Object System.Drawing.Size (600, 15)
    $RmPCInstructionsLabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Regular)
    $RmPCInstructionsLabel.Text = "Use this form to remove a Computer from AD. Click the 'Find' Button before clicking 'Delete'."
    $RemoveComputersPage.Controls.Add($RmPCInstructionsLabel)
    #ComputerName
    $RmPCNameTextbox = New-Object System.Windows.Forms.TextBox
    $RmPCNameTextbox.Location = New-Object System.Drawing.Size(325, 35)
    $RmPCNameTextbox.Size = New-Object System.Drawing.Size(150, 20)    
    $RmPCNameLabel = New-Object System.Windows.Forms.Label
    $RmPCNameLabel.Text = "Computer Name:"
    $RmPCNameLabel.Location = New-Object System.Drawing.Size (200, 35)
    $RmPCNameLabel.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $RmPCNameLabel.Size = New-Object System.Drawing.Size(150, 20)
    $RemoveComputersPage.Controls.Add($RmPCNameTextbox)
    $RemoveComputersPage.Controls.Add($RmPCNameLabel)
    #Groupbox
    $global:RmPCGroup = New-Object System.Windows.Forms.GroupBox
    $global:RmPCGroup.Location = New-Object System.Drawing.Size(195,85)
    $global:RmPCGroup.Size = New-Object System.Drawing.Size (300, 150)
    $global:RmPCGroup.Text = "Computer found:"
    #$RmPCGroup.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Regular)
    $RemoveComputersPage.Controls.Add($global:RmPCGroup)

    $Form.ShowDialog()
}

Function CheckUser #Takes input from Add User form and assigns it to the global variables and then calls the UserExist Function
{

    $global:lastName = $AddLastNameTextbox.Text
    $global:firstName = $AddFirstNameTextbox.Text
    $global:initial = $AddInitialTextbox.Text
    $global:email = $AddEmailTextbox.Text
    $global:Idnumber = $AddIdnumberTextbox.Text
    $global:symbol = $AddOfficeTextbox.Text
    $global:bldg = $AddBldgTextbox.Text
    $global:title = $AddJobTextbox.Text
    $global:city = $AddCityTextbox.Text
    $global:state = $AddStateTextbox.Text
    $global:zip = $AddZipTextbox.Text
    $global:phone = $AddPhoneTextbox.Text
    $global:org = $AddDescriptionTextbox.Text
    $global:company = $AddCompanyTextbox.Text
    $global:affl = $AfflTextbox.Selectidnumtem
    $global:rank = $AddRankTextbox.Text
    $global:country = "US"
    $global:password = (ConvertTo-SecureString -AsPlainText 1qaz2wsx!QAZ@WSX -force)
    $global:kerberos = "AES256"
    $global:OU = GetUserOU ($UserOUListbox.Selectidnumtem) 
    
    UserExist # Check to see if user already has an account
    
}

Function UserExist #Checks if User has an account already
{
    #Retrieve-UserInfo
    If ( $exists = (Get-ADUser -Filter "UserPrincipalName -eq '$global:Idnumber@mil'").DistinguishedName ) { # set variable equal to a user's DN property from AD       
        [System.Windows.Forms.MessageBox]::Show("User already has an account. `n$exists") # If $exists is not $null then return user's current OU
    }

    Else {        
        [System.Windows.Forms.MessageBox]::Show("User does not have an account.") # If $exists is $null then display that the account is clear to be created
    }
}

Function CreateUser #Actually create the AD User object with all required information
{  
    $displayName = ("$global:lastName, $global:firstName $global:initial $global:rank $global:affl USA") -replace '\s+', ' '
                    
    New-ADUser -Name "$global:lastName, $global:firstName $global:initial" -GivenName $global:firstName -Surname $global:lastName `
    -Initials $global:initial -DisplayName $displayName -Path $global:OU -AccountPassword $global:password -SamAccountName "$global:Idnumber.$global:affl" `
    -UserPrincipalName $global:Idnumber -EmailAddress $global:email -StreetAddress "Bldg $global:bldg" -City $global:city -State $global:state `
    -Country $global:country -PostalCode $global:zip -Office $global:symbol -Department $global:org -Description $global:org -Title $global:title `
    -OfficePhone $global:phone -Company $global:company -SmartcardLogonRequired $True -Enabled $True -PasswordNeverExpires $False `
    -KerberosEncryptionType $global:kerberos -PassThru

    ValidateUser #Check if the user was created
}

Function ValidateUser #Queries AD with the information provided in form to verify if user has been created
{
    # If successful, display message
    If ( $exists = (Get-ADUser -Filter "UserPrincipalName -eq '$global:Idnumber@mil'").DistinguishedName ) { # set variable equal to a user's DN property from AD       
        AddUserClear
        [System.Windows.Forms.MessageBox]::Show("Account created successfully!") # If $exists is not $null then return user's current OU        
    }

    Else {        
        [System.Windows.Forms.MessageBox]::Show("Error occurred while creating account for user. Please try again.") # If $exists is $null then display that the account is clear to be created
    }
}

Function AddUserClear #Clear Add User Form
{
    # Clear textbox inputs
    $AddLastNameTextbox.Text = $null
    $AddFirstNameTextbox.Text = $null
    $AddInitialTextbox.Text = $null
    $AddEmailTextbox.Text = $null
    $AddIdnumberTextbox.Text = $null
    $AddOfficeTextbox.Text = $null
    $AddBldgTextbox.Text = $null
    $AddJobTextbox.Text = $null
    $AddPhoneTextbox.Text = $null
    $AddDescriptionTextbox.Text = $null
    $AfflTextbox.Selectidnumtem = $null
    $AddRankTextbox.Text = $null
    $UserOUListbox.Selectidnumtem = $null

    # Reset global variables to $null
    $global:lastName = $null
    $global:firstName = $null
    $global:initial = $null
    $global:email = $null
    $global:Idnumber = $null
    $global:symbol = $null
    $global:bldg = $null
    $global:title = $null
    $global:city = $null
    $global:state = $null
    $global:zip = $null
    $global:phone = $null
    $global:org = $null
    $global:company = $null
    $global:affl = $null
    $global:rank = $null
    $global:country = $null
    $global:password = $null
    $global:kerberos = $null
    $global:OU = $null
}

Function GetUserOU
{
    param($OU)

    $PATH = $OU # Pass in the paramter OU and assign it to $PATH

    switch ( $PATH ) # SWITCH the original value passed in for an OU path that corresponds to that value
    {
        "Item1" { $result = 'OU1' }
        "Item2" { $result = 'OU2' }
        "Item3" { $result = 'OU3' }
        "Item4" { $result = 'OU4' }
        "Item5" { $result = 'OU5' }
        "Item6" { $result = 'OU6' }
        "Item7" { $result = 'OU7' }        
    }

    return $result #returns the OU path
}

Function RemoveUserClear #Clear Remove User Form
{
    $global:REMOVidnumsplayName.Text = $null
    $global:REMOVEmail.Text = $null
    $global:REMOVEidnum.Text = $null
    $global:REMOVEou.Text = $null

    $global:group.Controls.Remove($global:REMOVidnumsplayName)
    $global:group.Controls.Remove($global:REMOVEidnum)
    $global:group.Controls.Remove($global:REMOVEmail)
    $global:group.Controls.Remove($global:REMOVEou)

    $global:REMOVEemail = $null
    $global:REMOVEIdnumber = $null
    $global:REMOVEDN = $null
    $global:REMOVEuser = $null     
}

Function DeleteUser #Actually delete user after prompting
{
    $answer = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete user $global:REMOVEuser ?", "Delete User", 4) #Prompt to verify delete request
    If ($answer -eq "Yes"){ #If YES delete
    Get-ADUser -Identity $global:REMOVEDN | Remove-ADUser -Confirm:$false
    ConfirmDelete #Check if user is no longer in directory
    }
    If ($answer -eq "No"){ #If NO do not delete
    [System.Windows.Forms.MessageBox]::Show("Delete request has been canceled")
    }
}

Function ConfirmDelete #Queries AD with the form input to verify user no longer has an account
{
    $exists = (Get-ADUser -Identity $global:REMOVEDN -Properties *).Name
    If (-not $exists) {
    RemoveUserClear
    [System.Windows.Forms.MessageBox]::Show("Successfully deleted account for $global:REMOVEuser.")    
    }    
}

Function FindUser #Takes info from Remove User tab and assigns variables and checks for user, returning the users Name, Email, and Idnumber along with OU
{
    $global:REMOVEemail = $RemEmailTextbox.Text
    $global:REMOVEIdnumber = $RemIdnumberTextbox.Text

    If ($global:REMOVEemail){ #If searching by email
        $name = (Get-ADUser -Filter "EmailAddress -eq '$global:REMOVEemail'").Name
        $mail1 = (Get-ADUser -Filter "EmailAddress -eq '$global:REMOVEemail'" -Properties *).EmailAddress
        $idnum1 = (Get-ADUser -Filter "EmailAddress -eq '$global:REMOVEemail'" -Properties *).UserPrincipalName
        $ou1 = (Get-ADUser -Filter "EmailAddress -eq '$global:REMOVEemail'" -Properties *).CanonicalName
        $global:REMOVEDN = (Get-ADUser -Filter "EmailAddress -eq '$global:REMOVEemail'" -Properties *).DistinguishedName
        If (-not $name) {#If the AD query returns no results 
        [System.Windows.Forms.MessageBox]::Show("Error finding user. Please verify search parameters.")
        } #end inner if
        Else{ #If AD query returns results
        $global:REMOVidnumsplayName = New-Object System.Windows.Forms.Label
        $global:REMOVidnumsplayName.Location = New-Object System.Drawing.Size(20,20)
        $global:REMOVidnumsplayName.Size = New-Object System.Drawing.Size(250,17)
        $global:REMOVidnumsplayName.Text = $name
        $global:group.Controls.Add($global:REMOVidnumsplayName)

        $global:REMOVEmail = New-Object System.Windows.Forms.Label
        $global:REMOVEmail.Location = New-Object System.Drawing.Size (20,35)
        $global:REMOVEmail.Size = New-Object System.Drawing.Size(250,17)
        $global:REMOVEmail.Text = $mail1
        $global:group.Controls.Add($global:REMOVEmail)

        $global:REMOVEidnum = New-Object System.Windows.Forms.Label
        $global:REMOVEidnum.Location = New-Object System.Drawing.Size (20,50)
        $global:REMOVEidnum.Size = New-Object System.Drawing.Size(250,17)
        $global:REMOVEidnum.Text = $idnum1
        $global:group.Controls.Add($global:REMOVEidnum)

        $global:REMOVEou = New-Object System.Windows.Forms.Label
        $global:REMOVEou.Location = New-Object System.Drawing.Size (20,75)
        $global:REMOVEou.Size = New-Object System.Drawing.Size (285,50)
        $global:REMOVEou.Text = $ou1
        $global:group.Controls.Add($global:REMOVEou)

        $global:REMOVEuser = (Get-ADUser -Identity $global:REMOVEDN -Properties *).Name # set the found User's name to a global variable to be called by prompt
        } # end inner else
    } #end outer if
    Elseif ($global:REMOVEIdnumber){ #If searching by Idnumber do this
        $name = (Get-ADUser -Filter "UserPrincipalName -eq '$global:REMOVEIdnumber@mil'").Name
        $mail1 = (Get-ADUser -Filter "UserPrincipalName -eq '$global:REMOVEIdnumber@mil'" -Properties *).EmailAddress
        $idnum1 = (Get-ADUser -Filter "UserPrincipalName -eq '$global:REMOVEIdnumber@mil'" -Properties *).UserPrincipalName
        $ou1 = (Get-ADUser -Filter "UserPrincipalName -eq '$global:REMOVEIdnumber@mil'" -Properties *).CanonicalName
        $global:REMOVEDN = (Get-ADUser -Filter "UserPrincipalName -eq '$global:REMOVEIdnumber@mil'" -Properties *).DistinguishedName
        If (-not $name){ #If query returns no results
        [System.Windows.Forms.MessageBox]::Show("Error finding user. Please verify search parameters.")
        } #end inner if
        Else{ #If query Returns result, diplay Name, Email, Idnumber and OU the user is currently in
        $global:REMOVidnumsplayName = New-Object System.Windows.Forms.Label
        $global:REMOVidnumsplayName.Location = New-Object System.Drawing.Size(20,20)
        $global:REMOVidnumsplayName.Size = New-Object System.Drawing.Size(250,17)
        $global:REMOVidnumsplayName.Text = $name
        $global:group.Controls.Add($global:REMOVidnumsplayName)

        $global:REMOVEmail = New-Object System.Windows.Forms.Label
        $global:REMOVEmail.Location = New-Object System.Drawing.Size (20,35)
        $global:REMOVEmail.Size = New-Object System.Drawing.Size(250,17)
        $global:REMOVEmail.Text = $mail1
        $global:group.Controls.Add($global:REMOVEmail)

        $global:REMOVEidnum = New-Object System.Windows.Forms.Label
        $global:REMOVEidnum.Location = New-Object System.Drawing.Size (20,50)
        $global:REMOVEidnum.Size = New-Object System.Drawing.Size(250,17)
        $global:REMOVEidnum.Text = $idnum1
        $global:group.Controls.Add($global:REMOVEidnum)

        $global:REMOVEou = New-Object System.Windows.Forms.Label
        $global:REMOVEou.Location = New-Object System.Drawing.Size (20,75)
        $global:REMOVEou.Size = New-Object System.Drawing.Size (285,50)
        $global:REMOVEou.Text = $ou1
        $global:group.Controls.Add($global:REMOVEou)
        $global:REMOVEuser = (Get-ADUser -Identity $global:REMOVEDN -Properties *).Name # set the found User's name to a global variable to be called by prompt
        } #end inner else   
    }#elseif
    Else { [System.Windows.Forms.MessageBox]::Show("Error finding user. Please verify search parameters.") } #If user could not be found, display notification
}

Function FindADComputer #Accept input from Remove Computers Tab and query AD 
{
    $find = $RmPCNameTextbox.Text
    $global:adPCName = (Get-ADComputer -SearchBase $global:OU -Filter "Name -eq '$find'" -Properties *).Name
    $global:dn = (Get-ADComputer -SearchBase $global:OU -Filter "Name -eq '$find'" -Properties *).DistinguishedName
    If (-not $global:adPCName) { [System.Windows.Forms.MessageBox]::Show("Error finding computer. Please verify name.") } #If not found then prompt 
    Else { #If found, return AD Name for Computer along with OU
    $global:foundPC = New-Object System.Windows.Forms.Label
    $global:foundPC.Location = New-Object System.Drawing.Size(10,20)
    $global:foundPC.Text = $global:adPCName
    $global:foundPC.Size = New-Object System.Drawing.Size(250,20)
    $RmPCGroup.Controls.Add($global:foundPC)

    $global:foundDN = New-Object System.Windows.Forms.Label
    $global:foundDN.Location = New-Object System.Drawing.Size(10,50)
    $global:foundDN.Text = $global:dn
    $global:foundDN.Size = New-Object System.Drawing.Size(275,75)
    $RmPCGroup.Controls.Add($global:foundDN)

    }
}

Function DeleteComputer #Actually delete the Computer Object from AD
{
    $answer = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete computer $global:adPCName ?", "Delete Computer", 4) #Prompt to verify delete is correct

    If ($answer -eq "Yes") #If YES then delete
    {
        Remove-ADObject -Identity $global:dn -Confirm:$false
        $check = (Get-ADComputer -Filter "Name -eq '$global:adPCName'" -Properties *).Name

        If (-not $check) { 
        RemoveComputerClear
        [System.Windows.Forms.MessageBox]::Show("Computer deleted successfully!") 
        } 
        If ($check) { 
            $confirm = [System.Windows.Forms.MessageBox]::Show("Computer has child objects. Do you want to continue?", "Confirm Delete", 4) 
            If ($confirm -eq "Yes"){
            Remove-ADObject -Identity $global:dn -Confirm:$false -Recursive 

            $check = (Get-ADComputer -Filter "Name -eq '$global:adPCName'" -Properties *).Name
            If (-not $check) { 
                RemoveComputerClear
                [System.Windows.Forms.MessageBox]::Show("Computer deleted successfully!") 
                } 
            }
            If ($confirm -eq "No"){ [System.Windows.Forms.MessageBox]::Show("Delete request has been canceled.") }
        }
    }
    If ($answer -eq "No") { [System.Windows.Forms.MessageBox]::Show("Delete request has been canceled.") } #if NO then do not delete
}

Function RemoveComputerClear #Clears the Remove Computers form tab
{
    $global:foundPC.Text = $null
    $global:foundDN.Text = $null

    $global:RmPCGroup.Controls.Remove($global:foundDN)
    $global:RmPCGroup.Controls.Remove($global:foundPC)

    $global:adPCName = $null
}

CreateForm