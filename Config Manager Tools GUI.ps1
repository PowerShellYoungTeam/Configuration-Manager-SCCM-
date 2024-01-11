<# To DO 11/01/2024 - Config Manager Tools GUI

Create Module for required fuctions
    - Get-CMCollectionMembership (with AD lookup, use get-SCCMcollectionmembership as a base)
    - Get-CMApplicationFolderDetails (one that via nested loops captures every dependency of every deployment type)
    - pop in connect-configmgrserver function

Create installer for the tool
    - Copy the modules from the FileShare to the correct location on the local machine
    - check the $profile and add the import-module command if it doesn't exist

Update the Gui
    - Connect the drop down menu's to pull folder/collection data from the SCCM server as per find-process
    - add a file selector to the select output folder button
    - Add a button to open the output folder in explorer

#>


Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Config Manager Tools"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"

# Create the section for Get Collection Membership
$groupBoxCollectionMembership = New-Object System.Windows.Forms.GroupBox
$groupBoxCollectionMembership.Location = New-Object System.Drawing.Point(10, 10)
$groupBoxCollectionMembership.Size = New-Object System.Drawing.Size(570, 100)
$groupBoxCollectionMembership.Text = "Get Collection Membership"
$form.Controls.Add($groupBoxCollectionMembership)

# Create the label and text field for collection name
$labelCollectionName = New-Object System.Windows.Forms.Label
$labelCollectionName.Location = New-Object System.Drawing.Point(10, 30)
$labelCollectionName.Size = New-Object System.Drawing.Size(100, 20)
$labelCollectionName.Text = "Collection Name:"
$groupBoxCollectionMembership.Controls.Add($labelCollectionName)

$textBoxCollectionName = New-Object System.Windows.Forms.TextBox
$textBoxCollectionName.Location = New-Object System.Drawing.Point(120, 30)
$textBoxCollectionName.Size = New-Object System.Drawing.Size(200, 20)
$groupBoxCollectionMembership.Controls.Add($textBoxCollectionName)

# Create the search button
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Location = New-Object System.Drawing.Point(330, 30)
$buttonSearch.Size = New-Object System.Drawing.Size(75, 20)
$buttonSearch.Text = "Search"
$groupBoxCollectionMembership.Controls.Add($buttonSearch)

# Create the drop down menu for collection membership
$comboBoxCollectionMembership = New-Object System.Windows.Forms.ComboBox
$comboBoxCollectionMembership.Location = New-Object System.Drawing.Point(10, 60)
$comboBoxCollectionMembership.Size = New-Object System.Drawing.Size(200, 20)
$groupBoxCollectionMembership.Controls.Add($comboBoxCollectionMembership)

# Create the button for getting collection membership
$buttonGetCollectionMembership = New-Object System.Windows.Forms.Button
$buttonGetCollectionMembership.Location = New-Object System.Drawing.Point(220, 60)
$buttonGetCollectionMembership.Size = New-Object System.Drawing.Size(200, 20)
$buttonGetCollectionMembership.Text = "Get Collection Membership"
$groupBoxCollectionMembership.Controls.Add($buttonGetCollectionMembership)

# Create the section for Get Application Details
$groupBoxApplicationDetails = New-Object System.Windows.Forms.GroupBox
$groupBoxApplicationDetails.Location = New-Object System.Drawing.Point(10, 120)
$groupBoxApplicationDetails.Size = New-Object System.Drawing.Size(570, 100)
$groupBoxApplicationDetails.Text = "Get Application Details"
$form.Controls.Add($groupBoxApplicationDetails)

# Create the label and drop down menu for software folders
$labelSoftwareFolders = New-Object System.Windows.Forms.Label
$labelSoftwareFolders.Location = New-Object System.Drawing.Point(10, 30)
$labelSoftwareFolders.Size = New-Object System.Drawing.Size(100, 20)
$labelSoftwareFolders.Text = "Software Folders:"
$groupBoxApplicationDetails.Controls.Add($labelSoftwareFolders)

$comboBoxSoftwareFolders = New-Object System.Windows.Forms.ComboBox
$comboBoxSoftwareFolders.Location = New-Object System.Drawing.Point(120, 30)
$comboBoxSoftwareFolders.Size = New-Object System.Drawing.Size(200, 20)
$groupBoxApplicationDetails.Controls.Add($comboBoxSoftwareFolders)

# Create the button for getting folder's applications
$buttonGetFolderApplications = New-Object System.Windows.Forms.Button
$buttonGetFolderApplications.Location = New-Object System.Drawing.Point(330, 30)
$buttonGetFolderApplications.Size = New-Object System.Drawing.Size(200, 20)
$buttonGetFolderApplications.Text = "Get Folder's Applications"
$groupBoxApplicationDetails.Controls.Add($buttonGetFolderApplications)

# Create the section for output folder
$groupBoxOutputFolder = New-Object System.Windows.Forms.GroupBox
$groupBoxOutputFolder.Location = New-Object System.Drawing.Point(10, 230)
$groupBoxOutputFolder.Size = New-Object System.Drawing.Size(570, 100)
$groupBoxOutputFolder.Text = "Output Folder"
$form.Controls.Add($groupBoxOutputFolder)

# Create the button for selecting output folder
$buttonSelectOutputFolder = New-Object System.Windows.Forms.Button
$buttonSelectOutputFolder.Location = New-Object System.Drawing.Point(10, 30)
$buttonSelectOutputFolder.Size = New-Object System.Drawing.Size(150, 20)
$buttonSelectOutputFolder.Text = "Select Output Folder"
$groupBoxOutputFolder.Controls.Add($buttonSelectOutputFolder)

# Add event handler for button click
$buttonSelectOutputFolder.Add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $result = $folderBrowserDialog.ShowDialog()
    if ($result -eq 'OK') {
        $textBoxOutputFolder.Text = $folderBrowserDialog.SelectedPath
    }
})

# Create the button for showing output folder in explorer
$buttonShowOutputFolder = New-Object System.Windows.Forms.Button
$buttonShowOutputFolder.Location = New-Object System.Drawing.Point(170, 30)
$buttonShowOutputFolder.Size = New-Object System.Drawing.Size(200, 20)
$buttonShowOutputFolder.Text = "Open Output Folder in Explorer"
$groupBoxOutputFolder.Controls.Add($buttonShowOutputFolder)

# Add event handler for button click
$buttonShowOutputFolder.Add_Click({
    $outputFolder = $textBoxOutputFolder.Text
    if (Test-Path $outputFolder) {
        Invoke-Item -Path $outputFolder
    } else {
        Write-Host "Output folder does not exist."
    }
})

# Create the label and text field for output folder path
$labelOutputFolder = New-Object System.Windows.Forms.Label
$labelOutputFolder.Location = New-Object System.Drawing.Point(10, 60)
$labelOutputFolder.Size = New-Object System.Drawing.Size(100, 20)
$labelOutputFolder.Text = "Output Folder:"
$groupBoxOutputFolder.Controls.Add($labelOutputFolder)

$textBoxOutputFolder = New-Object System.Windows.Forms.TextBox
$textBoxOutputFolder.Location = New-Object System.Drawing.Point(120, 60)
$textBoxOutputFolder.Size = New-Object System.Drawing.Size(400, 20)
$textBoxOutputFolder.ReadOnly = $true
$groupBoxOutputFolder.Controls.Add($textBoxOutputFolder)

# Show the form
$form.ShowDialog()
