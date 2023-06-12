Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create form and controls
$form = New-Object System.Windows.Forms.Form
$form.Text = "Active Directory Search v3.1"
$form.Size = New-Object System.Drawing.Size(1230, 660)
$form.FormBorderStyle = 'Fixed3D'
$form.MaximizeBox = $False
$form.StartPosition = 'CenterScreen'

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10, 10)
$label2.Size = New-Object System.Drawing.Size(500, 20)
$label2.Text = "Enter: [firstnameSPACElastname], [firstnameDOTlastname], [EmployeeID], or [Computer Number]"
$form.Controls.Add($label2)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(10, 30)
$label3.Size = New-Object System.Drawing.Size(490, 20)
$label3.Text = "Example:         [John Doe]                           [john.doe]"
$form.Controls.Add($label3)

# Create the text box
$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10, 50)
$textBox2.Size = New-Object System.Drawing.Size(480, 20)
$form.Controls.Add($textBox2)

# Create the list box
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 315)
$listBox.Size = New-Object System.Drawing.Size(590, 300)
# Set the SelectionMode property to allow multiple items to be selected
$ListBox.SelectionMode = "MultiExtended"
# Add the event to handle Ctrl+C key press
$listBox.Add_KeyDown({
    if ($_.KeyCode -eq "C" -and $_.Control) {
        $selectedItems = ""
        foreach ($item in $listBox.SelectedItems) {
            $selectedValue = $item.ToString().Split(":")[1].Trim()
            $selectedItems += "$selectedValue`r`n"
        }
        Set-Clipboard -Value $selectedItems
    }
})

$listBox.TabIndex = 1

$listBox.TabIndex = 1

# Create the list box2
$listbox2 = New-Object System.Windows.Forms.ListBox
$listbox2.Location = New-Object System.Drawing.Point(610, 80)
$listbox2.Size = New-Object System.Drawing.Size(590, 530)
# Set the SelectionMode property to allow multiple items to be selected
$ListBox2.SelectionMode = "MultiExtended"
$listBox2.Add_KeyDown({ if ($_.KeyCode -eq "C") { $selectedItems = $listBox2.SelectedItems; $text = ""; foreach ($item in $selectedItems) { $text += $item + "`r`n" }; [System.Windows.Forms.Clipboard]::SetText($text) } })

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(500, 50)
$button.Size = New-Object System.Drawing.Size(100, 20)
$button.Text = "Search"
$form.Controls.Add($button)

$form.AcceptButton = $button

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Point(1047, 50)
$Button2.Size = New-Object System.Drawing.Size(150, 20)
$Button2.Text = "Open Active Directory"
$Button2.Add_Click({
        Start-Process "powershell.exe" -ArgumentList "-Command `"Start-Process $env:SystemRoot\system32\dsa.msc -Verb RunAs`""
    })
$Form.Controls.Add($Button2)

$button.Add_Click({
        $searchTerm2 = $textBox2.Text.Trim()

        $form.Controls.Add($listBox)
        #        $form.Controls.Add($listbox2)
        $form.Controls.Remove($listbox2)


        # Clear the results table
        if ($form.Controls.ContainsKey("resultTable")) {
            $form.Controls.RemoveByKey("resultTable")
        }
        
        # Clear all boxes
        #        $textBox.Clear()
        $textBox2.Clear()
        $listBox.Items.Clear()
        $listBox2.Items.Clear()

        $resultTable = New-Object System.Windows.Forms.ListView
        $resultTable.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::None
        $resultTable.Scrollable = $false
        $resultTable.View = [System.Windows.Forms.View]::Details
        $resultTable.FullRowSelect = $True # Allow full row selection
        $resultTable.Location = New-Object System.Drawing.Point(10, 80)
        $resultTable.Size = New-Object System.Drawing.Size(590, 215)
        $resultTable.Name = "resultTable"
                
        # Add a SelectedIndexChanged event handler
        $resultTable.add_SelectedIndexChanged({
                $selectedItems = $resultTable.SelectedItems
                if ($selectedItems.Count -gt 0) {
                    $selectedText = $selectedItems[0].SubItems | ForEach-Object { $_.Text } | Out-String
                    [System.Windows.Forms.Clipboard]::SetText($selectedText)
                }
            })

        # Add a KeyDown event handler to copy the selected value to the clipboard
        $resultTable.add_KeyDown({
                param($sender, $e)
                # Check for the "control + c" key combination
                if (($e.KeyCode -eq "C") -and ($e.Control)) {
                    $selectedItems = $sender.SelectedItems
                    if ($selectedItems.Count -gt 0) {
                        $selectedText = $selectedItems[0].SubItems[1].Text
                        [System.Windows.Forms.Clipboard]::SetText($selectedText)
                    }
                }
            })
                
        # Add columns to table
        [void]$resultTable.Columns.Add("")
        [void]$resultTable.Columns.Add("")    

        if (![string]::IsNullOrEmpty($searchTerm2)) {
            if ($searchTerm2 -Match "\b[0-9]{6}\b|\.") {
                if ($searchTerm2 -match "^\d+$") {
                    $result = Get-ADUser -Filter "EmployeeID -like '$searchTerm2*'" -Properties DisplayName, GivenName, Surname, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, Department, City, StreetAddress, PasswordExpired, PasswordLastSet, LastBadPasswordAttempt | Select-Object DisplayName, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, Department, City, StreetAddress
                    $employeeID = Get-ADUser -Filter "EmployeeID -like '$searchTerm2*'" -Properties * | Select-Object -ExpandProperty EmployeeID
                }
                # Check Firstname.Lastname
                else {
                    $result = Get-ADUser -Filter "mail -Like '$searchTerm2*'" -Properties DisplayName, GivenName, Surname, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, Department, City, StreetAddress, PasswordExpired, PasswordLastSet, LastBadPasswordAttempt | Select-Object DisplayName, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, Department, City, StreetAddress
                    $employeeID = Get-ADUser -Filter "mail -like '$searchTerm2*'" -Properties * | Select-Object -ExpandProperty EmployeeID
                }        
                if ($result) {
                    # Check if account is active
                    $netUserOutput = net user /domain $employeeID
                    $accountActive = $netUserOutput | Select-String "Account active" | ForEach-Object { $_.ToString().Trim() } | ForEach-Object { ($_ -replace "Account active", "").Trim() }
                    $passwordLastSet = $netUserOutput | Select-String "Password last set" | ForEach-Object { $_.ToString().Trim() } | ForEach-Object { ($_ -replace "Password last set", "").Trim() }
                    $passwordExpires = $netUserOutput | Select-String "Password expires" | ForEach-Object { $_.ToString().Trim() } | ForEach-Object { ($_ -replace "Password expires", "").Trim() }
                
                    if ($accountActive -eq "Yes") {
                        $result | Add-Member -MemberType NoteProperty -Name 'AccountActive' -Value $accountActive -PassThru |
                        Add-Member -MemberType NoteProperty -Name 'PasswordLastSet' -Value $passwordLastSet -PassThru |
                        Add-Member -MemberType NoteProperty -Name 'PasswordExpires' -Value $passwordExpires -PassThru |
                        Select-Object DisplayName, GivenName, Surname, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, 
                        Department, City, StreetAddress, AccountActive, PasswordLastSet, PasswordExpires
            
                        # Check and show MembersOf
                        $form.Controls.Add($listbox2)
                        function Search-AD($employeeID) {
                            $user = Get-ADUser -Filter { EmployeeID -eq $employeeID } -Properties memberOf
                            if ($user -ne $null) {
                                return $user.memberOf | ForEach-Object { $_.Split(',')[0].Substring(3) }
                            }
                            else {
                                return $null
                            }
                        }
                        $groups = Search-AD $employeeID
                        if ($groups -ne $null) {
                            $listbox2.BeginUpdate()
                            $listbox2.Items.Clear()
                            foreach ($group in ($groups | Sort-Object)) {
                                $listbox2.Items.Add($group)
                            }
                            $listbox2.EndUpdate()
                        }
                        else {
                            return $null
                            #                [System.Windows.Forms.MessageBox]::Show("Employee ID not found in Active Directory.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        }
                        function Resize-ListViewColumns {
                            param([System.Windows.Forms.ListView]$ListView)
                            $ListView.Columns[0].Width = -2
                            $ListView.Columns[1].Width = -2
                            $ListView.AutoResizeColumns('HeaderSize')
                            $ListView.Columns[0].Width = [System.Math]::Max($ListView.Columns[0].Width, 120)
                            $ListView.Columns[1].Width = [System.Math]::Max($ListView.Columns[1].Width, 185)
                        }
                    
                        # Display result
                        if ($result) { 
                            # Add rows to table
                            foreach ($property in $result.PSObject.Properties) {
                                $row = New-Object System.Windows.Forms.ListViewItem($property.Name)
                                if ($property.Value -ne $null) {
                                    $row.SubItems.Add($property.Value.ToString())
                                }
                                else {
                                    $row.SubItems.Add("")
                                }
                                if ($resultTable -ne $null) {
                                    [void]$resultTable.Items.Add($row)
                                }
                            }
                    
                            $form.Controls.Add($resultTable)
                            Resize-ListViewColumns $resultTable          
                    
                            # Get computer details
                            $searchResults = Get-ADComputer -Filter "Description -like '*$employeeID*' -or CN -like '*$employeeID*'" -Properties CN, Description, LastLogonDate, IPv4Address | Select-Object CN, Description, LastLogonDate, IPv4Address
                            foreach ($result in $searchResults) {
                                $listBox.Items.Add("CN: " + $result.CN)
                                $listBox.Items.Add("Description: " + $result.Description)
                                $listBox.Items.Add("Last Logon Date: " + $result.LastLogonDate)
                                $listBox.Items.Add("IPv4Address: " + $result.IPv4Address)
                                $listBox.Items.Add("") # Add a blank line between results
                            }
                        }
                        else {
                            [System.Windows.Forms.MessageBox]::Show("Account is not active", "Result", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        }
                    } 
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("No Result Found", "Result", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
        
            }
            else {
                $searchResults = Get-ADComputer -Filter "Description -like '*$searchTerm2*' -or CN -like '*$searchTerm2*'" -Properties CN, Description, LastLogonDate, IPv4Address | Select-Object CN, Description, LastLogonDate, IPv4Address
                if ($searchResults) {
                    if (![string]::IsNullOrEmpty($searchResults)) {
                        $employeeID = (Get-ADComputer -Filter "Description -Like '*$searchTerm2*' -OR CN -Like '*$searchTerm2*'" -Properties Description).Description.Split("|")[0] -replace "[^\d]+", ""
                        #                $emailAddress = Get-ADUser -Filter "EmployeeID -like '$employeeID*'" -Properties * | Select-Object -ExpandProperty EmailAddress
                        $resultMachine = Get-ADUser -Filter "EmployeeID -like '$employeeID*'" -Properties DisplayName, GivenName, Surname, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, Department, City, StreetAddress, PasswordExpired, PasswordLastSet, LastBadPasswordAttempt | Select-Object DisplayName, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, Department, City, StreetAddress
                        #Check if $searchTerm2 is a string
                        if ($searchTerm2 -cmatch '^[a-zA-Z\s]+$') {              
                            foreach ($resultMachine in $searchResults) {
                                $listBox.Items.Add("CN: " + $resultMachine.CN)
                                $listBox.Items.Add("Description: " + $resultMachine.Description)
                                $listBox.Items.Add("Last Logon Date: " + $resultMachine.LastLogonDate)
                                $listBox.Items.Add("IPv4Address: " + $resultMachine.IPv4Address)
                                $listBox.Items.Add("") # Add a blank line between results
                            }
                            
                            #                        [System.Windows.Forms.MessageBox]::Show("Multiple machines and/or accounts found.", "Result", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                            
                        }
                        else {
                            $searchResults = Get-ADComputer -Filter "Description -like '*$employeeID*' -or CN -like '*$employeeID*'" -Properties CN, Description, LastLogonDate, IPv4Address | Select-Object CN, Description, LastLogonDate, IPv4Address
                            
                            foreach ($resultMachine in $searchResults) {
                                $listBox.Items.Add("CN: " + $resultMachine.CN)
                                $listBox.Items.Add("Description: " + $resultMachine.Description)
                                $listBox.Items.Add("Last Logon Date: " + $resultMachine.LastLogonDate)
                                $listBox.Items.Add("IPv4Address: " + $resultMachine.IPv4Address)
                                $listBox.Items.Add("") # Add a blank line between results
                            }
            
                            $result = Get-ADUser -Filter "EmployeeID -like '$employeeID*'" -Properties DisplayName, GivenName, Surname, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, Department, City, StreetAddress, PasswordExpired, PasswordLastSet, LastBadPasswordAttempt | Select-Object DisplayName, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, Department, City, StreetAddress
            
                            $netUserOutput = net user /domain $employeeID
                            $accountActive = $netUserOutput | Select-String "Account active" | ForEach-Object { $_.ToString().Trim() } | ForEach-Object { ($_ -replace "Account active", "").Trim() }
                            $passwordLastSet = $netUserOutput | Select-String "Password last set" | ForEach-Object { $_.ToString().Trim() } | ForEach-Object { ($_ -replace "Password last set", "").Trim() }
                            $passwordExpires = $netUserOutput | Select-String "Password expires" | ForEach-Object { $_.ToString().Trim() } | ForEach-Object { ($_ -replace "Password expires", "").Trim() }
            
                            $result | Add-Member -MemberType NoteProperty -Name 'AccountActive' -Value $accountActive -PassThru |
                            Add-Member -MemberType NoteProperty -Name 'PasswordLastSet' -Value $passwordLastSet -PassThru |
                            Add-Member -MemberType NoteProperty -Name 'PasswordExpires' -Value $passwordExpires -PassThru |
                            Select-Object DisplayName, GivenName, Surname, EmployeeID, EmailAddress, OfficePhone, MobilePhone, Title, 
                            Department, City, StreetAddress, AccountActive, PasswordLastSet, PasswordExpires
            
                            # Check and show MembersOf
                            $form.Controls.Add($listbox2)
                            function Search-AD($employeeID) {
                                $user = Get-ADUser -Filter { EmployeeID -eq $employeeID } -Properties memberOf
                                if ($user -ne $null) {
                                    return $user.memberOf | ForEach-Object { $_.Split(',')[0].Substring(3) }
                                }
                                else {
                                    return $null
                                }
                            }
            
                            $groups = Search-AD $employeeID
                            if ($groups -ne $null) {
                                $listbox2.BeginUpdate()
                                $listbox2.Items.Clear()
                                foreach ($group in ($groups | Sort-Object)) {
                                    $listbox2.Items.Add($group)
                                }
                                $listbox2.EndUpdate()
                            }
                            else {
                                #                            [System.Windows.Forms.MessageBox]::Show("Employee ID not found in Active Directory.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            }
            
                            function Resize-ListViewColumns {
                                param([System.Windows.Forms.ListView]$ListView)
                                $ListView.Columns[0].Width = -2
                                $ListView.Columns[1].Width = -2
                                $ListView.AutoResizeColumns('HeaderSize')
                                $ListView.Columns[0].Width = [System.Math]::Max($ListView.Columns[0].Width, 120)
                                $ListView.Columns[1].Width = [System.Math]::Max($ListView.Columns[1].Width, 185)
                            }
            
                            # Add rows to table
                            foreach ($property in $result.PSObject.Properties) {
                                $row = New-Object System.Windows.Forms.ListViewItem($property.Name)
                                if ($property.Value -ne $null) {
                                    $row.SubItems.Add($property.Value.ToString())
                                }
                                else {
                                    $row.SubItems.Add("")
                                }
                                if ($resultTable -ne $null) {
                                    [void]$resultTable.Items.Add($row)
                                }
                            }
                
                            $form.Controls.Add($resultTable)
                            #            [System.Windows.Forms.MessageBox]::Show("Search found.", "Result", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                            Resize-ListViewColumns $resultTable
                        }
                    }
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("No Result Found", "Result", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            }
            


        }

    })

# Show form
$form.ShowDialog() | Out-Null