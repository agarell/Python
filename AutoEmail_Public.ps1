<#Written by Elliott Agar on 30/10/2024
This script creates a pop up to quickly send emails to Company for warranty and insurance pick ups.

The numbers after System.Drawing.Point are X & Y values to move the words and boxes around

If you've lost the exe, you can convert this script to one using the below command: 
Install-Module -Name ps2exe
Import-Module ps2exe
Invoke-PS2EXE C:\path\to\this\script.ps1 C:\path\to\the\exe\Lazy.exe 


#>

Add-Type -AssemblyName System.Windows.Forms

# Create the pop up
$form = New-Object System.Windows.Forms.Form
$form.Text = "Enter Serial Number"
$form.Width = 300
$form.Height = 250
$form.StartPosition = "CenterScreen"

# Create the Label for Service Number
$labelSN = New-Object System.Windows.Forms.Label
$labelSN.Text = "S/N"
$labelSN.AutoSize = $true
$labelSN.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($labelSN)

# Create the Label for weight selection
$labelWeight = New-Object System.Windows.Forms.Label
$labelWeight.Text = "Weight"
$labelWeight.AutoSize = $true
$labelWeight.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($labelWeight)

# Create the Label for date selection
$labelDate = New-Object System.Windows.Forms.Label
$labelDate.Text = "Date"
$labelDate.AutoSize = $true
$labelDate.Location = New-Object System.Drawing.Point(10, 100)
$form.Controls.Add($labelDate)

# Create the TextBox for input SN input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Width = 200
$textBox.Location = New-Object System.Drawing.Point(50, 20)
$form.Controls.Add($textBox)

# Create the drop down options for weight
$weightBox = New-Object System.Windows.Forms.ComboBox
$weightBox.Width = 200
$weightBox.Location = New-Object System.Drawing.Point(50, 60)
$weightBox.Items.AddRange(@("Keyboard - 1KG", "Laptop - 2KG"))
$weightBox.SelectedIndex = 0 # Set default option to first
$form.Controls.Add($weightBox)

# Create the second drop down for dates
$dateBox = New-Object System.Windows.Forms.ComboBox
$dateBox.Width = 200
$dateBox.Location = New-Object System.Drawing.Point(50, 100) 
$dateBox.Items.AddRange(@("Today", "Tomorrow"))
$dateBox.SelectedIndex = 0  # Set default to the first option
$form.Controls.Add($dateBox)

# Create OK/Send Email Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Send Email"
$okButton.Location = New-Object System.Drawing.Point(50, 120)
$okButton.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})
$form.Controls.Add($okButton)

# Create Cancel Button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Location = New-Object System.Drawing.Point(150, 120)
$cancelButton.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})
$form.Controls.Add($cancelButton)

# Show the Form as a dialog
$result = $form.ShowDialog()

# Check the result and send the email if confirmed
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $serviceNumber = $textBox.Text
    
    if ($dateBox.SelectedItem -eq "Today") {
        $getDate = Get-Date -Format "dd/MM"  # Format for today
    } elseif ($dateBox.SelectedItem -eq "Tomorrow") {
        $getDate = (Get-Date).AddDays(1).ToString("dd/MM")  # Format for tomorrow
    }

    # Extract just the weight value based on the selected item
    $selectedWeight = ""
    switch ($weightBox.SelectedItem) {
        "Keyboard - 1KG" { $selectedWeight = "1kg" }
        "Laptop - 2KG" { $selectedWeight = "2kg" }
    }

    # Initialize the Outlook COM object
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0) # 0 indicates a Mail item

    # Set email properties
    $Mail.To = "customerservice@qCompany.com "
    $Mail.Subject = "Pickup"

    # Format the email body, change this section if the email template needs to be changed
    $Mail.Body = @"

Hi Company,
	 
Just want to book a collection directly to our site,
	 
Details:
	 
SR number: $serviceNumber
Name: Staff Name
Address: 123 Fake Street
	 
Length: 60cm
Width: 30cm
Height: 10cm
Weight: $selectedWeight
Date and Time: $getDate anytime between 9:30am - 3:00pm
Comments: Pick up from Location's Mail room.
	 
If you need any other information, please let me know.
The package already has a return label.
	 
Kind regards,
"@

    # Send the email
    $Mail.Send()

    # Confirm email was sent
    [System.Windows.Forms.MessageBox]::Show("Email sent with Service Number: $serviceNumber", "Confirmation")
} else {
    Write-Output "Operation canceled."
}
