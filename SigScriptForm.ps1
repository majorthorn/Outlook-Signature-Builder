 
     Write-Output "This Command console is not used by this script. Please see Dialog."
    Add-Type -AssemblyName System.Windows.Forms 
    Add-Type -AssemblyName System.Drawing 
    $MyForm = New-Object System.Windows.Forms.Form 
    $MyForm.Text="MyForm" 
    $MyForm.Size = New-Object System.Drawing.Size(300,100) 
     
 
        $musernameTextBox = New-Object System.Windows.Forms.TextBox 
                $musernameTextBox.Text="" 
                $musernameTextBox.Top="30" 
                $musernameTextBox.Left="10" 
                $musernameTextBox.Anchor="Left,Top" 
        $musernameTextBox.Size = New-Object System.Drawing.Size(150,23) 
        $MyForm.Controls.Add($musernameTextBox) 
         
 
        $mexecuteButton = New-Object System.Windows.Forms.Button 
                $mexecuteButton.Text="Execute" 
                $mexecuteButton.Top="29" 
                $mexecuteButton.Left="182" 
                $mexecuteButton.Anchor="Left,Top" 
        $mexecuteButton.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mexecuteButton) 
         
 
        $mUsernameLabel = New-Object System.Windows.Forms.Label 
                $mUsernameLabel.Text="Username" 
                $mUsernameLabel.Top="15" 
                $mUsernameLabel.Left="10" 
                $mUsernameLabel.Anchor="Left,Top" 
        $mUsernameLabel.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mUsernameLabel) 
        $MyForm.ShowDialog()
