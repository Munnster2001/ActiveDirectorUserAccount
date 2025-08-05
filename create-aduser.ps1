# Import the Active Directory module
Import-Module ActiveDirectory

# Load WPF assembly
Add-Type -AssemblyName PresentationFramework

# Define the XAML for the GUI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Create AD Account" Height="750" Width="900" Background="#F5F5F5">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Width" Value="200"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="Margin" Value="5,0,0,0"/>
            <Setter Property="TabIndex" Value="-1"/> <!-- Disable tab for text boxes -->
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Width" Value="200"/>
            <Setter Property="Height" Value="30"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="Margin" Value="5,0,0,0"/>
            <Setter Property="TabIndex" Value="-1"/> <!-- Disable tab for combo box -->
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#4682B4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Width" Value="100"/>
            <Setter Property="Margin" Value="5,0,0,0"/>
        </Style>
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="Margin" Value="0,0,15,0"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="60"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Create AD Account" FontSize="16" FontWeight="Bold" Grid.ColumnSpan="3" Margin="0,0,0,20" Foreground="#333333"/>
        <StackPanel Grid.Row="1" Grid.Column="0">
            <TextBlock Text="User Information" FontSize="14" FontWeight="Bold" Margin="0,0,0,15"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="First Name:" Width="100"/>
                <TextBox x:Name="FirstNameTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Surname:" Width="100"/>
                <TextBox x:Name="SurnameTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Display Name:" Width="100"/>
                <TextBox x:Name="DisplayNameTextBox" IsReadOnly="True"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Description:" Width="100"/>
                <TextBox x:Name="DescriptionTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Username Format:" Width="100"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <RadioButton x:Name="FirstInitialSurnameRadio" Content="FirstInitialSurname" GroupName="UsernameFormat" IsChecked="True" TabIndex="1"/>
                <RadioButton x:Name="SurnameFirstInitialRadio" Content="SurnameFirstInitial" GroupName="UsernameFormat" TabIndex="2"/>
                <RadioButton x:Name="FirstNameSurnameRadio" Content="FirstName.Surname" GroupName="UsernameFormat" TabIndex="3"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock x:Name="UsernameLabel" Text="Username:" Width="100"/>
                <TextBox x:Name="UsernameTextBox" IsReadOnly="True"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Password:" Width="100"/>
                <TextBox x:Name="PasswordTextBox" Width="200"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <Button x:Name="GeneratePasswordButton" Content="Generate" TabIndex="4"/>
                <Button x:Name="CopyPasswordButton" Content="Copy" TabIndex="5"/>
            </StackPanel>
            <TextBlock Text="Contact Details" FontSize="14" FontWeight="Bold" Margin="0,20,0,15"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Telephone:" Width="100"/>
                <TextBox x:Name="TelephoneTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Email:" Width="100"/>
                <TextBox x:Name="EmailTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Mobile:" Width="100"/>
                <TextBox x:Name="MobileTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="IP Phone:" Width="100"/>
                <TextBox x:Name="IPPhoneTextBox"/>
            </StackPanel>
        </StackPanel>
        <StackPanel Grid.Row="1" Grid.Column="2">
            <TextBlock Text="Organization Details" FontSize="14" FontWeight="Bold" Margin="0,0,0,15"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Job Title:" Width="100"/>
                <TextBox x:Name="JobTitleTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Department:" Width="100"/>
                <TextBox x:Name="DepartmentTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Company:" Width="100"/>
                <TextBox x:Name="CompanyTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Manager:" Width="100"/>
                <ComboBox x:Name="ManagerComboBox" IsEditable="True"/>
            </StackPanel>
            <TextBlock Text="Profile Info" FontSize="14" FontWeight="Bold" Margin="0,20,0,15"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Logon Script:" Width="100"/>
                <TextBox x:Name="LogonScriptTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Home Directory:" Width="100"/>
                <TextBox x:Name="HomeDirectoryTextBox"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15" VerticalAlignment="Bottom">
                <Button x:Name="ResetButton" Content="Reset Fields" TabIndex="6"/>
                <Button x:Name="CreateButton" Content="Create Account" TabIndex="7"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"@

# Load the XAML into a window object
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find GUI elements
$firstNameTextBox = $window.FindName("FirstNameTextBox")
$surnameTextBox = $window.FindName("SurnameTextBox")
$displayNameTextBox = $window.FindName("DisplayNameTextBox")
$descriptionTextBox = $window.FindName("DescriptionTextBox")
$firstInitialSurnameRadio = $window.FindName("FirstInitialSurnameRadio")
$surnameFirstInitialRadio = $window.FindName("SurnameFirstInitialRadio")
$firstNameSurnameRadio = $window.FindName("FirstNameSurnameRadio")
$usernameLabel = $window.FindName("UsernameLabel")
$usernameTextBox = $window.FindName("UsernameTextBox")
$passwordTextBox = $window.FindName("PasswordTextBox")
$generatePasswordButton = $window.FindName("GeneratePasswordButton")
$copyPasswordButton = $window.FindName("CopyPasswordButton")
$telephoneTextBox = $window.FindName("TelephoneTextBox")
$emailTextBox = $window.FindName("EmailTextBox")
$mobileTextBox = $window.FindName("MobileTextBox")
$ipPhoneTextBox = $window.FindName("IPPhoneTextBox")
$jobTitleTextBox = $window.FindName("JobTitleTextBox")
$departmentTextBox = $window.FindName("DepartmentTextBox")
$companyTextBox = $window.FindName("CompanyTextBox")
$managerComboBox = $window.FindName("ManagerComboBox")
$logonScriptTextBox = $window.FindName("LogonScriptTextBox")
$homeDirectoryTextBox = $window.FindName("HomeDirectoryTextBox")
$createButton = $window.FindName("CreateButton")
$resetButton = $window.FindName("ResetButton")

if (-not ($firstNameTextBox -and $surnameTextBox -and $displayNameTextBox -and $firstInitialSurnameRadio -and $surnameFirstInitialRadio -and $firstNameSurnameRadio -and $usernameTextBox -and $passwordTextBox -and $generatePasswordButton -and $copyPasswordButton -and $telephoneTextBox -and $emailTextBox -and $mobileTextBox -and $ipPhoneTextBox -and $jobTitleTextBox -and $departmentTextBox -and $companyTextBox -and $managerComboBox -and $logonScriptTextBox -and $homeDirectoryTextBox -and $createButton -and $resetButton)) {
    throw "One or more GUI elements not found."
}

# Populate manager ComboBox with AD users
$users = Get-ADUser -Filter * -Properties SamAccountName, DisplayName | Select-Object @{Name="Name";Expression={"$($_.SamAccountName) ($($_.DisplayName))"}}
$managerComboBox.ItemsSource = $users | Select-Object -ExpandProperty Name

# Function to update display name and username with uniqueness check
function Update-DisplayNameAndUsername {
    $firstName = $firstNameTextBox.Text
    $surname = $surnameTextBox.Text
    if ($firstName -or $surname) {
        $displayNameTextBox.Text = "$firstName $surname".Trim()
    }
    if ($firstName -and $surname) {
        if ($firstInitialSurnameRadio.IsChecked) {
            $usernameTextBox.Text = "$($firstName[0])$surname".ToLower()
        } elseif ($surnameFirstInitialRadio.IsChecked) {
            $usernameTextBox.Text = "$surname$($firstName[0])".ToLower()
        } else {
            $usernameTextBox.Text = "$firstName.$surname".ToLower()
        }
        # Check for existing username in AD
        $existingUser = Get-ADUser -Filter {SamAccountName -eq $usernameTextBox.Text} -ErrorAction SilentlyContinue
        $usernameLabel.Foreground = if ($existingUser) { [System.Windows.Media.Brushes]::DarkRed } else { [System.Windows.Media.Brushes]::Black }
    } else {
        $usernameTextBox.Text = ""
        $usernameLabel.Foreground = [System.Windows.Media.Brushes]::Black
    }
}

# Function to generate a readable password (no o, 0, l, I, 1)
function Generate-Password {
    $chars = "abcdefghijkmnpqrstuvwxyz23456789!@#$%^&*()"
    $password = -join (1..12 | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
    $passwordTextBox.Text = $password
}

# Function to copy password to clipboard
function Copy-Password {
    [System.Windows.Clipboard]::SetText($passwordTextBox.Text)
}

# Pre-generate a password for new user
Generate-Password

# Event handler for manager autocomplete
$managerComboBox.Add_TextChanged({
    $text = $managerComboBox.Text
    if ($text) {
        $matches = $users | Where-Object { $_.Name -like "*$text*" -or $_.SamAccountName -like "*$text*" -or $_.DisplayName -like "*$text*" } | Select-Object -First 10 | Select-Object @{Name="Name";Expression={"$($_.SamAccountName) ($($_.DisplayName))"}}
        $managerComboBox.ItemsSource = $matches | Select-Object -ExpandProperty Name
        $managerComboBox.IsDropDownOpen = $true
    } else {
        $managerComboBox.ItemsSource = $users | Select-Object @{Name="Name";Expression={"$($_.SamAccountName) ($($_.DisplayName))"}}
    }
})

# Event handlers for text changes
$firstNameTextBox.Add_TextChanged({ Update-DisplayNameAndUsername })
$surnameTextBox.Add_TextChanged({ Update-DisplayNameAndUsername })
$firstInitialSurnameRadio.Add_Checked({ Update-DisplayNameAndUsername })
$surnameFirstInitialRadio.Add_Checked({ Update-DisplayNameAndUsername })
$firstNameSurnameRadio.Add_Checked({ Update-DisplayNameAndUsername })

# Event handler for password generation
$generatePasswordButton.Add_Click({ Generate-Password })

# Event handler for password copy
$copyPasswordButton.Add_Click({ Copy-Password })

# Event handler for reset button
$resetButton.Add_Click({
    $firstNameTextBox.Text = ""
    $surnameTextBox.Text = ""
    $displayNameTextBox.Text = ""
    $descriptionTextBox.Text = ""
    $usernameTextBox.Text = ""
    $passwordTextBox.Text = ""
    $telephoneTextBox.Text = ""
    $emailTextBox.Text = ""
    $mobileTextBox.Text = ""
    $ipPhoneTextBox.Text = ""
    $jobTitleTextBox.Text = ""
    $departmentTextBox.Text = ""
    $companyTextBox.Text = ""
    $managerComboBox.Text = ""
    $logonScriptTextBox.Text = ""
    $homeDirectoryTextBox.Text = ""
    $firstInitialSurnameRadio.IsChecked = $true
    $usernameLabel.Foreground = [System.Windows.Media.Brushes]::Black
})

# Event handler for create/update button
$createButton.Add_Click({
    try {
        $username = $usernameTextBox.Text
        $existingUser = Get-ADUser -Filter {SamAccountName -eq $username} -ErrorAction SilentlyContinue
        $securePassword = if ($passwordTextBox.Text) { ConvertTo-SecureString $passwordTextBox.Text -AsPlainText -Force } else { $null }
        $managerText = $managerComboBox.Text
        $managerSam = if ($managerText -and $managerText.Contains(' ')) { $managerText.Split(' ')[0] } else { $managerText }
        $manager = Get-ADUser -Filter { SamAccountName -eq $managerSam -or DisplayName -eq $managerText } -ErrorAction SilentlyContinue
        $managerDN = if ($manager) { $manager.DistinguishedName } else { $null }

        if ($existingUser) {
            # Update existing user with non-account-dependent fields
            $setParams = @{
                Identity = $existingUser
                GivenName = $firstNameTextBox.Text
                Surname = $surnameTextBox.Text
                DisplayName = $displayNameTextBox.Text
                SamAccountName = $usernameTextBox.Text
                Description = $descriptionTextBox.Text
                OfficePhone = $telephoneTextBox.Text
                EmailAddress = $emailTextBox.Text
                MobilePhone = $mobileTextBox.Text
                Title = $jobTitleTextBox.Text
                Department = $departmentTextBox.Text
                Company = $companyTextBox.Text
                ErrorAction = 'Stop'
            }
            if ($ipPhoneTextBox.Text) { $setParams['OtherAttributes'] = @{'otherTelephone' = $ipPhoneTextBox.Text } }
            Set-ADUser @setParams
            if ($securePassword) { Set-ADAccountPassword -Identity $existingUser -NewPassword $securePassword -Reset -ErrorAction Stop }
            if ($managerDN) { Set-ADUser -Identity $existingUser -Manager $managerDN -ErrorAction Stop }
            if ($logonScriptTextBox.Text) { Set-ADUser -Identity $existingUser -ScriptPath $logonScriptTextBox.Text -ErrorAction Stop }
            if ($homeDirectoryTextBox.Text) { Set-ADUser -Identity $existingUser -HomeDirectory $homeDirectoryTextBox.Text -ErrorAction Stop }
            [System.Windows.MessageBox]::Show("User '$username' updated successfully.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } else {
            # Create new user
            if ($username -and $firstNameTextBox.Text -and $surnameTextBox.Text) {
                $adParams = @{
                    Name = $displayNameTextBox.Text
                    GivenName = $firstNameTextBox.Text
                    Surname = $surnameTextBox.Text
                    SamAccountName = $usernameTextBox.Text
                    UserPrincipalName = "$($usernameTextBox.Text)@domain.com"
                    Enabled = $true
                    ChangePasswordAtLogon = $true
                    Description = $descriptionTextBox.Text
                }
                if ($securePassword) { $adParams.AccountPassword = $securePassword }
                if ($jobTitleTextBox.Text) { $adParams.Title = $jobTitleTextBox.Text }
                if ($departmentTextBox.Text) { $adParams.Department = $departmentTextBox.Text }
                if ($companyTextBox.Text) { $adParams.Company = $companyTextBox.Text }
                if ($managerDN) { $adParams.Manager = $managerDN }
                if ($telephoneTextBox.Text) { $adParams.OfficePhone = $telephoneTextBox.Text }
                if ($emailTextBox.Text) { $adParams.EmailAddress = $emailTextBox.Text }
                if ($mobileTextBox.Text) { $adParams.MobilePhone = $mobileTextBox.Text }
                if ($ipPhoneTextBox.Text) { $adParams.OtherAttributes = @{'otherTelephone' = $ipPhoneTextBox.Text } }
                if ($logonScriptTextBox.Text) { $adParams.ScriptPath = $logonScriptTextBox.Text }
                if ($homeDirectoryTextBox.Text) { $adParams.HomeDirectory = $homeDirectoryTextBox.Text }
                New-ADUser @adParams -ErrorAction Stop
                [System.Windows.MessageBox]::Show("Account '$username' created successfully.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            } else {
                [System.Windows.MessageBox]::Show("First Name, Surname, and Username are required to create a new account.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
        }
    } catch {
        [System.Windows.MessageBox]::Show("Error: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Ensure the window is shown only if it hasn't been closed
if ($window.IsLoaded -eq $false) {
    $window.ShowDialog() | Out-Null
}
