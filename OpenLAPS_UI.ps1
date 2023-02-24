#OPENLAPS_UI
#Thomas Hahnke
#A basic GUI that queries AD for LAPS passwords and allows users to update the expiration time.

#PS2EXE commands
#ps2exe .\OpenLAPS_UI.ps1 -noOutput -noConsole -version 1

#REQUIREMENTS
#RSAT installed : https://www.microsoft.com/en-us/download/details.aspx?id=45520

#Imports the necessary components for the GUI
Add-Type -AssemblyName PresentationFramework, System.Windows.Forms

#Imports the active directory module
Set-Variable ProgressPreference SilentlyContinue
import-module activedirectory

#XAML GUI

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

[xml]$XAML=@"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="OpenLAPS" Height="300" Width="365">
    <Grid  Background="gainsboro">
        <Label Content="Computer Name" HorizontalAlignment="Left" Height="27" Margin="16,15,0,0" VerticalAlignment="Top" Width="190" HorizontalContentAlignment="Left"/>
        <TextBox Name="txtComputerName" HorizontalAlignment="Left" Margin="20,37,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="250"/>
        <Button IsDefault="True" Name= "btnSearch" Content="Search" HorizontalAlignment="Left" Margin="275,37,0,0" VerticalAlignment="Top" Width="50"/>
        <Label Content="Password" HorizontalAlignment="Left" Height="27" Margin="16,60,0,0" VerticalAlignment="Top" Width="190" HorizontalContentAlignment="Left"/>
        <TextBox Name="txtPassword" 
                 HorizontalAlignment="Left" 
                 Margin="20,85,0,0" 
                 TextWrapping="Wrap" 
                 VerticalAlignment="Top" 
                 Width="250"
                 Background="gainsboro"
                 FontFamily="Consolas"
                 FontSize="24"
                 FontStretch="UltraExpanded"
                 TextAlignment="Left"
                 />
        <Button Name="btnClippy" Content="Copy" HorizontalAlignment="Left" Margin="275,90,0,0" VerticalAlignment="Top" Width="50"/>
        <Label Content="Password Expiration" HorizontalAlignment="Left" Height="27" Margin="16,120,0,0" VerticalAlignment="Top" Width="190" HorizontalContentAlignment="Left"/>
        <TextBox Name="txtPasswordExp"
                 HorizontalAlignment="Left" 
                 Margin="20,142,0,0" 
                 TextWrapping="Wrap" 
                 VerticalAlignment="Top" 
                 Width="250"
                 Background="gainsboro"
                 TextAlignment="Left"
                 />
        <Label Content="Set New Expiration Time" HorizontalAlignment="Left" Height="27" Margin="16,165,0,0" VerticalAlignment="Top" Width="190" HorizontalContentAlignment="Left"/>
        <DatePicker Name="dateExp" HorizontalAlignment="Left" Margin="20,187,0,0" VerticalAlignment="Top" Width="250"/>
        <Button Name="btnSet" Content="Set" HorizontalAlignment="Left" Margin="275,189,0,0" VerticalAlignment="Top" Width="50"/>
        <TextBox Name="txtError"
                 HorizontalAlignment="Left" 
                 Margin="-1,240,0,0" 
                 TextWrapping="Wrap" 
                 VerticalAlignment="Top" 
                 Width="365"
                 Background="gainsboro"
                 TextAlignment="Left"
                 Height="35"
                 />
    </Grid>
</Window>
"@

#Read XAML
$reader = New-Object System.Xml.XmlNodeReader($xaml)
try { 
    $form = [Windows.Markup.XamlReader]::Load($reader) 
} 
catch { 
    $_
    #exit 
}

#initializes each component of the Window
$xaml.SelectNodes('//*[@Name]') | 
    ForEach-Object{
        Set-Variable -Name $_.Name -Value $Form.FindName($_.Name)
        echo $_
    }

#Sets the password and password expiration text boxes to read only
$txtPassword.IsReadOnly = $True;
$txtPasswordExp.IsReadOnly = $True;
$dateExp.SelectedDate = Get-Date;


#Sets Button calls
$btnSearch.add_Click(
    {
        if ([string]::IsNullOrEmpty($txtComputerName.Text))
        {
            $txterror.Text="Computer Name cannot be empty"
        }
        else
        {
            LapsInfo -Computer $txtComputerName.Text
        }
    }
)

$btnClippy.add_Click(
    {
        if ([string]::IsNullOrEmpty($txtComputerName.Text))
        {
            $txterror.Text="Computer Name cannot be empty"
        }
        elseif ([string]::IsNullOrEmpty($txtPassword.Text))
        {
            $txterror.Text="Password field is empty"
        }
        else
        {
            try 
            { 
                Write-Output $txtPassword.Text | Set-Clipboard
            } 
            catch 
            { 
                $txterror.Text="Unkown Error Occured"
            }
        }
    }
)

$btnSet.add_Click(
    {
        if ([string]::IsNullOrEmpty($txtComputerName.Text))
        {
            $txterror.Text="Computer Name cannot be empty"
        }
        else
        {
            Set-LAPSInfo -Computer $txtComputerName.Text -Date $dateExp.Text
            Write-Host (Get-Date $dateExp.Text).TofileTime()
        } 
    }
)

#Function that is called to query LAPS information
Function Get-LAPSInfo {
    [CmdletBinding()]
    param (
    [Parameter(Mandatory)]
    [string]$Computer
    )
        try
        {
            $LAPSInfo = Get-ADComputer $Computer -Property ms-Mcs-AdmPwd, ms-Mcs-AdmPwdExpirationTime
            $txtPassword.Text=$LAPSInfo.'ms-Mcs-AdmPwd'
            $txtPasswordExp.Text=[DateTime]::FromFiletime([Int64]::Parse($LAPSInfo.'ms-Mcs-AdmPwdExpirationTime'))
        }
        catch 
        { 
            $txterror.Text="Computer not found"
            $txtPassword.Text=""
            $txtPasswordExp.Text=""
        }
}

#Function that is called to update the expiration date
Function Set-LAPSInfo {
    [CmdletBinding()]
    param (
    [Parameter(Mandatory)]
    [string]$Computer,
    [Parameter(Mandatory)]
    [string]$Date
    )
        try
        {
            Get-ADComputer $Computer | Set-ADObject -Replace @{"ms-mcs-AdmPwdExpirationTime"=(Get-Date $Date)}
        }
        catch 
        { 
            $txterror.Text="Unkown Error Occured"
        }
}
# Shows the form
$Form.ShowDialog()