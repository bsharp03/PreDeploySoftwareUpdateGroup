############################################################################################################################################## 
# AUTHOR  : Brian C. Sharp   
# DATE    : 04/22/2019
# COMMENT :
# Script will create a collection with random workstations in each subnet to help deploy Microsoft Updates via BranchCache
# For this script to work install the following components from https://www.microsoft.com/en-us/download/confirmation.aspx?id=43339
# Microsoft® System CLR Types for Microsoft® SQL Server® 2012 (SQLSysClrTypes.msi)
# Microsoft® SQL Server® 2012 Shared Management Objects (SharedManagementObjects.msi)
# Microsoft® Windows PowerShell Extensions for Microsoft® SQL Server® 2012 (PowerShellTools.msi)
############################################################################################################################################## 




#Set Parameters
    $SCCMServer = "EMO-CMCGTVPR-01"
    $CMSiteCode = "CGT"
    $Database = "ConfigMgr"
    $DatabaseServer = "EMO-CMCGTVDB-01"
    $Database = "ConfigMgr"
    $DeploymentTime = "22:00"
    $TodaysDate = (Get-Date)
    $CollectionName = "BranchCache Collection for Monthly Patch Deployment"
    
    
   #Select an IP Address from Each Subnet using SQL
   Import-Module sqlps -DisableNameChecking

$PreBranchCastHosts=@()
$SelectedSystems=@()
$IPSubnets = invoke-sqlcmd -ServerInstance "$DatabaseServer"  -Database $Database -Query "Select Distinct IP_Subnets0 from v_RA_System_IPSubnets"
ForEach ($IPSubnet in $IPSubnets)
        {
            If ($IPSubnet.IP_Subnets0 -notmatch "169.254")
                {
                    $IPSubnetWithQuotes = "'"+ $IPSubnet.IP_Subnets0 + "'"
                    $SelectedSystems = Invoke-Sqlcmd -ServerInstance "$DatabaseServer"  -Database $Database -Query "Select Top 2 v_r_system.Name0, v_r_system.ResourceID from v_R_System Join v_RA_System_IPSubnets on v_RA_System_IPSubnets.ResourceID = v_R_System.ResourceID  Where Operating_System_Name_and0 = 'Microsoft Windows NT Workstation 10.0' and Client0 = '1' and Active0 = '1' and Obsolete0 = '0' and (DATEDIFF(dd,Last_Logon_Timestamp0,GetDate())) < 30 and IP_Subnets0 = $IPSubnetWithQuotes"
                    $IPSubnetWithQuotes | Out-File C:\Users\bcsharp-admin\Desktop\selectedSystems.txt -Append
                    If ($SelectedSystems -ne $null) {ForEach($SelectedSystem in $SelectedSystems) {$PreBranchCastHosts+=$SelectedSystem}}
                }
        }


Set-Location c:\
Import-Module "\\$sccmServer\d$\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager\ConfigurationManager.psd1"
New-PSDrive -Name $CMSiteCode -PSProvider CMSite -Root $SCCMServer
$CMSiteCode += ":"
Set-Location $CMSiteCode


$Sugs = Get-CMSoftwareUpdateGroup

#Build Form
Add-Type -assembly System.Windows.Forms
$main_form = New-Object System.Windows.Forms.Form
$main_form.Width = 400
$main_form.Height = 200

$SugLabel = New-Object System.Windows.Forms.Label
$SugLabel.Text = "Software Update Groups"
$SugLabel.Location = New-Object System.Drawing.Point(75,5)
$SugLabel.AutoSize = $true
$SugComboBox = New-Object System.Windows.Forms.ComboBox
$SugComboBox.Width = 300
ForEach ($Sug in $Sugs) {$SugComboBox.Items.Add($Sug.LocalizedDisplayName)}
$SugComboBox.Location  = New-Object System.Drawing.Point(75,25)

$DatePickerLabel = New-Object System.Windows.Forms.Label
$DatePickerLabel.Text = "Deployment Time"
$DatePickerLabel.Location = New-Object System.Drawing.Point(75,55)
$DatePickerLabel.AutoSize = $true

$datepicker = New-Object System.Windows.Forms.DateTimePicker
$datepicker.Location = New-Object System.Drawing.Point(75,75)
$datepicker.Format="Custom"


$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(100,120)
$Button.Size = New-Object System.Drawing.Size(120,34)
$Button.Text = "Generate Collection and Deploy SUG"
$Button.Add_Click(
                    {
                        #If Collection Already Exist, delete the Collection and Associated Deployments
                        $CMDeviceCollections = $null
                        $CMDeviceCollections =  Get-CMDeviceCollection -Name $CollectionName
                        If ($CMDeviceCollections -ne $null) {Remove-CMCollection -Name $CollectionName -Force}
                        $Schedule = New-CMSchedule -Start "01/01/2016 12:00 AM" -RecurInterval Days -RecurCount 1
                        New-CMDeviceCollection -LimitingCollectionId SMS00001 -Name $CollectionName -RefreshType Periodic -RefreshSchedule $Schedule
                        
                        #Add Selected Host to Collection
                            ForEach ($PreBranchCastHost in $PreBranchCastHosts) {Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId (Get-CMDevice -Name $PreBranchCastHost.Name0).ResourceID}
                        
                        $SelectedSUG = (Get-CMSoftwareUpdateGroup -Name $SugComboBox).SelectedItem.CI_ID
                        $DeploymentDate = [datetime]$datepicker.text + $DeploymentTime
                        New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $SugComboBox.SelectedItem -CollectionName $CollectionName -DeploymentName "BranchCache Collection for Monthly Patch Deployment $TodaysDate" -Description "Software Update Group Deployment to Pre-Stage Updates for BranchCache" -UserNotification DisplaySoftwareCenterOnly -DeploymentType Required -VerbosityLevel OnlyErrorMessages -RestartServer $False -RestartWorkstation $False -DownloadFromMicrosoftUpdate $False -AvailableDateTime $DeploymentDate -DeadlineDateTime $DeploymentDate -AcceptEula -UseBranchCache $True -UnprotectedType UnprotectedDistributionPoint -Verbose
                        
                        $Main_Form.Close()
                    }
                   )
    

$main_form.Controls.Add($SugLabel)
$main_form.Controls.Add($SugComboBox)
$main_form.Controls.Add($DatePickerLabel)
$main_form.Controls.Add($datepicker)
$main_form.Controls.Add($Button)
$main_form.ShowDialog()


    

