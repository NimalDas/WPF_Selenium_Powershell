<#
.NOTES
    Title: FSToolkitDemo
    Author: Nimal Das
    GitHub:
    Version: 1.0.0.1
    Description: Service Now automation using Selenium and PowerShell
#>

# Set working directory
$location = "C:\Work\FSToolkit"
Set-Location -Path $location

function Close-AvailableRunspace{
    $available_runspaces = Get-Runspace | ?{$_.RunspaceAvailability -eq 'Available'}
    $available_runspaces | % {$_.dispose()}
}

# Adding env path of chrome binaries
$envpaths = $env:path.Split(';')
if(!($envpaths -contains $location))
    {
    $env:path += "$location;"
    }
if(!($envpaths -contains "$location\Chrome"))
    {
    $env:path += "$location\Chrome;"
    }

# Adding Selenium webdriver DLLs
Add-Type -Path "$location\WebDriver.dll"
Add-Type -Path "$location\WebDriver.Support.dll"

# Add required GUI assemblies
#Add-Type -AssemblyName presentationframework, presentationcore

# Create synchronised hashtable to hold WPF controls
$hash = [hashtable]::Synchronized(@{})
$hash.Host = $Host
$hash.location = $location

$XAML = Get-Content -Path .\MainWindow.xaml
$CleanXAML = $XAML -replace 'x:Class=".*?"', '' -replace 'x:N', 'N' -replace 'mc:Ignorable="d"', ''

# Convert XAML to XML
[XML]$XML = $CleanXAML

# Create XML reader
$XMLReader = New-Object System.Xml.XmlNodeReader $XML

# Create Window object
$window = [Windows.markup.xamlreader]::Load($XMLReader)
$hash.Window = $window

# Add control objects to the hashtable
$XML.SelectNodes("//*[@Name]") | %{$hash.$($_.name) = $($hash.window.findname($_.name))}

# Querying database
Import-Module SimplySql
Open-SQLiteConnection -ConnectionName dbconnect -DataSource C:\temp\TestDB.db
$teams = Invoke-SqlQuery -ConnectionName dbconnect -Query "select * from teams"
$fgroups = Invoke-SqlQuery -ConnectionName dbconnect -Query "select * from fgroups"
$engineers = Invoke-SqlQuery -ConnectionName dbconnect -Query "select * from Engineers"

# Create datacontectext for textblock
$datacontext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$observedValue = ''
$datacontext.Add($observedValue)
$hash.status_text_block.DataContext = $datacontext

# Create binding for viewbox
$binding = New-Object System.Windows.Data.Binding
$binding.Path = '[0]'
$binding.Mode = [System.Windows.Data.BindingMode]::OneWay
[void][System.Windows.Data.BindingOperations]::SetBinding($hash.status_text_block, [System.Windows.Controls.TextBlock]::TextProperty, $binding)

# Handling combo box
$hash.team_combo_box.AddText('Select Team')
$teams.name |%{$hash.team_combo_box.AddText($_)}
$hash.team_combo_box.SelectedIndex = 0

$hash.team_combo_box.add_DropDownClosed({
$hash.fgroup_list_box.Clear()
foreach($team in $teams)
    {
    if($team.name -eq $hash.team_combo_box.SelectedValue)
        {
        $hash.fgroup_list_box.Items.Clear()
        $FG = ($fgroups | ?{$_.UID_TEAM -eq $team.UID_TEAM}).name
        $FG | %{
            $checkbox = New-Object -TypeName System.Windows.Controls.CheckBox
            $checkbox.Content = $_
                    
            # Event if checkbox is checked
            $checkbox.Add_Checked({
                $engg_list = New-Object System.Collections.Generic.List[String]
                $s_fg = $this.content
                })
            # Event if checkbox is unchecked
            $checkbox.Add_Unchecked({
                $hash.engg_list_box.Items.Clear()
                })
            $hash.fgroup_list_box.AddChild($checkbox)
        }
    }
    if($hash.team_combo_box.SelectedValue -eq 'Select Team')
        {
        $hash.fgroup_list_box.Items.Clear()
        }   
    }
})

# Handling Rota button
$hash.Rota_btn.add_click({
$hash.engg_list_box.items.Clear()
$engg_list = @()
$hash.FG_selected = ($hash.fgroup_list_box.items | ?{$_.isChecked}).content
$hash.FG_selected = $fgroups | ?{$hash.FG_selected -contains $_.Name}
foreach($grp in $hash.FG_selected)
    {
    $FGRP_engg = @()
    $queue = "$($grp.Name)"+"_Q"
    for($i=0;$i -lt $engineers.Length;$i++)
        {
        $e_FGRPs = $engineers[$i].UID_FGRP.Split(',')
        if($e_FGRPs -contains $grp.UID_FGRP)
            {   
            $FGRP_engg += $engineers[$i].name
            #Write-Host $engg_list -ForegroundColor Cyan
            }
        }
    $hash.($queue) = $FGRP_engg | select -Unique
    $engg_list += $FGRP_engg
    }
    $engg_list | select -Unique | %{
    $e_checkbox = New-Object -TypeName System.Windows.Controls.CheckBox
    $e_checkbox.Content = $_
    $hash.engg_list_box.AddChild($e_checkbox)
    }
})

# Handling Start_btn
$hash.Start_Btn.Add_Click({
$hash.Engg_selected = ($hash.engg_list_box.items | ? {$_.isChecked}).content
$hash.qs = $hash.GetEnumerator() | ?{$_.Key -like "*_Q"}
$hash.qs | %{
    $hash.(($_.name)+'_index') = 0
    }
# Setting up queues
for($i=0;$i -lt $hash.qs.Count;$i++)
    {
    $q_e = @()
    $e = $hash.qs[$i].Value
    $e | % {
        for($j=0;$j -lt $hash.Engg_selected.Count;$j++)
            {
            if($_ -eq $hash.Engg_selected[$j])
                {
                $q_e += $_
                } 
            }
        }
    $hash.qs[$i].Value = $q_e
    $hash.($($hash.qs[$i].name)+'_index') = 0
    Write-Host $q_e -ForegroundColor DarkCyan
    if($hash.qs[$i].Value.count -eq 0)
        {
        [System.Windows.MessageBox]::Show("No engineer assigned to $($hash.qs[$i].name.TrimEnd('_Q'))")
        Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- No engineer assigned to $($hash.qs[$i].name)"
        }
    }
#Write-Host $hash.FG_selected.Name -ForegroundColor Cyan
#Write-Host $hash.qs.value -ForegroundColor Green
#<#
$codeBlock = {
$URL = '<Service Now URL>'
$chrome_driver_service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService()
$chrome_driver_service.SuppressInitialDiagnosticInformation = $true
$chrome_driver_service.HideCommandPromptWindow = $true
$chrome_options = New-Object -TypeName OpenQA.Selenium.Chrome.ChromeOptions
$chrome_options.AddArgument('--ignore-certificate-errors-spki-list')
$chrome_options.AddArgument('--ignore-urlfetcher-cert-requests')
$chrome_options.AddArgument('--start-maximized')
$chrome = New-Object -TypeName OpenQA.Selenium.Chrome.ChromeDriver -ArgumentList @($chrome_driver_service,$chrome_options)
<#
# WScript to handle certificate pop-up in chrome
$x = @"
Start-Sleep -Seconds 2
(New-Object -ComObject WScript.Shell).AppActivate('ServiceNow - Google Chrome')
Start-Sleep -Milliseconds 500
(New-Object -ComObject WScript.Shell).SendKeys('{ENTER}')
"@
Start-Process Powershell.exe -ArgumentList "-command $x"
#>
$chrome.Url = $URL
$hash.Chrome = $chrome
# functions
function Time-Stamp
    {
    return Get-Date -Format "dd-MM-yyyy @ HH:mm:ss"
    }

# Setting variables
$hash.flag = $true
$response = 0
$scriptroot = $hash.location
$log = "$scriptroot\logs.log"
$verboselogs = "$scriptroot\verboselogs.log"
$username = ($chrome.FindElementsByCssSelector('span[class="user-name hidden-xs hidden-sm hidden-md"]')).text

# Main
while($hash.flag)
    {
    $element = $content = ''
    sleep -Seconds 10
    $chrome.Url = $URL
    Write-Verbose "Refreshed  $(Get-Date)"
    $datacontext[0] =  "Refreshed  $(Get-Date)"
    start-sleep -Seconds 3
    try
        {
        [void]$chrome.SwitchTo().Frame(0)
        }
    catch
        {
        $response += 1
        if($response -gt 3)
            {
            Write-Verbose 'Script not working'
            $datacontext[0] = 'Script not working'
            Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- Script stopped working"
            }
        exit
        }

    # Expand collapsed groups
    $collapse_btns = $chrome.FindElementsByCssSelector('button[aria-label="Expand group"]')
    foreach($btn in $collapse_btns)
        {
        $btn.Click()
        }
    Start-Sleep -Seconds 2

    # Extract data
    $element = ($chrome.FindElementsByCssSelector("td")).text | Select-Object -First 1
    $element | Out-File "$scriptroot\records.txt" -Force
    $content = Get-Content "$scriptroot\records.txt"

    # Parse extracted data
    $parsedcontent = @()
    for($i=0;$i -lt $content.Count;$i++)
        {
        if($content[$i].StartsWith('Select record for action:'))
            {
            $parsedcontent += $content[$i+1]
            }
        else
            {
            continue
            }
        }
    #Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- Parsed extracted content"

    # Filter based on queue and assigned_to
    # Use REGeX to extract information from Service Now page
    $unassigned = @()
    foreach($line in $parsedcontent)
        {
        $line -match '(?<incident>^INC[0-9]+)\s+?(?<assigned_to>.+)?(?<queue>XXXX.+?)\s' | Out-Null
        $incident = $Matches['incident']
        $assigned_to = $Matches['assigned_to']
        $queue = $Matches['queue']
        if(($assigned_to -match '\(empty\)') -and ($hash.FG_selected.name -contains $queue))
            {
            $ticket = @{
                        Incident = $incident
                        Queue = $queue
                        }
            $unassigned += $ticket
            }
        }
    #Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- Filter parsed content`r`n`t`t$($unassigned.incident)"
    foreach($record in $unassigned)
        {
        [void]$chrome.SwitchTo().DefaultContent()
        ($chrome.FindElementById('sysparm_search')).Clear()
        ($chrome.FindElementById('sysparm_search')).SendKeys("$($record.incident)")
        Start-Sleep -Seconds 2
        ($chrome.FindElementById('sysparm_search')).SendKeys([OpenQA.Selenium.Keys]::Enter)
        Start-Sleep -Seconds 3
        [void]$chrome.SwitchTo().Frame(0)

        # Wait for "az_assign_to_me" button to be clickable
        [OpenQA.Selenium.Support.UI.WebDriverWait]$wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($chrome, [System.TimeSpan]::FromSeconds(15))
        $wait.PollingInterval = 2
        try 
            {
            [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable([OpenQA.Selenium.By]::CssSelector('button[id="az_assign_to_me"]')))
            }
        catch 
            {
            Write-Verbose ("Exception with 'az_assign_to_me': {0} ...`n(ignored)" -f $_.Exception.Message)
            Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $($_.Exception.Message)"
            continue
            }

        # Click 'az_assign_to_me' button
        $assigntomebtn = $chrome.FindElementByCssSelector('button[id="az_assign_to_me"]')
        $assigntomebtn.click()
        Start-Sleep -Seconds 2
        Write-Verbose "Assigntome button Clicked"
        Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- Assign to me button clicked"

        # Wait for provide info button to load to make sure ticket is self assigned
        [OpenQA.Selenium.Support.UI.WebDriverWait]$wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($chrome, [System.TimeSpan]::FromSeconds(15))
        $wait.PollingInterval = 2
        try 
            {
            [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable([OpenQA.Selenium.By]::CssSelector('button[id="az_provider_info_required"]')))
            }
        catch 
            {
            Write-verbose ("Exception with 'az_provider_info_required': {0} ...`n(ignored)" -f $_.Exception.Message)
            Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- Unable to take ownership of $($record.incident)`r`n`t`t$($_.Exception.Message)"
            continue
            }
        Write-Verbose "$($record.incident) self assigned"
        Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $($record.incident) self assigned"

        # Selecting queue
        $inci_q = $hash.qs | ?{$_.name.trimend('_Q') -eq $record.queue}

        # Extract worknotes in the incident to check work history
        $worknotes = ($chrome.FindElementsByCssSelector('li[class="h-card h-card_md h-card_comments"]')).text  | ?{$_ -like "*Work notes*"}

        # Loop through rota list to see if the Eng has previously worked; loop breaks after finding the first match
        :Outer foreach($worknote in $worknotes)
            {
            foreach($Eng in $inci_q.Value)
                {
                $history = $pEng = ''
                $history =  $worknote | ?{$_ -match $Eng}
                if($history)
                    {
                    Write-Verbose "$Eng has the recent update on $($record.incident)"
                    Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $Eng has the recent update on $($record.incident)"
                    Add-Content -Path $log -Value "$(Time-Stamp) -- $Eng has the recent update on $($record.incident)"
                    $pEng = $Eng
                    break Outer
                    }
                }
            }
        if($pEng)
            {
            if($pEng -eq $username)
                {
                Write-Verbose "$($record.incident) already assigned"
                $datacontext[0] = "$($record.incident) assigned to $pEng"
                Add-Content -Path $log -Value "$(Time-Stamp) -- $($record.incident) Assigned to -- $pEng"
                Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $($record.incident) Assigned to -- $pEng"
                continue
                }
            else
                {
                Write-Verbose "$pEng already worked on $($record.incident)"
                Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $pEng already worked on $($record.incident)"
                $assigned_to_field = ''
                $assigned_to_field  = $chrome.FindElementByCssSelector('input[id="sys_display.incident.assigned_to"]')
                $assigned_to_field.clear()
                Start-Sleep -Seconds 2
                $assigned_to_field.SendKeys($pEng)
                Start-Sleep -Seconds 2
        
                # Click save button
                $save = $chrome.FindElementByCssSelector('button[id="az_submit"]')
                $save.click()
                Start-Sleep -Seconds 2
        
                # Wait for the page to load after Save_button is clicked, look for az_inc_assign_to_me button
                [OpenQA.Selenium.Support.UI.WebDriverWait]$wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($chrome, [System.TimeSpan]::FromSeconds(15))
                $wait.PollingInterval = 2
                try 
                    {
                    [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::CssSelector('button[id="az_inc_assign_to_me"]')))
                    [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::TextToBePresentInElementValue([OpenQA.Selenium.By]::CssSelector('input[id="sys_display.incident.assigned_to"]'), $pEng))
                    }
                catch
                    {
                    Write-Verbose ("Exception with 'Assign_to_field loading': {0} ...`n(ignored)" -f $_.Exception.Message)
                    Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- Unable to assign $($record.incident) to $pEng"
                    Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $($_.Exception.Message)" 
                    continue
                    }
                Write-Verbose "Page loaded after clicking save button and successfully assigned to $pEng" 
                Add-Content -Path $log -Value "$(Time-Stamp) -- successfully assigned $($record.incident) to $pEng"
                Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- successfully assigned $($record.incident) to $pEng"
                $datacontext[0] = "$($record.incident) assigned to $pEng"
                }
            }
        else
            {
            Write-Verbose "No work history found on $($record.incident)"
            Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- No work history found on $($record.incident)"
            $c_q = $record.queue
            $index = $hash.($($c_q)+'_Q_index')
            if($inci_q.Value[$index] -eq $username)
                {
                Write-Verbose "$($record.incident) already assigned to $($inci_q.Value[$index])" 
                Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $($record.incident)  Assigned to -- $($inci_q.Value[$index])"
                Add-Content -Path $log -Value "$(Time-Stamp) -- $($record.incident)  Assigned to -- $($inci_q.Value[$index])"
                $datacontext[0] = "$($record.incident) assigned to $($inci_q.Value[$index])"
                $index += 1
                if($index -eq $inci_q.Value.Count){$index = 0}
                $hash.($($c_q)+'_Q_index') = $index
                Write-Verbose "Next in queue is $($inci_q.Value[($index+1)])"
                Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- Next in queue is $($inci_q.Value[($index+1)])"
                continue
                }
            else{
                $assigned_to_field = ''
                $assigned_to_field  = $chrome.FindElementByCssSelector('input[id="sys_display.incident.assigned_to"]')
                $assigned_to_field.clear()
                Start-Sleep -Seconds 2
                $assigned_to_field.SendKeys($inci_q.Value[$index])
                Start-Sleep -Seconds 2
        
                # Click save button
                $save = $chrome.FindElementByCssSelector('button[id="az_submit"]')
                $save.click()
                Start-Sleep -Seconds 2
        
                # Wait for the page to load after Save_button is clicked, look for az_inc_assign_to_me button
                [OpenQA.Selenium.Support.UI.WebDriverWait]$wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($chrome, [System.TimeSpan]::FromSeconds(15))
                $wait.PollingInterval = 2
                try 
                    {
                    [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::CssSelector('button[id="az_inc_assign_to_me"]')))
                    [void]$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::TextToBePresentInElementValue([OpenQA.Selenium.By]::CssSelector('input[id="sys_display.incident.assigned_to"]'), $($inci_q.Value[$index])))
                    $datacontext[0] = "$($record.incident) assigned to $($inci_q.Value[$index])"
                    $index += 1
                    
                    }
                catch 
                    {
                    Write-Verbose ("Exception with 'Assign_to_field loading': {0} ...`n(ignored)" -f $_.Exception.Message)
                    Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- Unable to assign $($record.incident) to $($inci_q.Value[$index])"
                    Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $($_.Exception.Message)"
                    continue
                    }
                Write-Verbose "Page loaded after clicking save button and successfully assigned to $($inci_q.Value[$index])"
                Add-Content -Path $verboselogs -Value "$(Time-Stamp) -- $($record.incident) Assigned to $($inci_q.Value[$index])"
                Add-Content -Path $log -Value "$(Time-Stamp) -- $($record.incident) Assigned to $($inci_q.Value[$index])"
                if($index -eq $inci_q.Value.Count){$index = 0}
                $hash.($($c_q)+'_Q_index') = $index
                Write-Verbose "Next in queue is $($inci_q.Value[($index)])"
                }
            }
        }# main Foreach
    #
    }
##
}
#>

# Create Runspace
$runspace = [runspacefactory]::CreateRunspace()
$rsid = $runspace.Id
$hash.rsid = $rsid
$runspace.ApartmentState = 'STA'
$runspace.ThreadOptions = 'ReuseThread'
$runspace.Open()
$runspace.SessionStateProxy.SetVariable('Hash', $hash)
$runspace.SessionStateProxy.SetVariable('datacontext', $datacontext)
   
$powershell = [powershell]::Create()
$powershell.runspace = $runspace
$powershell.addscript($codeblock)
$handle = $powershell.begininvoke()
$hash.Start_Btn.IsEnabled = $false
})

# Event handling of Stop_btn
$hash.stop_btn.Add_Click({
   $rs = Get-Runspace -Id $hash.rsid
   if($rs){
       $rs.Close()
       $rs.Dispose()
       Close-AvailableRunspace
        }
   $hash.chrome.quit()
   $hash.Start_Btn.IsEnabled = $true
})

Close-SqlConnection -ConnectionName dbconnect
$hash.window.showdialog()

#[void]$Hash.Window.Dispatcher.InvokeAsync{$Hash.Window.ShowDialog()}.Wait()

