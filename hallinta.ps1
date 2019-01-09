param(
    [String]$path,
    [Switch]$debug
)

class Host
{
    # Keeps track of the hosts and performs actions on them
    static [Host[]]$Hosts = @()
    [String]$Name
    [String]$Mac
    [Bool]$Status
    [Int]$Column
    [Int]$Row

    Host([String]$name, [String]$mac, [Int]$column, [Int]$row)
    {
        $this.Name = $name
        $this.Mac = $mac
        $this.Column = $column
        $this.Row = $row
    }

    static [void] Populate([String]$path, [String]$delimiter)
    {
        # Creates [Host] objects from given .csv file
        Write-Host -NoNewline "Populating from "
        Write-Host -ForegroundColor Yellow $path
        [Host]::Hosts = @()
        Import-Csv $path -Delimiter $delimiter | ForEach-Object {
            $h = [Host]::new($_.Name, $_.Mac, [Int]$_.Column, [Int]$_.Row)
            $pingJob = Test-Connection -ComputerName $h.Name -Count 1 -AsJob
            $h | Add-Member -NotePropertyName "pingJob" -NotePropertyValue $pingJob -Force
            [Host]::Hosts += $h
        }
        $needToExport = $false
        foreach($h in [Host]::Hosts)
        {
            $h.pingJob | Wait-Job # Wait for the first ping to complete
            if((Receive-Job $h.pingJob).StatusCode -eq 0){ $h.Status = $true } # else $false
            if(!$h.Mac) # Try to get missing mac-address if not populated from the file
            {
                Write-Host -NoNewline -ForegroundColor Red "Missing mac-address of "
                Write-Host -NoNewline -ForegroundColor Gray $h.Name
                if($h.Status)
                {
                    Write-Host -ForegroundColor Red ", retrieving and saving to file"
                    $needToExport = $true
                    $h.Mac = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $h.Name | Select-Object -First 1 -ExpandProperty MACAddress
                }
                else
                {
                    Write-Host -ForegroundColor Red ", unable to connect to offline host!"
                }
            }
            # Display populate status in console
            Write-Host -NoNewline -ForegroundColor Gray $h.Name
            Write-Host -NoNewline (": mac={0}, status=" -f $h.Mac)
            if($h.Status)
            {
                $color = "Green"
            }
            else
            {
                $color = "Red"
            }
            Write-Host -NoNewline -ForegroundColor $color $h.Status
            Write-Host (", column={0}, row={1}" -f $h.Column, $h.Row)
            $h.pingJob | Remove-Job
        }
        if($needToExport){ [Host]::Export($path, $delimiter) } # Save received mac-address back to file
    }

    static [void] Display()
    {
        # Displays hosts in the $script:table
        Write-Host "Displaying"
        $cellSize = 100
        $script:table.Rows | ForEach-Object {$_.Cells | ForEach-Object { $_.Value = ""; $_.ToolTipText = "" }} # Reset cells
        $script:table.ColumnCount = ([Host]::Hosts | ForEach-Object {$_.Column} | Measure-Object -Maximum).Maximum
        $script:table.RowCount = ([Host]::Hosts | ForEach-Object {$_.Row} | Measure-Object -Maximum).Maximum
        $script:table.Columns | ForEach-Object {
            $_.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
            $_.HeaderText = [Char]($_.Index + 65) # Sets the column headers to A, B, C...
            $_.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
            $_.Width = $cellSize
        }
        $script:table.Rows | ForEach-Object {
            $_.HeaderCell.Value = [String]($_.Index + 1) # Sets the row headers to 1, 2, 3...
            $_.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
            $_.Height = $cellSize
        }
        $script:root.MinimumSize = [System.Drawing.Size]::new(($cellSize * $script:table.ColumnCount + $script:table.RowHeadersWidth + 20), ($cellSize * $script:table.RowCount + $script:table.ColumnHeadersHeight + 65))
        $script:root.Size = [System.Drawing.Size]::new(315, $script:root.MinimumSize.Height) # The size of the window can't be smaller than the minimum size (sorry for the hardcode btw)
        foreach($h in [Host]::Hosts)
        {
            $cell = $script:table[($h.Column - 1), ($h.Row - 1)]
            $cell.Value = $h.Name
            $cell.Style.Font = [System.Drawing.Font]::new($cell.InheritedStyle.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
            # $cell.ToolTipText = $h.Mac
            if($h.Status)
            {
                $cell.Style.ForeColor = [System.Drawing.Color]::Green
                $cell.Style.SelectionForeColor = [System.Drawing.Color]::Green
            }
            else
            {
                $cell.Style.ForeColor = [System.Drawing.Color]::Red
                $cell.Style.SelectionForeColor = [System.Drawing.Color]::Red
            }
        }
    }

    static [void] Export([String]$path, [String]$delimiter)
    {
        [Host]::Hosts | Select-Object Name, Mac, Column, Row | Export-Csv $path -Delimiter $delimiter -NoTypeInformation
    }

    static [String[]] GetActive()
    {
        # Returns the names of the hosts that are online and selected
        return ([Host]::Hosts | Where-Object {$_.Status -and ($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Name})
    }

    static [String[]] GetMacs()
    {
        # Returns the mac addresses of selected hosts
        return ([Host]::Hosts | Where-Object {($script:table[($_.Column - 1), ($_.Row - 1)]).Selected -and $_} | ForEach-Object {$_.Mac})
    }
}

function StartVBS3Form
{
    # Command with GUI to run VBS3 with specified startup parameters
    $states = [ordered]@{
        "Kokonäyttö" = ""
        "Ikkuna" = "-window"
        "Palvelin" = "-server"
        "Simulation Client" = "simulationClient=0"
        "After Action Review" = "simulationClient=1"
        "SC + AAR" = "simulationClient=2"
    }
    $form = [System.Windows.Forms.Form]::new()
    $form.AutoSize = $true
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
    $form.Text = "VBS3 - Käynnistä"

    $grid = [System.Windows.Forms.TableLayoutPanel]::new()
    $grid.AutoSize = $true
    $grid.ColumnCount = 2
    $grid.Padding = [System.Windows.Forms.Padding]::new(10)
    $grid.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Inset

    $statePanel = [System.Windows.Forms.FlowLayoutPanel]::new()
    $statePanel.AutoSize = $true
    $statePanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $states.Keys | ForEach-Object {
        $r = [System.Windows.Forms.RadioButton]::new()
        $r.AutoSize = $true
        $r.Margin = [System.Windows.Forms.Padding]::new(0)
        $r.Text = $_
        if($r.Text -eq "Kokonäyttö"){ $r.Checked = $true } 
        $statePanel.Controls.Add($r)
    }
    $grid.SetCellPosition($statePanel, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 0)) 
    $grid.Controls.Add($statePanel)

    $adminCheckBox = [System.Windows.Forms.CheckBox]::new()
    $adminCheckBox.Text = "Admin"
    $grid.SetCellPosition($adminCheckBox, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 1)) 
    $grid.Controls.Add($adminCheckBox)

    $multicastCheckBox = [System.Windows.Forms.CheckBox]::new()
    $multicastCheckBox.Text = "Multicast"
    $multicastCheckBox.Checked = $true
    $grid.SetCellPosition($multicastCheckBox, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 2)) 
    $grid.Controls.Add($multicastCheckBox)

    $configLabel = [System.Windows.Forms.Label]::new()
    $configLabel.Text = "cfg="
    $configLabel.AutoSize = $true
    $configLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $grid.SetCellPosition($configLabel, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(0, 3)) 
    $grid.Controls.Add($configLabel)
    $configTextBox = [System.Windows.Forms.TextBox]::new()
    $configTextBox.Width = 200
    $grid.SetCellPosition($configTextBox, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 3)) 
    $grid.Controls.Add($configTextBox)

    $connectLabel = [System.Windows.Forms.Label]::new()
    $connectLabel.Text = "connect="
    $connectLabel.AutoSize = $true
    $connectLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $grid.SetCellPosition($connectLabel, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(0, 4)) 
    $grid.Controls.Add($connectLabel)
    $connectTextBox = [System.Windows.Forms.TextBox]::new()
    $connectTextBox.Width = 200
    $grid.SetCellPosition($connectTextBox, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 4)) 
    $grid.Controls.Add($connectTextBox)

    $cpuCountLabel = [System.Windows.Forms.Label]::new()
    $cpuCountLabel.Text = "cpuCount="
    $cpuCountLabel.AutoSize = $true
    $cpuCountLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $grid.SetCellPosition($cpuCountLabel, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(0, 5)) 
    $grid.Controls.Add($cpuCountLabel)
    $cpuCountTextBox = [System.Windows.Forms.TextBox]::new()
    $cpuCountTextBox.Width = 200
    $grid.SetCellPosition($cpuCountTextBox, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 5)) 
    $grid.Controls.Add($cpuCountTextBox)

    $exThreadsLabel = [System.Windows.Forms.Label]::new()
    $exThreadsLabel.Text = "exThreads="
    $exThreadsLabel.AutoSize = $true
    $exThreadsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $grid.SetCellPosition($exThreadsLabel, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(0, 6)) 
    $grid.Controls.Add($exThreadsLabel)
    $exThreadsTextBox = [System.Windows.Forms.TextBox]::new()
    $exThreadsTextBox.Width = 200
    $grid.SetCellPosition($exThreadsTextBox, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 6)) 
    $grid.Controls.Add($exThreadsTextBox)

    $maxMemLabel = [System.Windows.Forms.Label]::new()
    $maxMemLabel.Text = "maxMem="
    $maxMemLabel.AutoSize = $true
    $maxMemLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $grid.SetCellPosition($maxMemLabel, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(0, 7)) 
    $grid.Controls.Add($maxMemLabel)
    $maxMemTextBox = [System.Windows.Forms.TextBox]::new()
    $maxMemTextBox.Width = 200
    $grid.SetCellPosition($maxMemTextBox, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 7)) 
    $grid.Controls.Add($maxMemTextBox)

    $parameterLabel = [System.Windows.Forms.Label]::new()
    $parameterLabel.Text = "Muut parametrit:"
    $parameterLabel.AutoSize = $true
    $parameterLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $grid.SetCellPosition($parameterLabel, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(0, 8)) 
    $grid.Controls.Add($parameterLabel)
    $parameterTextBox = [System.Windows.Forms.TextBox]::new()
    $parameterTextBox.Width = 200
    $grid.SetCellPosition($parameterTextBox, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(1, 8)) 
    $grid.Controls.Add($parameterTextBox)

    $runButton = [System.Windows.Forms.Button]::new()
    $runButton | Add-Member @{StatePanel=$statePanel; States=$states; AdminCheckbox=$adminCheckBox; MulticastCheckBox=$multicastCheckBox; ConfigTextBox=$configTextBox; ConnectTextBox=$connectTextBox; CpuCountTextBox=$cpuCountTextBox; ExThreadsTextBox=$exThreadsTextBox; MaxMemTextBox=$maxMemTextBox; ParameterTextBox=$parameterTextBox; Form=$form} -PassThru -Force
    $runButton.Text = "Käynnistä"
    $runButton.Add_Click({
        $state = $this.StatePanel.Controls | Where-Object {$_.Checked} | Select-Object -ExpandProperty Text
        $argument = $this.States[$state]
        if($this.AdminCheckbox.Checked){ $argument = ("{0} -admin" -f $argument)}
        if(!$this.MulticastCheckBox.Checked){ $argument = ("{0} -multicast=0" -f $argument)}
        if($this.ConfigTextBox.Text){ $argument = ("{0} -cfg={1}" -f $argument, $this.ConfigTextBox.Text)}
        if($this.ConnectTextBox.Text){ $argument = ("{0} -connect={1}" -f $argument, $this.ConnectTextBox.Text)}
        if($this.CpuCountTextBox.Text){ $argument = ("{0} -cpuCount={1}" -f $argument, $this.CpuCountTextBox.Text)}
        if($this.ExThreadsTextBox.Text){ $argument = ("{0} -exThreads={1}" -f $argument, $this.ExThreadsTextBox.Text)}
        if($this.MaxMemTextBox.Text){ $argument = ("{0} -maxMem={1}" -f $argument, $this.MaxMemTextBox.Text)}
        if($this.ParameterTextBox.Text){ $argument = ("{0} {1}") -f $argument, $this.ParameterTextBox.Text }
        Start-ProgramOnTarget -target ([Host]::GetActive()) -executable "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\VBS3_64.exe" -argument $argument
        $this.Form.Close()
    })
    $runButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $form.AcceptButton = $runButton
    $grid.SetCellPosition($runButton, [System.Windows.Forms.TableLayoutPanelCellPosition]::new(0, 9))
    $grid.SetColumnSpan($runButton, 2)
    $grid.Controls.Add($runButton)
    $form.Controls.Add($grid)
    $this | Add-Member @{StartVBS3Form=$form} -PassThru -Force
}

function Invoke-CommandOnTarget([String[]]$target, [Scriptblock]$command, [Object]$params = @(), [Bool]$asJob = $true, [Bool]$output = $true)
{
    if(!$target){ return }
    if($output)
    {
        Write-Host -NoNewline "Running "
        Write-Host -NoNewline -ForegroundColor Yellow $command
        Write-Host -NoNewline " on "
        Write-Host -ForegroundColor Gray -Separator ", " $target
    }
    if($asJob)
    {
        Invoke-Command -ComputerName $target -Credential $script:credential -ScriptBlock $command -ArgumentList $params -AsJob
    }
    else
    {
        Invoke-Command -ComputerName $target -Credential $script:credential -ScriptBlock $command -ArgumentList $params | Write-Host
    }
}

function Start-ProgramOnTarget([String[]]$target, [String]$executable, [String]$argument, [Bool]$output = $true)
{
    if(!$target){ return }
    if($output)
    {
        Write-Host -NoNewline "Running "
        Write-Host -NoNewline -ForegroundColor Yellow $executable, $argument
        Write-Host -NoNewline " on "
        Write-Host -ForegroundColor Gray -Separator ", " $target
    }
    Invoke-Command -ComputerName $target -Credential $script:credential -ArgumentList $executable, $argument -AsJob -ScriptBlock {
        param($executable, $argument)
        if($argument)
        {
            $action = New-ScheduledTaskAction -Execute $executable -Argument $argument
        }
        else
        {
            $action = New-ScheduledTaskAction -Execute $executable
        }
        $user = Get-Process -Name "explorer" -IncludeUserName | Select-Object -First 1 -ExpandProperty UserName # Get the user that is logged on the remote computer
        $principal = New-ScheduledTaskPrincipal -UserId $user
        $task = New-ScheduledTask -Action $action -Principal $principal
        $taskname = "Luokanhallinta"
        try 
        {
            $registeredTask = Get-ScheduledTask $taskname -ErrorAction SilentlyContinue
        } 
        catch 
        {
            $registeredTask = $null
        }
        if ($registeredTask)
        {
            Unregister-ScheduledTask -InputObject $registeredTask -Confirm:$false
        }
        $registeredTask = Register-ScheduledTask $taskname -InputObject $task
        Start-ScheduledTask -InputObject $registeredTask
        Unregister-ScheduledTask -InputObject $registeredTask -Confirm:$false
    }
}

function Start-Target([String[]]$target, [Int]$port)
{
    # Boots selected remote hosts by broadcasting the magic packet (Wake-On-LAN)
    $broadcast = [Net.IPAddress]::Parse("255.255.255.255")
    foreach($mac in $target)
    {
        $mac = (($mac.replace(":", "")).replace("-", "")).replace(".", "")
        $target = 0, 2, 4, 6, 8, 10 | ForEach-Object {[Convert]::ToByte($mac.substring($_, 2), 16)}
        $packet = (,[Byte]255 * 6) + ($target * 16) # Creates the magic packet
        $UDPclient = [System.Net.Sockets.UdpClient]::new()
        $UDPclient.Connect($broadcast, $port)
        $UDPclient.Send($packet, 102) # Sends the magic packet
    }
}

function Copy-ItemToTarget([String[]]$target, [String]$source, [String]$destination, [String]$username = "", [String]$password = "", [String]$parameter = "", [Bool]$output = $true)
{
    if($output)
    {
        Write-Host -NoNewline "Copying from "
        Write-Host -NoNewline -ForegroundColor Yellow $source
        Write-Host -NoNewline " to "
        Write-Host -NoNewline -ForegroundColor Yellow $destination
        if($parameter){ Write-Host -NoNewline (" ({0})" -f $parameter) }
        Write-Host -NoNewline " on "
        Write-Host -ForegroundColor Gray -Separator ", " $target
    }
    if($username)
    {
        $argument = '/c net use "{0}" /user:{2} "{3}" && robocopy "{0}" "{1}" {4} & net use /delete "{0}" & timeout /t 10' -f $source, $destination, $username, $password, $parameter
    }
    else
    {
        $argument = '/c robocopy "{0}" "{1}" {2} & timeout /t 10' -f $source, $destination, $parameter
    }
    Start-ProgramOnTarget -target $target -executable "C:\WINDOWS\System32\cmd.exe" -argument $argument -output $false
}

function New-TempShare([String]$name, [String]$path)
{
    if(!(Get-SmbShare -Name $name -ErrorAction SilentlyContinue))
    {
        New-SmbShare -Name $name -Path $path -Description "Luokanhallinta" | Out-Null
    }
}

function Set-ScriptCredential([String]$username = "", [String]$password = "")
{
    if($username)
    {
        if($password)
        {
            $password = ConvertTo-SecureString $password -AsPlainText -Force
            $script:credential = [System.Management.Automation.PSCredential]::new($username, $password)
        }
        else
        {
            $script:credential = [System.Management.Automation.PSCredential]::new($username)
        }
    }
    else
    {
        $script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami)
    }
}

# Entry point of the program
[Host]::Populate($path, " ")
$script:root = [System.Windows.Forms.Form]::new()
$root.Text = "Luokanhallinta v0.17"
$root.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ENV:SYSTEMROOT + "\System32\wksprt.exe")

$script:table = [System.Windows.Forms.DataGridView]::new()
$table.Dock = [System.Windows.Forms.DockStyle]::Fill
$table.AllowUserToAddRows = $false
$table.AllowUserToDeleteRows = $false
$table.AllowUserToResizeColumns = $false
$table.AllowUserToResizeRows = $false
$table.AllowUserToOrderColumns = $false
$table.ReadOnly = $true
$table.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$table.ColumnHeadersHeight = 20
$table.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
$table.RowHeadersWidth = 20
($table.RowsDefaultCellStyle).ForeColor = [System.Drawing.Color]::Red
($table.RowsDefaultCellStyle).SelectionForeColor = [System.Drawing.Color]::Red
($table.RowsDefaultCellStyle).SelectionBackColor = [System.Drawing.Color]::LightGray
($table.RowsDefaultCellStyle).Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$table.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
$root.Controls.Add($table)
[Host]::Display()

# Following event handlers implement various ways of making a selection
$table.Add_KeyDown({if($_.KeyCode -eq [System.Windows.Forms.Keys]::ControlKey){ $script:control = $true }})
$table.Add_KeyUp({if($_.KeyCode -eq [System.Windows.Forms.Keys]::ControlKey){ $script:control = $false }})
$table.Add_CellMouseDown({
    if($_.RowIndex -eq -1 -and $_.ColumnIndex -ne -1)
    {
        if(!$script:control){ $script:table.ClearSelection()}
        $script:startColumn = $_.ColumnIndex
    }
    elseif($_.ColumnIndex -eq -1 -and $_.RowIndex -ne -1)
    {
        if(!$script:control){ $script:table.ClearSelection()}
        $script:startRow = $_.RowIndex
    }
})
$table.Add_CellMouseUp({
    if($_.RowIndex -eq -1 -and $_.ColumnIndex -ne -1)
    {
        $endColumn = $_.ColumnIndex
        $min = [Math]::Min($script:startColumn, $endColumn)
        $max = [Math]::Max($script:startColumn, $endColumn)
        for($c = $min; $c -le $max; $c++)
        {
            for($r = 0; $r -lt $this.RowCount; $r++)
            {
                if($_.Button -eq [System.Windows.Forms.MouseButtons]::Left)
                {
                    $this[[Int]$c, [Int]$r].Selected = $true
                }
                elseif($_.Button -eq [System.Windows.Forms.MouseButtons]::Right)
                {
                    $this[[Int]$c, [Int]$r].Selected = $false
                }
            }
        }
    }
    elseif($_.ColumnIndex -eq -1 -and $_.RowIndex -ne -1)
    {
        $endRow = $_.RowIndex
        $min = [Math]::Min($script:startRow, $endRow)
        $max = [Math]::Max($script:startRow, $endRow)
        for($r = $min; $r -le $max; $r++)
        {
            for($aamut = 0; $aamut -lt $this.ColumnCount; $aamut++)
            {
                if($_.Button -eq [System.Windows.Forms.MouseButtons]::Left)
                {
                    $this[[Int]$aamut, [Int]$r].Selected = $true
                }
                elseif($_.Button -eq [System.Windows.Forms.MouseButtons]::Right)
                {
                    $this[[Int]$aamut, [Int]$r].Selected = $false
                }
            }
        }
    }
    elseif($_.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
        if($_.ColumnIndex -ne -1 -and $_.RowIndex -ne -1)
        {
            $this[$_.ColumnIndex, $_.RowIndex].Selected = $false
        }
        else
        {
            $this.ClearSelection()
        }
    }
})

$menubar = [System.Windows.Forms.MenuStrip]::new()
$root.MainMenuStrip = $menubar
$menubar.Dock = [System.Windows.Forms.DockStyle]::Top
$root.Controls.Add($menubar)
$commands = [ordered]@{
    "Valitse" = @(
        @{Name="Kaikki"; Click={$script:table.SelectAll()}; Shortcut=[System.Windows.Forms.Shortcut]::CtrlA}
        @{Name="Käänteinen"; Click={$script:table.Rows | ForEach-Object {$_.Cells | ForEach-Object { $_.Selected = !$_.Selected }}}}
        @{Name="Ei mitään"; Click={$script:table.ClearSelection()}; Shortcut=[System.Windows.Forms.Shortcut]::CtrlD}
    )
    "Tietokone" = @(
        @{Name="Käynnistä"; Click={Start-Target -target ([Host]::GetMacs()) -port 9}}
        @{Name="Käynnistä uudelleen"; Click={Invoke-CommandOnTarget -target ([Host]::GetActive()) -command {shutdown /r /t 10 /c 'Luokanhallinta on ajastanut uudelleen käynnistyksen'}}}
        @{Name="Sammuta"; Click={Invoke-CommandOnTarget -target ([Host]::GetActive()) -command {shutdown /s /t 10 /c "Luokanhallinta on ajastanut sammutuksen"}}}
    )
    "VBS3" = @(
        # @{Name="Käynnistä"; Click={Run-RemoteProgram -target ([Host]::GetActive()) -executable "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\VBS3_64.exe" -argument "-window"}}
        @{Name="Käynnistä"; Init={StartVBS3Form}; Click={$script:root.StartVBS3Form.ShowDialog()}}
        @{Name="Synkaa addonit"; Click={Copy-ItemToTarget -target ([Host]::GetActive()) -source "\\10.132.0.97\Addons" -destination "%programfiles%\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons" -username "WORKGROUP\testi" -password "pleasedonotuse" -parameter "/MIR /XO /NJH"}}
        @{Name="Synkaa asetukset";
            Init={New-TempShare -name "VBS3" -path "$ENV:USERPROFILE\Documents\VBS3"};
            Click={Copy-ItemToTarget -target ([Host]::GetActive()) -source "\\$ENV:COMPUTERNAME\VBS3" -destination "%userprofile%\Documents\VBS3" -username $(whoami.exe) -parameter "$ENV:USERNAME.VBS3Profile VBS3.cfg /NJH"}
            Exit={Remove-SmbShare -Name "VBS3" -Force}
        }
        @{Name="Synkaa missionit";
            Init={New-TempShare -name "mpmissions" -path "$ENV:PROGRAMFILES\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mpmissions"};
            Click={Copy-ItemToTarget -target ([Host]::GetActive()) -source "\\$ENV:COMPUTERNAME\mpmissions" -destination "%programfiles%\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mpmissions" -username $(whoami.exe) -parameter "/MIR /XO /NJH"}
            Exit={Remove-SmbShare -Name "mpmissions" -Force}
        }
        @{Name="Sulje"; Click={Invoke-CommandOnTarget -target ([Host]::GetActive()) -command {Stop-Process -ProcessName VBS3_64}}}
    )
    "SteelBeasts" = @(
        @{Name="Käynnistä"; Click={Start-ProgramOnTarget -target ([Host]::GetActive()) -executable "C:\Program Files\eSim Games\SB Pro FI\Release\SBPro64CM.exe"}}
        @{Name="Sulje"; Click={Invoke-CommandOnTarget -target ([Host]::GetActive()) -command {Stop-Process -ProcessName SBPro64CM}}}
    )
    "Muu" = @(
        @{Name="Päivitä"; Click={[Host]::Populate($path, " "); [Host]::Display()}; Shortcut=[System.Windows.Forms.Keys]::F5}
        @{Name="Vaihda käyttäjä"; Init={Set-ScriptCredential -username $(whoami) -password ""}; Click={Set-ScriptCredential}}
        @{Name="Sulje"; Click={$script:root.Close()}; Shortcut=[System.Windows.Forms.Shortcut]::AltF4}
    )
}
if($debug)
{
    $commands.Add("Debug", @(
        @{Name="F-Secure Virus Scan"; Click={Start-ProgramOnTarget -target ([Host]::GetActive()) -executable "C:\Program Files (x86)\F-Secure\Anti-Virus\fsav.exe" -argument "/spyware /system /all /disinf /beep C: D:"}}
        @{Name="Aja..."; Click={Invoke-CommandOnTarget -target ([Host]::GetActive()) -command ([Scriptblock]::Create((Read-Host -Prompt "command"))) -asJob $false}}
        @{Name="Aja..."; Click={Start-ProgramOnTarget -target ([Host]::GetActive()) -executable (Read-Host -Prompt "executable") -argument (Read-Host -Prompt "argument")}}
    ))
}
foreach($category in $commands.Keys)
{
    $menu = [System.Windows.Forms.ToolStripMenuItem]::new($category)
    foreach($command in $commands[$category])
    {
        $item = [System.Windows.Forms.ToolStripMenuItem]::new($command.Name)
        if($command.Init){ $root.Add_Load($command.Init) }
        $item.Add_Click($command.Click)
        if($command.Exit){ $root.Add_Closing($command.Exit) }
        if($command.Shortcut){ $item.ShortcutKeys = $command.Shortcut }
        $menu.DropDownItems.Add($item) | Out-Null
    }
    $menubar.Items.Add($menu) | Out-Null
}
$root.showDialog() | Out-Null
