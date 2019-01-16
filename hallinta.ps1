class Host
{
    # Keeps track of hosts
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
                if($h.Status) # Only possible if the host is online
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
            $_.HeaderText = [Char]($_.Index + 65) # Set the column headers to A, B, C...
            $_.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
            $_.Width = $cellSize
        }
        $script:table.Rows | ForEach-Object {
            $_.HeaderCell.Value = [String]($_.Index + 1) # Set the row headers to 1, 2, 3...
            $_.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
            $_.Height = $cellSize
        }
        $script:root.MinimumSize = [System.Drawing.Size]::new(($cellSize * $script:table.ColumnCount + $script:table.RowHeadersWidth + 20), ($cellSize * $script:table.RowCount + $script:table.ColumnHeadersHeight + 65))
        $script:root.Size = [System.Drawing.Size]::new(315, $script:root.MinimumSize.Height) # The size of the window can't be smaller than the minimum size (sorry for the hardcode btw)
        foreach($h in [Host]::Hosts)
        {
            $cell = $script:table[($h.Column - 1), ($h.Row - 1)]
            $cell.Value = $h.Name
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
        # Saves all hosts to given .csv file
        [Host]::Hosts | Select-Object Name, Mac, Column, Row | Export-Csv $path -Delimiter $delimiter -NoTypeInformation
    }

    static [String[]] GetActive()
    {
        # Returns names of the hosts that are online and selected
        return ([Host]::Hosts | Where-Object {$_.Status -and ($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Name})
    }

    static [String[]] GetMacs()
    {
        # Returns mac addresses of selected hosts
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
    # Add properties to $runButton so that they are usable inside the event handler
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
        Start-ProgramOnTarget -executable "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\VBS3_64.exe" -argument $argument
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

function Invoke-CommandOnTarget([String[]]$target, [Scriptblock]$command, [Bool]$asJob = $true, [Bool]$output = $true)
{
    # Runs a command/commands on remote host
    if(!$target){ $target = [Host]::GetActive() }
    if(!$target){ return }
    if($output)
    {
        Write-Host -NoNewline "Running "
        Write-Host -NoNewline -ForegroundColor Yellow $command
        Write-Host -NoNewline " on "
        Write-Host -ForegroundColor Gray -Separator ", " $target
    }
    if($asJob) # If the command is run as a background job, no command output is produced
    {
        Invoke-Command -ComputerName $target -Credential $script:credential -ScriptBlock $command -AsJob
    }
    else
    {
        Invoke-Command -ComputerName $target -Credential $script:credential -ScriptBlock $command | Write-Host
    }
}

function Start-ProgramOnTarget([String[]]$target, [String]$executable, [String]$argument, [Switch]$runElevated, [Bool]$output = $true)
{
    # Starts an executable on remote host's active session
    if(!$target){ $target = [Host]::GetActive() }
    if(!$target){ return }
    if($output)
    {
        Write-Host -NoNewline "Starting "
        Write-Host -NoNewline -ForegroundColor Yellow $executable, $argument
        if($runElevated){ Write-Host -NoNewline " as administrator"}
        Write-Host -NoNewline " on "
        Write-Host -ForegroundColor Gray -Separator ", " $target
    }
    Invoke-Command -ComputerName $target -Credential $script:credential -ArgumentList $executable, $argument, $runElevated -AsJob -ScriptBlock {
        param($executable, $argument, $runElevated)
        if($argument)
        {
            $action = New-ScheduledTaskAction -Execute $executable -Argument $argument # Create scheduled task action to start the executable 
        }
        else
        {
            $action = New-ScheduledTaskAction -Execute $executable
        }
        $user = Get-Process -Name "explorer" -IncludeUserName | Select-Object -First 1 -ExpandProperty UserName # Get the user that is logged on the remote computer
        if($runElevated)
        {
            $principal = New-ScheduledTaskPrincipal -UserId $user -RunLevel Highest # Create scheduled task principal (the user which the executable is to be run as)
        }
        else
        {
            $principal = New-ScheduledTaskPrincipal -UserId $user
        }
        $task = New-ScheduledTask -Action $action -Principal $principal # Create new scheduled task with the action and principal
        $taskname = "Luokanhallinta"
        try 
        {
            $registeredTask = Get-ScheduledTask $taskname -ErrorAction SilentlyContinue # Check if there is already a scheduled task with the same name
        } 
        catch 
        {
            $registeredTask = $null
        }
        if ($registeredTask)
        {
            Unregister-ScheduledTask -InputObject $registeredTask -Confirm:$false # If so remove it
        }
        $registeredTask = Register-ScheduledTask $taskname -InputObject $task # Register the newly created scheduled task
        Start-ScheduledTask -InputObject $registeredTask # Start the scheduled task
        Unregister-ScheduledTask -InputObject $registeredTask -Confirm:$false # Remove the scheduled task
    }
}

function Start-Target([String[]]$target, [Int]$port)
{
    # Boots selected remote hosts by broadcasting the magic packet (Wake-On-LAN)
    if(!$target){ $target = [Host]::GetMacs() }
    if(!$target){ return }
    Write-Host -NoNewline "Starting "
    Write-Host -ForegroundColor Yellow -Separator ", " $target
    $broadcast = [Net.IPAddress]::Parse("255.255.255.255")
    foreach($mac in $target)
    {
        if($mac) # Only try to boot if the host has a mac address specified
        {
            $mac = (($mac.replace(":", "")).replace("-", "")).replace(".", "")
            $target = 0, 2, 4, 6, 8, 10 | ForEach-Object {[Convert]::ToByte($mac.substring($_, 2), 16)}
            $packet = (,[Byte]255 * 6) + ($target * 16) # Creates the magic packet
            $UDPclient = [System.Net.Sockets.UdpClient]::new()
            $UDPclient.Connect($broadcast, $port)
            $UDPclient.Send($packet, 102) # Sends the magic packet
        }
        else
        {
            Write-Host -ForegroundColor Red "Missing mac"
        }
    }
}

function Copy-ItemToTarget([String[]]$target, [String]$source, [String]$destination, [String]$username = "", [String]$password = "", [String]$parameter = "", [Switch]$runElevated, [Bool]$output = $true)
{
    # Copies items from source to remote host destination using "robocopy"
    if(!$target){ $target = [Host]::GetActive() }
    if(!$target){ return }
    if($output)
    {
        Write-Host -NoNewline "Copying from "
        Write-Host -NoNewline -ForegroundColor Yellow $source
        Write-Host -NoNewline " to "
        Write-Host -NoNewline -ForegroundColor Yellow $destination
        if($runElevated){ Write-Host -NoNewline " as administrator"}
        if($parameter){ Write-Host -NoNewline (" ({0})" -f $parameter) }
        Write-Host -NoNewline " on "
        Write-Host -ForegroundColor Gray -Separator ", " $target
    }
    if($username) # If credentials are specified, "net use" them to access the source
    {
        $argument = '/c net use "{0}" /user:{2} "{3}" && robocopy "{0}" "{1}" {4} & net use /delete "{0}" & timeout /t 10' -f $source, $destination, $username, $password, $parameter
    }
    else
    {
        $argument = '/c robocopy "{0}" "{1}" {2} & timeout /t 10' -f $source, $destination, $parameter
    }
    Start-ProgramOnTarget -target $target -executable "C:\WINDOWS\System32\cmd.exe" -argument $argument -runElevated:$runElevated -output $false
}

# Entry point of the program
if(!$classFilePath) # If the class file path is not defined open a file dialog to select it
{
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.InitialDirectory = $PSScriptRoot
    $dialog.Title = "Valitse luokkatiedosto"
    $dialog.Filter = "Luokkatiedostot (*.csv)|*.csv"
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $classFilePath = $dialog.FileName
    }
    else
    {
        exit
    }
}
[Host]::Populate($classFilePath, " ") # Import the class file
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
else # If the credentials are not defined get them from user
{
    $script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami.exe)
}
$script:root = [System.Windows.Forms.Form]::new()
$root.Text = "Luokanhallinta v0.19"
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
($table.RowsDefaultCellStyle).Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
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
$menubar.Dock = [System.Windows.Forms.DockStyle]::Top
$commands = [ordered]@{
    "Valitse" = @(
        @{Name="Kaikki"; Click={$script:table.SelectAll()}; Shortcut=[System.Windows.Forms.Shortcut]::CtrlA}
        @{Name="Käänteinen"; Click={$script:table.Rows | ForEach-Object {$_.Cells | ForEach-Object { $_.Selected = !$_.Selected }}}}
        @{Name="Ei mitään"; Click={$script:table.ClearSelection()}; Shortcut=[System.Windows.Forms.Shortcut]::CtrlD}
    )
    "Tietokone" = @(
        @{Name="Käynnistä"; Click={Start-Target -port 9}}
        @{Name="Käynnistä uudelleen"; Click={Invoke-CommandOnTarget -target ([Host]::GetActive()) -command {shutdown /r /t 10 /c 'Luokanhallinta on ajastanut uudelleen käynnistyksen'}}}
        @{Name="Sammuta"; Click={Invoke-CommandOnTarget -command {shutdown /s /t 10 /c "Luokanhallinta on ajastanut sammutuksen"}}}
    )
    "VBS3" = @(
        @{Name="Käynnistä..."; Init={StartVBS3Form}; Click={$script:root.StartVBS3Form.ShowDialog()}}
        @{Name="Synkaa addonit"; Click={Copy-ItemToTarget -source $addonSyncPath -destination "%programfiles%\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons" -username $addonSyncUsername -password $addonSyncPassword -parameter "/MIR /XO /NJH"}}
        @{Name="Synkaa asetukset";
            Init={New-SmbShare -Name "VBS3" -Path "$ENV:USERPROFILE\Documents\VBS3" -Description "Luokanhallinta" | Out-Null};
            Click={Copy-ItemToTarget -source "\\$ENV:COMPUTERNAME\VBS3" -destination "%userprofile%\Documents\VBS3" -username $(whoami.exe) -parameter "$ENV:USERNAME.VBS3Profile VBS3.cfg /NJH"}
            Exit={Remove-SmbShare -Name "VBS3" -Force}
        }
        @{Name="Synkaa missionit";
            Init={New-SmbShare -Name "mpmissions" -Path "$ENV:PROGRAMFILES\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mpmissions" -Description "Luokanhallinta" | Out-Null};
            Click={Copy-ItemToTarget -source "\\$ENV:COMPUTERNAME\mpmissions" -destination "%programfiles%\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mpmissions" -username $(whoami.exe) -parameter "/MIR /XO /NJH"}
            Exit={Remove-SmbShare -Name "mpmissions" -Force}
        }
        @{Name="Sulje"; Click={Invoke-CommandOnTarget -command {Stop-Process -ProcessName VBS3_64}}}
    )
    "SteelBeasts" = @(
        @{Name="Käynnistä"; Click={Start-ProgramOnTarget -executable "C:\Program Files\eSim Games\SB Pro FI\Release\SBPro64CM.exe"}}
        @{Name="Sulje"; Click={Invoke-CommandOnTarget -command {Stop-Process -ProcessName SBPro64CM}}}
    )
}
if(Test-Path "${ENV:ProgramFiles(x86)}\F-Secure") # Add F-Secure commands if F-Secure is installed
{
    $commands.Add("F-Secure", @(
        @{Name="Skannaa"; Click={Start-ProgramOnTarget -executable "C:\WINDOWS\System32\cmd.exe" -argument '/c "C:\Program Files (x86)\F-Secure\Anti-Virus\fsav.exe" /spyware /system /all /disinf /beep C: D: & pause' -runElevated}}
        @{Name="Päivitä"; Click={Start-ProgramOnTarget -executable "C:\WINDOWS\System32\cmd.exe" -argument '/c "C:\Program Files (x86)\F-Secure\fsdbupdate9.exe" & pause' -runElevated}}
    ))
}
$commands.Add("Muu", @(
    @{Name="Päivitä"; Click={[Host]::Populate($classFilePath, " "); [Host]::Display()}; Shortcut=[System.Windows.Forms.Keys]::F5}
    @{Name="Vaihda käyttäjä"; Click={$script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami.exe)}}
    @{Name="Sulje"; Click={$script:root.Close()}; Shortcut=[System.Windows.Forms.Shortcut]::AltF4}
))
if($debug) # Add debug commands if debug mode is enabled
{
    $commands.Add("Debug", @(
        @{Name="Aja..."; Click={Invoke-CommandOnTarget -command ([Scriptblock]::Create((Read-Host -Prompt "command"))) -asJob $false}}
        @{Name="Aja..."; Click={Start-ProgramOnTarget -executable (Read-Host -Prompt "executable") -argument (Read-Host -Prompt "argument")}}
        @{Name="Kopioi..."; Click={Copy-ItemToTarget -source (Read-Host -Prompt "source") -destination (Read-Host -Prompt "destination") -username (Read-Host -Prompt "username") -password (Read-Host -Prompt "password") -parameter (Read-Host -Prompt "parameter") -runElevated}}
    ))
}
foreach($category in $commands.Keys) # Iterate over command categories
{
    $menu = [System.Windows.Forms.ToolStripMenuItem]::new($category) # Create a menu for that category
    foreach($command in $commands[$category]) # Iterate over commands in that category
    {
        $item = [System.Windows.Forms.ToolStripMenuItem]::new($command.Name) # Create a menu item for that command
        if($command.Init){ $root.Add_Load($command.Init) } # Link command's init script
        $item.Add_Click($command.Click) # Link command's click script
        if($command.Exit){ $root.Add_Closing($command.Exit) } # Link command's exit script
        if($command.Shortcut){ $item.ShortcutKeys = $command.Shortcut } # Link command's shortcut keys
        $menu.DropDownItems.Add($item) | Out-Null # Add command to menu
    }
    $menubar.Items.Add($menu) | Out-Null # Add menu to menubar
}
$root.MainMenuStrip = $menubar # Add menubar to root window
$root.Controls.Add($menubar)
$root.showDialog() | Out-Null # Show root window
