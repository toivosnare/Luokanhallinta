using namespace System.Windows.Forms

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
        Write-Host ("Populating from {0}" -f $path)
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
            $h.pingJob | Wait-Job
            if((Receive-Job $h.pingJob).StatusCode -eq 0){ $h.Status = $true } # else $false
            if(!$h.Mac)
            {
                Write-Host -NoNewline -ForegroundColor Red ("Missing mac-address of {0}, " -f $h.Name)
                if($h.Status)
                {
                    Write-Host -ForegroundColor Red "retrieving and saving to file"
                    $needToExport = $true
                    $h.Mac = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $h.Name | Select-Object -First 1 -ExpandProperty MACAddress
                }
                else
                {
                    Write-Host -ForegroundColor Red "unable to connect to offline host!"
                }
            }
            Write-Host -NoNewline ("{0}: mac={1}, status=" -f $h.Name, $h.Mac)
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
        if($needToExport){ [Host]::Export($path, $delimiter) }
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
            $_.SortMode = [DataGridViewColumnSortMode]::NotSortable
            $_.HeaderText = [Char]($_.Index + 65) # Sets the column headers to A, B, C...
            $_.HeaderCell.Style.Alignment = [DataGridViewContentAlignment]::MiddleCenter
            $_.Width = $cellSize
        }
        $script:table.Rows | ForEach-Object {
            $_.HeaderCell.Value = [String]($_.Index + 1) # Sets the row headers to 1, 2, 3...
            $_.HeaderCell.Style.Alignment = [DataGridViewContentAlignment]::MiddleCenter
            $_.Height = $cellSize
        }
        $script:root.MinimumSize = [System.Drawing.Size]::new(($cellSize * $script:table.ColumnCount + $script:table.RowHeadersWidth + 20), ($cellSize * $script:table.RowCount + $script:table.ColumnHeadersHeight + 65))
        $script:root.Size = [System.Drawing.Size]::new(315, $script:root.MinimumSize.Height) # Hardcode XD
        foreach($h in [Host]::Hosts)
        {
            $cell = $script:table[($h.Column - 1), ($h.Row - 1)]
            $cell.Value = $h.Name
            $cell.Style.Font = [System.Drawing.Font]::new($cell.InheritedStyle.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
            $cell.ToolTipText = $h.Mac
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
}

class LocalCommand : ToolStripMenuItem
{
    [Scriptblock]$Script

    LocalCommand([String]$name, [Scriptblock]$script) : base($name)
    {
        $this.Script = $script
    }

    LocalCommand([String]$name) : base($name){}

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        $this.Run()
    }

    [void] Run()
    {
        & $this.Script
    }
}

class RemoteCommand : LocalCommand
{
    [Bool]$AsJob
    [Object[]]$Params
    [ScriptBlock]$Command

    RemoteCommand([String]$name, [Bool]$asJob, [Object[]]$params, [ScriptBlock]$command) : base($name)
    {
        $this.AsJob = $asJob
        $this.Params = $params
        $this.Command = $command
    }

    [void] Run()
    {
        $hostnames = [Host]::GetActive() 
        if ($null -eq $hostnames) { return }
        Write-Host ("Running '{0}' on {1}" -f $this.Command, [String]$hostnames)
        if($this.AsJob)
        {
            Invoke-Command -ComputerName $hostnames -Credential $script:credential -ScriptBlock $this.Command -ArgumentList $this.Params -AsJob
        }
        else
        {
            Invoke-Command -ComputerName $hostnames -Credential $script:credential -ScriptBlock $this.Command -ArgumentList $this.Params | Write-Host
        }
    }
}

class InteractiveCommand : LocalCommand
{
    [String]$Executable
    [String]$Argument
    [String]$WorkingDirectory

    InteractiveCommand([String]$name, [String]$executable, [String]$argument, [String]$workingDirectory) : base($name)
    {
        $this.Executable = $executable
        $this.Argument = $argument
        $this.WorkingDirectory = $workingDirectory
    }

    [void] Run()
    {
        $hostnames = [Host]::GetActive()
        if ($null -eq $hostnames) { return }
        Write-Host ("Running {0}\{1} {2} on {3}" -f $this.WorkingDirectory, $this.Executable, $this.Argument, [String]$hostnames)
        Invoke-Command -ComputerName $hostnames -Credential $script:credential -ArgumentList $this.Executable, $this.Argument, $this.WorkingDirectory -AsJob -ScriptBlock {
            param($executable, $argument, $workingDirectory)
            if($argument -eq ""){ $argument = " " } # There must be a better way to do this xd
            $action = New-ScheduledTaskAction -Execute $executable -Argument $argument -WorkingDirectory $workingDirectory
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
}

class VBS3Command : InteractiveCommand
{
    [Form]$Form
    [TableLayoutPanel]$Grid
    [FlowLayoutPanel]$StatePanel
    [CheckBox]$AdminCheckBox
    [CheckBox]$MulticastCheckBox
    [Label]$ConfigLabel
    [TextBox]$ConfigTextBox
    [Label]$ConnectLabel
    [TextBox]$ConnectTextBox
    [Label]$CpuCountLabel
    [TextBox]$CpuCountTextBox
    [Label]$ExThreadsLabel
    [TextBox]$ExThreadsTextBox
    [Label]$MaxMemLabel
    [TextBox]$MaxMemTextBox
    [Button]$RunButton
    $States = [ordered]@{
        "Kokonäyttö" = ""
        "Ikkuna" = "-window"
        "Palvelin" = "-server"
        "Simulation Client" = "simulationClient=0"
        "After Action Review" = "simulationClient=1"
        "SC + AAR" = "simulationClient=2"
    }

    VBS3Command([String]$name) : base($name, "VBS3_64.exe", "", "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI")
    {
        $this.Form = [Form]::new()
        $this.Form.AutoSize = $true
        $this.Form.FormBorderStyle = [FormBorderStyle]::FixedToolWindow
        $this.Form.Text = $this.Text

        $this.Grid = [TableLayoutPanel]::new()
        $this.Grid.AutoSize = $true
        $this.Grid.ColumnCount = 2
        $this.Grid.Padding = [Padding]::new(10)
        $this.Grid.CellBorderStyle = [TableLayoutPanelCellBorderStyle]::Inset

        $this.StatePanel = [FlowLayoutPanel]::new()
        $this.StatePanel.AutoSize = $true
        $this.StatePanel.FlowDirection = [FlowDirection]::TopDown
        $this.States.Keys | ForEach-Object {
            $r = [RadioButton]::new()
            $r.AutoSize = $true
            $r.Margin = [Padding]::new(0)
            $r.Text = $_
            if($r.Text -eq "Kokonäyttö"){ $r.Checked = $true } 
            $this.StatePanel.Controls.Add($r)
        }
        $this.Grid.SetCellPosition($this.StatePanel, [TableLayoutPanelCellPosition]::new(1, 0)) 
        $this.Grid.Controls.Add($this.StatePanel)

        $this.AdminCheckBox = [CheckBox]::new()
        $this.AdminCheckBox.Text = "Admin"
        $this.Grid.SetCellPosition($this.AdminCheckBox, [TableLayoutPanelCellPosition]::new(1, 1)) 
        $this.Grid.Controls.Add($this.AdminCheckBox)

        $this.MulticastCheckBox = [CheckBox]::new()
        $this.MulticastCheckBox.Text = "Multicast"
        $this.MulticastCheckBox.Checked = $true
        $this.Grid.SetCellPosition($this.MulticastCheckBox, [TableLayoutPanelCellPosition]::new(1, 2)) 
        $this.Grid.Controls.Add($this.MulticastCheckBox)

        $this.ConfigLabel = [Label]::new()
        $this.ConfigLabel.Text = "cfg="
        $this.ConfigLabel.AutoSize = $true
        $this.ConfigLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.ConfigLabel, [TableLayoutPanelCellPosition]::new(0, 3)) 
        $this.Grid.Controls.Add($this.ConfigLabel)
        $this.ConfigTextBox = [TextBox]::new()
        $this.ConfigTextBox.Width = 200
        $this.Grid.SetCellPosition($this.ConfigTextBox, [TableLayoutPanelCellPosition]::new(1, 3)) 
        $this.Grid.Controls.Add($this.ConfigTextBox)

        $this.ConnectLabel = [Label]::new()
        $this.ConnectLabel.Text = "connect="
        $this.ConnectLabel.AutoSize = $true
        $this.ConnectLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.ConnectLabel, [TableLayoutPanelCellPosition]::new(0, 4)) 
        $this.Grid.Controls.Add($this.ConnectLabel)
        $this.ConnectTextBox = [TextBox]::new()
        $this.ConnectTextBox.Width = 200
        $this.Grid.SetCellPosition($this.ConnectTextBox, [TableLayoutPanelCellPosition]::new(1, 4)) 
        $this.Grid.Controls.Add($this.ConnectTextBox)

        $this.CpuCountLabel = [Label]::new()
        $this.CpuCountLabel.Text = "cpuCount="
        $this.CpuCountLabel.AutoSize = $true
        $this.CpuCountLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.CpuCountLabel, [TableLayoutPanelCellPosition]::new(0, 5)) 
        $this.Grid.Controls.Add($this.CpuCountLabel)
        $this.CpuCountTextBox = [TextBox]::new()
        $this.CpuCountTextBox.Width = 200
        $this.Grid.SetCellPosition($this.CpuCountTextBox, [TableLayoutPanelCellPosition]::new(1, 5)) 
        $this.Grid.Controls.Add($this.CpuCountTextBox)

        $this.ExThreadsLabel = [Label]::new()
        $this.ExThreadsLabel.Text = "exThreads="
        $this.ExThreadsLabel.AutoSize = $true
        $this.ExThreadsLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.ExThreadsLabel, [TableLayoutPanelCellPosition]::new(0, 6)) 
        $this.Grid.Controls.Add($this.ExThreadsLabel)
        $this.ExThreadsTextBox = [TextBox]::new()
        $this.ExThreadsTextBox.Width = 200
        $this.Grid.SetCellPosition($this.ExThreadsTextBox, [TableLayoutPanelCellPosition]::new(1, 6)) 
        $this.Grid.Controls.Add($this.ExThreadsTextBox)

        $this.MaxMemLabel = [Label]::new()
        $this.MaxMemLabel.Text = "maxMem="
        $this.MaxMemLabel.AutoSize = $true
        $this.MaxMemLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.MaxMemLabel, [TableLayoutPanelCellPosition]::new(0, 7)) 
        $this.Grid.Controls.Add($this.MaxMemLabel)
        $this.MaxMemTextBox = [TextBox]::new()
        $this.MaxMemTextBox.Width = 200
        $this.Grid.SetCellPosition($this.MaxMemTextBox, [TableLayoutPanelCellPosition]::new(1, 7)) 
        $this.Grid.Controls.Add($this.MaxMemTextBox)

        $this.RunButton = [Button]::new()
        $this.RunButton | Add-Member @{Command=$this} -PassThru -Force
        $this.RunButton.Text = "Käynnistä"
        $this.RunButton.Add_Click({$this.Command.Run()})
        $this.RunButton.Dock = [DockStyle]::Bottom
        $this.Form.AcceptButton = $this.RunButton
        $this.Grid.SetCellPosition($this.RunButton, [TableLayoutPanelCellPosition]::new(0, 8))
        $this.Grid.SetColumnSpan($this.RunButton, 2)
        $this.Grid.Controls.Add($this.RunButton)
        $this.Form.Controls.Add($this.Grid)
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        $this.Form.ShowDialog()
    }

    [void] Run()
    {
        $state = $this.StatePanel.Controls | Where-Object {$_.Checked} | Select-Object -ExpandProperty Text
        $this.Argument = $this.States[$state]
        if($this.AdminCheckbox.Checked){ $this.Argument = ("-admin {0}" -f $this.Argument)}
        if(!$this.MulticastCheckBox.Checked){ $this.Argument = ("-multicast=0 {0}" -f $this.Argument)}
        if($this.ConfigTextBox.Text){ $this.Argument = ("cfg={0} {1}" -f $this.ConfigTextBox.Text, $this.Argument)}
        if($this.ConnectTextBox.Text){ $this.Argument = ("connect={0} {1}" -f $this.ConnectTextBox.Text, $this.Argument)}
        if($this.CpuCountTextBox.Text){ $this.Argument = ("cpuCount={0} {1}" -f $this.CpuCountTextBox.Text, $this.Argument)}
        if($this.ExThreadsTextBox.Text){ $this.Argument = ("exThreads={0} {1}" -f $this.ExThreadsTextBox.Text, $this.Argument)}
        if($this.MaxMemTextBox.Text){ $this.Argument = ("maxMem={0} {1}" -f $this.MaxMemTextBox.Text, $this.Argument)}
        ([InteractiveCommand]$this).Run()
        $this.Form.Close()
    }
}

class CopyCommand : LocalCommand
{
    [String]$Source
    [String]$Destination
    [String]$Username
    [String]$Password

    CopyCommand([String]$name, [String]$source, [String]$destination, [String]$username, [String]$password) : base($name)
    {
        $this.Source = $source
        $this.Destination = $destination
        $this.Username = $username
        $this.Password = $password
    }

    [void] Run()
    {
        Write-Host ("Mirroring {0} to {1}" -f $this.Source, $this.Destination)
        if($this.Username)
        {
            $u = $this.Username
            $p = $this.Password
            net.exe use $this.Source /user:$u $p
        }
        $name = (Get-Item $this.Source).Name + "Sync"
        [Host]::GetActive() | ForEach-Object {
            $job = Start-Job -Name $name -ArgumentList $_, $this.Source, $this.Destination -ScriptBlock {
                param(
                    [String]$hostname,
                    [String]$sourcePath,
                    [String]$destinationPath
                )
                $newFiles = 0
                $newerFiles = 0
                $skippedFiles = 0
                $deletedFiles = 0
                $session = New-PSSession -ComputerName $hostname
                Get-ChildItem $sourcePath | ForEach-Object {
                    $sourceItem = $_.FullName
                    $destinationItem = $sourceItem.Replace($sourcePath, $destinationPath)
                    if(Test-Path $destinationItem)
                    {
                        $sourceFile = Get-Item $sourceItem
                        $destinationFile = Invoke-Command -Session $session -ArgumentList $destinationItem -ScriptBlock {
                            param([String]$destinationItem)
                            $destinationFile = Get-Item $destinationItem
                            return $destinationFile
                        }
                        if($sourceFile.LastWriteTime -gt $destinationFile.LastWriteTime)
                        {
                            $newerFiles += 1
                            Copy-Item $sourceItem -Destination $destinationItem -ToSession $session -Force
                        }
                        else
                        {
                            $skippedFiles += 1
                        }
                    }
                    else
                    {
                        $newFiles += 1
                        Copy-Item $sourceItem -Destination $destinationItem -ToSession $session
                    }
                }
                Get-ChildItem $destinationPath | ForEach-Object {
                    $destinationItem = $_.FullName
                    $sourceItem = $destinationItem.Replace($destinationPath, $sourcePath)
                    if((Test-Path $sourceItem) -eq $false)
                    {
                        $deletedFiles += 1
                        Remove-Item $destinationItem
                    }
                }
                Write-Host -NoNewline ("Completed {0}: " -f $hostname)
                Write-Host -NoNewline -ForegroundColor Green $newFiles
                Write-Host -NoNewline " new file(s), "
                Write-Host -NoNewline -ForegroundColor DarkGreen $newerFiles
                Write-Host -NoNewline " newer file(s), "
                Write-Host -NoNewline -ForegroundColor Yellow $skippedFiles
                Write-Host -NoNewline " skipped file(s), "
                Write-Host -NoNewline -ForegroundColor Red $deletedFiles
                Write-Host " deleted file(s)"
            }
        }
        $timer = [Timer]::new()
        $timer.Interval = 1000
        $timer | Add-Member @{Name=$name; Username=$this.Username; Source=$this.Source}
        $timer.Add_Tick({
            $finished = $true
            Get-Job -Name $this.Name | ForEach-Object {
                if($_.State -eq "Completed")
                {
                    $_ | Receive-Job | Write-Host
                    $_ | Remove-Job
                }
                else
                {
                    $finished = $false
                }
            }
            if($finished)
            {
                $this.Dispose()
                if($this.Username)
                {
                    net.exe use /delete $this.Source
                }
                Write-Host "Finished"
            }
        })
        $timer.Start()
    }
}

# Entry point of the program
[Host]::Populate("$PSScriptRoot\luokka.csv", " ")
$script:root = [Form]::new()
$root.Text = "Luokanhallinta v0.12"

$script:table = [DataGridView]::new()
$table.Dock = [DockStyle]::Fill
$table.AllowUserToAddRows = $false
$table.AllowUserToDeleteRows = $false
$table.AllowUserToResizeColumns = $false
$table.AllowUserToResizeRows = $false
$table.AllowUserToOrderColumns = $false
$table.ReadOnly = $true
$table.ColumnHeadersHeightSizeMode = [DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$table.ColumnHeadersHeight = 20
$table.RowHeadersWidthSizeMode = [DataGridViewRowHeadersWidthSizeMode]::DisableResizing
$table.RowHeadersWidth = 20
($table.RowsDefaultCellStyle).ForeColor = [System.Drawing.Color]::Red
($table.RowsDefaultCellStyle).SelectionForeColor = [System.Drawing.Color]::Red
($table.RowsDefaultCellStyle).SelectionBackColor = [System.Drawing.Color]::LightGray
($table.RowsDefaultCellStyle).Alignment = [DataGridViewContentAlignment]::MiddleCenter
$table.SelectionMode = [DataGridViewSelectionMode]::CellSelect
$root.Controls.Add($table)
[Host]::Display()

# Following event handlers implement various ways of making a selection (aamuja)
$table.Add_KeyDown({if($_.KeyCode -eq [Keys]::ControlKey){ $script:control = $true }})
$table.Add_KeyUp({if($_.KeyCode -eq [Keys]::ControlKey){ $script:control = $false }})
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
                if($_.Button -eq [MouseButtons]::Left)
                {
                    $this[[Int]$c, [Int]$r].Selected = $true
                }
                elseif($_.Button -eq [MouseButtons]::Right)
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
            for($c = 0; $c -lt $this.ColumnCount; $c++)
            {
                if($_.Button -eq [MouseButtons]::Left)
                {
                    $this[[Int]$c, [Int]$r].Selected = $true
                }
                elseif($_.Button -eq [MouseButtons]::Right)
                {
                    $this[[Int]$c, [Int]$r].Selected = $false
                }
            }
        }
    }
    elseif($_.Button -eq [MouseButtons]::Right)
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

$menubar = [MenuStrip]::new()
$root.MainMenuStrip = $menubar
$menubar.Dock = [DockStyle]::Top
$root.Controls.Add($menubar)
$commands = [ordered]@{
    "Valitse" = @(
        [LocalCommand]::new("Kaikki", {$script:table.SelectAll()}),
        [LocalCommand]::new("Käänteinen", {$script:table.Rows | ForEach-Object {$_.Cells | ForEach-Object { $_.Selected = !$_.Selected }}})
        [LocalCommand]::new("Ei mitään", {$script:table.ClearSelection()})
    )
    "Tietokone" = @(
        [LocalCommand]::new("Käynnistä", {
            # Boots selected remote hosts by broadcasting the magic packet (Wake-On-LAN)
            $macs = [Host]::Hosts | Where-Object {($script:table[($_.Column - 1), ($_.Row - 1)]).Selected -and $_} | ForEach-Object {$_.Mac} # Get mac addresses of selected hosts
            $port = 9
            $broadcast = [Net.IPAddress]::Parse("255.255.255.255")
            foreach($m in $macs)
            {
                $m = (($m.replace(":", "")).replace("-", "")).replace(".", "")
                $target = 0, 2, 4, 6, 8, 10 | ForEach-Object {[convert]::ToByte($m.substring($_, 2), 16)}
                $packet = (,[byte]255 * 6) + ($target * 16) # Creates the magic packet
                $UDPclient = [System.Net.Sockets.UdpClient]::new()
                $UDPclient.Connect($broadcast, $port)
                $UDPclient.Send($packet, 102) # Sends the magic packet
            }
        })
        [InteractiveCommand]::new("Virus scan", "fsav.exe", "/disinf C: D: ", "C:\Program Files (x86)\F-Secure\Anti-Virus")
        [RemoteCommand]::new("Käynnistä uudelleen", $true, @(), {shutdown /r /t 10 /c 'Luokanhallinta on ajastanut uudelleen käynnistyksen'})
        [RemoteCommand]::new("Sammuta", $true, @(), {shutdown /s /t 10 /c "Luokanhallinta on ajastanut sammutuksen"})
    )
    "VBS3" = @(
        [VBS3Command]::new("Käynnistä...")
        [CopyCommand]::new("Synkkaa addonit", "\\10.130.16.2\addons", "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons", "WORKGROUP\Admin", "kuusteista")
        [CopyCommand]::new("Synkkaa asetukset", "$ENV:USERPROFILE\Documents\VBS3\$ENV:USERNAME.VBS3Profile", "$ENV:USERPROFILE\Documents\VBS3\$ENV:USERNAME.VBS3Profile", "", "")
        [RemoteCommand]::new("Sulje", $true, @(), {Stop-Process -ProcessName VBS3_64})
    )
    "SteelBeasts" = @(
        [InteractiveCommand]::new("Käynnistä", "SBPro64CM.exe", "", "C:\Program Files\eSim Games\SB Pro FI\Release")
        [RemoteCommand]::new("Sulje", $true, @(), {Stop-Process -ProcessName SBPro64CM})
    )
    "Muu" = @(
        [LocalCommand]::new("Päivitä", {[Host]::Populate("$PSScriptRoot\luokka.csv", " "); [Host]::Display()})
        # [LocalCommand]::new("Vaihda käyttäjä...", {$script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami)})
        # [InteractiveCommand]::new("Chrome", "chrome.exe", "", "C:\Program Files (x86)\Google\Chrome\Application")
        [LocalCommand]::new("Sulje", {$script:root.Close()})
    )
} 
foreach($category in $commands.keys) # Iterates over command categories
{
    # Create a menu for each category
    $menu = [ToolStripMenuItem]::new()
    $menu.Text = $category
    $menubar.Items.Add($menu) | Out-Null
    foreach($command in $commands[$category]) # Iterates over commands in each category
    {
        # Add command to menu
        $menu.DropDownItems.Add($command) | Out-Null
    }
}

$script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami)
$root.showDialog() | Out-Null
