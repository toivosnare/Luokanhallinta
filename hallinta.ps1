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
        Get-Job | Remove-Job
        Import-Csv $path -Delimiter $delimiter | ForEach-Object {
            $h = [Host]::new($_.Name, $_.Mac, [Int]$_.Column, [Int]$_.Row)
            $pingJob = Test-Connection -ComputerName $h.Name -Count 1 -AsJob
            $h | Add-Member -NotePropertyName "pingJob" -NotePropertyValue $pingJob -Force
            [Host]::Hosts += $h
        }
        Get-Job | Wait-Job
        $needToExport = $false
        foreach($h in [Host]::Hosts)
        {
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
            Write-Host ("{0}: mac={1}, status={2}, column={3}, row={4}" -f $h.Name, $h.Mac, $h.Status, $h.Column, $h.Row)
        }
        if($needToExport){ [Host]::Export($path, $delimiter) }
        Get-Job | Remove-Job
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

    static [void] Run([ScriptBlock]$command, [Bool]$AsJob)
    {
        # Runs a specified commands on all selected remote hosts
        $hostnames = [Host]::Hosts | Where-Object {$_.Status -and ($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Name} # Gets the names of the hosts that are online and selected
        if ($null -eq $hostnames) { return }
        Write-Host ("Running '{0}' on {1}" -f $command, [String]$hostnames)
        if($AsJob)
        {
            Invoke-Command -ComputerName $hostnames -Credential $script:credential -ScriptBlock $command -AsJob
        }
        else
        {
            Invoke-Command -ComputerName $hostnames -Credential $script:credential -ScriptBlock $command | Write-Host
        }
    }

    static [void] Run([String]$executable, [String]$argument, [String]$workingDirectory)
    {
        # Runs a specified interactive program on all selected remote hosts by creating a scheduled task on currently logged on user then running it and finally deleting it
        $hostnames = [Host]::Hosts | Where-Object {$_.Status -and ($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Name}
        if ($null -eq $hostnames) { return }
        Write-Host ("Running {0}\{1} {2} on {3}" -f $workingDirectory, $executable, $argument, [String]$hostnames)
        Invoke-Command -ComputerName $hostnames -Credential $script:credential -ArgumentList $executable, $argument, $workingDirectory -AsJob -ScriptBlock {
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

    static [void] Wake()
    {
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
    }
}

class BaseCommand : ToolStripMenuItem
{
    # Defines base functionality for a command. Also contains some static members which handle command initialization.

    [Scriptblock]$Script
    static $Commands = [ordered]@{
        "Valitse" = @(
            [BaseCommand]::new("Kaikki", {$script:table.SelectAll()}),
            [BaseCommand]::new("Käänteinen", {$script:table.Rows | ForEach-Object {$_.Cells | ForEach-Object { $_.Selected = !$_.Selected }}})
            [BaseCommand]::new("Ei mitään", {$script:table.ClearSelection()})
        )
        "Tietokone" = @(
            [BaseCommand]::new("Käynnistä", {[Host]::Wake()})
            [BaseCommand]::new("Käynnistä uudelleen", {[Host]::Run({shutdown /r /t 10 /c "Luokanhallinta on ajastanut uudelleen käynnistyksen"}, $true)})
            [BaseCommand]::new("Sammuta", {[Host]::Run({shutdown /s /t 10 /c "Luokanhallinta on ajastanut sammutuksen"}, $true)})
        )
        "VBS3" = @(
            [VBS3Command]::new("Käynnistä")
            [BaseCommand]::new("Synkkaa addonit", {
                [Host]::Run({
                    $source = "\\10.130.16.2\Addons"
                    $domain = "WORKGROUP"
                    $user = "Admin"
                    $password = "kuusteista"
                    net.exe use $source /user:$domain\$user $password
                    Robocopy.exe $source "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons" /MIR /XO /R:0
                }, $true)
            })
            [BaseCommand]::new("Synkkaa asetukset", {
                if(!(Get-SmbShare -Name "VBS3" -ErrorAction SilentlyContinue))
                {
                    $path = "$ENV:USERPROFILE\Documents\VBS3"
                    Write-Host -ForegroundColor Red "$path not shared, creating SMB share..."
                    New-SmbShare -Name "VBS3" -Path $path -Description "Tarvitaan luokanhallintaohjelman asetussynkkiin"
                }
                [Host]::Run({
                    $source = "\\127.0.0.1\VBS3"
                    $domain = "VKY00093"
                    $user = "Uzer"
                    $password = """" # Empty password should be represented as ""
                    net.exe use $source /user:$domain\$user $password
                    Robocopy.exe $source "$ENV:USERPROFILE\Documents\VBS3" "$user.VBS3Profile"
                    Rename-Item -Path "$ENV:USERPROFILE\Documents\VBS3\$user.VBS3Profile" -NewName "$ENV:USERNAME.VBS3Profile"
                }, $true)
            })
            [BaseCommand]::new("Sulje", {[Host]::Run({Stop-Process -ProcessName VBS3_64}, $true)})
        )
        "SteelBeasts" = @(
            [BaseCommand]::new("Käynnistä", {[Host]::Run("SBPro64CM.exe", "", "C:\Program Files\eSim Games\SB Pro FI\Release")})
            [BaseCommand]::new("Sulje", {[Host]::Run({Stop-Process -ProcessName SBPro64CM}, $true)})
        )
        "Muu" = @(
            [BaseCommand]::new("Päivitä", {[Host]::Populate("$PSScriptRoot\luokka.csv", " "); [Host]::Display()})
            [BaseCommand]::new("Vaihda käyttäjä...", {$script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami)})
            # [BaseCommand]::new("Aja...", {[Host]::Run([Scriptblock]::Create((Read-Host "Komento")), $false)})
            # [InteractiveCommand]::new("Aja", "chrome.exe" ,"", "C:\Program Files (x86)\Google\Chrome\Application")
            [BaseCommand]::new("Sulje", {$script:root.Close()})
        )
    } 

    BaseCommand([String]$name, [Scriptblock]$script) : base($name)
    {
        $this.Script = $script
    }

    BaseCommand([String]$name) : base($name){}

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        & $this.Script
    }

    static [void] Display()
    {
        foreach($category in [BaseCommand]::Commands.keys) # Iterates over command categories
        {
            # Create a menu for each category
            $menu = [ToolStripMenuItem]::new()
            $menu.Text = $category
            $script:menubar.Items.Add($menu)
            foreach($command in [BaseCommand]::Commands[$category]) # Iterates over commands in each category
            {
                # Add command to menu
                $menu.DropDownItems.Add($command)
            }
        }
    }
}

class InteractiveCommand : BaseCommand
{
    # Command with three fields for running an interactive programs on remote hosts

    [Object[]]$Widgets = @(
        (New-Object Label -Property @{
            Text = "Ohjelma:"
            AutoSize = $true
            Anchor = [AnchorStyles]::Right
        }),
        (New-Object TextBox -Property @{
            Width = 300
            Anchor = [AnchorStyles]::Left
        }),
        (New-Object Label -Property @{
            Text = "Parametri:"
            AutoSize = $true
            Anchor = [AnchorStyles]::Right
        }),
        (New-Object TextBox -Property @{
            Width = 300
            Anchor = [AnchorStyles]::Left
        }),
        (New-Object Label -Property @{
            Text = "Polku:"
            AutoSize = $true
            Anchor = [AnchorStyles]::Right
        }),
        (New-Object TextBox -Property @{
            Width = 300
            Anchor = [AnchorStyles]::Left
        })
    )

    InteractiveCommand([String]$name, [String]$executable, [String]$argument, [String]$workingDirectory) : base($name + "...")
    {
        # Sets default values for the fields
        ($this.Widgets[1]).Text = $executable
        ($this.Widgets[3]).Text = $argument
        ($this.Widgets[5]).Text = $workingDirectory
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        $form = [Form]::new()
        $form.Text = $this.Text
        $form.FormBorderStyle = [FormBorderStyle]::FixedToolWindow
        $form.Width = 410
        $form.Height = 175

        $grid = [TableLayoutPanel]::new()
        $grid.CellBorderStyle = [TableLayoutPanelCellBorderStyle]::Inset
        $grid.Location = [System.Drawing.Point]::new(0, 0)
        $grid.AutoSize = $true
        $grid.Padding = [Padding]::new(10)
        $grid.ColumnCount = 2
        $grid.RowCount = 4
        $grid.Controls.AddRange($this.Widgets)

        $button = [Button]::new()
        $button.Text = "Aja"
        $button.Dock = [DockStyle]::Bottom
        $button = $button | Add-Member @{Widgets=$this.Widgets} -PassThru -Force
        $button.Add_Click({
            $executable = ($this.Widgets[1]).Text
            $argument = ($this.Widgets[3]).Text
            $workingDirectory = ($this.Widgets[5]).Text
            [Host]::Run($executable, $argument, $workingDirectory)
        })
        $grid.Controls.Add($button)
        $grid.SetColumnSpan($button, 2)

        $form.Controls.Add($grid)
        $form.ShowDialog()
    }
}

class VBS3Command : BaseCommand
{
    $States = [ordered]@{
        "Kokonäyttö" = ""
        "Ikkuna" = "-window"
        "Palvelin" = "-server"
        "Simulation Client" = "simulationClient=0"
        "After Action Review" = "simulationClient=1"
        "SC + AAR" = "simulationClient=2"
    }

    VBS3Command([String]$name) : base($name + "..."){}

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        $form = [Form]::new()
        $form.AutoSize = $true
        $form.FormBorderStyle = [FormBorderStyle]::FixedToolWindow
        $form.Text = $this.Text

        $grid = [TableLayoutPanel]::new()
        $grid.AutoSize = $true
        $grid.ColumnCount = 2
        $grid.Padding = [Padding]::new(10)
        $grid.CellBorderStyle = [TableLayoutPanelCellBorderStyle]::Inset

        $statePanel = [FlowLayoutPanel]::new()
        $statePanel.AutoSize = $true
        $statePanel.FlowDirection = [FlowDirection]::TopDown
        $this.States.Keys | ForEach-Object {
            $r = [RadioButton]::new()
            $r.AutoSize = $true
            $r.Margin = [Padding]::new(0)
            $r.Text = $_
            if($r.Text -eq "Kokonäyttö"){ $r.Checked = $true } 
            $statePanel.Controls.Add($r)
        }
        $grid.SetCellPosition($statePanel, [TableLayoutPanelCellPosition]::new(1, 0)) 
        $grid.Controls.Add($statePanel)

        $adminCheckBox = [CheckBox]::new()
        $adminCheckBox.Text = "Admin"
        $grid.SetCellPosition($adminCheckBox, [TableLayoutPanelCellPosition]::new(1, 1)) 
        $grid.Controls.Add($adminCheckBox)

        $multicastCheckBox = [CheckBox]::new()
        $multicastCheckBox.Text = "Multicast"
        $multicastCheckBox.Checked = $true
        $grid.SetCellPosition($multicastCheckBox, [TableLayoutPanelCellPosition]::new(1, 2)) 
        $grid.Controls.Add($multicastCheckBox)

        $configLabel = [Label]::new()
        $configLabel.Text = "cfg="
        $configLabel.AutoSize = $true
        $configLabel.Anchor = [AnchorStyles]::Right
        $grid.SetCellPosition($configLabel, [TableLayoutPanelCellPosition]::new(0, 3)) 
        $grid.Controls.Add($configLabel)

        $configTextBox = [TextBox]::new()
        $configTextBox.Width = 200
        $grid.SetCellPosition($configTextBox, [TableLayoutPanelCellPosition]::new(1, 3)) 
        $grid.Controls.Add($configTextBox)

        $connectLabel = [Label]::new()
        $connectLabel.Text = "connect="
        $connectLabel.AutoSize = $true
        $connectLabel.Anchor = [AnchorStyles]::Right
        $grid.SetCellPosition($connectLabel, [TableLayoutPanelCellPosition]::new(0, 4)) 
        $grid.Controls.Add($connectLabel)

        $connectTextBox = [TextBox]::new()
        $connectTextBox.Width = 200
        $grid.SetCellPosition($connectTextBox, [TableLayoutPanelCellPosition]::new(1, 4)) 
        $grid.Controls.Add($connectTextBox)

        $cpuCountLabel = [Label]::new()
        $cpuCountLabel.Text = "cpuCount="
        $cpuCountLabel.AutoSize = $true
        $cpuCountLabel.Anchor = [AnchorStyles]::Right
        $grid.SetCellPosition($cpuCountLabel, [TableLayoutPanelCellPosition]::new(0, 5)) 
        $grid.Controls.Add($cpuCountLabel)

        $cpuCountTextBox = [TextBox]::new()
        $cpuCountTextBox.Width = 200
        $grid.SetCellPosition($cpuCountTextBox, [TableLayoutPanelCellPosition]::new(1, 5)) 
        $grid.Controls.Add($cpuCountTextBox)

        $exThreadsLabel = [Label]::new()
        $exThreadsLabel.Text = "exThreads="
        $exThreadsLabel.AutoSize = $true
        $exThreadsLabel.Anchor = [AnchorStyles]::Right
        $grid.SetCellPosition($exThreadsLabel, [TableLayoutPanelCellPosition]::new(0, 6)) 
        $grid.Controls.Add($exThreadsLabel)

        $exThreadsTextBox = [TextBox]::new()
        $exThreadsTextBox.Width = 200
        $grid.SetCellPosition($exThreadsTextBox, [TableLayoutPanelCellPosition]::new(1, 6)) 
        $grid.Controls.Add($exThreadsTextBox)

        $maxMemLabel = [Label]::new()
        $maxMemLabel.Text = "maxMem="
        $maxMemLabel.AutoSize = $true
        $maxMemLabel.Anchor = [AnchorStyles]::Right
        $grid.SetCellPosition($maxMemLabel, [TableLayoutPanelCellPosition]::new(0, 7)) 
        $grid.Controls.Add($maxMemLabel)

        $maxMemTextBox = [TextBox]::new()
        $maxMemTextBox.Width = 200
        $grid.SetCellPosition($maxMemTextBox, [TableLayoutPanelCellPosition]::new(1, 7)) 
        $grid.Controls.Add($maxMemTextBox)

        $runButton = [Button]::new()
        $runButton = $runButton | Add-Member @{States=$this.States; StatePanel=$statePanel; AdminCheckbox=$adminCheckBox; MulticastCheckBox=$multicastCheckBox; ConfigTextBox=$configTextBox; ConnectTextBox=$connectTextBox; CpuCountTextBox=$cpuCountTextBox; ExThreadsTextBox=$exThreadsTextBox; MaxMemTextBox=$maxMemTextBox} -PassThru -Force
        $runButton.Text = "Käynnistä"
        $runButton.Add_Click({
            $state = $this.StatePanel.Controls | Where-Object {$_.Checked} | Select-Object -ExpandProperty Text
            $argument = $this.States[$state]
            if($this.AdminCheckbox.Checked){ $argument = "-admin $argument" }
            if(!$this.MulticastCheckBox.Checked){ $argument = "-multicast=0 $argument"}
            if($this.ConfigTextBox.Text){ $argument = ("cfg={0} {1}" -f $this.ConfigTextBox.Text, $argument)}
            if($this.ConnectTextBox.Text){ $argument = ("connect={0} {1}" -f $this.ConnectTextBox.Text, $argument)}
            if($this.CpuCountTextBox.Text){ $argument = ("cpuCount={0} {1}" -f $this.CpuCountTextBox.Text, $argument)}
            if($this.ExThreadsTextBox.Text){ $argument = ("exThreads={0} {1}" -f $this.ExThreadsTextBox.Text, $argument)}
            if($this.MaxMemTextBox.Text){ $argument = ("maxMem={0} {1}" -f $this.MaxMemTextBox.Text, $argument)}
            [Host]::Run("VBS3_64.exe", $argument, "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI")
        })
        $runButton.Dock = [DockStyle]::Bottom
        $form.AcceptButton = $runButton
        $grid.SetCellPosition($runButton, [TableLayoutPanelCellPosition]::new(0, 8))
        $grid.SetColumnSpan($runButton, 2)
        $grid.Controls.Add($runButton)
        $form.Controls.Add($grid)
        $form.ShowDialog()
    }
}

[Host]::Populate("$PSScriptRoot\luokka.csv", " ")
$script:root = [Form]::new()
$root.Text = "Luokanhallinta v0.8"

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

$script:menubar = [MenuStrip]::new()
$root.MainMenuStrip = $menubar
$menubar.Dock = [DockStyle]::Top
$root.Controls.Add($menubar)
[BaseCommand]::Display()

$script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami)
[void]$root.showDialog()
