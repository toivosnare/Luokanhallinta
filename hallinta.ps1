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

    static [void] Run([Bool]$AsJob, [Object[]]$Params, [ScriptBlock]$command)
    {
        # Runs a specified commands on all selected remote hosts
        $hostnames = [Host]::Hosts | Where-Object {$_.Status -and ($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Name} # Gets the names of the hosts that are online and selected
        if ($null -eq $hostnames) { return }
        Write-Host ("Running '{0}' on {1}" -f $command, [String]$hostnames)
        if($AsJob)
        {
            Invoke-Command -ComputerName $hostnames -Credential $script:credential -ScriptBlock $command -ArgumentList $Params -AsJob
        }
        else
        {
            Invoke-Command -ComputerName $hostnames -Credential $script:credential -ScriptBlock $command -ArgumentList $Params | Write-Host
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
            [BaseCommand]::new("Virus scan", {[Host]::Run("fsav.exe", "/disinf C: D: ", "C:\Program Files (x86)\F-Secure\Anti-Virus")})
            [BaseCommand]::new("Käynnistä uudelleen", {[Host]::Run($true, @(), {shutdown /r /t 10 /c "Luokanhallinta on ajastanut uudelleen käynnistyksen"})})
            [BaseCommand]::new("Sammuta", {[Host]::Run($true, @(), {shutdown /s /t 10 /c "Luokanhallinta on ajastanut sammutuksen"})})
        )
        "VBS3" = @(
            [VBS3Command]::new("Käynnistä...")
            #[CopyCommand]::new("Synkkaa addonit...", "\\10.130.16.2\Addons", "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons", "WORKGROUP\Admin", "kuusteista", "/MIR")
            #[CopyCommand]::new("Synkkaa asetukset...", ("\\{0}\VBS3" -f (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress.Length -gt 1 -and $_.IPAddress -ne "127.0.0.1"} | Select-Object -ExpandProperty IPAddress)), "$ENV:USERPROFILE\Documents\VBS3", $(whoami), "", "$ENV:USERNAME.VBS3Profile")
            [NewCopyCommand]::new("Synkkaa addonit", "\\10.130.16.2\addons", "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons", "WORKGROUP\Admin", "kuusteista")
            [NewCopyCommand]::new("Synkkaa asetukset", "$ENV:USERPROFILE\Documents\VBS3\$ENV:USERNAME.VBS3Profile", "$ENV:USERPROFILE\Documents\VBS3\$ENV:USERNAME.VBS3Profile", "", "")
            [BaseCommand]::new("Sulje", {[Host]::Run($true, @(), {Stop-Process -ProcessName VBS3_64})})
        )
        "SteelBeasts" = @(
            [BaseCommand]::new("Käynnistä", {[Host]::Run("SBPro64CM.exe", "", "C:\Program Files\eSim Games\SB Pro FI\Release")})
            [BaseCommand]::new("Sulje", {[Host]::Run($true, @(), {Stop-Process -ProcessName SBPro64CM})})
        )
        "Muu" = @(
            [BaseCommand]::new("Päivitä", {[Host]::Populate("$PSScriptRoot\luokka.csv", " "); [Host]::Display()})
            [BaseCommand]::new("Vaihda käyttäjä...", {$script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami)})
            [BaseCommand]::new("Aja...", {[Host]::Run($false, @(), [Scriptblock]::Create((Read-Host "Komento")))})
            [InteractiveCommand]::new("Aja...", "chrome.exe", "", "C:\Program Files (x86)\Google\Chrome\Application")
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

    [Form]$Form
    [TableLayoutPanel]$Grid
    [Label]$ProgramLabel
    [TextBox]$ProgramTextBox
    [Label]$ArgumentLabel
    [TextBox]$ArgumentTextBox
    [Label]$PathLabel
    [TextBox]$PathTextBox
    [Button]$RunButton
    [Scriptblock]$ClickScript = { [Host]::Run($this.ProgramTextBox.Text, $this.ArgumentTextBox.Text, $this.PathTextBox.Text) }

    InteractiveCommand([String]$name, [String]$program, [String]$argument, [String]$path) : base($name)
    {
        $this.Form = [Form]::new()
        $this.Form.Text = $this.Text
        $this.Form.FormBorderStyle = [FormBorderStyle]::FixedToolWindow
        $this.Form.AutoSize = $true

        $this.Grid = [TableLayoutPanel]::new()
        $this.Grid.CellBorderStyle = [TableLayoutPanelCellBorderStyle]::Inset
        $this.Grid.AutoSize = $true
        $this.Grid.Padding = [Padding]::new(10)
        $this.Grid.ColumnCount = 2

        $this.ProgramLabel = [Label]::new()
        $this.ProgramLabel.Text = "Ohjelma:"
        $this.ProgramLabel.AutoSize = $true
        $this.ProgramLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.ProgramLabel, [TableLayoutPanelCellPosition]::new(0, 0))
        $this.Grid.Controls.Add($this.ProgramLabel)
        $this.ProgramTextBox = [TextBox]::new()
        $this.ProgramTextBox.Width = 300
        $this.ProgramTextBox.Anchor = [AnchorStyles]::Left
        $this.ProgramTextBox.Text = $program
        $this.Grid.SetCellPosition($this.ProgramTextBox, [TableLayoutPanelCellPosition]::new(1, 0))
        $this.Grid.Controls.Add($this.ProgramTextBox)

        $this.ArgumentLabel = [Label]::new()
        $this.ArgumentLabel.Text = "Parametri:"
        $this.ArgumentLabel.AutoSize = $true
        $this.ArgumentLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.ArgumentLabel, [TableLayoutPanelCellPosition]::new(0, 1))
        $this.Grid.Controls.Add($this.ArgumentLabel)
        $this.ArgumentTextBox = [TextBox]::new()
        $this.ArgumentTextBox.Width = 300
        $this.ArgumentTextBox.Anchor = [AnchorStyles]::Left
        $this.ArgumentTextBox.Text = $argument
        $this.Grid.SetCellPosition($this.ArgumentTextBox, [TableLayoutPanelCellPosition]::new(1, 1))
        $this.Grid.Controls.Add($this.ArgumentTextBox)

        $this.PathLabel = [Label]::new()
        $this.PathLabel.Text = "Polku:"
        $this.PathLabel.AutoSize = $true
        $this.PathLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.PathLabel, [TableLayoutPanelCellPosition]::new(0, 2))
        $this.Grid.Controls.Add($this.PathLabel)
        $this.PathTextBox = [TextBox]::new()
        $this.PathTextBox.Width = 300
        $this.PathTextBox.Anchor = [AnchorStyles]::Left
        $this.PathTextBox.Text = $path
        $this.Grid.SetCellPosition($this.PathTextBox, [TableLayoutPanelCellPosition]::new(1, 2))
        $this.Grid.Controls.Add($this.PathTextBox)

        $this.RunButton = [Button]::new()
        $this.RunButton.Text = "Aja"
        $this.RunButton.Dock = [DockStyle]::Bottom
        $this.RunButton = $this.RunButton | Add-Member @{ProgramTextBox=$this.ProgramTextBox; ArgumentTextBox=$this.ArgumentTextBox; PathTextBox=$this.PathTextBox} -PassThru -Force
        $this.RunButton.Add_Click($this.ClickScript)
        $this.Grid.SetCellPosition($this.RunButton, [TableLayoutPanelCellPosition]::new(0, 3))
        $this.Grid.SetColumnSpan($this.RunButton, 2)
        $this.Grid.Controls.Add($this.RunButton)
        $this.Form.Controls.Add($this.Grid)
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        $this.Form.ShowDialog()
    }
}

class VBS3Command : BaseCommand
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
    [Scriptblock]$ClickScript = {
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
    }

    VBS3Command([String]$name) : base($name)
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
        $this.RunButton = $this.RunButton | Add-Member @{States=$this.States; StatePanel=$this.StatePanel; AdminCheckbox=$this.AdminCheckBox; MulticastCheckBox=$this.MulticastCheckBox; ConfigTextBox=$this.ConfigTextBox; ConnectTextBox=$this.ConnectTextBox; CpuCountTextBox=$this.CpuCountTextBox; ExThreadsTextBox=$this.ExThreadsTextBox; MaxMemTextBox=$this.MaxMemTextBox} -PassThru -Force
        $this.RunButton.Text = "Käynnistä"
        $this.RunButton.Add_Click($this.ClickScript)
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
}

class CopyCommand : BaseCommand
{
    [Form]$Form
    [TableLayoutPanel]$Grid
    [Label]$SourceLabel
    [TextBox]$sourceTextBox
    [Label]$DestinationLabel
    [TextBox]$DestinationTextBox
    [Label]$UsernameLabel
    [TextBox]$UsernameTextBox
    [Label]$PasswordLabel
    [TextBox]$PasswordTextBox
    [Label]$ArgumentLabel
    [TextBox]$ArgumentTextBox
    [Button]$RunButton
    [ScriptBlock]$ClickScript = {
        [Host]::Run($false, @($this.Source.Text, $this.Destination.Text, $this.Username.Text, $this.Password.Text, $this.Argument.Text), {
            param(
                [String]$source,
                [String]$destination,
                [String]$username,
                [String]$password,
                [String]$argument
            )
            # Write-Host ("{0}, {1}, {2}, {3}, {4}" -f $source, $destination, $username, $password, $argument)
            if(!$password){ $password = """" }
            net.exe use $source /user:$username $password
            Robocopy.exe "$source" "$destination" "$argument"
            net.exe use /delete $source
        })
    }

    CopyCommand([String]$name, [String]$source, [String]$destination, [String]$username, [String]$password, [String]$argument) : base($name)
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

        $this.SourceLabel = [Label]::new()
        $this.SourceLabel.Text = "Lähde:"
        $this.SourceLabel.AutoSize = $true
        $this.SourceLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.SourceLabel, [TableLayoutPanelCellPosition]::new(0, 0)) 
        $this.Grid.Controls.Add($this.SourceLabel)
        $this.SourceTextBox = [TextBox]::new()
        $this.SourceTextBox.Width = 200
        $this.SourceTextBox.Text = $source
        $this.Grid.SetCellPosition($this.SourceTextBox, [TableLayoutPanelCellPosition]::new(1, 0)) 
        $this.Grid.Controls.Add($this.SourceTextBox)

        $this.DestinationLabel = [Label]::new()
        $this.DestinationLabel.Text = "Kohde:"
        $this.DestinationLabel.AutoSize = $true
        $this.DestinationLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.DestinationLabel, [TableLayoutPanelCellPosition]::new(0, 1)) 
        $this.Grid.Controls.Add($this.DestinationLabel)
        $this.DestinationTextBox = [TextBox]::new()
        $this.DestinationTextBox.Width = 200
        $this.DestinationTextBox.Text = $destination
        $this.Grid.SetCellPosition($this.DestinationTextBox, [TableLayoutPanelCellPosition]::new(1, 1)) 
        $this.Grid.Controls.Add($this.DestinationTextBox)

        $this.UsernameLabel = [Label]::new()
        $this.UsernameLabel.Text = "Käyttäjänimi:"
        $this.UsernameLabel.AutoSize = $true
        $this.UsernameLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.UsernameLabel, [TableLayoutPanelCellPosition]::new(0, 2)) 
        $this.Grid.Controls.Add($this.UsernameLabel)
        $this.UsernameTextBox = [TextBox]::new()
        $this.UsernameTextBox.Width = 200
        $this.UsernameTextBox.Text = $username
        $this.Grid.SetCellPosition($this.UsernameTextBox, [TableLayoutPanelCellPosition]::new(1, 2)) 
        $this.Grid.Controls.Add($this.UsernameTextBox)

        $this.PasswordLabel = [Label]::new()
        $this.PasswordLabel.Text = "Salasana:"
        $this.PasswordLabel.AutoSize = $true
        $this.PasswordLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.PasswordLabel, [TableLayoutPanelCellPosition]::new(0, 3)) 
        $this.Grid.Controls.Add($this.PasswordLabel)
        $this.PasswordTextBox = [TextBox]::new()
        $this.PasswordTextBox.Width = 200
        $this.PasswordTextBox.Text = $password
        $this.PasswordTextBox.PasswordChar = "*"
        $this.Grid.SetCellPosition($this.PasswordTextBox, [TableLayoutPanelCellPosition]::new(1, 3)) 
        $this.Grid.Controls.Add($this.PasswordTextBox)

        $this.ArgumentLabel = [Label]::new()
        $this.ArgumentLabel.Text = "Parametri:"
        $this.ArgumentLabel.AutoSize = $true
        $this.ArgumentLabel.Anchor = [AnchorStyles]::Right
        $this.Grid.SetCellPosition($this.ArgumentLabel, [TableLayoutPanelCellPosition]::new(0, 4)) 
        $this.Grid.Controls.Add($this.ArgumentLabel)
        $this.ArgumentTextBox = [TextBox]::new()
        $this.ArgumentTextBox.Width = 200
        $this.ArgumentTextBox.Text = $argument
        $this.Grid.SetCellPosition($this.ArgumentTextBox, [TableLayoutPanelCellPosition]::new(1, 4)) 
        $this.Grid.Controls.Add($this.ArgumentTextBox)

        $this.RunButton = [Button]::new()
        $this.RunButton.Text = "Kopioi"
        $this.RunButton = $this.RunButton | Add-Member @{Source=$this.SourceTextBox; Destination=$this.DestinationTextBox; Username=$this.UsernameTextBox; Password=$this.PasswordTextBox; Argument=$this.ArgumentTextBox} -PassThru -Force
        $this.RunButton.Add_Click($this.ClickScript)
        $this.RunButton.Dock = [DockStyle]::Bottom
        $this.Grid.SetCellPosition($this.RunButton, [TableLayoutPanelCellPosition]::new(0, 5))
        $this.Grid.SetColumnSpan($this.RunButton, 2)
        $this.Form.AcceptButton = $this.RunButton
        $this.Grid.Controls.Add($this.RunButton)
        $this.Form.Controls.Add($this.Grid)
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        $this.Form.ShowDialog()
    }
}

class NewCopyCommand : BaseCommand
{
    [String]$Source
    [String]$Destination
    [String]$Username
    [String]$Password

    NewCopyCommand([String]$name, [String]$source, [String]$destination, [String]$username, [String]$password) : base($name)
    {
        $this.Source = $source
        $this.Destination = $destination
        $this.Username = $username
        $this.Password = $password
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        Write-Host ("Mirroring {0} to {1}" -f $this.Source, $this.Destination)
        if($this.Username)
        {
            $u = $this.Username
            $p = $this.Password
            net.exe use $this.Source /user:$u $p
        }
        [Host]::Hosts | Where-Object {$_.Status -and ($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {
            Write-Host -ForegroundColor Magenta $_.Name
            $session = New-PSSession -ComputerName $_.Name
            Get-ChildItem $this.Source | ForEach-Object {
                $sourcePath = $_.FullName
                $destinationPath = $sourcePath.Replace($this.Source, $this.Destination)

                Write-Host -NoNewline ("{0}: " -f $_)
                if(Test-Path $destinationPath)
                {
                    $sourceFile = Get-Item $sourcePath
                    $destinationFile = Invoke-Command -Session $session -ArgumentList $destinationPath -ScriptBlock {
                        param([String]$destinationPath)
                        $destinationFile = Get-Item $destinationPath
                        return $destinationFile
                    }
                    if($sourceFile.LastWriteTime -gt $destinationFile.LastWriteTime)
                    {
                        Write-Host -ForegroundColor Yellow "newer version, copying"
                        Copy-Item $sourcePath -Destination $destinationPath -ToSession $session -Force
                    }
                    else
                    {
                        Write-Host -ForegroundColor Green "already in place, skipping"  
                    }
                }
                else
                {
                    Write-Host -ForegroundColor Yellow "not found, copying"
                    Copy-Item $sourcePath -Destination $destinationPath -ToSession $session
                }
            }
            Get-ChildItem $destination | ForEach-Object {
                $destinationPath = $_.FullName
                $sourcePath = $destinationPath.Replace($this.Destination, $this.Source)
                if((Test-Path $sourcePath) -eq $false)
                {
                    Write-Host -ForegroundColor Red ("{0} not in {1}, removing" -f $_, $this.Source)
                    Remove-Item $destinationPath
                }
            }
        }
        if($this.Username)
        {
            net.exe use /delete $this.Source
        }
        Write-Host "Done"
    }
}

[Host]::Populate("$PSScriptRoot\luokka.csv", " ")
# if(!(Get-SmbShare -Name "VBS3" -ErrorAction SilentlyContinue))
# {
#     $path = "$ENV:USERPROFILE\Documents\VBS3"
#     Write-Host -ForegroundColor Red "$path not shared, creating SMB share..."
#     New-SmbShare -Name "VBS3" -Path $path -Description "Tarvitaan luokanhallintaohjelman asetussynkkiin"
# }
$script:root = [Form]::new()
$root.Text = "Luokanhallinta v0.10"

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
