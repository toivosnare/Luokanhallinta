using namespace System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms

class Host
{
    static [Host[]]$Hosts = @()
    [String]$Name
    [String]$Mac
    [Bool]$Status
    [Int]$Column
    [Int]$Row

    Host([String]$name, [String]$mac, [Int]$column, [Int]$row)
    {
        $this.Name = $name
        $this.Status = Test-Connection -ComputerName $this.Name -Count 1 -Quiet
        if($mac)
        {
            $this.Mac = $mac
        }
        elseif($this.Status)
        {
            $this.Mac = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $this.Name | Select-Object -First 1 -ExpandProperty MACAddress
        }
        $this.Column = $column
        $this.Row = $row
    }

    static [void] Populate([String]$path, [String]$delimiter)
    {
        [Host]::Hosts = @()
        Import-Csv $path -Delimiter $delimiter | ForEach-Object {[Host]::Hosts += [Host]::new($_.Nimi, $_.Mac, [Int]$_.Sarake, [Int]$_.Rivi)}
    }

    static [void] Display()
    {
        $script:table.ColumnCount = ([Host]::Hosts | ForEach-Object {$_.Column} | Measure-Object -Maximum).Maximum
        $script:table.RowCount = ([Host]::Hosts | ForEach-Object {$_.Row} | Measure-Object -Maximum).Maximum
        $script:table.Columns | ForEach-Object {
            $_.HeaderText = [Char]($_.Index + 65)
            $_.SortMode = [DataGridViewColumnSortMode]::NotSortable
        }
        $script:table.Rows | ForEach-Object { $_.HeaderCell.Value = [String]($_.Index + 1) }
        foreach($h in [Host]::Hosts)
        {
            $cell = $script:table[($h.Column - 1), ($h.Row - 1)]
            $cell.Value = $h.Name
            $cell.ToolTipText = $h.Mac
            if($h.Status)
            {
                $cell.Style.ForeColor = [System.Drawing.Color]::Green
                $cell.Style.SelectionForeColor = [System.Drawing.Color]::Green
            }
        }
    }

    static [String] Run([ScriptBlock]$command, [Bool]$AsJob=$false)
    {
        $hostnames = [Host]::Hosts | Where-Object {$_.Status -and ($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Name}
        if ($null -eq $hostnames) { return ""}
        $session = New-PSSession -ComputerName $hostnames -Credential $script:credential
        if ($session.Availability -ne [System.Management.Automation.Runspaces.RunspaceAvailability]::Available){ return "Virhe" }
        if($AsJob)
        {
            Invoke-Command -Session $session -ScriptBlock $command -AsJob
            return "Job started"
        }
        else
        {
            return Invoke-Command -Session $session -ScriptBlock $command
        }
    }

    static [Bool] Run([String]$executable, [String]$argument, [String]$workingDirectory)
    {
        $hostnames = [Host]::Hosts | Where-Object {$_.Status -and ($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Name}
        if ($null -eq $hostnames) { return ""}
        $session = New-PSSession -ComputerName $hostnames -Credential $script:credential
        if ($session.Availability -ne [System.Management.Automation.Runspaces.RunspaceAvailability]::Available){ return "Virhe" }
        Invoke-Command -Session $session -ArgumentList $executable, $argument, $workingDirectory -ScriptBlock {
            param($executable, $argument, $workingDirectory)
            if($argument -eq ""){ $argument = " " }
            $action = New-ScheduledTaskAction -Execute $executable -Argument $argument -WorkingDirectory $workingDirectory
            #$principal = New-ScheduledTaskPrincipal -GroupId "BUILTIN\Administrators" -RunLevel Highest
            $task = New-ScheduledTask -Action $action #-Principal $principal
            $taskname = "LKNHLNT"
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
        return $true
    }

    static [void] Wake()
    {
        $macs = [Host]::Hosts | Where-Object {($script:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Mac}
        $port = 9
        $broadcast = [Net.IPAddress]::Parse("255.255.255.255")
        foreach($m in $macs)
        {
            $m = (($m.replace(":", "")).replace("-", "")).replace(".", "")
            $target = 0, 2, 4, 6, 8, 10 | % {[convert]::ToByte($m.substring($_, 2), 16)}
            $packet = (,[byte]255 * 6) + ($target * 16)
            $UDPclient = [System.Net.Sockets.UdpClient]::new()
            $UDPclient.Connect($broadcast, $port)
            $UDPclient.Send($packet, 102)
        }
    }
}

class BaseCommand : ToolStripMenuItem
{
    [Scriptblock]$Script
    static $Commands = [ordered]@{
        "Valitse" = @(
            [BaseCommand]::new("Kaikki", {$script:table.SelectAll()}),
            [BaseCommand]::new("Ei mitään", {$script:table.ClearSelection()})
        )
        "Tietokone" = @(
            [BaseCommand]::new("Käynnistä", {[Host]::Wake()})
            [BaseCommand]::new("Käynnistä uudelleen", {[Host]::Run({shutdown /r /t 10 /c "Luokanhallinta on ajastanut uudelleen käynnistyksen"}, $false)})
            [BaseCommand]::new("Sammuta", {[Host]::Run({shutdown /s /t 10 /c "Luokanhallinta on ajastanut sammutuksen"}, $false)})
        )
        "VBS3" = @(
            [InteractiveCommand]::new("Käynnistä", "VBS3_64.exe", "", "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI")
            [BaseCommand]::new("Synkkaa addonit", {Write-Host ([Host]::Run({robocopy '\\PSPR-Storage' 'C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons' /MIR /XO /R:2 /W:10}, $true))})
            [BaseCommand]::new("Sulje", {[Host]::Run({Stop-Process -ProcessName VBS3_64}, $false)})
        )
        "SteelBeasts" = @(
            [BaseCommand]::new("Käynnistä", {[Host]::Run("SBPro64CM.exe", "", "C:\Program Files\eSim Games\SB Pro FI\Release")})
            [BaseCommand]::new("Sulje", {[Host]::Run({Stop-Process -ProcessName SBPro64CM}, $false)})
        )
        "Muu" = @(
            [BaseCommand]::new("Sulje", {$script:root.Close()})
            [InteractiveCommand]::new("Chrome", "chrome.exe", "", "C:\Program Files (x86)\Google\Chrome\Application")
            [BaseCommand]::new("Aja", {Write-Host ([Host]::Run([Scriptblock]::Create((Read-Host "Komento")), $false))})
            [BaseCommand]::new("Vaihda käyttäjä", {$script:credential = Get-Credential -Message "Hallitse luokkaa seuraavalla käyttäjällä" -UserName $(whoami)})
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
        foreach($category in [BaseCommand]::Commands.keys)
        {
            $menu = [ToolStripMenuItem]::new()
            $menu.Text = $category
            $script:menubar.Items.Add($menu)
            foreach($command in [BaseCommand]::Commands[$category])
            {
                $menu.DropDownItems.Add($command)
            }
        }
    }
}

class PopUpCommand : BaseCommand
{
    [Object[]]$Widgets
    [Scriptblock]$ClickScript
    [Scriptblock]$RunScript

    PopUpCommand([String]$name, [Object[]]$widgets, [ScriptBlock]$clickScript, [Scriptblock]$runScript) : base($name + "...")
    {
        $this.Widgets = $widgets
        $this.ClickScript = $clickScript
        $this.RunScript = $runScript
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([ToolStripMenuItem]$this).OnClick($e)
        $form = [Form]::new()
        $form.Text = $this.Text
        $form.AutoSize = $true
        $form.FormBorderStyle = [FormBorderStyle]::FixedToolWindow
        $button = [RunButton]::new($this, $form)
        & $this.ClickScript
        $form.ShowDialog()
    }

    [void] Run(){ & $this.RunScript }
}

class InteractiveCommand : PopUpCommand
{
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
    
    [Scriptblock]$ClickScript = {
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
        $button.Dock = [DockStyle]::Bottom
        $grid.Controls.Add($button)
        $grid.SetColumnSpan($button, 2)
        $form.Controls.Add($grid)
    }
    [Scriptblock]$RunScript = {
        $executable = ($this.Widgets[1]).Text
        $argument = ($this.Widgets[3]).Text
        $workingDirectory = ($this.Widgets[5]).Text
        [Host]::Run($executable, $argument, $workingDirectory)
    }

    InteractiveCommand([String]$name, [String]$executable, [String]$argument, [String]$workingDirectory) : base($name, $this.Widgets, $this.ClickScript, $this.RunScript)
    {
        ($this.Widgets[1]).Text = $executable
        ($this.Widgets[3]).Text = $argument
        ($this.Widgets[5]).Text = $workingDirectory
    }
}

class RunButton : Button
{
    [BaseCommand]$Command
    [Form]$Form

    RunButton([BaseCommand]$command, [Form]$form)
    {
        $this.Command = $command
        $this.Form = $form
        $this.Text = "Aja"
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([Button]$this).OnClick($e)
        $this.Command.Run()
        $this.Form.Close()
    }
}

[Host]::Populate("$PSScriptRoot\luokka.csv", " ")
$script:root = [Form]::new()
$root.Text = "Luokanhallinta"
$root.Width = 1280
$root.Height = 720

$script:table = [DataGridView]::new()
$table.Dock = [DockStyle]::Fill
$table.AllowUserToAddRows = $false
$table.AllowUserToDeleteRows = $false
$table.AllowUserToResizeColumns = $false
$table.AllowUserToResizeRows = $false
$table.AllowUserToOrderColumns = $false
$table.ReadOnly = $true
$table.AutoSizeColumnsMode = [DataGridViewAutoSizeColumnsMode]::AllCells
$table.ColumnHeadersHeightSizeMode = [DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$table.AutoSizeRowsMode = [DataGridViewAutoSizeRowsMode]::AllCells
$table.RowHeadersWidthSizeMode = [DataGridViewRowHeadersWidthSizeMode]::DisableResizing
($table.RowsDefaultCellStyle).Padding = [Padding]::new(30)
($table.RowsDefaultCellStyle).ForeColor = [System.Drawing.Color]::Red
($table.RowsDefaultCellStyle).SelectionForeColor = [System.Drawing.Color]::Red
($table.RowsDefaultCellStyle).SelectionBackColor = [System.Drawing.Color]::LightGray
$table.SelectionMode = [DataGridViewSelectionMode]::CellSelect
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
        $min = ($script:startColumn, $endColumn | Measure-Object -Min).Minimum
        $max = ($script:startColumn, $endColumn | Measure-Object -Max).Maximum
        for($c = $min; $c -le $max; $c++)
        {
            for($r = 0; $r -lt $this.RowCount; $r++)
            {
                $this[[Int]$c, [Int]$r].Selected = $true
            }
        }
    }
    elseif($_.ColumnIndex -eq -1 -and $_.RowIndex -ne -1)
    {
        $endRow = $_.RowIndex
        $min = ($script:startRow, $endRow | Measure-Object -Min).Minimum
        $max = ($script:startRow, $endRow | Measure-Object -Max).Maximum
        for($r = $min; $r -le $max; $r++)
        {
            for($c = 0; $c -lt $this.ColumnCount; $c++)
            {
                $this[[Int]$c, [Int]$r].Selected = $true
            }
        }
    }
    elseif($_.Button -eq [MouseButtons]::Right)
    {
        $this[$_.ColumnIndex, $_.RowIndex].Selected = $false
    }
})
$root.Controls.Add($table)
[Host]::Display()

$script:menubar = [MenuStrip]::new()
$root.MainMenuStrip = $menubar
$menubar.Dock = [DockStyle]::Top
$root.Controls.Add($menubar)
[BaseCommand]::Display()

$script:credential = Get-Credential -Message "Hallitse luokkaa seuraavalla käyttäjällä" -UserName $(whoami)
[void]$root.showDialog()
