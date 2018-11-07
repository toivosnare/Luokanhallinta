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
        $this.Mac = $mac
        $this.Status = Test-Connection -ComputerName $this.Name -Count 1 -Quiet
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
        foreach($h in [Host]::Hosts)
        {
            $cell = $global:table[($h.Column - 1), ($h.Row - 1)]
            $cell.Value = $h.Name
            $cell.ToolTipText = $h.Mac
            if($h.Status)
            {
                $cell.Style.ForeColor = [System.Drawing.Color]::Green
                $cell.Style.SelectionForeColor = [System.Drawing.Color]::Green
            }
        }
    }

    hidden static [System.Management.Automation.Runspaces.PSSession] GetSession()
    {
        $hostnames = [Host]::Hosts | Where-Object {$_.Status -and ($global:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Name}
        if($null -eq $hostnames){ return $null }
        return New-PSSession -ComputerName $hostnames
    }

    static [String] Run([ScriptBlock]$command, [Bool]$AsJob=$false)
    {
        $session = [Host]::GetSession()
        if (($null -eq $session) -or ($session.Availability -ne [System.Management.Automation.Runspaces.RunspaceAvailability]::Available)){ return "Virhe" }
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
        $session = [Host]::GetSession()
        if (($null -eq $session) -or ($session.Availability -ne [System.Management.Automation.Runspaces.RunspaceAvailability]::Available)){ return $false }
        Invoke-Command -Session $session -ArgumentList $executable, $argument, $workingDirectory -ScriptBlock {
            param($executable, $argument, $workingDirectory)
            if($argument -eq ""){ $argument = " " }
            $action = New-ScheduledTaskAction -Execute $executable -Argument $argument -WorkingDirectory $workingDirectory
            $principal = New-ScheduledTaskPrincipal -userid $(whoami)
            $task = New-ScheduledTask -Action $action -Principal $principal
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
        $macs = [Host]::Hosts | Where-Object {($global:table[($_.Column - 1), ($_.Row - 1)]).Selected} | ForEach-Object {$_.Mac}
        $port = 9
        $broadcast = [Net.IPAddress]::Parse("255.255.255.255")
        foreach($m in $macs)
        {
            $m = (($m.replace(":", "")).replace("-", "")).replace(".", "")
            $target = 0, 2, 4, 6, 8, 10 | % {[convert]::ToByte($m.substring($_, 2), 16)}
            $packet = (,[byte]255 * 6) + ($target * 16)
            $UDPclient = [System.Net.Sockets.UdpClient]::new()
            $UDPclient.Connect($broadcast, $port)
            [void]$UDPclient.Send($packet, 102)
        }
    }
}

class BaseCommand : System.Windows.Forms.ToolStripMenuItem
{
    [Scriptblock]$Script
    static $Commands = [ordered]@{
        "Valitse" = @(
            [BaseCommand]::new("Kaikki", {$global:table.SelectAll()}),
            [BaseCommand]::new("Ei mitään", {$global:table.ClearSelection()})
        )
        "Tietokone" = @(
            [BaseCommand]::new("Käynnistä", {[Host]::Wake()})
            [BaseCommand]::new("Käynnistä uudelleen", {[Host]::Run({shutdown /r}, $false)})
            [BaseCommand]::new("Sammuta", {[Host]::Run({shutdown /s}, $false)})
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
            [BaseCommand]::new("Sulje", {$global:root.Close()})
            [InteractiveCommand]::new("Chrome", "chrome.exe", "", "C:\Program Files (x86)\Google\Chrome\Application")
        )
    } 

    BaseCommand([String]$name, [Scriptblock]$script) : base($name)
    {
        $this.Script = $script
    }

    BaseCommand([String]$name) : base($name){}

    [void] OnClick([System.EventArgs]$e)
    {
        ([System.Windows.Forms.ToolStripMenuItem]$this).OnClick($e)
        & $this.Script
    }

    static [void] Display()
    {
        foreach($category in [BaseCommand]::Commands.keys)
        {
            $menu = [System.Windows.Forms.ToolStripMenuItem]::new()
            $menu.Text = $category
            $global:menubar.Items.Add($menu)
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

    PopUpCommand([String]$name, [Object[]]$widgets, [ScriptBlock]$clickScript, [Scriptblock]$runScript) : base($name)
    {
        $this.Widgets = $widgets
        $this.ClickScript = $clickScript
        $this.RunScript = $runScript
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([System.Windows.Forms.ToolStripMenuItem]$this).OnClick($e)
        $form = [System.Windows.Forms.Form]::new()
        $form.Text = $this.Text
        $form.AutoSize = $true
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
        $button = [RunButton]::new($this, $form)
        & $this.ClickScript
        $form.ShowDialog()
    }

    [void] Run(){ & $this.RunScript }
}

class InteractiveCommand : PopUpCommand
{
    [Object[]]$Widgets = @(
        (New-Object System.Windows.Forms.Label -Property @{
            Text = "Ohjelma:"
            AutoSize = $true
            Anchor = [System.Windows.Forms.AnchorStyles]::Right
        }),
        (New-Object System.Windows.Forms.TextBox -Property @{
            Width = 300
            Anchor = [System.Windows.Forms.AnchorStyles]::Left
        }),
        (New-Object System.Windows.Forms.Label -Property @{
            Text = "Parametri:"
            AutoSize = $true
            Anchor = [System.Windows.Forms.AnchorStyles]::Right
        }),
        (New-Object System.Windows.Forms.TextBox -Property @{
            Width = 300
            Anchor = [System.Windows.Forms.AnchorStyles]::Left
        }),
        (New-Object System.Windows.Forms.Label -Property @{
            Text = "Polku:"
            AutoSize = $true
            Anchor = [System.Windows.Forms.AnchorStyles]::Right
        }),
        (New-Object System.Windows.Forms.TextBox -Property @{
            Width = 300
            Anchor = [System.Windows.Forms.AnchorStyles]::Left
        })
    )
    
    [Scriptblock]$ClickScript = {
        $form.Width = 410
        $form.Height = 175
        $grid = [System.Windows.Forms.TableLayoutPanel]::new()
        $grid.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Inset
        $grid.Location = [System.Drawing.Point]::new(0, 0)
        $grid.AutoSize = $true
        $grid.Padding = [System.Windows.Forms.Padding]::new(10)
        $grid.ColumnCount = 2
        $grid.RowCount = 4
        $grid.Controls.AddRange($this.Widgets)
        $button.Dock = [System.Windows.Forms.DockStyle]::Bottom
        $grid.Controls.Add($button)
        $grid.SetColumnSpan($button, 2)
        $form.Controls.Add($grid)
    }
    [Scriptblock]$RunScript = {
        $executable = ($this.Widgets[1]).Text
        $argument = ($this.Widgets[3]).Text
        $workingDirectory = ($this.Widgets[5]).Text
        Write-Host ([Host]::Run($executable, $argument, $workingDirectory))
    }

    InteractiveCommand([String]$name, [String]$executable, [String]$argument, [String]$workingDirectory) : base($name, $this.Widgets, $this.ClickScript, $this.RunScript)
    {
        ($this.Widgets[1]).Text = $executable
        ($this.Widgets[3]).Text = $argument
        ($this.Widgets[5]).Text = $workingDirectory
    }
}

class RunButton : System.Windows.Forms.Button
{
    [BaseCommand]$Command
    [System.Windows.Forms.Form]$Form

    RunButton([BaseCommand]$command, [System.Windows.Forms.Form]$form)
    {
        $this.Command = $command
        $this.Form = $form
        $this.Text = "Aja"
    }

    [void] OnClick([System.EventArgs]$e)
    {
        ([System.Windows.Forms.Button]$this).OnClick($e)
        $this.Command.Run()
        $this.Form.Close()
    }
}
