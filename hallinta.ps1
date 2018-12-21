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
            $h.pingJob | Wait-Job # Wait for the ping to complete
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
}

class LocalCommand : ToolStripMenuItem
{
    # Basic command that runs specified script onclick
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
    # Runs specified script on remote hosts (WinRM)
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
        Write-Host -NoNewline "Running "
        Write-Host -NoNewline -ForegroundColor Yellow $this.Command
        Write-Host -NoNewline " on "
        Write-Host -ForegroundColor Gray -Separator ", " $hostnames
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
    # Runs specified program interactively on the active session of remote hosts (WinRM)
    [String]$Executable
    [String]$Argument

    InteractiveCommand([String]$name, [String]$executable, [String]$argument) : base($name)
    {
        $this.Executable = $executable
        $this.Argument = $argument
    }

    [void] Run()
    {
        # Creates, runs and removes Windows scheduled task that runs specified program interactively on the locally logged on user
        $hostnames = [Host]::GetActive()
        if ($null -eq $hostnames) { return }
        Write-Host -NoNewline "Running "
        Write-Host -NoNewline -ForegroundColor Yellow $this.Executable, $this.Argument
        Write-Host -NoNewline " on "
        Write-Host -ForegroundColor Gray -Separator ", " $hostnames
        Invoke-Command -ComputerName $hostnames -Credential $script:credential -ArgumentList $this.Executable, $this.Argument -AsJob -ScriptBlock {
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
}

class VBS3Command : InteractiveCommand
{
    # Command with GUI to run VBS3 with specified startup parameters
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

    VBS3Command([String]$name) : base($name, "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\VBS3_64.exe", "")
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

# class OldCopyCommand : LocalCommand
# {
#     # Mirrors files from specified source to remote hosts in parallel (SMB)
#     [String]$Source
#     [String]$Destination
#     [String]$Username
#     [String]$Password

#     OldCopyCommand([String]$name, [String]$source, [String]$destination, [String]$username, [String]$password) : base($name)
#     {
#         $this.Source = $source
#         $this.Destination = $destination
#         $this.Username = $username
#         $this.Password = $password
#     }

#     [void] Run()
#     {
#         $hostnames = [Host]::GetActive()
#         if ($null -eq $hostnames) { return }
#         if($this.Username)
#         {
#             $u = $this.Username
#             $p = $this.Password
#             net.exe use $this.Source /user:$u $p
#             if($LASTEXITCODE -ne 0){ return } # Check that the source is available and the credentials are accepted
#         }
#         else
#         {
#             if(!(Test-Path $this.Source))
#             {
#                 Write-Host ("Cannot find source: {0}" -f $this.Source)
#                 return
#             }    
#         }
#         $description = "Mirroring {0} to {1}" -f $this.Source, $this.Destination
#         Write-Host $description
#         Write-Progress -Activity $description -Status "Starting" -PercentComplete 0
#         $sourceItems = Get-ChildItem $this.Source
#         $jobs = @()
#         $hostnames | ForEach-Object {
#             $jobs += Start-Job -ArgumentList $_, $sourceItems, $this.Destination -ScriptBlock {
#                 param(
#                     [String]$hostname,
#                     [Object[]]$sourceItems,
#                     [String]$destination
#                 )
#                 $session = New-PSSession -ComputerName $hostname
#                 $items = Invoke-Command -Session $session -ArgumentList $sourceItems, $destination -ScriptBlock {
#                     param(
#                         [Object[]]$sourceItems,
#                         [String]$destination
#                     )
#                     $items = New-Object PsObject -Property @{New=@(); Newer=@(); Skip=@(); Remove=@()}
#                     $sourceNames = @()
#                     $sourceItems | ForEach-Object {
#                         $sourceNames += $_.Name
#                         if((Get-Item $destination).PSIsContainer)
#                         {
#                             $destinationItem = Join-Path -Path $destination -ChildPath $_.Name
#                         }
#                         else
#                         {
#                             $destinationItem = $destination    
#                         }
#                         if(Test-Path $destinationItem)
#                         {
#                             $sourceTime = $_.LastWriteTime
#                             $destinationTime = (Get-Item $destinationItem).LastWriteTime
#                             if($sourceTime -gt $destinationTime)
#                             {
#                                 $items.Newer += $_
#                             }
#                             else
#                             {
#                                 $items.Skip += $_
#                             }
#                         }
#                         else
#                         {
#                             $items.New += $_
#                         }
#                     }
#                     Get-ChildItem $destination | ForEach-Object {
#                         if(!($sourceNames.Contains($_.Name)))
#                         {
#                             $items.Remove += $_
#                             Remove-Item $_.FullName
#                         }
#                     }
#                     return $items
#                 }
#                 $neededItems = $items.New + $items.Newer
#                 $totalSize = ($neededItems | Measure-Object -Property Length -Sum).Sum
#                 $processedSize = 0
#                 $neededItems | ForEach-Object {
#                     (100*$processedSize/$totalSize)
#                     if((Get-Item $destination).PSIsContainer)
#                     {
#                         $destinationPath = Join-Path -Path $destination -ChildPath $_.Name
#                     }
#                     else
#                     {
#                         $destinationPath = $destination
#                     }
#                     Copy-Item $_.FullName -Destination $destinationPath -ToSession $session
#                     $processedSize += $_.Length
#                 }
#                 $items | Add-Member @{Hostname=$hostname}
#                 Remove-PSSession $session
#                 $items
#             }
#         }
#         $timer = [Timer]::new()
#         $timer.Interval = 100
#         $jobs | ForEach-Object {$_ | Add-Member @{P=0}}
#         $timer | Add-Member @{Jobs=$jobs; CompletedJobs=@(); Username=$this.Username; Source=$this.Source; Description=$description} # Attach to the timer object so that they are accessible inside the event handler
#         $timer.Add_Tick({
#             $finished = $true
#             $this.Jobs | Where-Object { !($this.CompletedJobs.Contains($_)) } | ForEach-Object {
#                 $output = Receive-Job $_
#                 if($_.State -ne "Completed")
#                 {
#                     $finished = $false
#                     if($output)
#                     {
#                         if($output.GetType() -eq [System.Int32] -or $output.GetType() -eq [System.Double])
#                         {
#                             $_.P = $output
#                         }
#                     }
#                 }
#                 else
#                 {
#                     $_.P = 100
#                     if($output)
#                     {
#                         Write-Host -NoNewline ("Completed {0}: " -f $output.Hostname)
#                         Write-Host -NoNewline -ForegroundColor Green $output.New.Count
#                         Write-Host -NoNewline " new file(s), "
#                         Write-Host -NoNewline -ForegroundColor DarkGreen $output.Newer.Count
#                         Write-Host -NoNewline " newer file(s), "
#                         Write-Host -NoNewline -ForegroundColor Yellow $output.Skip.Count
#                         Write-Host -NoNewline " skipped file(s), "
#                         Write-Host -NoNewline -ForegroundColor Red $output.Remove.Count
#                         Write-Host " removed file(s)"
#                         $this.CompletedJobs += $_
#                     }
#                 }
#             }
#             if(!$finished)
#             {
#                 $average = ($this.Jobs | Measure-Object -Property P -Average).Average
#                 Write-Progress -Activity $this.Description -Status "$average%" -PercentComplete $average
#             }
#             else
#             {
#                 $this.Jobs | Remove-Job
#                 if($this.Username)
#                 {
#                     net.exe use /delete $this.Source
#                 }
#                 Write-Progress -Activity "Sync" -Status "Finished" -Completed
#                 Write-Host "Finished"
#                 $this.Dispose()
#             }
#         })
#         $timer.Start()
#     }
# }
# 
class CopyCommand : InteractiveCommand
{
    CopyCommand([String]$name, [String]$source, [String]$destination, [String]$username, [String]$password, [String]$argument) : base($name, "C:\WINDOWS\System32\cmd.exe", $this.GetParameter($source, $destination, $username, $password, $argument)){}

    [String] GetParameter([String]$source, [String]$destination, [String]$username, [String]$password, [String]$argument)
    {
        if($username)
        {
            return '/c net use "{0}" /user:{2} "{3}" && robocopy "{0}" "{1}" {4} & net use /delete "{0}" & timeout /t 10' -f $source, $destination, $username, $password, $argument
        }
        else
        {
            return '/c robocopy "{0}" "{1}" {2} & timeout /t 10' -f $source, $destination, $argument
        }
    }
}

# Entry point of the program
[Host]::Populate("$PSScriptRoot\luokka.csv", " ")
$script:root = [Form]::new()
$root.Text = "Luokanhallinta v0.16"
$root.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ENV:SystemRoot + "\System32\wksprt.exe")

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
        [RemoteCommand]::new("Käynnistä uudelleen", $true, @(), {shutdown /r /t 10 /c 'Luokanhallinta on ajastanut uudelleen käynnistyksen'})
        [RemoteCommand]::new("Sammuta", $true, @(), {shutdown /s /t 10 /c "Luokanhallinta on ajastanut sammutuksen"})
    )
    "VBS3" = @(
        [VBS3Command]::new("Käynnistä...")
        [CopyCommand]::new("Synkkaa addonit", "\\10.132.0.97\Addons", "%programfiles%\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons", "WORKGROUP\Admin", "kuusteista", "/MIR /XO /NJH")
        [CopyCommand]::new("Synkkaa asetukset", "\\$ENV:COMPUTERNAME\VBS3", "%userprofile%\Documents\VBS3", $(whoami), "", "$ENV:USERNAME.VBS3Profile /NJH")
        [CopyCommand]::new("Synkkaa missionit", "\\$ENV:COMPUTERNAME\mpmissions", "%programfiles%\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mpmissions", $(whoami), "", "/MIR /XO /NJH")
        [RemoteCommand]::new("Sulje", $true, @(), {Stop-Process -ProcessName VBS3_64})
    )
    "SteelBeasts" = @(
        [InteractiveCommand]::new("Käynnistä", "C:\Program Files\eSim Games\SB Pro FI\Release\SBPro64CM.exe", "")
        [RemoteCommand]::new("Sulje", $true, @(), {Stop-Process -ProcessName SBPro64CM})
    )
    "Muu" = @(
        [LocalCommand]::new("Päivitä", {[Host]::Populate("$PSScriptRoot\luokka.csv", " "); [Host]::Display()})
        # [InteractiveCommand]::new("Virus scan?", "C:\Program Files (x86)\F-Secure\Anti-Virus\fsav.exe", "/spyware /system /all /disinf /beep C: D:")
        # [InteractiveCommand]::new("Update F-Secure?", "C:\Program Files (x86)\F-Secure\FSGUI\postinstall.exe", "")
        [LocalCommand]::new("Vaihda käyttäjä...", {$script:credential = Get-Credential -Message "Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin" -UserName $(whoami)})
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
        $menu.DropDownItems.Add($command) | Out-Null # Add command to menu
    }
}
$shares = @{
    "VBS3" = "$ENV:USERPROFILE\Documents\VBS3"
    "mpmissions" = "$ENV:ProgramFiles\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mpmissions"
}
foreach($share in $shares.Keys)
{
    if(!(Get-SmbShare -Name $share -ErrorAction SilentlyContinue))
    {
        New-SmbShare -Name $share -Path $shares[$share] -Description "Luokanhallinta" | Out-Null
    }
}

$username = $(whoami)
$password = ""
# If default credentials are specified, use them instead of getting them from Get-Credential
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
$root.showDialog() | Out-Null
# After the root window has been closed
$shares.Keys | ForEach-Object { Remove-SmbShare -Name $_ -Force }
