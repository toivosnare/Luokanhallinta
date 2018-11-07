Add-Type -assembly System.Windows.Forms
. $PSScriptRoot\hallinta.ps1

[Host]::Populate("luokka.csv", " ")
$global:root = [System.Windows.Forms.Form]::new()
$root.Text = "Luokanhallinta"
$root.Width = 1280
$root.Height = 720

$global:table = [System.Windows.Forms.DataGridView]::new()
$table.Dock = [System.Windows.Forms.DockStyle]::Fill
$table.AllowUserToAddRows = $false
$table.AllowUserToDeleteRows = $false
$table.AllowUserToResizeColumns = $false
$table.AllowUserToResizeRows = $false
$table.AllowUserToOrderColumns = $false
$table.ReadOnly = $true
$table.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
$table.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$table.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
$table.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
($table.RowsDefaultCellStyle).Padding = [System.Windows.Forms.Padding]::new(30)
($table.RowsDefaultCellStyle).ForeColor = [System.Drawing.Color]::Red
($table.RowsDefaultCellStyle).SelectionForeColor = [System.Drawing.Color]::Red
($table.RowsDefaultCellStyle).SelectionBackColor = [System.Drawing.Color]::LightGray
$table.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
$table.ColumnCount = ([Host]::Hosts | ForEach-Object {$_.Column} | Measure-Object -Maximum).Maximum
$table.RowCount = ([Host]::Hosts | ForEach-Object {$_.Row} | Measure-Object -Maximum).Maximum
$table.Columns | ForEach-Object {
    $_.HeaderText = [Char]($_.Index + 65)
    $_.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
}
$table.Rows | ForEach-Object {$_.HeaderCell.Value = [String]($_.Index + 1)}
$root.Controls.Add($table)
[Host]::Display()

$global:menubar = [System.Windows.Forms.MenuStrip]::new()
$root.MainMenuStrip = $menubar
$menubar.Dock = [System.Windows.Forms.DockStyle]::Top
$root.Controls.Add($menubar)
[BaseCommand]::Display()

$root.showDialog()