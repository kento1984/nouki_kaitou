# ファイル選択ダイアログを表示
Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "10PM.XLSを選択してください"
$dialog.Filter = "Excelファイル (*.xls;*.xlsx)|*.xls;*.xlsx|すべてのファイル (*.*)|*.*"
$dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dialog.FileName
} else {
    exit 1
}
