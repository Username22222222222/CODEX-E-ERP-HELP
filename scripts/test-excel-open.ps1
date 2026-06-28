param(
    [string[]]$Paths = @(
        "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
        "C:\Users\micha\Documents\X-ERP-HELP\backups\X-ERP-HELP.before-view-sheets-20260624-235519.xlsx",
        "C:\Users\micha\Documents\X-ERP-HELP\backups\X-ERP-HELP.before-outline-repair-20260624-234749.xlsx"
    )
)

$ErrorActionPreference = 'Continue'
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    foreach ($path in $Paths) {
        $ok = $false
        $message = ''
        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($path, 0, $true)
            $ok = $true
            $message = "opened"
        }
        catch {
            $message = $_.Exception.Message
        }
        finally {
            if ($wb -ne $null) {
                $wb.Close($false)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            }
        }
        [pscustomobject]@{
            Path = $path
            ExcelOpenOk = $ok
            Message = $message
        }
    }
}
finally {
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
