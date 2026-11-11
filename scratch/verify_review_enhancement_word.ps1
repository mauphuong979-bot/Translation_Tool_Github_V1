param([string]$SampleDirectory = (Join-Path $PSScriptRoot 'review_enhancement_qa'))

$ErrorActionPreference = 'Stop'
$directory = [IO.Path]::GetFullPath($SampleDirectory)
Add-Type -ReferencedAssemblies ([Management.Automation.PSObject].Assembly.Location) -TypeDefinition @'
using System;
using System.Reflection;
using System.Management.Automation;
public static class ReviewWordInterop {
    public static object Invoke(object target, string name, object[] arguments) {
        var values = new object[arguments.Length];
        for (int index = 0; index < arguments.Length; index++) {
            var wrapped = arguments[index] as PSObject;
            values[index] = wrapped == null ? arguments[index] : wrapped.BaseObject;
        }
        var wrappedTarget = target as PSObject;
        if (wrappedTarget != null) target = wrappedTarget.BaseObject;
        return target.GetType().InvokeMember(name,
            BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, target, values);
    }
}
'@

function Invoke-WordMethod($target, $name, [object[]]$arguments) {
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        try {
            return [ReviewWordInterop]::Invoke($target, $name, $arguments)
        } catch {
            if ($_.Exception.ToString() -notmatch '80010001' -or $attempt -eq 29) { throw "Word method $name failed: $($_.Exception.Message)" }
            Start-Sleep -Milliseconds 500
        }
    }
}

$word = $null
$owned = $false
try {
    $word = New-Object -ComObject Word.Application
    if ($word.Documents.Count -ne 0) { throw 'Refusing to use a Word instance with existing documents.' }
    $owned = $true
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $worksheetPath = Join-Path $directory '02_Worksheet.docx'
    $worksheet = $word.Documents.Open([ref]([object]$worksheetPath))
    Invoke-WordMethod $worksheet 'ExportAsFixedFormat' @((Join-Path $directory '02_Worksheet.pdf'), 17)
    Invoke-WordMethod $worksheet 'Close' @(0)
    foreach ($action in @('Accept', 'Reject')) {
        $reportPath = Join-Path $directory '04_Tracked_Report.docx'
        $document = $word.Documents.Open([ref]([object]$reportPath))
        $before = $document.Revisions.Count
        Invoke-WordMethod $document ($action + 'AllRevisions') @()
        if ($document.Revisions.Count -ne 0) { throw "Revisions remain after $action" }
        $stem = '05_' + $action + 'ed_Report'
        Invoke-WordMethod $document 'SaveAs2' @((Join-Path $directory ($stem + '.docx')), 16)
        if ($action -eq 'Accept') {
            Invoke-WordMethod $document 'ExportAsFixedFormat' @((Join-Path $directory ($stem + '.pdf')), 17)
        }
        Invoke-WordMethod $document 'Close' @(0)
        Write-Output "$action verified: $before revisions resolved"
    }
} finally {
    if ($word -and $owned) { Invoke-WordMethod $word 'Quit' @(0) }
    if ($word) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
}
