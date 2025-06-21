<#
.SYNOPSIS
    PowerShell 7's ConvertTo-Csv functionality wrapper for PowerShell 5.1. (PS7-Compatibility Enhanced)
    PowerShell 5.1 の ConvertTo-Csv を PowerShell 7 相当の機能にするためのラッパー関数です。(PS7互換強化版)

.DESCRIPTION
    This function wraps the standard ConvertTo-Csv cmdlet in PowerShell 5.1 to provide features available in PowerShell 7.
    This version handles all quoting modes manually to ensure full compatibility with PowerShell 7's behavior, especially regarding newline characters in 'Always' quote mode.

    この関数は、PowerShell 5.1 の ConvertTo-Csv をラップし、PowerShell 7で利用可能な機能を提供します。
    このバージョンでは、すべての引用モードを手動で処理することで、特に'Always'モードでの改行文字の扱いに関して、PowerShell 7の挙動との完全な互換性を確保します。
#>
function ConvertTo-Csv7 {
    [CmdletBinding(DefaultParameterSetName = 'Delimiter')]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject[]]$InputObject,

        [Parameter()]
        [ValidateSet('AsNeeded', 'Always', 'Never')]
        [string]$UseQuotes = 'AsNeeded',

        [Parameter()]
        [switch]$NoHeader,

        [Parameter()]
        [switch]$UseLf,

        [Parameter(ParameterSetName = 'Delimiter')]
        [string]$Delimiter = ',',

        [Parameter(ParameterSetName = 'UseCulture')]
        [switch]$UseCulture
    )
    begin {
        $list = [System.Collections.Generic.List[psobject]]::new()
    }
    process {
        $list.AddRange($InputObject)
    }
    end {
        if ($list.Count -eq 0) { return }

        $d = if ($UseCulture) { (Get-Culture).TextInfo.ListSeparator } else { $Delimiter }
        $lines = [System.Collections.Generic.List[string]]::new()
        
        # PS7との互換性を確保するため、すべてのモードを手動で構築します
        $headers = $list[0].PSObject.Properties.Name
        if (-not $NoHeader) {
            # ヘッダーは常に引用符で囲まれます
            $lines.Add(($headers | ForEach-Object { "`"$($_ -replace '`"','`"`"')`"" } | Join-String -Separator $d))
        }

        foreach ($item in $list) {
            $row = $headers | ForEach-Object {
                $val = $item.$_
                $str = if ($null -eq $val) { '' } else { $val.ToString() }
                
                $quoteIt = $false
                if ($UseQuotes -eq 'Always') {
                    $quoteIt = $true
                } 
                elseif ($UseQuotes -eq 'AsNeeded') {
                    if ($str -match "[$([regex]::Escape($d))\""`r`n]") {
                        $quoteIt = $true
                    }
                }
                # 'Never' の場合、$quoteItはfalseのままです

                if ($quoteIt) {
                    "`"$($str -replace '`"','`"`"')`""
                } else {
                    $str
                }
            }
            $lines.Add(($row | Join-String -Separator $d))
        }
        
        # 指定された改行コードで最終的な文字列を生成します
        $lines -join ($UseLf.IsPresent ? "`n" : "`r`n")
    }
}
