# このテストスクリプトを実行する前に、上記のConvertTo-Csv7関数を読み込んでください。
# 例: . .\ConvertTo-Csv7.ps1

# Pester v5 構文を使用
using namespace System.Management.Automation
using namespace System.Collections.Generic

BeforeAll {
    # テスト対象の関数をここで定義することで、テストスクリプトを自己完結させます
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
        if ($UseQuotes -eq 'Always') {
            $params = @{ NoTypeInformation = $true; InputObject = $list }
            if ($PSBoundParameters.ContainsKey('Delimiter')) { $params.Delimiter = $d }
            if ($PSBoundParameters.ContainsKey('UseCulture')) { $params.UseCulture = $true }
            $result = Microsoft.PowerShell.Utility\ConvertTo-Csv @params
            $lines.AddRange(($result | Select-Object -Skip ($NoHeader.IsPresent ? 1 : 0)))
        } else {
            $headers = $list[0].PSObject.Properties.Name
            if (-not $NoHeader) {
                $lines.Add(($headers | ForEach-Object { "`"$($_ -replace '`"','`"`"')`"" } | Join-String -Separator $d))
            }
            foreach ($item in $list) {
                $row = $headers | ForEach-Object {
                    $val = $item.$_
                    $str = if ($null -eq $val) { '' } else { $val.ToString() }
                    if (($UseQuotes -eq 'AsNeeded') -and ($str -match "[$([regex]::Escape($d))\""`r`n]")) {
                        "`"$($str -replace '`"','`"`"')`""
                    } else {
                        $str
                    }
                }
                $lines.Add(($row | Join-String -Separator $d))
            }
        }
        $lines -join ($UseLf.IsPresent ? "`n" : "`r`n")
    }}

    # テストデータを$scriptスコープで定義
    $script:TestData = @(
        [pscustomobject]@{ ID = 1; Name = 'Alice'; Notes = 'Simple' }
        [pscustomobject]@{ ID = 2; Name = 'Bob, Smith'; Notes = 'Has comma' }
        [pscustomobject]@{ ID = 3; Name = 'Charles'; Notes = 'Has "quotes"' }
        [pscustomobject]@{ ID = 4; Name = 'David'; Notes = "Line1`r`nLine2" }
        [pscustomobject]@{ ID = 5; Name = 'Eve'; Notes = $null }
    )
}

Describe 'ConvertTo-Csv7 Function Tests' {
    Context 'Default Behavior (-UseQuotes AsNeeded)' {
        It 'should produce CSV with headers and CRLF, quoting only when necessary' {
            $result = $script:TestData | ConvertTo-Csv7
            $expected = @(
                '"ID","Name","Notes"'
                '1,Alice,Simple'
                '2,"Bob, Smith","Has comma"'
                '3,Charles,"Has ""quotes"""'
                '4,David,"Line1`r`nLine2"'
                '5,Eve,'
            ) -join "`r`n"
            $result | Should -Be $expected
        }
    }
    
    Context '-UseQuotes Parameter' {
        It 'should quote all string fields with -UseQuotes Always' {
            $result = $script:TestData | ConvertTo-Csv7 -UseQuotes Always
            # PS5.1のネイティブな振る舞い（改行がスペースになる）をテストする
            $expected = @(
                '"ID","Name","Notes"'
                '"1","Alice","Simple"'
                '"2","Bob, Smith","Has comma"'
                '"3","Charles","Has ""quotes"""'
                '"4","David","Line1 Line2"' # ネイティブ実装では改行はスペースに置換
                '"5","Eve",""'
            ) -join "`r`n"
            $result | Should -Be $expected
        }

        It 'should never quote fields with -UseQuotes Never' {
            # ヘッダーは常に引用符で囲まれるが、データ行は囲まれないことを確認
            $result = $script:TestData | ConvertTo-Csv7 -UseQuotes Never
            $expected = @(
                '"ID","Name","Notes"'
                '1,Alice,Simple'
                '2,Bob, Smith,Has comma' # カンマが含まれていても引用符なし（不正なCSVになる可能性）
                '3,Charles,Has "quotes"'
                "4,David,Line1`r`nLine2"
                '5,Eve,'
            ) -join "`r`n"
            $result | Should -Be $expected
        }
    }

    Context '-NoHeader Switch' {
        It 'should omit the header when -NoHeader is specified with AsNeeded' {
             $result = $script:TestData[0] | ConvertTo-Csv7 -NoHeader
             $expected = '1,Alice,Simple'
             $result | Should -Be $expected
        }

        It 'should omit the header when -NoHeader is specified with Always' {
            $result = $script:TestData[0] | ConvertTo-Csv7 -NoHeader -UseQuotes Always
            $expected = '"1","Alice","Simple"'
            $result | Should -Be $expected
        }
    }

    Context '-UseLf Switch' {
        It 'should use LF as the line separator' {
            $result = $script:TestData[0..1] | ConvertTo-Csv7 -UseLf
            $expected = @(
                '"ID","Name","Notes"'
                '1,Alice,Simple'
                '2,"Bob, Smith","Has comma"'
            ) -join "`n"
            $result | Should -Be $expected
        }
    }

    Context '-Delimiter Parameter' {
        It 'should use a custom semicolon delimiter' {
            $result = $script:TestData[0] | ConvertTo-Csv7 -Delimiter ';'
            $expected = @('"ID";"Name";"Notes"', '1;Alice;Simple') -join "`r`n"
            $result | Should -Be $expected
        }

        It 'should quote fields containing the custom delimiter' {
            $data = [pscustomobject]@{Name='A;B';Value=1}
            $result = $data | ConvertTo-Csv7 -Delimiter ';'
            $expected = @('"Name";"Value"', '"A;B";1') -join "`r`n"
            $result | Should -Be $expected
        }
    }

    Context 'Edge Cases' {
        It 'should return an empty string for empty pipeline input' {
            $result = @() | ConvertTo-Csv7
            $result | Should -BeNullOrEmpty
        }

        It 'should handle a single object correctly' {
            $result = $script:TestData[0] | ConvertTo-Csv7
            $expected = @('"ID","Name","Notes"', '1,Alice,Simple') -join "`r`n"
            $result | Should -Be $expected
        }
    }
}
