<#
.SYNOPSIS
PowerShell 5-compatible wrapper for Invoke-RestMethod with extended support for PowerShell 7 features.

.DESCRIPTION
This function emulates PowerShell 7's Invoke-RestMethod behavior as closely as possible within the constraints of PowerShell 5.
It supports most major options including certificate handling, proxy configuration, custom headers, authentication, and response parsing.

.FEATURE SUPPORT STATUS

Supported:
- Uri                      ✅
- Method                   ✅ (GET, POST, PUT, DELETE, PATCH, HEAD, OPTIONS)
- Headers                  ✅
- Body                     ✅ (JSON serialization supported)
- ContentType              ✅
- Credential               ✅
- UserAgent                ✅
- TimeoutSec               ✅
- SslProtocol              ✅ (Tls, Tls11, Tls12, Tls13 via numeric override)
- SkipCertificateCheck     ✅ (temporary override via callback)
- Proxy                    ✅
- ProxyCredential          ✅
- TransferEncoding         ✅ (chunked supported)
- DisableKeepAlive         ✅
- MaximumRedirection       ✅
- ResponseHeadersVariable  ✅ (via [ref] parameter)
- UseDefaultCredentials    ✅
- Form                     ✅ (application/x-www-form-urlencoded)
- CertificateThumbprint    ✅ (client certificate from CurrentUser store)

Partially Supported:
- Authentication           ⚠️ (Basic supported; OAuth requires manual token handling)
- HttpVersion              ⚠️ (HttpWebRequest limited to HTTP/1.1)
- Multipart/Form-Data      ⚠️ (not natively supported; requires manual boundary construction)

Not Supported:
- SkipHeaderValidation     ❌ (not available in PowerShell 5)
- Custom HTTP handlers     ❌ (HttpClient pipeline not available)

.NOTES
Author: Microsoft Copilot
Compatibility: Windows PowerShell 5.1
Limitations: Some advanced PowerShell 7 features may not be fully reproducible due to .NET Framework constraints.

.EXAMPLE
# Example 1: Basic GET request with default settings

$result = Invoke-RestMethod5 -Uri "https://jsonplaceholder.typicode.com/posts/1"
$result | Format-List

.EXAMPLE
# Example 2: POST request with JSON body and custom headers

$responseHeaders = $null
$result = Invoke-RestMethod5 -Uri "https://jsonplaceholder.typicode.com/posts" `
    -Method POST `
    -Body @{title="test"; body="hello"; userId=99} `
    -Headers @{Authorization="Bearer dummy"} `
    -ContentType "application/json" `
    -UserAgent "CopilotTest/1.0" `
    -TimeoutSec 30 `
    -IgnoreCertificateError `
    -ResponseHeadersVariable ([ref]$responseHeaders)

$result | Format-List
$responseHeaders | Format-List

.EXAMPLE
# Example 3: Form data submission using application/x-www-form-urlencoded

$formData = "name=弘起&game=Factorio"
$result = Invoke-RestMethod5 -Uri "https://httpbin.org/post" `
    -Method POST `
    -Form $formData `
    -ContentType "application/x-www-form-urlencoded"

$result | Format-List

.EXAMPLE
# Example 4: Using a proxy with credentials

$proxyCred = Get-Credential
$result = Invoke-RestMethod5 -Uri "https://jsonplaceholder.typicode.com/posts/1" `
    -Proxy "http://proxy.example.com:8080" `
    -ProxyCredential $proxyCred

$result | Format-List

.EXAMPLE
# Example 5: Using a client certificate by thumbprint

$result = Invoke-RestMethod5 -Uri "https://secure.example.com/api" `
    -CertificateThumbprint "ABCD1234EF567890ABCD1234EF567890ABCD1234"

$result | Format-List

.EXAMPLE
# Example 6: GET request with response headers captured

$responseHeaders = $null
$result = Invoke-RestMethod5 -Uri "https://jsonplaceholder.typicode.com/posts/1" `
    -ResponseHeadersVariable ([ref]$responseHeaders)

$result | Format-List
$responseHeaders | Format-List
#>
function Invoke-RestMethod5 {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Uri,

        [ValidateSet("GET", "POST", "PUT", "DELETE", "PATCH", "HEAD", "OPTIONS")]
        [string]$Method = "GET",

        [hashtable]$Headers = @{},
        [object]$Body,
        [string]$ContentType = "application/json",
        [switch]$IgnoreCertificateError,
        [int]$TimeoutSec = 100,
        [System.Net.ICredentials]$Credential,
        [string]$UserAgent,
        [string]$SslProtocol = "Tls12",
        [string]$Proxy,
        [System.Net.ICredentials]$ProxyCredential,
        [ref]$ResponseHeadersVariable,
        [int]$MaximumRedirection = 5,
        [switch]$DisableKeepAlive,
        [string]$TransferEncoding,
        [switch]$UseDefaultCredentials,
        [string]$Form,
        [string]$CertificateThumbprint
    )

    # SSL/TLS プロトコル設定
    switch ($SslProtocol) {
        "Tls"   { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls }
        "Tls11" { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 }
        "Tls12" { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 }
        "Tls13" { [System.Net.ServicePointManager]::SecurityProtocol = 12288 } # Tls13 は定数未定義
    }

    # 証明書検証の一時無効化
    $callback = $null
    if ($IgnoreCertificateError) {
        $callback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }

    try {
        $request = [System.Net.HttpWebRequest]::Create($Uri)
        $request.Method = $Method
        $request.Timeout = $TimeoutSec * 1000
        $request.AllowAutoRedirect = $MaximumRedirection -gt 0
        $request.MaximumAutomaticRedirections = $MaximumRedirection
        $request.KeepAlive = -not $DisableKeepAlive

        if ($TransferEncoding) {
            $request.SendChunked = $true
            $request.TransferEncoding = $TransferEncoding
        }

        if ($UserAgent) {
            $request.UserAgent = $UserAgent
        }

        if ($UseDefaultCredentials) {
            $request.UseDefaultCredentials = $true
        } elseif ($Credential) {
            $request.Credentials = $Credential
        }

        if ($Proxy) {
            $webProxy = New-Object System.Net.WebProxy($Proxy, $true)
            if ($ProxyCredential) {
                $webProxy.Credentials = $ProxyCredential
            }
            $request.Proxy = $webProxy
        }

        foreach ($key in $Headers.Keys) {
            $request.Headers[$key] = $Headers[$key]
        }

        if ($CertificateThumbprint) {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
            $store.Open("ReadOnly")
            $cert = $store.Certificates | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
            if ($cert) {
                $request.ClientCertificates.Add($cert)
            }
            $store.Close()
        }

        if ($Body -or $Form) {
            if ($Form) {
                $request.ContentType = "application/x-www-form-urlencoded"
                $bodyString = $Form
            } elseif ($ContentType -eq "application/json" -and ($Body -isnot [string])) {
                $request.ContentType = $ContentType
                $bodyString = $Body | ConvertTo-Json -Depth 10
            } else {
                $request.ContentType = $ContentType
                $bodyString = $Body
            }

            $bytes = [System.Text.Encoding]::UTF8.GetBytes($bodyString)
            $request.ContentLength = $bytes.Length
            $stream = $request.GetRequestStream()
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Close()
        }

        $response = $request.GetResponse()
        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
        $result = $reader.ReadToEnd()
        $reader.Close()

        if ($ResponseHeadersVariable) {
            $ResponseHeadersVariable.Value = $response.Headers
        }

        try {
            return $result | ConvertFrom-Json
        } catch {
            return $result
        }
    } catch {
        Write-Warning "Request failed: $($_.Exception.Message)"
        return $null
    } finally {
        if ($IgnoreCertificateError) {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $callback
        }
    }
}
#
Export-ModuleMember -Function Invoke-RestMethod5
