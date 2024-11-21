
# Generic function to avoid code duplication
<#
.SYNOPSIS
Sets properties on an input object based on a configuration JSON file.

.DESCRIPTION
The `Set-SC365PropertiesFromConfigJson` function takes an input object and updates its properties using values from a provided JSON object.
It handles general properties, as well as properties specific to routing, options, and regions, ensuring configuration consistency across components.

.PARAMETER InputObject
The object whose properties will be updated based on the JSON configuration.

.PARAMETER Json
A JSON object containing the configuration settings. The function expects the JSON to include general properties and optionally, 
nested `Routing`, `Option`, or `Region` sections.

.PARAMETER Routing
An optional parameter specifying the routing type. If provided, the function will apply settings from the corresponding 
`Routing` section in the JSON object.

.PARAMETER Option
An optional array of configuration options. The function applies properties from the `Option` section in the JSON 
for the specified options.

.PARAMETER Region
An optional parameter specifying the geographic region. If provided, the function applies settings from the 
`Region` section in the JSON object.

.OUTPUTS
None
The function modifies the `InputObject` in place.

.EXAMPLE
PS> $config = Get-Content "config.json" | ConvertFrom-Json
PS> $obj = New-Object PSObject
PS> Set-SC365PropertiesFromConfigJson -InputObject $obj -Json $config -Routing "Hybrid" -Option @("Option1", "Option2") -Region "NorthAmerica"

This example updates the properties of `$obj` using the configuration in `config.json`, applying settings for the `Hybrid` routing type, 
`Option1` and `Option2`, and the `NorthAmerica` region.

.NOTES
- This function assumes the JSON structure includes general properties and optionally, nested sections for `Routing`, `Option`, and `Region`.
- The function skips certain predefined keys (`Name`, `Option`, `Routing`, `Region`) for general property updates.

.REQUIREMENTS
- PowerShell 5.1 or later.
- A properly structured JSON configuration file.

.LINK
For information on managing JSON data in PowerShell, see:
https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/json

#>
function Set-SC365PropertiesFromConfigJson
{
    [CmdLetBinding()]
    Param
    (
        [psobject] $InputObject,
        [psobject] $Json,
        [SC365.MailRouting] $Routing,
        [SC365.ConfigOption[]] $Option,
        [SC365.GeoRegion] $Region
    )

    # Set all properties that aren't version specific
    $json.psobject.properties | Foreach-Object {
        if ($_.Name -notin @("Name", "Option", "Routing", "Region"))
        { $InputObject.$($_.Name) = $_.Value }
    }

    if($routing -and $json.Routing)
    {
        $json.Routing.$Routing.psobject.properties | Foreach-Object {
            $InputObject.$($_.Name) = $_.Value
        }
    }

    if($Option -and $json.Option)
    {
        $Option | Where-Object {$json.Option.$_} | ForEach-Object{
            $Json.Option.$_.psobject.properties | ForEach-Object{
                $InputObject.$($_.Name) = $_.Value
            }
        }
    }


    if($Region -and $json.Region)
    {
        $json.Region.$Region.psobject.properties | %Foreach-Object {
            $InputObject.$($_.Name) = $_.Value
        }
    }
}

<#
.SYNOPSIS
Retrieves inbound connector settings for a specified routing type.

.DESCRIPTION
The `Get-SC365InboundConnectorSettings` function reads inbound connector settings from a JSON configuration file 
and retrieves the settings specific to the provided routing type. The function ensures case-insensitivity for 
routing type lookups and returns the relevant configuration as a hashtable.

.PARAMETER Routing
Specifies the routing type for which inbound connector settings are to be retrieved. 
This parameter is mandatory.

.PARAMETER Option
Specifies additional options for the function. This parameter is optional.

.OUTPUTS
Hashtable
A hashtable containing the inbound connector settings for the specified routing type.

.EXAMPLE
PS> Get-SC365InboundConnectorSettings -Routing "HybridRouting"

This example retrieves the inbound connector settings for the routing type `HybridRouting`.

.EXAMPLE
PS> Get-SC365InboundConnectorSettings -Routing "DirectRouting" -Option $customOption

This example retrieves the inbound connector settings for the routing type `DirectRouting`, 
taking into account additional custom options provided via the `Option` parameter.

.NOTES
- The function reads the JSON file `InBound.json` located in the `ExOConfig\Connectors\` directory relative to the script root.
- Ensure the JSON file is properly formatted and includes a `routing` section with routing-specific settings.
- The `ToLower()` method ensures case-insensitive lookup for the routing type.

.REQUIREMENTS
- PowerShell 5.1 or later.
- A valid `InBound.json` file located in the `ExOConfig\Connectors\` directory.

.LINK
For more information on configuring inbound connectors in Exchange Online, visit:
https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow
#>
function Get-SC365InboundConnectorSettings
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $routing,
        $option
    )

    Write-Verbose "Loading inbound connector settings for routingtype $Routing"
    $inBoundRaw = (Get-Content "$PSScriptRoot\..\ExOConfig\Connectors\InBound.json" -Raw|Convertfrom-Json -AsHashtable)
    $ret = $inBoundRaw.routing.($routing.Tolower())

    return $ret
}
<#
.SYNOPSIS
Retrieves outbound connector settings for a specified routing type.

.DESCRIPTION
The `Get-SC365OutboundConnectorSettings` function loads and returns outbound connector configuration settings
from a JSON file based on the specified routing type. It parses the JSON file and retrieves the relevant settings 
for the provided routing option, ensuring case-insensitivity for routing type lookup.

.PARAMETER Routing
Specifies the routing type for which outbound connector settings are to be retrieved. 
This parameter is mandatory and must be provided.

.PARAMETER Option
Specifies additional configuration options that may influence the settings retrieval.
This parameter is optional.

.OUTPUTS
Hashtable
A hashtable containing the outbound connector settings for the specified routing type.

.EXAMPLE
PS> Get-SC365OutboundConnectorSettings -Routing "HybridRouting"

This example retrieves the outbound connector settings for the routing type `HybridRouting`.

.EXAMPLE
PS> Get-SC365OutboundConnectorSettings -Routing "DirectRouting" -Option $customOption

This example retrieves the outbound connector settings for the routing type `DirectRouting`, 
taking into account the additional custom option.

.NOTES
- The function reads outbound connector settings from the `OutBound.json` file located in the 
`$PSScriptRoot\..\ExOConfig\Connectors\` directory.
- The JSON file must be properly formatted and include a `routing` section with routing-specific settings.
- The `ToLower()` method ensures case-insensitive lookup for the routing type.

.REQUIREMENTS
- PowerShell 5.1 or later.
- A valid `OutBound.json` file located in the `ExOConfig\Connectors\` directory.

.LINK
For more information on Exchange Online connectors, visit:
https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow

#>
function Get-SC365OutboundConnectorSettings
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $routing,
        $option
    )

    Write-Verbose "Loading outbound connector settings"
    $outBoundRaw = (Get-Content "$PSScriptRoot\..\ExOConfig\Connectors\OutBound.json" -Raw|Convertfrom-Json -AsHashtable)
    $ret= $outBoundRaw.routing.($routing.ToLower())
    return $ret
}

function Get-SC365TransportRuleSettings
{
    [CmdLetBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $routing,
        [Parameter(Mandatory = $true)]
        [string] $file,
        [switch] $IncludeSkipped
    )

    begin {
        $ret = $null
        $raw = $null
    }
    process {
        $raw = (Get-Content $File -Raw|Convertfrom-Json -AsHashtable)
        $ret = $raw.routing.($routing.ToLower())
    }
    end {
        return $ret    
    }
}
<#
.SYNOPSIS
Retrieves the cloud configuration for a specified geographic region.

.DESCRIPTION
The `Get-SC365CloudConfig` function loads and returns cloud configuration settings for a specific region by reading and parsing a JSON file. 
The function retrieves the relevant configuration based on the provided region and ensures that the region name is case-insensitive.

.PARAMETER Region
The geographic region for which the cloud configuration is to be retrieved. 
This parameter is mandatory and must be provided as a string.

.OUTPUTS
PSObject
An object containing the cloud configuration settings for the specified region.

.EXAMPLE
PS> Get-SC365CloudConfig -Region "ch"

This example retrieves the cloud configuration settings for the `NorthAmerica` region.

.EXAMPLE
PS> Get-SC365CloudConfig -Region "eu"

This example retrieves the cloud configuration settings for the `Europe` region. The region name is not case-sensitive.

.NOTES
- The function reads from a file named `GeoRegion.json` located in the `CloudConfig` directory relative to the script root.
- Ensure that the JSON file is properly formatted and includes a `GeoRegion` section with region-specific configuration.

.REQUIREMENTS
- PowerShell 5.1 or later.
- A valid `GeoRegion.json` file located at `$PSScriptRoot\..\ExOConfig\CloudConfig\GeoRegion.json`.

#>
function Get-SC365CloudConfig
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [String]$Region
    )

    Write-Verbose "Loading inbound connector settings for region $Region"

    $ret = (ConvertFrom-Json (Get-Content -Path "$PSScriptRoot\..\ExOConfig\CloudConfig\GeoRegion.json" -Raw)).GeoRegion.($region.ToLower())
    return $ret
}
<#
.SYNOPSIS
Converts a raw numeric value into a human-readable file size format (kB, MB, GB, TB).

.DESCRIPTION
The `Convertto-SC365Numberformat` function takes an integer input representing a raw numeric value (e.g., bytes)
and converts it into a human-readable format, such as kilobytes (kB), megabytes (MB), gigabytes (GB), or terabytes (TB),
based on the size of the input number.

.PARAMETER RawNumber
An integer value representing the size to be converted. The value is categorized into the appropriate unit 
based on its length and then formatted to two decimal places.

.OUTPUTS
String
A formatted string indicating the size in the appropriate unit (e.g., "1.23 MB").

.EXAMPLE
PS> Convertto-SC365Numberformat -RawNumber 123456

This example converts 123,456 into the string "120.56 kB".

.EXAMPLE
PS> Convertto-SC365Numberformat -RawNumber 1234567890

This example converts 1,234,567,890 into the string "1.15 GB".

.NOTES
- The function uses the `switch` statement to determine the appropriate size category.
- The thresholds for determining the unit are based on the number of digits in the input number.

#>
function Convertto-SC365Numberformat 
{
    param (
        [Int64]$rawnumber
    )
    $ConvertedNumber = switch ($rawNumber.ToString().Length) {
                           {($_ -le 5)} {($rawNumber/1KB).ToString("N2") + " kB"} 
         {(($_ -gt 5) -and ($_ -le 9))} {($rawNumber/1MB).ToString("N2") + " MB"} 
        {(($_ -gt 9) -and ($_ -le 12))} {($rawNumber/1GB).ToString("N2") + " GB"} 
                          {($_ -gt 12)} {($rawNumber/1TB).ToString("N2") + " TB"} 
    }
    return $ConvertedNumber
}
<#
.SYNOPSIS
Generates a hash value from a given input string using a specified hashing algorithm.

.DESCRIPTION
The `Get-SC365StringHash` function computes a hash value for an input string using the specified algorithm. 
It supports common hashing algorithms such as MD5, RIPEMD160, SHA1, SHA256, SHA384, and SHA512. 
The function outputs the computed hash as a hexadecimal string.

.PARAMETER String
The input string to be hashed. This parameter is mandatory and can accept input from the pipeline.

.PARAMETER HashName
Specifies the hashing algorithm to use. 
Valid values are "MD5", "RIPEMD160", "SHA1", "SHA256", "SHA384", and "SHA512". 
The default is "SHA1" if no value is specified.

.OUTPUTS
String
The function returns a hexadecimal string representing the hash of the input.

.EXAMPLE
PS> Get-SC365StringHash -String "HelloWorld"

This example computes the SHA1 hash of the string "HelloWorld" and outputs the hash value.

.EXAMPLE
PS> "MyString" | Get-SC365StringHash -HashName SHA256

This example pipes the string "MyString" to the function and computes its SHA256 hash.

.NOTES
- This function uses the .NET `System.Security.Cryptography` library to perform hashing.
- Ensure the input string is not null or empty to avoid errors.

#>
Function Get-SC365StringHash {
    [cmdletbinding()]
    [OutputType([String])]
    param(
      [parameter(ValueFromPipeline, Mandatory = $true, Position = 0)]
      [String]$String,
      
      [parameter(ValueFromPipelineByPropertyName, Mandatory = $false, Position = 1)]
      [ValidateSet("MD5", "RIPEMD160", "SHA1", "SHA256", "SHA384", "SHA512")]
      [String]$HashName = 'SHA1'
    )
    begin {
  
    }
    Process {
      $StringBuilder = New-Object System.Text.StringBuilder
      [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($String))| foreach-object {
      [Void]$StringBuilder.Append($_.ToString("x2"))
      }
      $output = $StringBuilder.ToString()
    }
    end {
      return $output
    }
}

<#
.SYNOPSIS
Removes domains with the `.onmicrosoft.com` suffix from a given list of domains.

.DESCRIPTION
The `Remove-SC365OnMicrosoftDomain` function filters out all domains in the input `DomainList` that have the `.onmicrosoft.com` suffix. 
It returns a new list containing only the domains that do not match this suffix. This is useful for cleaning up domain lists when
you want to exclude temporary or system-generated Microsoft 365 domains.

.PARAMETER DomainList
A mandatory parameter that accepts an array list of domain names to process. 
The list must be provided as a `[System.Collections.ArrayList]` object.

.RETURNS
[System.Collections.ArrayList]
A new array list containing only the domains that do not have the `.onmicrosoft.com` suffix.

.EXAMPLE
PS> $domains = [System.Collections.ArrayList]@('example.com', 'tenant.onmicrosoft.com', 'anotherdomain.com')
PS> $filteredDomains = Remove-SC365OnMicrosoftDomain -DomainList $domains
PS> $filteredDomains

This example filters out the `tenant.onmicrosoft.com` domain, returning:
example.com
anotherdomain.com

.EXAMPLE
PS> $domains = [System.Collections.ArrayList]@('domain1.com', 'domain2.onmicrosoft.com')
PS> Remove-SC365OnMicrosoftDomain -DomainList $domains

This example returns only `domain1.com`.

.NOTES
- The function uses the `-NotLike` operator to filter out `.onmicrosoft.com` domains.
- The function returns a new array list while leaving the original `DomainList` unchanged.

.REQUIREMENTS
- PowerShell 5.1 or later.

.LINK
For more information on Microsoft 365 domains, see:
https://learn.microsoft.com/en-us/microsoft-365/

#>
Function Remove-SC365OnMicrosoftDomain {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.ArrayList]$DomainList
    )
    [System.Collections.ArrayList]$NewDomainList= @()
    Foreach ($domain in $DomainList) {
        if ($domain -Notlike '*.onmicrosoft.com') {
                [void]$NewDomainList.Add($Domain)     
        }
    }
    return $NewDomainList    
}

#Beginning with v 1.4.0 this function is obsolete
<#function Get-ExoHTMLData {
    param (
        [Parameter(
              Mandatory = $true,
            HelpMessage = 'Enter Cmdlte to ')]
        [string]$ExoCmd
    )
    try {
        $allCmd = $exoCmd.Split('|')[0].Trim()
        $htmlSelectCmd = $exoCmd.Split('|')[1].Trim()

        $rawData = Invoke-Expression -Command $allCmd
        if ($null -eq $rawData) {
            $ExoHTMLData = New-object -type PSobject -property @{Result = '--- no information available ---'}|Convertto-HTML -Fragment
        } else {
            $ExoHTMLCmd = "{0}|{1}" -f  $allcmd,$htmlSelectCmd
            $ExoHTMLData = Invoke-expression -Command $ExoHTMLCmd |Convertto-HTML -Fragment
            if ($jsonBackup) {
                $script:JsonData += '---' + $AllCmd + '---'|Convertto-Json
                $script:JsonData += $rawData|ConvertTo-Json
            }
        } 
        return $ExoHTMLData
    }
    catch {
        Write-Warning "Could not fetch data from command '$exoCmd'"
    }    
}#>

<#
.SYNOPSIS
Generates a unique report filename based on the current time and the default domain name.

.DESCRIPTION
The `New-SelfGeneratedReportName` function creates a unique, self-generated report filename. 
The filename includes the current time in `HHm-ddMMyyy` format and the default email domain name, 
retrieved from the output of the `Get-AcceptedDomain` cmdlet. The filename is appended with `.html`.

.EXAMPLE
PS> New-SelfGeneratedReportName

This will return a string similar to `1507-10112024defaultdomain.com.html`, where:
- `1507` represents the current time in hours and minutes.
- `10112024` represents the date in `ddMMyyyy` format.
- `defaultdomain.com` is the default email domain.

.PARAMETER None
The function does not accept parameters.

.RETURNS
String
A string representing the self-generated filename.

.NOTES
- Ensure the `Get-AcceptedDomain` cmdlet is available and provides a `default` property to identify the default domain.
- The function requires the `-ExpandProperty` flag in `Select-Object` to retrieve the `Domainname` property.

.REQUIREMENTS
- PowerShell 5.1 or later
- Exchange Online PowerShell module or other modules providing the `Get-AcceptedDomain` cmdlet.
#>
function New-SelfGeneratedReportName {
    Write-Verbose "Creating self-generated report filename."
    return ("{0:HHm-ddMMyyy}" -f (Get-Date)) + (Get-AcceptedDomain|where-object default -eq $true|select-object -expandproperty Domainname) + '.html'
}

# SIG # Begin signature block
# MIIVzAYJKoZIhvcNAQcCoIIVvTCCFbkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCARsZw+Jo5ecDqC
# GWb0fd0k6rFExMG7yEPbumYo37XinqCCEggwggVvMIIEV6ADAgECAhBI/JO0YFWU
# jTanyYqJ1pQWMA0GCSqGSIb3DQEBDAUAMHsxCzAJBgNVBAYTAkdCMRswGQYDVQQI
# DBJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcMB1NhbGZvcmQxGjAYBgNVBAoM
# EUNvbW9kbyBDQSBMaW1pdGVkMSEwHwYDVQQDDBhBQUEgQ2VydGlmaWNhdGUgU2Vy
# dmljZXMwHhcNMjEwNTI1MDAwMDAwWhcNMjgxMjMxMjM1OTU5WjBWMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS0wKwYDVQQDEyRTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQCN55QSIgQkdC7/FiMCkoq2rjaFrEfUI5ErPtx94jGgUW+s
# hJHjUoq14pbe0IdjJImK/+8Skzt9u7aKvb0Ffyeba2XTpQxpsbxJOZrxbW6q5KCD
# J9qaDStQ6Utbs7hkNqR+Sj2pcaths3OzPAsM79szV+W+NDfjlxtd/R8SPYIDdub7
# P2bSlDFp+m2zNKzBenjcklDyZMeqLQSrw2rq4C+np9xu1+j/2iGrQL+57g2extme
# me/G3h+pDHazJyCh1rr9gOcB0u/rgimVcI3/uxXP/tEPNqIuTzKQdEZrRzUTdwUz
# T2MuuC3hv2WnBGsY2HH6zAjybYmZELGt2z4s5KoYsMYHAXVn3m3pY2MeNn9pib6q
# RT5uWl+PoVvLnTCGMOgDs0DGDQ84zWeoU4j6uDBl+m/H5x2xg3RpPqzEaDux5mcz
# mrYI4IAFSEDu9oJkRqj1c7AGlfJsZZ+/VVscnFcax3hGfHCqlBuCF6yH6bbJDoEc
# QNYWFyn8XJwYK+pF9e+91WdPKF4F7pBMeufG9ND8+s0+MkYTIDaKBOq3qgdGnA2T
# OglmmVhcKaO5DKYwODzQRjY1fJy67sPV+Qp2+n4FG0DKkjXp1XrRtX8ArqmQqsV/
# AZwQsRb8zG4Y3G9i/qZQp7h7uJ0VP/4gDHXIIloTlRmQAOka1cKG8eOO7F/05QID
# AQABo4IBEjCCAQ4wHwYDVR0jBBgwFoAUoBEKIz6W8Qfs4q8p74Klf9AwpLQwHQYD
# VR0OBBYEFDLrkpr/NZZILyhAQnAgNpFcF4XmMA4GA1UdDwEB/wQEAwIBhjAPBgNV
# HRMBAf8EBTADAQH/MBMGA1UdJQQMMAoGCCsGAQUFBwMDMBsGA1UdIAQUMBIwBgYE
# VR0gADAIBgZngQwBBAEwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybC5jb21v
# ZG9jYS5jb20vQUFBQ2VydGlmaWNhdGVTZXJ2aWNlcy5jcmwwNAYIKwYBBQUHAQEE
# KDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZI
# hvcNAQEMBQADggEBABK/oe+LdJqYRLhpRrWrJAoMpIpnuDqBv0WKfVIHqI0fTiGF
# OaNrXi0ghr8QuK55O1PNtPvYRL4G2VxjZ9RAFodEhnIq1jIV9RKDwvnhXRFAZ/ZC
# J3LFI+ICOBpMIOLbAffNRk8monxmwFE2tokCVMf8WPtsAO7+mKYulaEMUykfb9gZ
# pk+e96wJ6l2CxouvgKe9gUhShDHaMuwV5KZMPWw5c9QLhTkg4IUaaOGnSDip0TYl
# d8GNGRbFiExmfS9jzpjoad+sPKhdnckcW67Y8y90z7h+9teDnRGWYpquRRPaf9xH
# +9/DUp/mBlXpnYzyOmJRvOwkDynUWICE5EV7WtgwggYaMIIEAqADAgECAhBiHW0M
# UgGeO5B5FSCJIRwKMA0GCSqGSIb3DQEBDAUAMFYxCzAJBgNVBAYTAkdCMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxLTArBgNVBAMTJFNlY3RpZ28gUHVibGljIENv
# ZGUgU2lnbmluZyBSb290IFI0NjAeFw0yMTAzMjIwMDAwMDBaFw0zNjAzMjEyMzU5
# NTlaMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzAp
# BgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYwggGiMA0G
# CSqGSIb3DQEBAQUAA4IBjwAwggGKAoIBgQCbK51T+jU/jmAGQ2rAz/V/9shTUxjI
# ztNsfvxYB5UXeWUzCxEeAEZGbEN4QMgCsJLZUKhWThj/yPqy0iSZhXkZ6Pg2A2NV
# DgFigOMYzB2OKhdqfWGVoYW3haT29PSTahYkwmMv0b/83nbeECbiMXhSOtbam+/3
# 6F09fy1tsB8je/RV0mIk8XL/tfCK6cPuYHE215wzrK0h1SWHTxPbPuYkRdkP05Zw
# mRmTnAO5/arnY83jeNzhP06ShdnRqtZlV59+8yv+KIhE5ILMqgOZYAENHNX9SJDm
# +qxp4VqpB3MV/h53yl41aHU5pledi9lCBbH9JeIkNFICiVHNkRmq4TpxtwfvjsUe
# dyz8rNyfQJy/aOs5b4s+ac7IH60B+Ja7TVM+EKv1WuTGwcLmoU3FpOFMbmPj8pz4
# 4MPZ1f9+YEQIQty/NQd/2yGgW+ufflcZ/ZE9o1M7a5Jnqf2i2/uMSWymR8r2oQBM
# dlyh2n5HirY4jKnFH/9gRvd+QOfdRrJZb1sCAwEAAaOCAWQwggFgMB8GA1UdIwQY
# MBaAFDLrkpr/NZZILyhAQnAgNpFcF4XmMB0GA1UdDgQWBBQPKssghyi47G9IritU
# pimqF6TNDDAOBgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNV
# HSUEDDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEsG
# A1UdHwREMEIwQKA+oDyGOmh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1
# YmxpY0NvZGVTaWduaW5nUm9vdFI0Ni5jcmwwewYIKwYBBQUHAQEEbzBtMEYGCCsG
# AQUFBzAChjpodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2Rl
# U2lnbmluZ1Jvb3RSNDYucDdjMCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0
# aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOCAgEABv+C4XdjNm57oRUgmxP/BP6YdURh
# w1aVcdGRP4Wh60BAscjW4HL9hcpkOTz5jUug2oeunbYAowbFC2AKK+cMcXIBD0Zd
# OaWTsyNyBBsMLHqafvIhrCymlaS98+QpoBCyKppP0OcxYEdU0hpsaqBBIZOtBajj
# cw5+w/KeFvPYfLF/ldYpmlG+vd0xqlqd099iChnyIMvY5HexjO2AmtsbpVn0OhNc
# WbWDRF/3sBp6fWXhz7DcML4iTAWS+MVXeNLj1lJziVKEoroGs9Mlizg0bUMbOalO
# hOfCipnx8CaLZeVme5yELg09Jlo8BMe80jO37PU8ejfkP9/uPak7VLwELKxAMcJs
# zkyeiaerlphwoKx1uHRzNyE6bxuSKcutisqmKL5OTunAvtONEoteSiabkPVSZ2z7
# 6mKnzAfZxCl/3dq3dUNw4rg3sTCggkHSRqTqlLMS7gjrhTqBmzu1L90Y1KWN/Y5J
# KdGvspbOrTfOXyXvmPL6E52z1NZJ6ctuMFBQZH3pwWvqURR8AgQdULUvrxjUYbHH
# j95Ejza63zdrEcxWLDX6xWls/GDnVNueKjWUH3fTv1Y8Wdho698YADR7TNx8X8z2
# Bev6SivBBOHY+uqiirZtg0y9ShQoPzmCcn63Syatatvx157YK9hlcPmVoa1oDE5/
# L9Uo2bC5a4CH2RwwggZzMIIE26ADAgECAhAMcJlHeeRMvJV4PjhvyrrbMA0GCSqG
# SIb3DQEBDAUAMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYw
# HhcNMjMwMzIwMDAwMDAwWhcNMjYwMzE5MjM1OTU5WjBqMQswCQYDVQQGEwJERTEP
# MA0GA1UECAwGQmF5ZXJuMSQwIgYDVQQKDBtTRVBQbWFpbCAtIERldXRzY2hsYW5k
# IEdtYkgxJDAiBgNVBAMMG1NFUFBtYWlsIC0gRGV1dHNjaGxhbmQgR21iSDCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAOapobQkNYCMP+Y33JcGo90Soe9Y
# /WWojr4bKHbLNBzKqZ6cku2uCxhMF1Ln6xuI4ATdZvm4O7GqvplG9nF1ad5t2Lus
# 5SLs45AYnODP4aqPbPU/2NGDRpfnceF+XhKeiYBwoIwrPZ04b8bfTpckj/tvenB9
# P8/9hAjWK97xv7+qsIz4lMMaCuWZgi8RlP6XVxsb+jYrHGA1UdHZEpunEFLaO9Ss
# OPqatPAL2LNGs/JVuGdq9p47GKzn+vl+ANd5zZ/TIP1ifX76vorqZ9l9a5mzi/HG
# vq43v2Cj3jrzIQ7uTbxtiLlPQUqkRzPRtiwTV80JdtRE+M+gTf7bT1CTvG2L3scf
# YKFk7S80M7NydxV/qL+l8blGGageCzJ8svju2Mo4BB+ALWr+gBmCGqrM8YKy/wXR
# tbvdEvBOLsATcHX0maw9xRCDRle2jO+ndYkTKZ92AMH6a/WdDfL0HrAWloWWSg62
# TxmJ/QiX54ILQv2Tlh1Al+pjGHN2evxS8i+XoWcUdHPIOoQd37yjnMjCN593wDzj
# XCEuDABYw9BbvfSp29G/uiDGtjttDXzeMRdVCJFgULV9suBVP7yFh9pK/mVpz+aC
# L2PvqiGYR41xRBKqwrfJEdoluRsqDy6KD985EdXkTvdIFKv0B7MfbcBCiGUBcm1r
# fLAbs8Q2lqvqM4bxAgMBAAGjggGpMIIBpTAfBgNVHSMEGDAWgBQPKssghyi47G9I
# ritUpimqF6TNDDAdBgNVHQ4EFgQUL96+KAGrvUgJnXwdVnA/uy+RlEcwDgYDVR0P
# AQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwSgYD
# VR0gBEMwQTA1BgwrBgEEAbIxAQIBAwIwJTAjBggrBgEFBQcCARYXaHR0cHM6Ly9z
# ZWN0aWdvLmNvbS9DUFMwCAYGZ4EMAQQBMEkGA1UdHwRCMEAwPqA8oDqGOGh0dHA6
# Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVTaWduaW5nQ0FSMzYu
# Y3JsMHkGCCsGAQUFBwEBBG0wazBEBggrBgEFBQcwAoY4aHR0cDovL2NydC5zZWN0
# aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdDQVIzNi5jcnQwIwYIKwYB
# BQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMB4GA1UdEQQXMBWBE3N1cHBv
# cnRAc2VwcG1haWwuY2gwDQYJKoZIhvcNAQEMBQADggGBAHnWpS4Jw/QiiLQi2EYv
# THCtwKsj7O3G7wAN7wijSJcWF7iCx6AoCuCIgGdWiQuEZcv9pIUrXQ6jOSRHsDNX
# SvIhCK9JakZJSseW/SCb1rvxZ4d0n2jm2SdkWf5j7+W+X4JHeCF9ZOw0ULpe5pFs
# IGTh8bmTtUr3yA11yw4vHfXFwin7WbEoTLVKiL0ZUN0Qk+yBniPPSRRlUZIX8P4e
# iXuw7lh9CMaS3HWRKkK89w//18PjUMxhTZJ6dszN2TAfwu1zxdG/RQqvxXUTTAxU
# JrrCuvowtnDQ55yXMxkkSxWUwLxk76WvXwmohRdsavsGJJ9+yxj5JKOd+HIZ1fZ7
# oi0VhyOqFQAnjNbwR/TqPjRxZKjCNLXSM5YSMZKAhqrJssGLINZ2qDK/CEcVDkBS
# 6Hke4jWMczny8nB8+ATJ84MB7tfSoXE7R0FMs1dinuvjVWIyg6klHigpeEiAaSaG
# 5KF7vk+OlquA+x4ohPuWdtFxobOT2OgHQnK4bJitb9aDazGCAxowggMWAgEBMGgw
# VDELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UE
# AxMiU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5nIENBIFIzNgIQDHCZR3nkTLyV
# eD44b8q62zANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgACh
# AoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAM
# BgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDW5D6Ff6aEntIIlQ9Q0smTdZQz
# zAgQoZAdmvUXzyWFGzANBgkqhkiG9w0BAQEFAASCAgBKFUiwg4PhNSUKPlO3JHOz
# dV9VYJsskF7spJ6pK+VfIovjQCm+YjJHALWMrTPpDlEkumGIwTQe8+zKvzxDqdq9
# n7NXD8hf2u8VYslxvW+QlOrJStIOdYQfJDbINdQI+gk5or5Vhx6RaMOFLqFtfxIW
# QbHoZxBFXxnfr1MzMXAq/T1t6NhTbXE2N5k3K+jqvrYzL2YdwDrZ610IlaYXU6oZ
# VgqbB92mgYsgr5aLqyJXID2BPDrgiB6KC3mOldbvaufe1deE9K7fIYOHzOOmDLBz
# Gm0X+Kme5tcgNxxOC1ufvrL9mj+RnrkYN6/MYPzG8B81s2SSltMesEWfiP6mLm1q
# v6/lM7iD2yazY1kd52PrKkdTGOfs4/crFSr2UIiG6hW2NTH3fd34Kbbi+Ma3jjC/
# Uh3rG+0psE0zAaKdhpRr39BFdwjuxqqLnY6GvsXjEFTWLVDopXgyPbIpiz0+r8jU
# JURVaCcpyWYoUiSuW6JjEAbrN7fS0oNRasVgkWoaK/FD8sfCa8ktOCSGnUvlOQea
# 4SWe1fd7mnBjJCCkCmQ7aRkxx5lnVxj+bp/EWs1EulVTZFEdNha/hyNkSSFj+mCg
# gqGIzOzKM7p7o5dpW62M8Y0324askVVye5fdhAk6od0ne680+nonn7msnMAzduoN
# YE6ZGQampMXGo/iQyOiCiQ==
# SIG # End signature block
