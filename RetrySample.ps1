<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.

Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>

$siteUrl = "https://tenant.sharepoint.com"

# Load the required assemblies. 
# Note that SharePoint Online CSOM 16.1.8361.1200 is the required version of this sample.
Add-Type -Path "C:\csom\lib\net45\Microsoft.SharePoint.Client.dll";
Add-Type -Path "C:\csom\lib\net45\Microsoft.SharePoint.Client.Runtime.dll";

$script:context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl);

# Input User
Write-Host "Please input user name : "
$username = Read-Host

# Input Password
Write-Host "Please input password : "
$password = Read-Host -AsSecureString

$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password);
$script:context.Credentials = $creds;

# Specify the UserAgent in use of ClientContext.
$script:context.add_ExecutingWebRequest({
    param ($source, $eventArgs);
    $request = $eventArgs.WebRequestExecutor.WebRequest;
    $request.UserAgent = "NONISV|Contoso|Application/1.0";
})

function ExecuteQueryWithIncrementalRetry {
    param (
        [parameter(Mandatory = $true)]
        [int]$retryCount,
        [int]$delay = 120
    );

    $RetryAfterHeaderName = "Retry-After";
    $retryAttempts = 0;
    $backoffInterval = $delay
    $retryAfterInterval = 0;
    $retry = $false;

    if ($retryCount -le 0) {
        throw "Provide a retry count greater than zero."
    }
    if ($delay -le 0) {
        throw "Provide a delay greater than zero."
    }



    while ($retryAttempts -lt $retryCount) {
        try {
            if (!$retry)
            {
                $script:context.ExecuteQuery();
                return;
            }
            else
            {
                if (($wrapper -ne $null) -and ($wrapper.Value -ne $null))
                {
                    $script:context.RetryQuery($wrapper.Value);
                    return;
                }
            }
        }
        catch [System.Net.WebException] {
            $response = $_.Exception.Response

            if (($null -ne $response) -and (($response.StatusCode -eq 429) -or ($response.StatusCode -eq 503))) {

                $wrapper = [Microsoft.SharePoint.Client.ClientRequestWrapper]($_.Exception.Data["ClientRequest"]);
                $retry = $true


                $retryAfterHeader = $response.GetResponseHeader($RetryAfterHeaderName);
                $retryAfterInMs = $DefaultRetryAfterInMs;

                if (-not [string]::IsNullOrEmpty($retryAfterHeader)) {
                    if (-not [int]::TryParse($retryAfterHeader, [ref]$retryAfterInterval)) {
                        $retryAfterInterval = $DefaultRetryAfterInMs;
                    }
                }
                else
                {
                    $retryAfterInterval = $backoffInterval;
                }

                Write-Output ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($retryAfterInterval))
                #Add delay.
                Start-Sleep -m ($retryAfterInterval * 1000)
                #Add to retry count.
                $retryAttempts++;
                $backoffInterval = $backoffInterval * 2;
            }
            else {
                throw;
            }
        }
    }

    throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}

# Start your implementation from here.
$web = $script:context.Web
$script:context.Load($web)
# Replace $context.ExecuteQuery() with the following line of code.
ExecuteQueryWithIncrementalRetry -retryCount 5
$web.Title = "RetryTest"
$web.Update()
# Replace $context.ExecuteQuery() with the following line of code.
ExecuteQueryWithIncrementalRetry -retryCount 5