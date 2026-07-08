<#
.Synopsis
Exchange Recipient Management Tools Local Web Server
.Description
Starts webserver as powershell process as the current user
Navigate to the web site to use Exchange Admin Tools
.Inputs
None
.Outputs
None
.Example
Start-ExchangeRecipientAdminCenter.ps1
.Notes
Author: Steve Goodman
Version: 1.1.0
#>

if (!(Get-PSSnapIn Microsoft.Exchange.Management.PowerShell.RecipientManagement -Registered -ErrorAction SilentlyContinue)) {
    throw "Please install the Exchange 2019 CU12 and above Management Tools-Only install. See: https://docs.microsoft.com/en-us/Exchange/manage-hybrid-exchange-recipients-with-management-tools"
    break
}

# Load Recipient Management PowerShell Tools
Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.RecipientManagement

# Define webserver details
$BASEDIR = $PSScriptRoot + "/web"
$BINDING = "http://localhost:$(Get-Random -Minimum 4000 -Maximum 10000)/"
$INDEX = "\index.html"

# MIME hash table for static content
$MIMEHASH = @{".avi" = "video/x-msvideo"; ".crt" = "application/x-x509-ca-cert"; ".css" = "text/css"; ".der" = "application/x-x509-ca-cert"; ".doc" = "application/msword"; ".flv" = "video/x-flv"; ".gif" = "image/gif"; ".htm" = "text/html"; ".html" = "text/html"; ".ico" = "image/x-icon"; ".jar" = "application/java-archive"; ".jpeg" = "image/jpeg"; ".jpg" = "image/jpeg"; ".js" = "application/javascript"; ".json" = "application/json"; ".mjs" = "application/javascript"; ".mov" = "video/quicktime"; ".mp3" = "audio/mpeg"; ".mp4" = "video/mp4"; ".mpeg" = "video/mpeg"; ".mpg" = "video/mpeg"; ".pdf" = "application/pdf"; ".pem" = "application/x-x509-ca-cert"; ".pl" = "application/x-perl"; ".png" = "image/png"; ".rss" = "application/rss+xml"; ".shtml" = "text/html"; ".txt" = "text/plain"; ".war" = "application/java-archive"; ".wmv" = "video/x-ms-wmv"; ".xml" = "application/xml"; ".xsl" = "application/xml" }

# Result Message Placeholders
$HTML_SUCCESS = "<div class=`"alert alert-success d-flex align-items-center`" role=`"alert`">{result}</div>"
$HTML_WARN = "<div class=`"alert alert-warning  d-flex align-items-center`" role=`"alert`">{result}</div>"

function Get-RemoteMailboxEditPage {
    # Renders editremotemailbox.html for a given mailbox identity. Shared by the
    # GET (initial view) and POST (after an update) handlers below.
    param(
        [Parameter(Mandatory)][string]$Identity,
        [string]$ResultHtml = ""
    )

    $Mailbox = Get-RemoteMailbox -Identity $Identity -ErrorAction SilentlyContinue

    if (-not $Mailbox) {
        return "<!doctype html><html><body>Remote mailbox '$($Identity)' not found</body></html>"
    }

    # Accepted domain list, used for the "add alias" domain picker
    $HTMLROWS_AD = ""
    foreach ($Item in (Get-AcceptedDomain)) {
        $HTMLROWS_AD += "`n<option value=`"$($Item.Name)`">$($Item.DomainName)</option>"
    }

    # Proxy address (EmailAddresses) rows, each with a remove button
    $HTMLROWS_PROXY = ""
    foreach ($Addr in $Mailbox.EmailAddresses) {
        $AddrString = $Addr.ToString()
        $IsPrimary = $AddrString.StartsWith("SMTP:")
        if ($IsPrimary) {
            $Badge = "<span class=`"badge text-bg-primary`">Primary</span>"
            $RemoveBtn = ""
        }
        else {
            $Badge = "<span class=`"badge text-bg-secondary`">Alias</span>"
            $RemoveBtn = "
            <form method=`"post`" action=`"/editremotemailbox`" onsubmit=`"return confirm('Remove $($AddrString)?')`">
            <input type=`"hidden`" name=`"id`" value=`"$($Mailbox.PrimarySmtpAddress)`">
            <input type=`"hidden`" name=`"Action`" value=`"removealias`">
            <input type=`"hidden`" name=`"alias`" value=`"$($AddrString)`">
            <button type=`"submit`" class=`"btn btn-sm btn-outline-danger`">Remove</button>
            </form>"
        }
        $HTMLROWS_PROXY += "
        <tr>
        <td>$($AddrString)</td>
        <td>$($Badge)</td>
        <td>$($RemoveBtn)</td>
        </tr>";
    }

    if ($Mailbox.HiddenFromAddressListsEnabled) {
        $HiddenBadgeClass = "text-bg-warning"
        $HiddenStatusText = "Hidden"
        $HiddenToggleValue = "false"
        $HiddenToggleLabel = "Unhide from GAL"
    }
    else {
        $HiddenBadgeClass = "text-bg-success"
        $HiddenStatusText = "Visible"
        $HiddenToggleValue = "true"
        $HiddenToggleLabel = "Hide from GAL"
    }

    $HTMLRESPONSE = Get-Content -Path "$($BASEDIR)\editremotemailbox.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DisplayName}", $Mailbox.DisplayName)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{PrimarySmtpAddress}", $Mailbox.PrimarySmtpAddress)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Alias}", $Mailbox.Alias)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{RemoteRoutingAddress}", $Mailbox.RemoteRoutingAddress)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{id}", $Mailbox.PrimarySmtpAddress)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{hidden_badge_class}", $HiddenBadgeClass)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{hidden_status_text}", $HiddenStatusText)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{hidden_toggle_value}", $HiddenToggleValue)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{hidden_toggle_label}", $HiddenToggleLabel)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_proxy} -->", $HTMLROWS_PROXY)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_ad} -->", $HTMLROWS_AD)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

# Starting the powershell webserver
"$(Get-Date -Format s) Starting Exchange Recipient Admin Webserver at: $($BINDING)"
$LISTENER = New-Object System.Net.HttpListener
$LISTENER.Prefixes.Add($BINDING)
$LISTENER.Start()
$Error.Clear()

Start-Process -FilePath $BINDING

try {
    "$(Get-Date -Format s) Powershell webserver started."
    $WEBLOG = "$(Get-Date -Format s) Powershell webserver started.`n"
    while ($LISTENER.IsListening) {
        # analyze incoming request
        $CONTEXT = $LISTENER.GetContext()
        $REQUEST = $CONTEXT.Request
        $RESPONSE = $CONTEXT.Response
        $RESPONSEWRITTEN = $FALSE

        # log to console
        "$(Get-Date -Format s) $($REQUEST.RemoteEndPoint.Address.ToString()) $($REQUEST.httpMethod) $($REQUEST.Url.PathAndQuery)"
        # and in log variable
        $WEBLOG += "$(Get-Date -Format s) $($REQUEST.RemoteEndPoint.Address.ToString()) $($REQUEST.httpMethod) $($REQUEST.Url.PathAndQuery)`n"
        $RECEIVED = '{0} {1}' -f $REQUEST.httpMethod, $REQUEST.Url.LocalPath
        # check for known commands
        switch ($RECEIVED) {
            
            "GET /" { 
                # Return the dashboard homepage
                $HTMLRESPONSE = Get-Content -Path "$($BASEDIR)\index.html"
                break
            }

            "GET /remotemailboxes" { 
                # Remote Mailbox Section
                
                # Process submitted form
                if ($REQUEST.Url.Query) {
                    $Table = @{}
                    foreach ($Item in [URI]::UnescapeDataString(($REQUEST.Url.Query.Replace("?", ""))).Split("&")) {
                        $Table.Add($Item.Split("=")[0], $Item.Split("=")[1])
                    }
                    try {
                        $Result = Enable-RemoteMailbox -Identity $Table['username'] -PrimarySMTPAddress "$($Table['primarysmtpaddress_local'])@$($Table['primarysmtpaddress_accepteddomain'])" -RemoteRoutingAddress "$($Table['remoteroutingaddress_local'])@$($Table['remoteroutingaddress_accepteddomain'])"
                        $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "User $($Table['username']) enabled as Remote Mailbox")
                    }
                    catch {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
                    }
                    
                }

                # Prepare user list for non-Exchange users
                $HTMLROWS_USERS = ""
                foreach ($Item in (Get-User -Filter "RecipientType -eq 'User' -and RecipientTypeDetails -ne 'DisabledUser'" | Where { $_.UserPrincipalName })) {
                    $HTMLROWS_USERS += "`n<option value=`"$($Item.UserPrincipalName)`">$($Item.UserPrincipalName)</option>"
                }

                # Prepare accepted domain list
                $HTMLROWS_AD = ""
                foreach ($Item in (Get-AcceptedDomain)) {
                    
                    if ($Item.Default) {
                        $HTMLROWS_AD += "`n<option selected value=`"$($Item.Name)`">$($Item.DomainName)</option>"
                    }
                    else {
                        $HTMLROWS_AD += "`n<option value=`"$($Item.Name)`">$($Item.DomainName)</option>"
                    }
                }

                # Prepare remote routing domain list
                $HTMLROWS_RRA = ""
                foreach ($Item in (Get-AcceptedDomain)) {
                    
                    if ($Item.DomainName -like "*.mail.onmicrosoft.com") {
                        $HTMLROWS_RRA += "`n<option selected value=`"$($Item.Name)`">$($Item.DomainName)</option>"
                    }
                    else {
                        $HTMLROWS_RRA += "`n<option value=`"$($Item.Name)`">$($Item.DomainName)</option>"
                    }
                }

                # Return remote mailbox list
                $HTMLROWS_MBX = ""
                foreach ($Item in (Get-RemoteMailbox | Select DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenChanged)) {
                    $HTMLROWS_MBX += "
                    <tr>
                    <th scope=`"row`">
                    <a href=`"/editremotemailbox?id=$($Item.PrimarySMTPAddress)`">$($Item.DisplayName)</a></th>
                    <td>$($Item.PrimarySMTPAddress)</td>
                    <td>$($Item.RecipientTypeDetails)</td>
                    <td>$($Item.WhenChanged)</td>
                    </tr>";
                }

                # Create response and replace template placeholders
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\remotemailboxes.html").Replace("<!-- {row_mbx} -->", $HTMLROWS_MBX).Replace("<!-- {row_ad} -->", $HTMLROWS_AD).Replace("<!-- {row_user} -->", $HTMLROWS_USERS).Replace("<!-- {row_rra} -->", $HTMLROWS_RRA).Replace("<!-- {result} -->", $HTML_RESULT)
                break
            }

            "GET /editremotemailbox" {
                # Edit Remote Mailbox Section
                $id = [URI]::UnescapeDataString($REQUEST.Url.Query.TrimStart('?').Split("=")[1])

                $HTMLRESPONSE = Get-RemoteMailboxEditPage -Identity $id
                break
            }

            "POST /editremotemailbox" {
                # Process Edit Remote Mailbox
                $reader = New-Object System.IO.StreamReader($REQUEST.InputStream, $REQUEST.ContentEncoding)
                $data = $reader.ReadToEnd()
                $reader.Close()
                $REQUEST.InputStream.Close()

                $params = @{}
                $data.Split('&') | ForEach-Object {
                    $key, $value = $_.Split('=')
                    $params[$key] = [System.Web.HttpUtility]::UrlDecode($value)
                }

                # "id" is set by the alias/hidden mini-forms; the main properties
                # form has no id field and keys off PrimarySmtpAddress instead.
                $Identity = if ($params.ContainsKey('id')) { $params['id'] } else { $params['PrimarySmtpAddress'] }

                try {
                    switch ($params['Action']) {
                        "addalias" {
                            $NewAlias = "$($params['alias_local'])@$($params['alias_accepteddomain'])"
                            Set-RemoteMailbox -Identity $Identity -EmailAddresses @{Add = $NewAlias }
                            $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Added alias $NewAlias")
                            break
                        }
                        "removealias" {
                            Set-RemoteMailbox -Identity $Identity -EmailAddresses @{Remove = $params['alias'] }
                            $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Removed alias $($params['alias'])")
                            break
                        }
                        "togglehidden" {
                            $NewHidden = $params['hidden'] -eq 'true'
                            Set-RemoteMailbox -Identity $Identity -HiddenFromAddressListsEnabled $NewHidden
                            $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Set 'Hidden from address lists' to $NewHidden")
                            break
                        }
                        default {
                            Set-RemoteMailbox -Identity $Identity -DisplayName $params['DisplayName'] -Alias $params['Alias'] -RemoteRoutingAddress $params['RemoteRoutingAddress']
                            $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Remote Mailbox updated successfully")
                            # DisplayName/Alias changes don't affect PrimarySmtpAddress, so it still identifies the mailbox below
                            $Identity = $params['PrimarySmtpAddress']
                        }
                    }
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
                }

                $HTMLRESPONSE = Get-RemoteMailboxEditPage -Identity $Identity -ResultHtml $HTML_RESULT
                break
            }

            "GET /distributiongroups" { 
                # Distribution Groups Section

                # Prepare Distribution Group lists split into tabs
                $HTMLROWS_DL = ""
                $HTMLROWS_MES = ""
                foreach ($Item in (Get-DistributionGroup | Select DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenCreated)) {
                    if ($Item.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                        $HTMLROWS_DL += "
                        <tr>
                        <th scope=`"row`">
                        <a href=`"/editdistributiongroup?id=$($Item.PrimarySMTPAddress)`">$($Item.DisplayName)</a></th>
                        <td>$($Item.PrimarySMTPAddress)</td>
                        <td>$($Item.WhenCreated)</td>
                        </tr>";
                    }
                    elseif ($Item.RecipientTypeDetails -eq "MailUniversalSecurityGroup") {
                        $HTMLROWS_MES += "
                        <tr>
                        <th scope=`"row`">
                        <a href=`"/editdistributiongroup?id=$($Item.PrimarySMTPAddress)`">$($Item.DisplayName)</a></th>
                        <td>$($Item.PrimarySMTPAddress)</td>
                        <td>$($Item.WhenCreated)</td>
                        </tr>";
                    }
                }

                # Create response and replace template placeholders
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\distributiongroups.html").Replace("<!-- {row_dl} -->", $HTMLROWS_DL).Replace("<!-- {row_mes} -->", $HTMLROWS_MES)
                break
            }

            "GET /editdistributiongroup" {
                # Edit Distribution Group Section
                $id = $REQUEST.Url.Query.Split("=")[1]
                $group = Get-DistributionGroup -Identity $id
                
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\editdistributiongroup.html")
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DisplayName}", $group.DisplayName)
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{PrimarySmtpAddress}", $group.PrimarySmtpAddress)
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Alias}", $group.Alias)
                break
            }

            "POST /editdistributiongroup" {
                # Process Edit Distribution Group
                $reader = New-Object System.IO.StreamReader($REQUEST.InputStream, $REQUEST.ContentEncoding)
                $data = $reader.ReadToEnd()
                $reader.Close()
                $REQUEST.InputStream.Close()

                $params = @{}
                $data.Split('&') | ForEach-Object {
                    $key, $value = $_.Split('=')
                    $params[$key] = [System.Web.HttpUtility]::UrlDecode($value)
                }

                try {
                    Set-DistributionGroup -Identity $params['PrimarySmtpAddress'] -DisplayName $params['DisplayName'] -Alias $params['Alias']
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Distribution Group updated successfully")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
                }

                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\distributiongroups.html").Replace("<!-- {result} -->", $HTML_RESULT)
                break
            }
            
            "GET /contacts" { 
                # Mail Contacts Section

                # Prepare contacts list
                $HTMLROWS = ""
                foreach ($Item in (Get-MailContact | Select DisplayName, PrimarySMTPAddress, RecipientType)) {
                    $HTMLROWS += "
                    <tr>
                    <th scope=`"row`">
                    <a href=`"/editcontact?id=$($Item.PrimarySMTPAddress)`">$($Item.DisplayName)</a></th>
                    <td>$($Item.PrimarySMTPAddress)</td>
                    <td>$($Item.RecipientType)</td>
                    </tr>";
                }

                # Create response and replace template placeholders
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\contacts.html").Replace("<!-- {row} -->", $HTMLROWS)
                break
            }

            "GET /editcontact" {
                # Edit Contact Section
                $id = $REQUEST.Url.Query.Split("=")[1]
                $contact = Get-MailContact -Identity $id
                
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\editcontact.html")
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DisplayName}", $contact.DisplayName)
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{PrimarySmtpAddress}", $contact.PrimarySmtpAddress)
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{ExternalEmailAddress}", $contact.ExternalEmailAddress)
                break
            }

            "POST /editcontact" {
                # Process Edit Contact
                $reader = New-Object System.IO.StreamReader($REQUEST.InputStream, $REQUEST.ContentEncoding)
                $data = $reader.ReadToEnd()
                $reader.Close()
                $REQUEST.InputStream.Close()

                $params = @{}
                $data.Split('&') | ForEach-Object {
                    $key, $value = $_.Split('=')
                    $params[$key] = [System.Web.HttpUtility]::UrlDecode($value)
                }

                try {
                    Set-MailContact -Identity $params['PrimarySmtpAddress'] -DisplayName $params['DisplayName'] -ExternalEmailAddress $params['ExternalEmailAddress']
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Contact updated successfully")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
                }

                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\contacts.html").Replace("<!-- {result} -->", $HTML_RESULT)
                break
            }

            "GET /emailaddresspolicies" { 
                # Email Address Policies Section

                # Prepare email address policies list
                $HTMLROWS = ""
                foreach ($Item in (Get-EmailAddressPolicy | Select Name, Priority, RecipientFilter)) {
                    $HTMLROWS += "
                    <tr>
                    <th scope=`"row`">
                    <a href=`"/editemailaddresspolicy?id=$($Item.Name)`">$($Item.Name)</a></th>
                    <td>$($Item.Priority)</td>
                    <td>$($Item.RecipientFilter)</td>
                    </tr>";
                }

                # Create response and replace template placeholders
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\emailaddresspolicies.html").Replace("<!-- {row} -->", $HTMLROWS)
                break
            }

            "GET /editemailaddresspolicy" {
                # Edit Email Address Policy Section
                $id = $REQUEST.Url.Query.Split("=")[1]
                $policy = Get-EmailAddressPolicy -Identity $id
                
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\editemailaddresspolicy.html")
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Name}", $policy.Name)
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Priority}", $policy.Priority)
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{RecipientFilter}", $policy.RecipientFilter)
                break
            }

            "POST /editemailaddresspolicy" {
                # Process Edit Email Address Policy
                $reader = New-Object System.IO.StreamReader($REQUEST.InputStream, $REQUEST.ContentEncoding)
                $data = $reader.ReadToEnd()
                $reader.Close()
                $REQUEST.InputStream.Close()

                $params = @{}
                $data.Split('&') | ForEach-Object {
                    $key, $value = $_.Split('=')
                    $params[$key] = [System.Web.HttpUtility]::UrlDecode($value)
                }

                try {
                    Set-EmailAddressPolicy -Identity $params['Name'] -Priority $params['Priority'] -RecipientFilter $params['RecipientFilter']
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Email Address Policy updated successfully")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
                }

                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\emailaddresspolicies.html").Replace("<!-- {result} -->", $HTML_RESULT)
                break
            }

            "GET /accepteddomains" { 
                # Accepted Domains section

                # Prepare list of accepted domains
                $HTMLROWS = ""
                foreach ($Item in (Get-AcceptedDomain)) {
                    $HTMLROWS += "
                    <tr>
                    <th scope=`"row`">
                    <a href=`"/editaccepteddomain?id=$($Item.Name)`">$($Item.Name)</a></th>
                    <td>$($Item.DomainName)</td>
                    <td>$($Item.DomainType)</td>
                    </tr>";
                }

                # Create response and replace template placeholders
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\accepteddomains.html").Replace("<!-- {row} -->", $HTMLROWS)
                break
            }

            "GET /editaccepteddomain" {
                # Edit Accepted Domain Section
                $id = $REQUEST.Url.Query.Split("=")[1]
                $domain = Get-AcceptedDomain -Identity $id
                
                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\editaccepteddomain.html")
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Name}", $domain.Name)
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DomainName}", $domain.DomainName)
                $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DomainType}", $domain.DomainType)
                break
            }

            "POST /editaccepteddomain" {
                # Process Edit Accepted Domain
                $reader = New-Object System.IO.StreamReader($REQUEST.InputStream, $REQUEST.ContentEncoding)
                $data = $reader.ReadToEnd()
                $reader.Close()
                $REQUEST.InputStream.Close()

                $params = @{}
                $data.Split('&') | ForEach-Object {
                    $key, $value = $_.Split('=')
                    $params[$key] = [System.Web.HttpUtility]::UrlDecode($value)
                }

                try {
                    Set-AcceptedDomain -Identity $params['Name'] -DomainName $params['DomainName'] -DomainType $params['DomainType']
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Accepted Domain updated successfully")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", $Error -join "<br />")
                }

                $HTMLRESPONSE = (Get-Content -Path "$($BASEDIR)\accepteddomains.html").Replace("<!-- {result} -->", $HTML_RESULT)
                break
            }

            "GET /exit" { 
                # Create response preparing for webserver shutdown
                $HTMLRESPONSE = "<!doctype html><html><body>Please close the browser window</body></html>"
                break
            }

            default {    
                # PowerShell webserver main code - this section should be updated if the main project is
                    
                # unknown command, check if path to file
 
                # create physical path based upon the base dir and url
                $CHECKDIR = $BASEDIR.TrimEnd("/\") + $REQUEST.Url.LocalPath
                $CHECKFILE = ""
                if (Test-Path $CHECKDIR -PathType Container) {
                    # physical path is a directory
                    $INDEX = "/index.html"
                    $CHECKFILE = $CHECKDIR.TrimEnd("/\") + $INDEX
                    if (Test-Path $CHECKFILE -PathType Leaf) {
                        # index file found, path now in $CHECKFILE
                        break
                    }
                    $CHECKFILE = ""
                        
                    if ($CHECKFILE -eq "") {
                        # do not generate directory listing - 404 
                        # no file to serve found, return error
                        $RESPONSE.StatusCode = 404
                        $HTMLRESPONSE = "<!doctype html><html><body>Page $($RECEIVED) not found</body></html>"
                    }
                }
                else {
                    # no directory, check for file
                    if (Test-Path $CHECKDIR -PathType Leaf) {
                        # file found, path now in $CHECKFILE
                        $CHECKFILE = $CHECKDIR
                    }
                }

                if ($CHECKFILE -ne "") {
                    # static content available
                    try {
                        # ... serve static content
                        $BUFFER = [System.IO.File]::ReadAllBytes($CHECKFILE)
                        $RESPONSE.ContentLength64 = $BUFFER.Length
                        $RESPONSE.SendChunked = $FALSE
                        $EXTENSION = [IO.Path]::GetExtension($CHECKFILE)
                        if ($MIMEHASH.ContainsKey($EXTENSION)) {
                            # known mime type for this file's extension available
                            $RESPONSE.ContentType = $MIMEHASH.Item($EXTENSION)
                        }
                        else {
                            # no, serve as binary download
                            $RESPONSE.ContentType = "application/octet-stream"
                            $FILENAME = Split-Path -Leaf $CHECKFILE
                            $RESPONSE.AddHeader("Content-Disposition", "attachment; filename=$FILENAME")
                        }
                        $RESPONSE.AddHeader("Last-Modified", [IO.File]::GetLastWriteTime($CHECKFILE).ToString('r'))
                        $RESPONSE.AddHeader("Server", "Powershell Webserver/1.2 on ")
                        $RESPONSE.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
                        # mark response as already given
                        $RESPONSEWRITTEN = $TRUE
                    }
                    catch {
                        # just ignore. Error handling comes afterwards since not every error throws an exception
                    }
                    if ($Error.Count -gt 0) {
                        # retrieve error message on error
                        $RESULT += "`nError while downloading '$CHECKFILE'`n`n"
                        $RESULT += $Error[0]
                        $Error.Clear()
                    }
                }
                else {
                    # no file to serve found, return error
                    if (!(Test-Path $CHECKDIR -PathType Container)) {
                        $RESPONSE.StatusCode = 404
                        $HTMLRESPONSE = "<!doctype html><html><body>Page $($RECEIVED) not found</body></html>"
                    }
                }
            }
        }

        # only send response if not already done
        if (!$RESPONSEWRITTEN) {
            # return HTML answer to caller
            $BUFFER = [Text.Encoding]::UTF8.GetBytes($HTMLRESPONSE)
            $RESPONSE.ContentLength64 = $BUFFER.Length
            $RESPONSE.AddHeader("Last-Modified", [DATETIME]::Now.ToString('r'))
            $RESPONSE.AddHeader("Server", "Powershell Webserver/1.2 on localhost")
            $RESPONSE.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
        }

        # and finish answer to client
        $RESPONSE.Close()

        # If exit was chosen, break out of loop and exit
        if ($RECEIVED -eq 'GET /exit') {
            # then break out of while loop
            "$(Get-Date -Format s) Stopping powershell webserver..."
            break;
        }
    }
}
finally {
    # Stop powershell webserver
    $LISTENER.Stop()
    $LISTENER.Close()
    "$(Get-Date -Format s) Powershell webserver stopped."
}