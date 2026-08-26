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

# MIME hash table for static content
$MIMEHASH = @{".avi" = "video/x-msvideo"; ".crt" = "application/x-x509-ca-cert"; ".css" = "text/css"; ".der" = "application/x-x509-ca-cert"; ".doc" = "application/msword"; ".flv" = "video/x-flv"; ".gif" = "image/gif"; ".htm" = "text/html"; ".html" = "text/html"; ".ico" = "image/x-icon"; ".jar" = "application/java-archive"; ".jpeg" = "image/jpeg"; ".jpg" = "image/jpeg"; ".js" = "application/javascript"; ".json" = "application/json"; ".mjs" = "application/javascript"; ".mov" = "video/quicktime"; ".mp3" = "audio/mpeg"; ".mp4" = "video/mp4"; ".mpeg" = "video/mpeg"; ".mpg" = "video/mpeg"; ".pdf" = "application/pdf"; ".pem" = "application/x-x509-ca-cert"; ".pl" = "application/x-perl"; ".png" = "image/png"; ".rss" = "application/rss+xml"; ".shtml" = "text/html"; ".txt" = "text/plain"; ".war" = "application/java-archive"; ".wmv" = "video/x-ms-wmv"; ".xml" = "application/xml"; ".xsl" = "application/xml" }

# Result Message Placeholders
$HTML_SUCCESS = "<div class=`"alert alert-success d-flex align-items-center`" role=`"alert`">{result}</div>"
$HTML_WARN = "<div class=`"alert alert-warning  d-flex align-items-center`" role=`"alert`">{result}</div>"

function ConvertTo-SafeHtml {
    # Every value that reaches a template comes from Active Directory and can legally
    # contain " ' & < >. Substituted raw, a double quote truncates an input's value
    # attribute (silently corrupting the field on save) and an apostrophe breaks out
    # of an inline confirm('...') string, which nulls the handler and removes the
    # confirmation prompt from a destructive action. Everything is encoded on the way in.
    param([Parameter(ValueFromPipeline)]$Value)
    process {
        if ($null -eq $Value) { return "" }
        [System.Net.WebUtility]::HtmlEncode([string]$Value)
    }
}

function ConvertFrom-HttpQuery {
    # Shared parser for GET query strings and POST form bodies.
    #
    # Replaces two broken implementations: the GET path used to unescape the WHOLE
    # query before splitting, so a %3D or %26 inside a value was promoted to a real
    # delimiter and silently truncated the value; and it used Hashtable.Add, which
    # throws on a duplicate key and blanked the page. Splitting first and unescaping
    # each half separately fixes both. '+' means space in form encoding; a literal
    # plus arrives as %2B and is unaffected by the pre-pass.
    param([string]$Raw)

    $Result = @{}
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $Result }

    foreach ($Pair in $Raw.TrimStart('?').Split('&')) {
        if ([string]::IsNullOrEmpty($Pair)) { continue }

        $Split = $Pair.IndexOf('=')
        if ($Split -lt 0) { $RawKey = $Pair; $RawValue = "" }
        else { $RawKey = $Pair.Substring(0, $Split); $RawValue = $Pair.Substring($Split + 1) }

        if ([string]::IsNullOrEmpty($RawKey)) { continue }

        $Key = [URI]::UnescapeDataString($RawKey.Replace('+', ' '))
        $Value = [URI]::UnescapeDataString($RawValue.Replace('+', ' '))

        # last value wins rather than throwing, so ?a=1&a=2 can't blank the page
        $Result[$Key] = $Value
    }
    return $Result
}

function Read-Template {
    # -Raw keeps the file as one string. Without it Get-Content returns a string[],
    # which UTF8.GetBytes() later joins with spaces - destroying every line break and
    # silently breaking any placeholder that spans two lines. Also fails loudly if a
    # template is missing rather than calling .Replace() on $null (blank page).
    param([Parameter(Mandatory)][string]$Name)

    $Path = Join-Path $BASEDIR $Name
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Template '$Name' not found at $Path"
    }
    return (Get-Content -LiteralPath $Path -Raw)
}

function Get-ErrorPage {
    # A recognisable page inside the site chrome instead of a bare white document,
    # so a mistyped or stale ?id= doesn't dead-end the operator with no way back.
    param([string]$Title = "Not found", [string]$Detail = "")

    $SafeTitle = ConvertTo-SafeHtml $Title
    $SafeDetail = ConvertTo-SafeHtml $Detail
    return @"
<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<title>Exchange Recipient Admin Center - $SafeTitle</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head><body class="p-4">
<div class="container">
<div class="alert alert-warning"><strong>$SafeTitle</strong><br />$SafeDetail</div>
<a class="btn btn-primary" href="/">Back to Dashboard</a>
<a class="btn btn-outline-secondary" href="/remotemailboxes">Remote Mailboxes</a>
</div></body></html>
"@
}

function Get-RemoteMailboxEditPage {
    # Renders editremotemailbox.html for a given mailbox identity. Shared by the
    # GET (initial view) and POST (after an update) handlers below.
    param(
        [string]$Identity,
        [string]$ResultHtml = ""
    )

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        return Get-ErrorPage -Title "No mailbox specified" -Detail "This page needs a mailbox to open. Pick one from the Remote Mailboxes list."
    }

    $Mailbox = Get-RemoteMailbox -Identity $Identity -ErrorAction SilentlyContinue

    if (-not $Mailbox) {
        return Get-ErrorPage -Title "Remote mailbox not found" -Detail "No remote mailbox matched '$Identity'."
    }

    # Accepted domain list, used for the "add alias" domain picker. The option VALUE
    # must be DomainName, not Name: the handler concatenates it after an @ to build an
    # address, and Name is a free-text AD object name that only matches DomainName by
    # convention.
    $HTMLROWS_AD = ""
    foreach ($Item in (Get-AcceptedDomain)) {
        $Domain = ConvertTo-SafeHtml $Item.DomainName
        $HTMLROWS_AD += "`n<option value=`"$Domain`">$Domain</option>"
    }

    # Proxy address (EmailAddresses) rows, each with a remove button
    $HTMLROWS_PROXY = ""
    $SafeId = ConvertTo-SafeHtml $Mailbox.PrimarySmtpAddress
    foreach ($Addr in $Mailbox.EmailAddresses) {
        $AddrString = $Addr.ToString()
        $SafeAddr = ConvertTo-SafeHtml $AddrString
        $IsPrimary = $AddrString.StartsWith("SMTP:")
        if ($IsPrimary) {
            $Badge = "<span class=`"badge text-bg-primary`">Primary</span>"
            $RemoveBtn = ""
        }
        else {
            $Badge = "<span class=`"badge text-bg-secondary`">Alias</span>"
            # The confirmation text lives in a data- attribute rather than inline in the
            # confirm() call. Interpolating an address straight into confirm('...') let an
            # apostrophe (legal in an SMTP local part) terminate the string, which nulled
            # the handler and removed the prompt entirely from a destructive action.
            $RemoveBtn = "
            <form method=`"post`" action=`"/editremotemailbox`" data-confirm=`"Remove ${SafeAddr}?`" onsubmit=`"return confirm(this.dataset.confirm)`">
            <input type=`"hidden`" name=`"id`" value=`"$SafeId`">
            <input type=`"hidden`" name=`"Action`" value=`"removealias`">
            <input type=`"hidden`" name=`"alias`" value=`"$SafeAddr`">
            <button type=`"submit`" class=`"btn btn-sm btn-outline-danger`">Remove</button>
            </form>"
        }
        $HTMLROWS_PROXY += "
        <tr>
        <td>$SafeAddr</td>
        <td>$Badge</td>
        <td>$RemoveBtn</td>
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

    $HTMLRESPONSE = Read-Template "editremotemailbox.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DisplayName}", (ConvertTo-SafeHtml $Mailbox.DisplayName))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{PrimarySmtpAddress}", $SafeId)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Alias}", (ConvertTo-SafeHtml $Mailbox.Alias))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{RemoteRoutingAddress}", (ConvertTo-SafeHtml $Mailbox.RemoteRoutingAddress))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{id}", $SafeId)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{hidden_badge_class}", $HiddenBadgeClass)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{hidden_status_text}", $HiddenStatusText)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{hidden_toggle_value}", $HiddenToggleValue)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{hidden_toggle_label}", $HiddenToggleLabel)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_proxy} -->", $HTMLROWS_PROXY)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_ad} -->", $HTMLROWS_AD)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Get-RemoteMailboxListPage {
    # Renders remotemailboxes.html. Used after disabling a Remote Mailbox, since
    # the mailbox no longer exists to redirect back to an edit page for.
    param(
        [string]$ResultHtml = ""
    )

    $HTMLROWS_USERS = ""
    foreach ($Item in (Get-User -Filter "RecipientType -eq 'User' -and RecipientTypeDetails -ne 'DisabledUser'" | Where-Object { $_.UserPrincipalName })) {
        $Upn = ConvertTo-SafeHtml $Item.UserPrincipalName
        $HTMLROWS_USERS += "`n<option value=`"$Upn`">$Upn</option>"
    }

    # Option values are DomainName (what actually goes after the @), not Name.
    $HTMLROWS_AD = ""
    foreach ($Item in (Get-AcceptedDomain)) {
        $Domain = ConvertTo-SafeHtml $Item.DomainName
        $Sel = if ($Item.Default) { " selected" } else { "" }
        $HTMLROWS_AD += "`n<option$Sel value=`"$Domain`">$Domain</option>"
    }

    $HTMLROWS_RRA = ""
    foreach ($Item in (Get-AcceptedDomain)) {
        $Domain = ConvertTo-SafeHtml $Item.DomainName
        $Sel = if ($Item.DomainName -like "*.mail.onmicrosoft.com") { " selected" } else { "" }
        $HTMLROWS_RRA += "`n<option$Sel value=`"$Domain`">$Domain</option>"
    }

    $HTMLROWS_MBX = ""
    foreach ($Item in (Get-RemoteMailbox | Select-Object DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenChanged)) {
        $Href = ConvertTo-SafeHtml ([URI]::EscapeDataString([string]$Item.PrimarySMTPAddress))
        $HTMLROWS_MBX += "
        <tr>
        <th scope=`"row`">
        <a href=`"/editremotemailbox?id=$Href`">$(ConvertTo-SafeHtml $Item.DisplayName)</a></th>
        <td>$(ConvertTo-SafeHtml $Item.PrimarySMTPAddress)</td>
        <td>$(ConvertTo-SafeHtml $Item.RecipientTypeDetails)</td>
        <td>$(ConvertTo-SafeHtml $Item.WhenChanged)</td>
        </tr>";
    }

    $HTMLRESPONSE = Read-Template "remotemailboxes.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_mbx} -->", $HTMLROWS_MBX)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_ad} -->", $HTMLROWS_AD)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_user} -->", $HTMLROWS_USERS)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_rra} -->", $HTMLROWS_RRA)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Get-AcceptedDomainsListPage {
    # Renders accepteddomains.html. Shared by GET /accepteddomains and every
    # mutating POST handler (add/edit/delete) so the row list - including the
    # per-row Delete button - isn't rebuilt separately at each call site.
    param(
        [string]$ResultHtml = ""
    )

    $HTMLROWS = ""
    foreach ($Item in (Get-AcceptedDomain)) {
        $SafeName = ConvertTo-SafeHtml $Item.Name
        $SafeDomain = ConvertTo-SafeHtml $Item.DomainName
        $Href = ConvertTo-SafeHtml ([URI]::EscapeDataString([string]$Item.Name))
        $HTMLROWS += "
        <tr>
        <th scope=`"row`">
        <a href=`"/editaccepteddomain?id=$Href`">$SafeName</a></th>
        <td>$SafeDomain</td>
        <td>$(ConvertTo-SafeHtml $Item.DomainType)</td>
        <td>
        <form method=`"post`" action=`"/deleteaccepteddomain`" data-confirm=`"Delete accepted domain ${SafeDomain}? This cannot be undone from here.`" onsubmit=`"return confirm(this.dataset.confirm)`">
        <input type=`"hidden`" name=`"name`" value=`"$SafeName`">
        <button type=`"submit`" class=`"btn btn-sm btn-outline-danger`">Delete</button>
        </form>
        </td>
        </tr>";
    }

    $HTMLRESPONSE = Read-Template "accepteddomains.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row} -->", $HTMLROWS)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Get-DistributionGroupsListPage {
    param([string]$ResultHtml = "")

    $HTMLROWS_DL = ""
    $HTMLROWS_MES = ""
    foreach ($Item in (Get-DistributionGroup | Select-Object DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenCreated)) {
        $Href = ConvertTo-SafeHtml ([URI]::EscapeDataString([string]$Item.PrimarySMTPAddress))
        $Row = "
        <tr>
        <th scope=`"row`">
        <a href=`"/editdistributiongroup?id=$Href`">$(ConvertTo-SafeHtml $Item.DisplayName)</a></th>
        <td>$(ConvertTo-SafeHtml $Item.PrimarySMTPAddress)</td>
        <td>$(ConvertTo-SafeHtml $Item.WhenCreated)</td>
        </tr>"
        if ($Item.RecipientTypeDetails -eq "MailUniversalDistributionGroup") { $HTMLROWS_DL += $Row }
        elseif ($Item.RecipientTypeDetails -eq "MailUniversalSecurityGroup") { $HTMLROWS_MES += $Row }
    }

    # Populates the Mail-Enable modal's picker. The template has always carried this
    # placeholder but nothing ever replaced it, so the dropdown was permanently empty -
    # and being 'required', the browser blocked the form on a control with no options.
    #
    # Wrapped defensively: this is the only place the page depends on Get-Group, and the
    # Recipient Management snap-in ships a reduced cmdlet set that varies by CU. If the
    # cmdlet is missing or errors, the rest of the Distribution Groups page must still
    # render - it worked before this picker existed and must keep working.
    $HTMLROWS_GROUPS = ""
    try {
        foreach ($Item in (Get-Group -ResultSize Unlimited -ErrorAction Stop |
                Where-Object { -not $_.WindowsEmailAddress -and $_.GroupType -notlike "*BuiltinLocal*" } |
                Select-Object -First 500 Name)) {
            $SafeName = ConvertTo-SafeHtml $Item.Name
            $HTMLROWS_GROUPS += "`n<option value=`"$SafeName`">$SafeName</option>"
        }
    }
    catch {
        "$(Get-Date -Format s) Could not enumerate AD groups for the mail-enable picker: $($_.Exception.Message)"
        $HTMLROWS_GROUPS = "`n<option value=`"`">(could not load groups - see console)</option>"
    }

    $HTMLRESPONSE = Read-Template "distributiongroups.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_dl} -->", $HTMLROWS_DL)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row_mes} -->", $HTMLROWS_MES)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {existing_groups} -->", $HTMLROWS_GROUPS)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Get-ContactsListPage {
    param([string]$ResultHtml = "")

    $HTMLROWS = ""
    foreach ($Item in (Get-MailContact | Select-Object DisplayName, PrimarySMTPAddress, RecipientType)) {
        $Href = ConvertTo-SafeHtml ([URI]::EscapeDataString([string]$Item.PrimarySMTPAddress))
        $HTMLROWS += "
        <tr>
        <th scope=`"row`">
        <a href=`"/editcontact?id=$Href`">$(ConvertTo-SafeHtml $Item.DisplayName)</a></th>
        <td>$(ConvertTo-SafeHtml $Item.PrimarySMTPAddress)</td>
        <td>$(ConvertTo-SafeHtml $Item.RecipientType)</td>
        </tr>";
    }

    $HTMLRESPONSE = Read-Template "contacts.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row} -->", $HTMLROWS)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Get-EmailAddressPoliciesListPage {
    param([string]$ResultHtml = "")

    $HTMLROWS = ""
    foreach ($Item in (Get-EmailAddressPolicy | Select-Object Name, Priority, RecipientFilter)) {
        $Href = ConvertTo-SafeHtml ([URI]::EscapeDataString([string]$Item.Name))
        $HTMLROWS += "
        <tr>
        <th scope=`"row`">
        <a href=`"/editemailaddresspolicy?id=$Href`">$(ConvertTo-SafeHtml $Item.Name)</a></th>
        <td>$(ConvertTo-SafeHtml $Item.Priority)</td>
        <td>$(ConvertTo-SafeHtml $Item.RecipientFilter)</td>
        </tr>";
    }

    $HTMLRESPONSE = Read-Template "emailaddresspolicies.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {row} -->", $HTMLROWS)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Read-RequestBody {
    # Every POST handler repeated these five lines verbatim.
    param($Request)
    $Reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    $Data = $Reader.ReadToEnd()
    $Reader.Close()
    $Request.InputStream.Close()
    return $Data
}

function Get-AcceptedDomainEditPage {
    # Renders editaccepteddomain.html for a given domain identity. Shared by
    # the GET (initial view) and POST (after an update) handlers below.
    param(
        [string]$Identity,
        [string]$ResultHtml = ""
    )

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        return Get-ErrorPage -Title "No accepted domain specified" -Detail "This page needs a domain to open. Pick one from the Accepted Domains list."
    }

    $Domain = Get-AcceptedDomain -Identity $Identity -ErrorAction SilentlyContinue

    if (-not $Domain) {
        return Get-ErrorPage -Title "Accepted domain not found" -Detail "No accepted domain matched '$Identity'."
    }

    $HTMLOPTIONS_TYPE = ""
    foreach ($Type in @("Authoritative", "InternalRelay", "ExternalRelay")) {
        $Sel = if ($Type -eq [string]$Domain.DomainType) { " selected" } else { "" }
        $HTMLOPTIONS_TYPE += "`n<option$Sel value=`"$Type`">$Type</option>"
    }

    $HTMLRESPONSE = Read-Template "editaccepteddomain.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Name}", (ConvertTo-SafeHtml $Domain.Name))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DomainName}", (ConvertTo-SafeHtml $Domain.DomainName))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {domaintype_options} -->", $HTMLOPTIONS_TYPE)
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Get-DistributionGroupEditPage {
    # Renders editdistributiongroup.html for a given group identity. Shared by
    # the GET (initial view) and POST (after an update) handlers below.
    param(
        [string]$Identity,
        [string]$ResultHtml = ""
    )

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        return Get-ErrorPage -Title "No distribution group specified" -Detail "This page needs a group to open. Pick one from the Distribution Groups list."
    }

    $Group = Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue

    if (-not $Group) {
        return Get-ErrorPage -Title "Distribution group not found" -Detail "No distribution group matched '$Identity'."
    }

    $HTMLRESPONSE = Read-Template "editdistributiongroup.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DisplayName}", (ConvertTo-SafeHtml $Group.DisplayName))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{PrimarySmtpAddress}", (ConvertTo-SafeHtml $Group.PrimarySmtpAddress))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Alias}", (ConvertTo-SafeHtml $Group.Alias))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Get-ContactEditPage {
    # Renders editcontact.html for a given contact identity. Shared by the GET
    # (initial view) and POST (after an update) handlers below.
    param(
        [string]$Identity,
        [string]$ResultHtml = ""
    )

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        return Get-ErrorPage -Title "No contact specified" -Detail "This page needs a contact to open. Pick one from the Contacts list."
    }

    $Contact = Get-MailContact -Identity $Identity -ErrorAction SilentlyContinue

    if (-not $Contact) {
        return Get-ErrorPage -Title "Contact not found" -Detail "No mail contact matched '$Identity'."
    }

    $HTMLRESPONSE = Read-Template "editcontact.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{DisplayName}", (ConvertTo-SafeHtml $Contact.DisplayName))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{PrimarySmtpAddress}", (ConvertTo-SafeHtml $Contact.PrimarySmtpAddress))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{ExternalEmailAddress}", (ConvertTo-SafeHtml $Contact.ExternalEmailAddress))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("<!-- {result} -->", $ResultHtml)
    return $HTMLRESPONSE
}

function Get-EmailAddressPolicyEditPage {
    # Renders editemailaddresspolicy.html for a given policy identity. Shared
    # by the GET (initial view) and POST (after an update) handlers below.
    param(
        [string]$Identity,
        [string]$ResultHtml = ""
    )

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        return Get-ErrorPage -Title "No email address policy specified" -Detail "This page needs a policy to open. Pick one from the Email Address Policies list."
    }

    $Policy = Get-EmailAddressPolicy -Identity $Identity -ErrorAction SilentlyContinue

    if (-not $Policy) {
        return Get-ErrorPage -Title "Email address policy not found" -Detail "No email address policy matched '$Identity'."
    }

    $HTMLRESPONSE = Read-Template "editemailaddresspolicy.html"
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Name}", (ConvertTo-SafeHtml $Policy.Name))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{Priority}", (ConvertTo-SafeHtml $Policy.Priority))
    $HTMLRESPONSE = $HTMLRESPONSE.Replace("{RecipientFilter}", (ConvertTo-SafeHtml $Policy.RecipientFilter))
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
      try {
        # analyze incoming request
        $CONTEXT = $LISTENER.GetContext()
        $REQUEST = $CONTEXT.Request
        $RESPONSE = $CONTEXT.Response
        $RESPONSEWRITTEN = $FALSE

        # Everything in this loop lives in one long-lived scope, so anything not reset
        # here carries into the next request. $HTMLRESPONSE leaking meant a request that
        # matched no route could serve the PREVIOUS caller's page verbatim, and
        # $HTML_RESULT leaking pinned a stale success/failure banner - an outcome claim
        # about a destructive Exchange operation - onto later, unrelated page loads.
        $HTMLRESPONSE = ""
        $HTML_RESULT = ""

        # $Error is PowerShell's CUMULATIVE error list. The catch blocks below report
        # the current failure only, but it is cleared here as well so nothing else
        # inspecting it sees errors from earlier requests.
        $Error.Clear()

        # log to console
        "$(Get-Date -Format s) $($REQUEST.RemoteEndPoint.Address.ToString()) $($REQUEST.httpMethod) $($REQUEST.Url.PathAndQuery)"
        # and in log variable, capped so a long-running server doesn't grow unbounded
        $WEBLOG += "$(Get-Date -Format s) $($REQUEST.RemoteEndPoint.Address.ToString()) $($REQUEST.httpMethod) $($REQUEST.Url.PathAndQuery)`n"
        if ($WEBLOG.Length -gt 262144) { $WEBLOG = $WEBLOG.Substring($WEBLOG.Length - 131072) }
        $RECEIVED = '{0} {1}' -f $REQUEST.httpMethod, $REQUEST.Url.LocalPath
        # check for known commands
        switch ($RECEIVED) {

            "GET /" {
                # Return the dashboard homepage
                $HTMLRESPONSE = Read-Template "index.html"
                break
            }

            "GET /remotemailboxes" {
                # Remote Mailbox Section. The list render is shared with
                # Get-RemoteMailboxListPage rather than duplicated here - the old inline
                # copy injected the loop-scoped $HTML_RESULT unconditionally, so a banner
                # from an earlier request stayed pinned to this page for the session.
                $Table = ConvertFrom-HttpQuery $REQUEST.Url.Query

                # Only treat this as a form submission when the form's own fields are
                # present. Testing "is there any query string" meant a stray ?foo fired a
                # real Enable-RemoteMailbox with a null identity and an address of "@".
                if ($Table.ContainsKey('username') -and -not [string]::IsNullOrWhiteSpace($Table['username'])) {
                    try {
                        $NewPrimary = "$($Table['primarysmtpaddress_local'])@$($Table['primarysmtpaddress_accepteddomain'])"
                        $NewRouting = "$($Table['remoteroutingaddress_local'])@$($Table['remoteroutingaddress_accepteddomain'])"
                        Enable-RemoteMailbox -Identity $Table['username'] -PrimarySMTPAddress $NewPrimary -RemoteRoutingAddress $NewRouting -ErrorAction Stop | Out-Null
                        $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "User $(ConvertTo-SafeHtml $Table['username']) enabled as Remote Mailbox ($(ConvertTo-SafeHtml $NewPrimary))")
                    }
                    catch {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                    }
                }

                $HTMLRESPONSE = Get-RemoteMailboxListPage -ResultHtml $HTML_RESULT
                break
            }

            "GET /editremotemailbox" {
                # Edit Remote Mailbox Section
                $id = (ConvertFrom-HttpQuery $REQUEST.Url.Query)['id']

                $HTMLRESPONSE = Get-RemoteMailboxEditPage -Identity $id
                break
            }

            "POST /editremotemailbox" {
                # Process Edit Remote Mailbox
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                # "id" is set by the alias/hidden/disable mini-forms; the main
                # properties form has no id field and keys off PrimarySmtpAddress.
                $Identity = if ($params.ContainsKey('id')) { $params['id'] } else { $params['PrimarySmtpAddress'] }
                $Disabled = $false

                try {
                    switch ($params['Action']) {
                        "addalias" {
                            # Exchange requires an explicit prefix on proxy addresses. Without
                            # one, Set-RemoteMailbox raises a non-terminating error and the
                            # address is silently not added, so accept the prefix if the
                            # operator typed it and supply "smtp:" if they didn't.
                            $AliasInput = "$($params['alias_local'])".Trim()
                            if (-not $AliasInput) {
                                $HTML_RESULT = $HTML_WARN.Replace("{result}", "Alias cannot be empty")
                                break
                            }

                            # Peel off any prefix so the address can be validated, keeping the
                            # operator's casing: lower-case "smtp:" is an alias, upper-case
                            # "SMTP:" makes it the primary reply address.
                            $Prefix = ''
                            if ($AliasInput -match '^(?<p>smtp):(?<rest>.*)$') {
                                $Prefix     = $Matches['p']
                                $AliasInput = $Matches['rest'].Trim()
                            }

                            # A full address typed into the local-part box wins over the
                            # domain picker, rather than being concatenated into nonsense.
                            $UsedPicker = ($AliasInput -notlike '*@*')
                            $Address = if ($UsedPicker) {
                                "$AliasInput@$($params['alias_accepteddomain'])"
                            } else {
                                $AliasInput
                            }

                            if ($Address -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                                $HTML_RESULT = $HTML_WARN.Replace("{result}", "'$Address' is not a valid email address")
                                break
                            }

                            if (-not $Prefix) { $Prefix = 'smtp' }
                            $NewAlias = "${Prefix}:$Address"

                            # -ErrorAction Stop so a failure reaches the catch below instead of
                            # falling through to the success message.
                            Set-RemoteMailbox -Identity $Identity -EmailAddresses @{Add = $NewAlias } -ErrorAction Stop

                            $AddedAs = if ($Prefix -ceq 'SMTP') { "primary address" } else { "alias" }
                            $PickerNote = if ($UsedPicker) { "" } else { " (domain taken from the address you typed)" }
                            $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Added $AddedAs $Address$PickerNote")
                            break
                        }
                        "removealias" {
                            Set-RemoteMailbox -Identity $Identity -EmailAddresses @{Remove = $params['alias'] } -ErrorAction Stop
                            $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Removed alias $($params['alias'])")
                            break
                        }
                        "togglehidden" {
                            $NewHidden = $params['hidden'] -eq 'true'
                            Set-RemoteMailbox -Identity $Identity -HiddenFromAddressListsEnabled $NewHidden -ErrorAction Stop
                            $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Set 'Hidden from address lists' to $NewHidden")
                            break
                        }
                        "disable" {
                            # Require the confirmation text to match the mailbox's actual
                            # PrimarySmtpAddress - checked server-side too, not just via the
                            # UI's type-to-confirm JS, since that's trivially bypassed.
                            $MailboxToDisable = Get-RemoteMailbox -Identity $Identity -ErrorAction SilentlyContinue
                            if (-not $MailboxToDisable) {
                                $HTML_RESULT = $HTML_WARN.Replace("{result}", "Remote mailbox '$($Identity)' not found")
                            }
                            elseif ($params['confirmAddress'] -ne $MailboxToDisable.PrimarySmtpAddress) {
                                $HTML_RESULT = $HTML_WARN.Replace("{result}", "Confirmation text didn't match $($MailboxToDisable.PrimarySmtpAddress) - no changes were made")
                            }
                            else {
                                Disable-RemoteMailbox -Identity $Identity -Confirm:$false -ErrorAction Stop
                                $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Remote Mailbox for $($MailboxToDisable.PrimarySmtpAddress) has been disabled")
                                $Disabled = $true
                            }
                            break
                        }
                        default {
                            Set-RemoteMailbox -Identity $Identity -DisplayName $params['DisplayName'] -Alias $params['Alias'] -RemoteRoutingAddress $params['RemoteRoutingAddress'] -ErrorAction Stop
                            $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Remote Mailbox updated successfully")
                            # DisplayName/Alias changes don't affect PrimarySmtpAddress, so it still identifies the mailbox below
                            $Identity = $params['PrimarySmtpAddress']
                        }
                    }
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                # A disabled mailbox no longer exists to render an edit page for -
                # send the user back to the list instead.
                if ($Disabled) {
                    $HTMLRESPONSE = Get-RemoteMailboxListPage -ResultHtml $HTML_RESULT
                }
                else {
                    $HTMLRESPONSE = Get-RemoteMailboxEditPage -Identity $Identity -ResultHtml $HTML_RESULT
                }
                break
            }

            "GET /distributiongroups" {
                # Distribution Groups Section
                $HTMLRESPONSE = Get-DistributionGroupsListPage
                break
            }

            "GET /editdistributiongroup" {
                # Edit Distribution Group Section
                $id = (ConvertFrom-HttpQuery $REQUEST.Url.Query)['id']
                $HTMLRESPONSE = Get-DistributionGroupEditPage -Identity $id
                break
            }

            "POST /editdistributiongroup" {
                # Process Edit Distribution Group
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    Set-DistributionGroup -Identity $params['PrimarySmtpAddress'] -DisplayName $params['DisplayName'] -Alias $params['Alias'] -ErrorAction Stop
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Distribution Group updated successfully")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-DistributionGroupEditPage -Identity $params['PrimarySmtpAddress'] -ResultHtml $HTML_RESULT
                break
            }
            
            "GET /contacts" {
                # Mail Contacts Section
                $HTMLRESPONSE = Get-ContactsListPage
                break
            }

            "GET /editcontact" {
                # Edit Contact Section
                $id = (ConvertFrom-HttpQuery $REQUEST.Url.Query)['id']
                $HTMLRESPONSE = Get-ContactEditPage -Identity $id
                break
            }

            "POST /editcontact" {
                # Process Edit Contact
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    Set-MailContact -Identity $params['PrimarySmtpAddress'] -DisplayName $params['DisplayName'] -ExternalEmailAddress $params['ExternalEmailAddress'] -ErrorAction Stop
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Contact updated successfully")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-ContactEditPage -Identity $params['PrimarySmtpAddress'] -ResultHtml $HTML_RESULT
                break
            }

            "GET /emailaddresspolicies" {
                # Email Address Policies Section
                $HTMLRESPONSE = Get-EmailAddressPoliciesListPage
                break
            }

            "GET /editemailaddresspolicy" {
                # Edit Email Address Policy Section
                $id = (ConvertFrom-HttpQuery $REQUEST.Url.Query)['id']
                $HTMLRESPONSE = Get-EmailAddressPolicyEditPage -Identity $id
                break
            }

            "POST /editemailaddresspolicy" {
                # Process Edit Email Address Policy
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    Set-EmailAddressPolicy -Identity $params['Name'] -Priority $params['Priority'] -RecipientFilter $params['RecipientFilter'] -ErrorAction Stop
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Email Address Policy updated successfully")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-EmailAddressPolicyEditPage -Identity $params['Name'] -ResultHtml $HTML_RESULT
                break
            }

            "GET /accepteddomains" {
                # Accepted Domains section
                $HTMLRESPONSE = Get-AcceptedDomainsListPage
                break
            }

            "GET /editaccepteddomain" {
                # Edit Accepted Domain Section
                $id = (ConvertFrom-HttpQuery $REQUEST.Url.Query)['id']
                $HTMLRESPONSE = Get-AcceptedDomainEditPage -Identity $id
                break
            }

            "POST /editaccepteddomain" {
                # Process Edit Accepted Domain
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    # -DomainName is a New-AcceptedDomain parameter only: the SMTP
                    # namespace is fixed when the domain is created, and Set-AcceptedDomain
                    # rejects it outright ("A parameter cannot be found that matches
                    # parameter name 'DomainName'"), so changing the type failed too.
                    Set-AcceptedDomain -Identity $params['Name'] -DomainType $params['DomainType'] -ErrorAction Stop
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Accepted Domain $(ConvertTo-SafeHtml $params['Name']) set to $(ConvertTo-SafeHtml $params['DomainType'])")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-AcceptedDomainEditPage -Identity $params['Name'] -ResultHtml $HTML_RESULT
                break
            }

            "POST /addaccepteddomain" {
                # Process Add Accepted Domain - the "Add Accepted Domain" modal on
                # accepteddomains.html posts here, but this route didn't exist
                # (pre-existing gap, not something Set-AcceptedDomain covers since
                # this is for a brand new domain, not editing an existing one).
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    New-AcceptedDomain -Name $params['domainName'] -DomainName $params['domainName'] -DomainType $params['domainType'] -ErrorAction Stop
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Accepted Domain $($params['domainName']) added successfully")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-AcceptedDomainsListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /deleteaccepteddomain" {
                # Process Delete Accepted Domain - one Delete button per row on
                # accepteddomains.html, guarded by a JS confirm() dialog.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    Remove-AcceptedDomain -Identity $params['name'] -Confirm:$false -ErrorAction Stop
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Accepted Domain $($params['name']) deleted")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-AcceptedDomainsListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /deletecontact" {
                # Danger Zone on editcontact.html. The typed confirmation is re-checked
                # here, not just in the browser, since client-side guards are bypassable.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)
                $Identity = $params['id']

                try {
                    $Target = Get-MailContact -Identity $Identity -ErrorAction SilentlyContinue
                    if (-not $Target) {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", "Contact '$(ConvertTo-SafeHtml $Identity)' not found - nothing was deleted")
                    }
                    elseif ($params['confirmText'] -ne [string]$Target.PrimarySmtpAddress) {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", "Confirmation text didn't match $(ConvertTo-SafeHtml $Target.PrimarySmtpAddress) - nothing was deleted")
                    }
                    else {
                        Remove-MailContact -Identity $Identity -Confirm:$false -ErrorAction Stop
                        $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Contact $(ConvertTo-SafeHtml $Target.PrimarySmtpAddress) deleted")
                    }
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-ContactsListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /disabledistributiongroup" {
                # Mail-disable: strips the Exchange attributes but keeps the AD group,
                # its membership and anything it grants access to. The reversible
                # counterpart to /deletedistributiongroup - the group can be mail-enabled
                # again from the Mail-Enable a Group modal.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)
                $Identity = $params['id']

                try {
                    $Target = Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
                    if (-not $Target) {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", "Distribution group '$(ConvertTo-SafeHtml $Identity)' not found - nothing was changed")
                    }
                    elseif ($params['confirmText'] -ne [string]$Target.PrimarySmtpAddress) {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", "Confirmation text didn't match $(ConvertTo-SafeHtml $Target.PrimarySmtpAddress) - nothing was changed")
                    }
                    else {
                        Disable-DistributionGroup -Identity $Identity -Confirm:$false -ErrorAction Stop
                        $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Distribution group $(ConvertTo-SafeHtml $Target.PrimarySmtpAddress) mail-disabled. The Active Directory group was kept and can be mail-enabled again from Mail-Enable a Group.")
                    }
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-DistributionGroupsListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /deletedistributiongroup" {
                # Danger Zone on editdistributiongroup.html. Remove-DistributionGroup
                # deletes the AD group object outright, so the confirmation is verified
                # server-side before anything is touched.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)
                $Identity = $params['id']

                try {
                    $Target = Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
                    if (-not $Target) {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", "Distribution group '$(ConvertTo-SafeHtml $Identity)' not found - nothing was deleted")
                    }
                    elseif ($params['confirmText'] -ne [string]$Target.PrimarySmtpAddress) {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", "Confirmation text didn't match $(ConvertTo-SafeHtml $Target.PrimarySmtpAddress) - nothing was deleted")
                    }
                    else {
                        Remove-DistributionGroup -Identity $Identity -Confirm:$false -ErrorAction Stop
                        $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Distribution group $(ConvertTo-SafeHtml $Target.PrimarySmtpAddress) deleted")
                    }
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-DistributionGroupsListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /deleteemailaddresspolicy" {
                # Danger Zone on editemailaddresspolicy.html. Keyed on Name, which is
                # what identifies a policy.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)
                $Identity = $params['id']

                try {
                    $Target = Get-EmailAddressPolicy -Identity $Identity -ErrorAction SilentlyContinue
                    if (-not $Target) {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", "Email address policy '$(ConvertTo-SafeHtml $Identity)' not found - nothing was deleted")
                    }
                    elseif ($params['confirmText'] -ne [string]$Target.Name) {
                        $HTML_RESULT = $HTML_WARN.Replace("{result}", "Confirmation text didn't match $(ConvertTo-SafeHtml $Target.Name) - nothing was deleted")
                    }
                    else {
                        Remove-EmailAddressPolicy -Identity $Identity -Confirm:$false -ErrorAction Stop
                        $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Email address policy $(ConvertTo-SafeHtml $Target.Name) deleted")
                    }
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-EmailAddressPoliciesListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /adddistributiongroup" {
                # "Add a Group" modal on distributiongroups.html. The form has always
                # posted here; the route never existed, so every submission 404'd and
                # discarded whatever the operator had typed.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    $GroupType = if ($params['groupType']) { $params['groupType'] } else { "Distribution" }
                    New-DistributionGroup -Name $params['groupName'] -PrimarySmtpAddress $params['groupEmail'] -Type $GroupType -ErrorAction Stop | Out-Null
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Group $(ConvertTo-SafeHtml $params['groupName']) created")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-DistributionGroupsListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /mailenablegroup" {
                # "Mail-Enable a Group" modal on distributiongroups.html - same missing
                # route. Takes an existing, non-mail-enabled AD group and mail-enables it.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    Enable-DistributionGroup -Identity $params['existingGroup'] -PrimarySmtpAddress $params['groupEmail'] -ErrorAction Stop | Out-Null
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Group $(ConvertTo-SafeHtml $params['existingGroup']) mail-enabled as $(ConvertTo-SafeHtml $params['groupEmail'])")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-DistributionGroupsListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /addcontact" {
                # "Add new contact" modal on contacts.html - same missing route.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    New-MailContact -Name $params['displayName'] -DisplayName $params['displayName'] -ExternalEmailAddress $params['externalEmailAddress'] -ErrorAction Stop | Out-Null
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Contact $(ConvertTo-SafeHtml $params['displayName']) created")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-ContactsListPage -ResultHtml $HTML_RESULT
                break
            }

            "POST /addemailaddresspolicy" {
                # "Add new policy" modal on emailaddresspolicies.html - same missing route.
                $params = ConvertFrom-HttpQuery (Read-RequestBody $REQUEST)

                try {
                    New-EmailAddressPolicy -Name $params['policyName'] -Priority $params['priority'] -RecipientFilter $params['recipientFilter'] -ErrorAction Stop | Out-Null
                    $HTML_RESULT = $HTML_SUCCESS.Replace("{result}", "Email Address Policy $(ConvertTo-SafeHtml $params['policyName']) created")
                }
                catch {
                    $HTML_RESULT = $HTML_WARN.Replace("{result}", (ConvertTo-SafeHtml $_.Exception.Message))
                }

                $HTMLRESPONSE = Get-EmailAddressPoliciesListPage -ResultHtml $HTML_RESULT
                break
            }

            "GET /favicon.ico" {
                # Browsers request this on every page. The icon lives in images/ next to
                # the script rather than under web/, so the static handler (rooted at
                # $BASEDIR) can never find it and every page load logged a 404.
                $ICON = Join-Path $PSScriptRoot "images\favicon.ico"
                if (Test-Path -LiteralPath $ICON -PathType Leaf) {
                    $BUFFER = [System.IO.File]::ReadAllBytes($ICON)
                    $RESPONSE.ContentType = "image/x-icon"
                    $RESPONSE.SendChunked = $FALSE
                    $RESPONSE.ContentLength64 = $BUFFER.Length
                    $RESPONSE.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
                    $RESPONSEWRITTEN = $TRUE
                }
                else {
                    $RESPONSE.StatusCode = 404
                    $HTMLRESPONSE = "not found"
                }
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
                    # physical path is a directory - serve its index.html if present.
                    # This used to 'break', which exits the SWITCH, not this block: it
                    # skipped the send below and shipped whatever $HTMLRESPONSE still held
                    # from the PREVIOUS request. Reachable from POST / and GET //.
                    $CANDIDATE = $CHECKDIR.TrimEnd("/\") + "/index.html"
                    if (Test-Path $CANDIDATE -PathType Leaf) {
                        $CHECKFILE = $CANDIDATE
                    }
                    else {
                        # do not generate a directory listing - 404
                        $RESPONSE.StatusCode = 404
                        $HTMLRESPONSE = "<!doctype html><html><body>Page $(ConvertTo-SafeHtml $RECEIVED) not found</body></html>"
                    }
                }
                else {
                    # no directory, check for file
                    if (Test-Path $CHECKDIR -PathType Leaf) {
                        # file found, path now in $CHECKFILE
                        $CHECKFILE = $CHECKDIR
                    }
                    else {
                        $RESPONSE.StatusCode = 404
                        $HTMLRESPONSE = "<!doctype html><html><body>Page $(ConvertTo-SafeHtml $RECEIVED) not found</body></html>"
                    }
                }

                if ($CHECKFILE -ne "") {
                    # static content available
                    try {
                        # ... serve static content
                        $BUFFER = [System.IO.File]::ReadAllBytes($CHECKFILE)
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
                        # ContentLength64 is set only once the bytes are in hand. Setting it
                        # before a read that then fails left the socket promising bytes it
                        # never sent, and the client hung until its own timeout.
                        $RESPONSE.SendChunked = $FALSE
                        $RESPONSE.ContentLength64 = $BUFFER.Length
                        $RESPONSE.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
                        # mark response as already given
                        $RESPONSEWRITTEN = $TRUE
                    }
                    catch {
                        # A failed read must not fall through to the HTML sender below -
                        # that served the previous request's page as if it were the asset.
                        "$(Get-Date -Format s) Static file error '$CHECKFILE': $($_.Exception.Message)"
                        $RESPONSEWRITTEN = $TRUE
                    }
                }
            }
        }

        # only send response if not already done
        if (!$RESPONSEWRITTEN) {
            # return HTML answer to caller
            $BUFFER = [Text.Encoding]::UTF8.GetBytes([string]$HTMLRESPONSE)
            # Templated responses used to go out with no Content-Type at all, leaving
            # every client to guess at the payload.
            $RESPONSE.ContentType = "text/html; charset=utf-8"
            $RESPONSE.ContentLength64 = $BUFFER.Length
            $RESPONSE.AddHeader("Last-Modified", [DATETIME]::Now.ToString('r'))
            $RESPONSE.AddHeader("Server", "Powershell Webserver/1.2 on localhost")
            $RESPONSE.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
        }

        # and finish answer to client
        $RESPONSE.Close()
      }
      catch {
        # A single request failing (e.g. client closed the connection before the
        # response finished writing - "network name is no longer available") should
        # not take down the whole webserver. Log it and keep listening.
        "$(Get-Date -Format s) Request error: $($_.Exception.Message)"
        $WEBLOG += "$(Get-Date -Format s) Request error: $($_.Exception.Message)`n"

        # Tell the operator something went wrong instead of returning a blank 200.
        try {
            if (-not $RESPONSEWRITTEN -and $RESPONSE -and $RESPONSE.OutputStream.CanWrite) {
                $ERRBUF = [Text.Encoding]::UTF8.GetBytes((Get-ErrorPage -Title "Something went wrong handling that request" -Detail $_.Exception.Message))
                $RESPONSE.StatusCode = 500
                $RESPONSE.ContentType = "text/html; charset=utf-8"
                $RESPONSE.ContentLength64 = $ERRBUF.Length
                $RESPONSE.OutputStream.Write($ERRBUF, 0, $ERRBUF.Length)
            }
        }
        catch { }
        try { $RESPONSE.Close() } catch { }
      }

      # Checked outside the per-request try so that an error while serving the exit
      # page can't leave the server running after the operator clicked Exit.
      if ($RECEIVED -eq 'GET /exit') {
          "$(Get-Date -Format s) Stopping powershell webserver..."
          break
      }
    }
}
finally {
    # Stop powershell webserver
    $LISTENER.Stop()
    $LISTENER.Close()
    "$(Get-Date -Format s) Powershell webserver stopped."
}