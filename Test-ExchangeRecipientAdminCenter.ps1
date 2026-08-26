<#
.Synopsis
Regression tests for Start-ExchangeRecipientAdminCenter.ps1
.Description
Boots the REAL, unmodified server script against stubbed Exchange cmdlets and
drives every route over HTTP. No Exchange, AD or admin rights required, and
nothing outside this folder is touched - the stubs shadow Get-PSSnapIn /
Add-PSSnapIn (so the snap-in guard passes), Get-Random (deterministic port) and
Start-Process (no browser). Everything else is the production code path.

The stub data deliberately contains characters a real directory allows and that
previously broke the app: an apostrophe, a double quote, an ampersand and HTML
markup in a DisplayName.
.Example
.\Test-ExchangeRecipientAdminCenter.ps1
.Example
.\Test-ExchangeRecipientAdminCenter.ps1 -Port 19200
#>
param(
    [string]$AppRoot = $PSScriptRoot,
    [int]$Port = 19100
)
$script = Join-Path $AppRoot "Start-ExchangeRecipientAdminCenter.ps1"
$base   = "http://localhost:$Port"

$prelude = {
    param($ScriptPath,$Port)
    # Values a real directory can legally hold: apostrophe, double quote, ampersand, markup.
    $DN   = 'Say "Hi" O''Brien & <b>Sons</b>'
    $SMTP = "o'brien@contoso.com"

    function Get-PSSnapIn {param([Parameter(ValueFromRemainingArguments)]$a) [pscustomobject]@{Name='s'}}
    function Add-PSSnapIn {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Start-Process {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Get-Random {param([Parameter(ValueFromRemainingArguments)]$a) $Port}
    function Get-User {param([Parameter(ValueFromRemainingArguments)]$a) [pscustomobject]@{UserPrincipalName="u@x.internal"}}
    function Get-RemoteMailbox {param([string]$Identity,[Parameter(ValueFromRemainingArguments)]$a)
        if ($Identity -and $Identity -notlike "*contoso.com*") { return $null }
        [pscustomobject]@{Name=$DN;DisplayName=$DN;Alias="ob";PrimarySmtpAddress=$SMTP
            RemoteRoutingAddress="ob@t.mail.onmicrosoft.com";RecipientTypeDetails="RemoteUserMailbox"
            WhenChanged="2026-08-26";HiddenFromAddressListsEnabled=$false
            EmailAddresses=@("SMTP:$SMTP","smtp:o'brien.alt@contoso.com")}}
    function Set-RemoteMailbox {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Enable-RemoteMailbox {param([string]$Identity,[string]$PrimarySMTPAddress,[string]$RemoteRoutingAddress,[Parameter(ValueFromRemainingArguments)]$a)
        Set-Content -Path (Join-Path $env:TEMP "erac_enable_test.txt") -Value "$PrimarySMTPAddress" }
    function Disable-RemoteMailbox {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Get-AcceptedDomain {param([string]$Identity,[Parameter(ValueFromRemainingArguments)]$a)
        @([pscustomobject]@{Name="contoso.com";DomainName="contoso.com";DomainType="InternalRelay";Default=$true},
          [pscustomobject]@{Name="t.mail.onmicrosoft.com";DomainName="t.mail.onmicrosoft.com";DomainType="InternalRelay";Default=$false}) |
          Where-Object { -not $Identity -or $_.Name -eq $Identity } }
    function Set-AcceptedDomain {param([Parameter(ValueFromRemainingArguments)]$a)}
    function New-AcceptedDomain {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Remove-AcceptedDomain {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Get-DistributionGroup {param([string]$Identity,[Parameter(ValueFromRemainingArguments)]$a)
        [pscustomobject]@{DisplayName=$DN;Alias="dl";PrimarySmtpAddress=$SMTP;RecipientTypeDetails="MailUniversalDistributionGroup";WhenCreated="2026-01-01"}}
    function Set-DistributionGroup {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Remove-DistributionGroup {param([string]$Identity,[Parameter(ValueFromRemainingArguments)]$a)
        Add-Content -Path (Join-Path $env:TEMP "erac_deletes_test.txt") -Value "group:$Identity" }
    function New-DistributionGroup {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Enable-DistributionGroup {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Get-Group {param([Parameter(ValueFromRemainingArguments)]$a)
        [pscustomobject]@{Name="Plain AD Group";DistinguishedName="CN=x";WindowsEmailAddress=$null}}
    function Get-MailContact {param([string]$Identity,[Parameter(ValueFromRemainingArguments)]$a)
        if ($Identity -and $Identity -notlike "*contoso.com*") { return $null }
        [pscustomobject]@{DisplayName=$DN;PrimarySmtpAddress=$SMTP;RecipientType="MailContact";ExternalEmailAddress="x@y.com"}}
    function Set-MailContact {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Remove-MailContact {param([string]$Identity,[Parameter(ValueFromRemainingArguments)]$a)
        Add-Content -Path (Join-Path $env:TEMP "erac_deletes_test.txt") -Value "contact:$Identity" }
    function New-MailContact {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Get-EmailAddressPolicy {param([string]$Identity,[Parameter(ValueFromRemainingArguments)]$a)
        [pscustomobject]@{Name="Default Policy";Priority=1;RecipientFilter="Company -eq 'A&B'"}}
    function Set-EmailAddressPolicy {param([Parameter(ValueFromRemainingArguments)]$a) throw "SIMULATED-FAILURE"}
    function New-EmailAddressPolicy {param([Parameter(ValueFromRemainingArguments)]$a)}
    function Remove-EmailAddressPolicy {param([string]$Identity,[Parameter(ValueFromRemainingArguments)]$a)
        Add-Content -Path (Join-Path $env:TEMP "erac_deletes_test.txt") -Value "policy:$Identity" }
    . $ScriptPath
}

$job = Start-Job -ScriptBlock $prelude -ArgumentList $script,$Port
foreach($i in 1..40){Start-Sleep -m 400; try{Invoke-WebRequest "$base/" -UseBasicParsing -TimeoutSec 3|Out-Null;break}catch{}}

$pass=0; $fail=0
function Req($m,$p,$b){
  $a=@{Uri="$base$p";Method=$m;UseBasicParsing=$true;TimeoutSec=20}
  if($b){$a.Body=$b;$a.ContentType="application/x-www-form-urlencoded"}
  try{ $r=Invoke-WebRequest @a
       $c = if($r.Content -is [byte[]]){[Text.Encoding]::UTF8.GetString($r.Content)}else{[string]$r.Content}
       [pscustomobject]@{Status=[int]$r.StatusCode;Body=$c;CT=[string]$r.Headers['Content-Type']} }
  catch{ $st=-1; if($_.Exception.Response){$st=[int]$_.Exception.Response.StatusCode}
       [pscustomobject]@{Status=$st;Body="";CT=""} } }
function Check($name,$cond,$detail){
  if($cond){$script:pass++; "  PASS  $name"} else {$script:fail++; "  FAIL  $name  --> $detail"} }

"=== 1. Dead routes now implemented (were 404, data lost on submit) ==="
foreach($r in @(
  @{p="/addcontact";           b="displayName=New+Contact&externalEmailAddress=n%40v.com"},
  @{p="/adddistributiongroup"; b="groupName=New+Group&groupEmail=g%40contoso.com&groupType=Distribution"},
  @{p="/mailenablegroup";      b="existingGroup=Plain+AD+Group&groupEmail=g%40contoso.com"},
  @{p="/addemailaddresspolicy";b="policyName=New+Policy&priority=2&recipientFilter=x+-eq+%27y%27"})){
  $x = Req POST $r.p $r.b
  Check "POST $($r.p) -> 200 and renders a page" ($x.Status -eq 200 -and $x.Body.Length -gt 500) "status=$($x.Status) bytes=$($x.Body.Length)"
}

"`n=== 2. Malformed URLs no longer return blank 200s ==="
foreach($u in @("/editremotemailbox","/editremotemailbox?id=","/editaccepteddomain?id=","/editemailaddresspolicy?id=","/editcontact?nonsense","/editdistributiongroup?id=")){
  $x = Req GET $u
  Check "GET $u -> real page, not blank" ($x.Status -eq 200 -and $x.Body.Length -gt 200) "bytes=$($x.Body.Length)"
}
$x = Req GET "/remotemailboxes?username=a&username=b"
Check "duplicate query key no longer blanks the page" ($x.Body.Length -gt 500) "bytes=$($x.Body.Length)"

"`n=== 3. No stale page served for unrouted requests ==="
$a = Req GET "/accepteddomains"
$b = Req POST "/" ""
Check "POST / does not echo the previous page" ($b.Body -ne $a.Body) "identical bodies ($($b.Body.Length) bytes)"
$c = Req GET "//"
Check "GET // does not echo the previous page" ($c.Body -ne $a.Body) "identical bodies"

"`n=== 4. Silent SMTP truncation fixed ==="
Req GET "/remotemailboxes?username=u%40x.internal&primarysmtpaddress_local=a%3Db&primarysmtpaddress_accepteddomain=contoso.com&remoteroutingaddress_local=r&remoteroutingaddress_accepteddomain=t.mail.onmicrosoft.com" | Out-Null
Start-Sleep -Milliseconds 300
$got = (Get-Content (Join-Path $env:TEMP "erac_enable_test.txt") -EA SilentlyContinue)
Check "'a=b' reaches Exchange intact as a=b@contoso.com" ($got -eq 'a=b@contoso.com') "Exchange got '$got'"
$x = Req GET "/remotemailboxes?foo=bar"
Check "stray query no longer fires a provisioning attempt" ($x.Body -notmatch 'enabled as Remote Mailbox') "banner appeared"

"`n=== 5. Escaping: guards intact, fields not truncated ==="
$x = Req GET "/editremotemailbox?id=o%27brien%40contoso.com"
Check "apostrophe alias keeps its confirm guard" ($x.Body -match 'data-confirm="Remove smtp:o&#39;brien\.alt') "guard markup not found"
Check "no raw apostrophe left inside a confirm() literal" ($x.Body -notmatch "confirm\('[^']*o'brien") "raw apostrophe still in JS"
Check "disable form carries expected addr in data-expect" ($x.Body -match 'data-expect="o&#39;brien@contoso\.com"') "data-expect missing"
Check "DisplayName double quote is encoded, not truncating" ($x.Body -match 'value="Say &quot;Hi&quot; O&#39;Brien &amp; &lt;b&gt;Sons&lt;/b&gt;"') "value attribute not encoded"
$x = Req GET "/contacts"
Check "markup in a DisplayName is inert in list rows" ($x.Body -notmatch '<b>Sons</b>' -and $x.Body -match '&lt;b&gt;Sons&lt;/b&gt;') "markup rendered live"

"`n=== 6. Error banner shows only the current failure ==="
$b1 = (Req POST "/editemailaddresspolicy" "Name=Default+Policy&Priority=1&RecipientFilter=x").Body
$b2 = (Req POST "/editemailaddresspolicy" "Name=Default+Policy&Priority=1&RecipientFilter=x").Body
$n1 = ([regex]::Matches($b1,'SIMULATED-FAILURE')).Count
$n2 = ([regex]::Matches($b2,'SIMULATED-FAILURE')).Count
Check "error is not repeated on the 2nd failure" ($n1 -eq 1 -and $n2 -eq 1) "1st=$n1 2nd=$n2 occurrences"

"`n=== 7. Result banner does not leak to later page loads ==="
Req POST "/addaccepteddomain" "domainName=x.com&domainType=InternalRelay" | Out-Null
$x = Req GET "/remotemailboxes"
Check "GET /remotemailboxes is clean after an unrelated POST" ($x.Body -notmatch 'class="alert') "stale banner present"

"`n=== 8. Headers, favicon, placeholders ==="
$x = Req GET "/accepteddomains"
Check "HTML responses declare Content-Type" ($x.CT -match 'text/html') "got '$($x.CT)'"
$x = Req GET "/favicon.ico"
Check "favicon served instead of 404" ($x.Status -eq 200) "status=$($x.Status)"
foreach($u in @("/","/remotemailboxes","/distributiongroups","/contacts","/emailaddresspolicies","/accepteddomains")){
  $x = Req GET $u
  $leak = [regex]::Matches($x.Body,'<!--\s*\{[a-z_]+\}\s*-->') | ForEach-Object {$_.Value}
  Check "no unreplaced placeholder on $u" ($leak.Count -eq 0) "$($leak -join ',')"
}
$x = Req GET "/distributiongroups"
Check "mail-enable group picker is populated" ($x.Body -match '<option value="Plain AD Group">') "dropdown still empty"

"`n=== 9. Still safe, still alive ==="
foreach($u in @("/../Start-ExchangeRecipientAdminCenter.ps1","/%2e%2e/Start-ExchangeRecipientAdminCenter.ps1")){
  $x = Req GET $u
  Check "traversal blocked: $u" ($x.Status -eq 404) "status=$($x.Status)"
}
$x = Req GET "/accepteddomains"
Check "server alive after the whole battery" ($x.Status -eq 200) "status=$($x.Status)"

"`n=== 10. Delete: guarded, server-verified, and actually performed ==="
$delLog = Join-Path $env:TEMP "erac_deletes_test.txt"
if (Test-Path $delLog) { Remove-Item $delLog -Force }

# every edit page must offer a delete behind a type-to-confirm guard
foreach($pg in @(
  @{u="/editcontact?id=o%27brien%40contoso.com";        a="/deletecontact"},
  @{u="/editdistributiongroup?id=o%27brien%40contoso.com"; a="/deletedistributiongroup"},
  @{u="/editemailaddresspolicy?id=Default%20Policy";    a="/deleteemailaddresspolicy"})){
  $x = Req GET $pg.u
  Check "delete form present on $($pg.u.Split('?')[0])" ($x.Body -match [regex]::Escape('action="' + $pg.a + '"')) "form missing"
  Check "  guarded by type-to-confirm + disabled button" ($x.Body -match 'data-expect="' -and $x.Body -match 'data-confirm-button') "guard attributes missing"
}

# a WRONG confirmation must delete nothing
Req POST "/deletecontact" "id=o%27brien%40contoso.com&confirmText=not-the-address" | Out-Null
Req POST "/deletedistributiongroup" "id=o%27brien%40contoso.com&confirmText=wrong" | Out-Null
Req POST "/deleteemailaddresspolicy" "id=Default+Policy&confirmText=wrong" | Out-Null
Start-Sleep -Milliseconds 200
Check "wrong confirmation deletes nothing" (-not (Test-Path $delLog)) "a Remove-* cmdlet ran: $(Get-Content $delLog -EA SilentlyContinue)"
$x = Req POST "/deletecontact" "id=o%27brien%40contoso.com&confirmText=not-the-address"
Check "wrong confirmation explains itself" ($x.Body -match "didn&#39;t match|didn't match") "no explanatory banner"

# the CORRECT confirmation must actually delete
Req POST "/deletecontact" "id=o%27brien%40contoso.com&confirmText=o%27brien%40contoso.com" | Out-Null
Req POST "/deletedistributiongroup" "id=o%27brien%40contoso.com&confirmText=o%27brien%40contoso.com" | Out-Null
Req POST "/deleteemailaddresspolicy" "id=Default+Policy&confirmText=Default+Policy" | Out-Null
Start-Sleep -Milliseconds 200
$done = @(Get-Content $delLog -EA SilentlyContinue)
Check "correct confirmation deletes the contact" ($done -contains "contact:o'brien@contoso.com") "log=$($done -join ',')"
Check "correct confirmation deletes the group"   ($done -contains "group:o'brien@contoso.com")   "log=$($done -join ',')"
Check "correct confirmation deletes the policy"  ($done -contains "policy:Default Policy")       "log=$($done -join ',')"
if (Test-Path $delLog) { Remove-Item $delLog -Force }

# deleting an object that is already gone must not explode
$x = Req POST "/deletecontact" "id=ghost%40nowhere.com&confirmText=ghost%40nowhere.com"
Check "deleting a missing object degrades gracefully" ($x.Status -eq 200 -and $x.Body -match 'not found') "status=$($x.Status)"

"`n================ $pass passed, $fail failed ================"
try{Invoke-WebRequest "$base/exit" -UseBasicParsing -TimeoutSec 5|Out-Null}catch{}
Start-Sleep -Seconds 1
Stop-Job $job -EA SilentlyContinue; Remove-Job $job -Force -EA SilentlyContinue
