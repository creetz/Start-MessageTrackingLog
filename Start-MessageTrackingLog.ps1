<#
    .SYNOPSIS
    Comfortable MessagetrackingLog
   
   	Christian Reetz
    (Updated by Christian Reetz to support Exchange Server 2013)
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	26.05.2017
	
    .DESCRIPTION

    Comfortable MessagetrackingLog
	
	PARAMETER start
    Start Date in english format | When not defindes today minus 7 days will be used
	
	PARAMETER end
	End Date in english format | When not defindes today 23:59 a clock will be used
	
	PARAMETER sender
	Sender E-Mail Address

    PARAMETER recipient	
    Recipient E-Mail Address

    PARAMETER subject (out-of-order)
    Subject | f.E. *project1*

    PARAMETER germantimeformat
    $true or keep clear
		
	EXAMPLES
    .\Start-MessagetrackingLog.ps1 -start 01/31/2016 -end 02/28/2016 -sender max.mustermann@contoso.local
    .\Start-MessagetrackingLog.ps1 -sender max.mustermann@contoso.local
    .\Start-MessagetrackingLog.ps1 -sender max.mustermann@contoso.local -mailboxservers server1, server2,server3
    .\Start-MessagetrackingLog.ps1 -sender max.mustermann@contoso.local -start "23.05.2017 07:00" -end "23.05.2017 08:00" -germantimeformat $true

    #>
#$mailboxserver = @()

Param (
	[Parameter(Mandatory=$false)] $start,
	[Parameter(Mandatory=$false)] $end,
	[Parameter(Mandatory=$false)] $sender,
    [Parameter(Mandatory=$false)] $recipient,
    [Parameter(Mandatory=$false)] $subject,
    [Parameter(Mandatory=$false)] $germantimeformat,
    [Parameter(Mandatory=$false)] [string[]]$mailboxserver
)
$export = "c:\temp\messagetrackinglog$(Get-Date -Format ddMMyyyyHHmm).log"

if (!($start) -and !($end) -and !($sender) -and !($recipient) -and !($subject))
{
Write-Host -ForegroundColor white "No parameter was specified! Please define."
$germantimeformat = read-host "German timeformat? (keep clear for no)"
$start = read-host "define start-date"
$end = read-host "define end-date"
$sender = read-host "define sender"
$recipient = read-host "define recipient"
$subject = read-host "define subject"
} 

if (($germantimeformat) -and $start)
{
    $startday = $start.Split('.')[0]
    $startmonth = $start.Split('.')[1]
    $startyear = $start.Split('.')[2].split(' ')[0]
    $starthour = $start.Split('.')[2].split(' ')[1].split(':')[0]
    $startminute = $start.Split('.')[2].split(' ')[1].split(':')[1]
    $start = get-date -Year $startyear -Month $startmonth -Day $startday -Hour $starthour -Minute $startminute
}

if (($germantimeformat) -and $end)
{
    $endday = $end.Split('.')[0]
    $endmonth = $end.Split('.')[1]
    $endyear = $end.Split('.')[2].split(' ')[0]
    $endhour = $end.Split('.')[2].split(' ')[1].split(':')[0]
    $endminute = $end.Split('.')[2].split(' ')[1].split(':')[1]
    $end = get-date -Year $endyear -Month $endmonth -Day $endday -Hour $endhour -Minute $endminute
}

if (-not ($start))
{
$start = (Get-Date).Adddays(-7).ToString("MM'/'dd'/'yyyy")
}

if (-not ($end))
{
$end = (Get-Date).Adddays(1).ToString("MM'/'dd'/'yyyy")
}

$maillog = @()

if ($mailboxserver)
{
    $mbxserver = $mailboxserver
}
else
{
    $mbxserver = (Get-Mailboxserver | Select-Object name).name
}

foreach ($server in $mbxserver)
{
    if ($sender -and !($recipient))
    {
        $maillog += Get-MessageTrackingLog -Server $server -Resultsize Unlimited -Start $start -End $end -Sender $Sender -erroraction inquire | Sort-Object Timestamp
    }

    if ($recipient -and !($sender))
    {
        $maillog += Get-MessageTrackingLog -Server $server -Resultsize Unlimited -Start $start -End $end -Recipient $Recipient -erroraction inquire | Sort-Object Timestamp
    }

    if ($sender -and $recipient)
    {
        $maillog += Get-MessageTrackingLog -Server $server -Resultsize Unlimited -Start $start -End $end -Recipient $recipient -Sender $Sender -erroraction inquire | Sort-Object Timestamp
    }

    if ($subject -and !($sender) -and !($recipient))
    {
        $maillog += Get-MessageTrackingLog -Server $server -Resultsize Unlimited -Start $start -End $end -MessageSubject "$subject" -erroraction inquire | Sort-Object Timestamp
    }
}

if ($subject -and ($recipient -or $sender))
{
    $maillog = $maillog | where-object {$_.MessageSubject -like "$subject"}
    $org_maillog = $maillog
}
else
{
    $org_maillog = $maillog
}

#Areyousure function. Alows user to select y or n when asked to exit. Y exits and N returns to main menu.  
 function areyousure {$areyousure = read-host "Are you sure you want to exit? (y/n)"  
           if ($areyousure -eq "y"){exit}  
           if ($areyousure -eq "n"){mainmenu}  
           else {write-host -foregroundcolor red "Invalid Selection"   
                 areyousure  
                }  
} 

#undo changes and filters
function undo{
$maillog = $org_maillog
$fl = "0"
$wrap = "0"
mainmenu 
}

#filter for deliver
function deliver{
 cls
 $maillog = $maillog | where-object {$_.EventID -eq "DELIVER"} 
mainmenu 
}

#filter for send
function send{
 cls
 $maillog = $maillog | where-object {$_.EventID -eq "send"} 
mainmenu 
}

#filter for dsn
function dsn{
 cls
 $maillog = $maillog | where-object {$_.EventID -eq "DSN"}
 if ($maillog.count -eq 0)
 {
 $maillog = $org_maillog
 $fl = "0"
 $wrap = "0"
 }
mainmenu 
}

#filter for fail
function fail{
 cls
 $maillog = $maillog | where-object {$_.EventID -eq "fail"}
 if ($maillog.count -eq 0)
 {
 $maillog = $org_maillog
 $fl = "0"
 $wrap = "0"
 }
mainmenu 
}

#filter for subject
function subject{
 $subject = read-host "Enter subject"
 $maillog = $maillog | where-object {$_.MessageSubject -like "$subject"}
 mainmenu 
}

#filter for subject
function export{
 cls
 $wrap = "1"
mainmenu 
}

#filter for subject
function fulllist{
 cls
 $fl = "1"
mainmenu 
}

function gridview{
cls
$maillog| select-object Timestamp,EventId,Sender,Recipients,MessageSubject,TotalBytes,MessageId,ClientIp,ClientHostname,ServerIp,ServerHostname,SourceContext,ConnectorId,Source,InternalMessageId,RecipientStatus,RecipientCount,ReturnPath,Directionality,TenantId,OriginalClientIp,MessageInfo,MessageLatency,MessageLatencyType,EventData | Out-GridView -Title "MessageTrackingLog"
mainmenu
}

$wrap = "0";

 #Mainmenu function. Contains the screen output for the menu and waits for and handles user input.  
 function mainmenu{
 cls
 if ($wrap -eq "1")
 {
 $maillog | ft Timestamp, EventId, Sender, Recipients, MessageSubject, TotalBytes, MessageId, serverhostname -wrap > "$export"
 $maillog | ft Timestamp, EventId, Sender, Recipients, MessageSubject, TotalBytes, MessageId, serverhostname
 Write-Host -ForegroundColor Yellow "Export done! ($export)"
 $wrap = "0"
 }
 else
 {  
    if ($fl -eq "1")
    {
    $maillog | fl
    $fl = "0"
    }
    else
    {
    $maillog | ft Timestamp, EventId, Sender, Recipients, MessageSubject 
    }
 }

 echo "---------------------------------------------------------"  
 echo "    1. Filter for EventID DELIVER"  
 echo "    2. Filter for EventID SEND"  
 echo "    3. Filter for EventID DSN"
 echo "    4. Filter for EventID FAIL"   
 echo "    5. Filter for Subject"  
 echo "    6. Export to Transfer"
 echo "    7. Show FullList"
 echo "    8. Output to GridView"
 echo "    9. undo Filter"  
 echo "    0. Exit"  
 echo "---------------------------------------------------------"  
 echo ""  

 echo ""  
 $answer = read-host "Please Make a Selection"  
 if ($answer -eq 1){deliver}  
 if ($answer -eq 2){send}  
 if ($answer -eq 3){dsn}  
 if ($answer -eq 4){fail} 
 if ($answer -eq 5){subject}  
 if ($answer -eq 6){export}  
 if ($answer -eq 7){fulllist}
 if ($answer -eq 8){gridview}
 if ($answer -eq 9){undo}  
 if ($answer -eq 0){areyousure}
 else {write-host -ForegroundColor red "Invalid Selection"  
       sleep 5  
       mainmenu  
      }  
                }  
 mainmenu