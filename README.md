![example](https://github.com/creetz/Start-MessageTrackingLog/blob/master/pic1.png)

# Start-MessageTrackingLog
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
