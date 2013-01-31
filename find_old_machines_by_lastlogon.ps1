################################################################################
## 				Stale Computer Account Finding Script			 			  ##
################################################################################
##Find all computer accounts with a last logon of 60 days or more using the   ##
##lastlogontimestamp attribute. Machine accounts with a creationdate of less  ##
##than three weeks ago should also be filtered out because the 				  ##	
##lastlogontimestamp might not be populated yet.					 		  ##	
##																			  ##	
##The script will create a CSV file with machines fitting this criteria,      ##
##and email that file.														  ##
##																			  ##
##This script will need to run in a context that has permission to disable any##
##user account under your root 						                          ##	
##The script requires the Quest AD Management tools installed				  ##	
##																			  ##
##Created by David Hahn : 9/10/2010											  ##
################################################################################

## add snapin if not already added. - can remove this if you're using the new
## built in AD stuff
Add-PSSnapin Quest.ActiveRoles.ADManagement 
##Define new export-csv function that allows for appending to CSV files, the 
##native export-csv does not do this.
<#
  This Export-CSV behaves exactly like native Export-CSV
  However it has one optional switch -Append
  Which lets you append new data to existing CSV file: e.g.
  Get-Process | Select ProcessName, CPU | Export-CSV processes.csv -Append

  For details, see

http://dmitrysotnikov.wordpress.com/2010/01/19/export-csv-append/

  (c) Dmitry Sotnikov
#>
function Export-CSV2 {
[CmdletBinding(DefaultParameterSetName='Delimiter',
  SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
 [Parameter(Mandatory=$true, ValueFromPipeline=$true,
           ValueFromPipelineByPropertyName=$true)]
 [System.Management.Automation.PSObject]
 ${InputObject},

 [Parameter(Mandatory=$true, Position=0)]
 [Alias('PSPath')]
 [System.String]
 ${Path},

 #region -Append (added by Dmitry Sotnikov)
 [Switch]
 ${Append},
 #endregion 

 [Switch]
 ${Force},

 [Switch]
 ${NoClobber},

 [ValidateSet('Unicode','UTF7','UTF8','ASCII','UTF32',
                  'BigEndianUnicode','Default','OEM')]
 [System.String]
 ${Encoding},

 [Parameter(ParameterSetName='Delimiter', Position=1)]
 [ValidateNotNull()]
 [System.Char]
 ${Delimiter},

 [Parameter(ParameterSetName='UseCulture')]
 [Switch]
 ${UseCulture},

 [Alias('NTI')]
 [Switch]
 ${NoTypeInformation})

begin
{
 # This variable will tell us whether we actually need to append
 # to existing file
 $AppendMode = $false

 try {
  $outBuffer = $null
  if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
  {
      $PSBoundParameters['OutBuffer'] = 1
  }
  $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Export-Csv',
    [System.Management.Automation.CommandTypes]::Cmdlet)

 #String variable to become the target command line
 $scriptCmdPipeline = ''

 # Add new parameter handling
 #region Dmitry: Process and remove the Append parameter if it is present
 if ($Append) {

  $PSBoundParameters.Remove('Append') | Out-Null

  if ($Path) {
   if (Test-Path $Path) {
    # Need to construct new command line
    $AppendMode = $true

    if ($Encoding.Length -eq 0) {
     # ASCII is default encoding for Export-CSV
     $Encoding = 'ASCII'
    }

    # For Append we use ConvertTo-CSV instead of Export
    $scriptCmdPipeline += 'ConvertTo-Csv -NoTypeInformation '

    # Inherit other CSV convertion parameters
    if ( $UseCulture ) {
     $scriptCmdPipeline += ' -UseCulture '
    }
    if ( $Delimiter ) {
     $scriptCmdPipeline += " -Delimiter '$Delimiter' "
    } 

    # Skip the first line (the one with the property names) 
    $scriptCmdPipeline += ' | Foreach-Object {$start=$true}'
    $scriptCmdPipeline += '{if ($start) {$start=$false} else {$_}} '

    # Add file output
    $scriptCmdPipeline += " | Out-File -FilePath '$Path'"
    $scriptCmdPipeline += " -Encoding '$Encoding' -Append "

    if ($Force) {
     $scriptCmdPipeline += ' -Force'
    }

    if ($NoClobber) {
     $scriptCmdPipeline += ' -NoClobber'
    }
   }
  }
 } 

 $scriptCmd = {& $wrappedCmd @PSBoundParameters }

 if ( $AppendMode ) {
  # redefine command line
  $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock(
      $scriptCmdPipeline
    )
 } else {
  # execute Export-CSV as we got it because
  # either -Append is missing or file does not exist
  $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock(
      [string]$scriptCmd
    )
 }

 # standard pipeline initialization
 $steppablePipeline = $scriptCmd.GetSteppablePipeline(
        $myInvocation.CommandOrigin)
 $steppablePipeline.Begin($PSCmdlet)

 } catch {
   throw
 }

}

process
{
  try {
      $steppablePipeline.Process($_)
  } catch {
      throw
  }
}

end
{
  try {
      $steppablePipeline.End()
  } catch {
      throw
  }
}
<#

.ForwardHelpTargetName Export-Csv
.ForwardHelpCategory Cmdlet

#>

}

##path to the output log file.
$reportfile = "c:\temp\oldmachines.csv" 
##This is the date that the machine accounts will be deleted. Use for Display in
##email.
$futuredate = (Get-Date).adddays(30).toShortDateString()
##smtp server name
$smtpServer = "smtpserver" 
##This is used to filter out machines that have been recently created, 
##so the lastlogontimestamp might not have been created yet.
$threeweeksago = (Get-Date).adddays(-21)
##define where in AD to begin the search
$searchroot = 'DC=contoso,DC=com'

$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

##from address
$msg.From = "fromaddress@contoso.com" 
##to address. separate multiple addresses with a comma
$msg.To.Add("recipient@contoso.com")
##subject of the email
$msg.Subject = "Stale AD Computer Accounts Report"
##create message body using a here string
$msg.Body = @"
The attached CSV contains details on computer accounts that have not booted on the network in the last 60 days and who's machine account has existed for at least 21 days. Machines with no age have never been booted on the network. 

These accounts have been disabled.

If there are accounts in this list that need to remain, please re-enable the accounts before $futuredate. If they are not re-enabled, they will be deleted.

Accounts still disabled in 30 days will be deleted in order to clean up Active Directory.

"@

##Delete old report file##
del $reportfile

##Fancy one-liner to get the old objects. I tried using the -inactive parameter for
##get-qadcomputer, however, it proved to be slower than using this oneliner.
##not sure why..
$oldmachines = get-qadcomputer -IncludedProperties lastlogontimestamp -searchroot "$searchroot" -SizeLimit 0 | 
Where-Object {($_.lastlogontimestamp -lt (get-date).AddDays(-60)) -and ($_.creationdate -lt $threeweeksago) } 
foreach ($oldmachine in $oldmachines) {
	select -input $oldmachine -property name,@{n="OU";e={$_.CanonicalName.Split("/")[4]}},osname,description,creationdate,lastlogontimestamp,@{n="Days Since Last Logged On";e={((get-date)- $_.lastlogontimestamp).days}} | 
	Export-CSV2 "$reportfile" -append
}

##instantiate new attachment object and send the email
$att = new-object Net.Mail.Attachment($reportfile)
$msg.Attachments.Add($att)
$smtp.Send($msg)
$att.Dispose() ##garbage collection

##Disable the objects after sending the email
##The current code only does this in "what if" mode. To actually disable the
##accounts, remove the -whatif paramter from disable-qadcomputer
##
##We couldn't put this in the pipeline above because disable-qadcomputer consumes
##objects and does not pass them through the pipeline

foreach($oldmachine in $oldmachines) {Disable-QADComputer $oldmachine -WhatIf}
