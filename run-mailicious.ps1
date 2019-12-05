param 
(
      [Parameter(Mandatory = $true)] [string] $MailboxName 
    , [Parameter(Mandatory = $false)] [System.Net.NetworkCredential] $AdminCred
    , [Parameter(Mandatory = $false)] [string] $FolderInName = "Inbox"
    , [Parameter(Mandatory = $false)] [string] $FolderFinalName = "Incoming-Completed"
    , [Parameter(Mandatory = $false)] [string] $logtoFolder  = ("Recalls {0:yyyy-MM}" -f (Get-Date).ToUniversalTime())
    , [Parameter(Mandatory = $false)] [switch] $UseImpersonation
    , [Parameter(Mandatory = $false)] [switch] $SkipTrace
    , [Parameter(Mandatory = $false)] [string] $TraceFile
    , [Parameter(Mandatory = $false)] [string[]] $IgnoredMailboxesInTrace 
    , [Parameter(Mandatory = $false)] [string] $splunkHEC = 
    , [Parameter(Mandatory = $false)] [string] $splunkHECToken = 
    , [Parameter(Mandatory = $false)] [string] $splunkIndex = "userevents"
    , [Parameter(Mandatory = $false)] [string] $splunkSourcetype = "maliciousmail:userreported"
    , [Parameter(Mandatory = $false)] [int] $traceWindowHours = 12
    , [Parameter(Mandatory = $false)] [string[]] $IgnoredURLs = ("http://schemas.microsoft.com/office/2004/12/omml", "http://www.w3.org/tr/rec-html40")

)

Function ConvertEmlToCdo ($EmlItem)
{
    $msgString = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetString($emlItem.Content)
    $cdoMsg = New-Object -ComObject CDO.Message
    $cdomsgstream = $cdoMsg.GetStream()
    $cdomsgstream.WriteText($msgString)
    $cdomsgstream.Flush()
    $cdomsgstream.Close()

    return $cdoMsg
}


Function ExtractLinks($body)
{
    $foundurls = @()

    foreach($link in [regex]::matches($body, '(www|http:|https:)[^\s^"]+[\w]') ) 
    {
        $link = $link.Value.ToLower().Trim()
        if( !($foundurls.Contains($link)) -and !($IgnoredURLs.Contains($link)))
        {
            $foundurls += $link
        }
    }
    return $foundurls
}

Function AnalyzeItem($itemToAnalyze)
{
    try
    {
        $itemToAnalyze.Load()
    }
    catch
    {
        Write-Warning ("Could not open attachment {0}.  Error details below." -f $itemToAnalyze.Name)
        Write-Host $Error[0].Exception.InnerException
        return
    }

    $body = $null
    if($itemToAnalyze.Item.InternetMessageId -ne $null )
    {
        # Outlook dragged-in attachment 
        $itemToAnalyze =  $itemToAnalyze.Item                # dereference

        if($itemToAnalyze.SentOn -ne $null)
        {
            $script:reportedMessage = $itemToAnalyze   # this is probably the "root" message
         
            $script:internetMessageId = $itemToAnalyze.InternetMessageId
            $script:authResults = $itemToAnalyze.InternetMessageHeaders.Find("Authentication-Results")             
            $script:receivedSPF = $itemToAnalyze.InternetMessageHeaders.Find("Received-SPF")
        }

    }
    elseif ($itemToAnalyze.ContentType -eq "message/rfc822")    
    {
        # eml file 
        $itemToAnalyze = ConvertEmlToCdo $itemToAnalyze
        if($itemToAnalyze.SentOn -ne $null)
        {
            $script:reportedMessage = $itemToAnalyze   # this is probably the "root" message

            $script:internetMessageId = $itemToAnalyze.Fields.Item("urn:schemas:mailheader:message-id").Value
            $script:authResults = $itemToAnalyze.Fields.Item("urn:schemas:mailheader:authentication-results").Value
            $script:receivedSPF = $itemToAnalyze.Fields.Item("urn:schemas:mailheader:received-spf").Value
        }

    }

    elseif ($itemToAnalyze.InternetMessageHeaders -ne $null) # simple FW: message
    {
        $script:internetMessageId = $itemToAnalyze.InternetMessageHeaders.Find("References").value
    }
    elseif ($itemToAnalyze.ContentType -notlike "message/*")   # file attachments
    {
        if ($itemToAnalyze.Name -like "*.htm*" )
        {
            # HTML attachments
            $body = [System.Text.Encoding]::UTF8.GetString($itemToAnalyze.Content)

        }
        elseif ($itemToAnalyze.ContentType -eq "application/octet-stream" -and ($itemToAnalyze.Name -like "*.eml" -or $itemToAnalyze.Name -like "*.msg"))
        {                  
            # or "fake"/unsent email attachments                
            $emlAttachment = ConvertEmlToCdo( $itemToAnalyze )
            $body =  $emlAttachment.TextBody 
        }
        elseif ($itemToAnalyze.Name -like "*.pdf" -or $itemToAnalyze.ContentType -like "application/*")
        {
            # PDFs
        }

        # this wasn't a real message so assume it was an attachment
        $script:attachments += $itemToAnalyze.Name
    }

    # if this is a message object with attachments then recursively analyze those
    foreach($a in $itemToAnalyze.Attachments | ?{$_.ContentType -notlike "image/*"})
    {
        if($a.Name -ne $null -and $script:attachments -notcontains $a.Name)  # don't analyze it twice
        {
            $script:attachments += $itemToAnalyze.Name
            AnalyzeItem $a
        }
    }


    # Body extraction logic here - need to further refactor
    if($body -eq $null)
    {
        if($itemToAnalyze.Body.BodyType -eq [Microsoft.Exchange.WebServices.Data.BodyType]::HTML)
        {
            $body =   $itemToAnalyze.Body.Text  
        }
        elseif($itemToAnalyze.HTMLBody.length -gt 0)
        {
            $body = $itemToAnalyze.HTMLBody
        }
            
        if($body -eq $null)
        {
            $body = $itemToAnalyze.TextBody
        }
    }

    $script:urls += ExtractLinks ($body)
}


#  variable init
$traces = @()  # force array -  use the append (+=) operator below to maintain array
$msgGroups = @()
$msgGroupsByID = @()
$fromIPs = @()
$removed = $null

# Load cmdlets needed for Exchange on-prem and O365 interaction
& $PSScriptRoot\Import-Exchange-Cmdlets.ps1

# Load libraries needed for message processing
$lib = "$PSScriptRoot\microsoft.exchange.webservices.dll"
Add-Type -Path ($lib) 

#Connect to Exchange Web Service
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
if($UseImpersonation)
{
    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
}

if($AdminCred -ne $null)
{
#    Write-Host ("Connecting to EWS as {0}\{1}" -f $AdminCred.Domain,$AdminCred.UserName)
    $Service.Credentials = $AdminCred
}

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

$Service.AutodiscoverUrl($MailboxName,{$true})

#Fetch IDs of Well-known folders
$RootFolderID = New-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, (New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $MailboxName)))
$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$RootFolderID)
#$InboxFolderID = New-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, (New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $MailboxName)))
$DraftsFolder = New-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Drafts, (New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $MailboxName)))

if($RootFolder -eq $null)
{  Throw "Could not connect to Root folder of $MailboxName"  }

#Create a Folder View
$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow

#Retrieve Folders from view
$response = $RootFolder.FindFolders($FolderView)

if(($folderPickup = $response | ?{$_.DisplayName -eq $FolderInName}) -eq $null)
{
    Throw "The folder $FolderInName not found"
}

if(($FolderFinal = $response | ?{$_.DisplayName -eq $FolderFinalName}) -eq $null)
{
    Throw "The folder $FolderFinalName not found"
}

if ($folderPickup.TotalCount -lt 1)
{
    Write-Host "No items in " + $folderPickup.Name
    exit
} 

# filter for unread
$sfUnread = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead,$false)  

# sorting by oldest
$itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  # fetch only one message
$itemView.OrderBy.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, "Ascending")

:submission foreach($submission in $service.FindItems($folderPickup.Id, $sfUnread, $itemview))
{
    $submission.Load()  # this is the message we received from the reporting user

    $reportedMessage = $null # this will be the "carrier" message that actually came to the user

    # One submission may have multiple (including nested) attachments
    $itemsToAnalyze = @()
    try 
    {
        # arrays to hold indicators from the reported message
        $urls = @()
        $attachments = @()

        $itemsToAnalyze += $submission.Attachments | ?{$_.IsInline -ne $true -or $_.GetType().FullName -eq "Microsoft.Exchange.WebServices.Data.EmailMessage"}
    
        # if no attachements found, analyze this message
        if ($itemsToAnalyze.Count -lt 1)
        {
            $itemsToAnalyze += $submission
        }

        # process all of the items found to extract indicators
        :itemtoanalyze foreach($itemToAnalyze in $itemsToAnalyze | ?{$_.IsInline -ne $true} )   # weird negation to handle nulls
        {
            AnalyzeItem $itemToAnalyze

        }

        #  if no message ID could be found, take the base message
        if($internetMessageId -eq $null)
        {
            $reportedMessage = $submission
        }

        
        ## SPF review
        Write-Warning  "Sender authentication results"
        "Subject:  {0}" -f $reportedMessage.Subject
        "Sender:  {0}" -f $reportedMessage.From
        "MessageID:  {0}" -f $internetMessageID
        $receivedSPF
        $authResults

        Write-Host 


        ## URL review
        # Analyst picks URLs
        if($urls.Count -gt 0)
        {
            Write-Warning ("URLs found: ")
#            $urls  -replace "^http","hXXp" | % {$index=0} {$_; $index++} | Format-Table -Property @{ Label="index";Expression={$index}; Width=5 },*
            $urlTable = @()
            for($tableIdx=0; $tableIdx -lt $urls.Count; $tableIdx++) 
            {
                $urlTable += [pscustomobject] @{
                    index = $tableIdx
                    url = $urls[$tableIdx] -replace "^http","hXXp" 
                } 
            }
            $urlTable | Format-Table
            $urlTable = $null

            Write-Host "Choose URLs to report and block (enter comma-separated list of index numbers or enter A for all or N for none): " -ForegroundColor Yellow -NoNewline             
            $picks =  @((Read-Host).Split(','))
            if($picks -eq "A")
            {
                # no action,  take them all
            }
            elseif($picks -eq "N")
            {
                $urls = @()
                # take none
            }
            else
            {
                $pickedUrls = @()
                foreach($pick in $picks)
                {
                    $idx = [int] $pick
                    $pickedUrls += $urls[$idx]
                }

                $urls = $pickedUrls
            }

            Write-Host "Will include these URLs:"  -ForegroundColor Yellow
            $urls -replace "^http","hXXp" 
            Write-Host
        }
        else
        {
                Write-Host "No URLs found`r`n"  -ForegroundColor Yellow
        }


        ## Tracing
        if($SkipTrace)
        {
            Write-Warning "Skipping O365 message trace for $internetMessageID`r`n"
            $senderAddr = $reportedMessage.Fields['urn:schemas:httpmail:fromemail'].OriginalValue -replace '[<>]',''
        }
        else  # normal tracing
        {
            if($TraceFile -gt "")
            {
                Write-Warning "Using traces from $TraceFile"

                $traces = Import-Csv -Path $TraceFile

                $senderAddr = $traces[0].SenderAddress.ToLower()

            }
            else
            {
                Write-Warning "Starting O365 message trace for $internetMessageID`r`n"

                try
                {
                    # trace the reported message so that we can find the from address format and delivery times.    
                    if(!($traceInitialReport = Get-O365MessageTrace  -MessageId $internetMessageID  -PageSize 1 -StartDate (Get-Date).AddDays(-10)  -EndDate (Get-Date)))  #  Look back the full 10 days O365 allows because we're using message ID
                    {
                        Write-error ("Could not find a trace result for {0} from {1:u} to today" -f $internetMessageID,(Get-Date).AddDays(-10) )
                        Exit
                    } 

                    $senderAddr = $traceInitialReport.SenderAddress.ToLower()


                    # Now search for similar messages.
                    # Try to handle any wackiness in timestamps but still do a time-ranged search for good performance

                    #  Trace for all messages +/- n hours of initial delivery
                    $traces += Get-O365MessageTrace  -SenderAddress $senderAddr  -StartDate $traceInitialReport.Received.AddHours(-1 * $traceWindowHours) -EndDate $traceInitialReport.Received.AddHours($traceWindowHours) -Status Quarantined,Delivered -ErrorAction Stop  `
                        | ?{$IgnoredMailboxesInTrace -notcontains $_.RecipientAddress -and $_.Subject -eq $reportedMessage.Subject }
                    
                }
                catch
                {
                    Write-Warning ("Attachment could not be traced.   Investigate further: `"{0}`"" -f $reportedMessage.Name )
                    $Error[0]

                    Write-Host "Continue with message analysis? [y/n]" -ForegroundColor Yellow -NoNewline             
                    if ( ( Read-Host ) -notmatch "[yY]" ) 
                    { 
                        continue submission
                    }
                }
            }

            if($traces.count -lt 1)
            {
                $reportedMessage | fl *

                $senderAddr = "#find in authentication header#"
                $fromIP = "#find in authentication header#"
            }
            else
            {

                $msgGroups = $traces | Group-Object RecipientAddress
                $msgGroupsByID = $traces | Group-Object MessageID

                Write-Host ("Trace completed,   found {0}" -f $traces.count )   -ForegroundColor Yellow

                $traces | format-table  Received,SenderAddress,RecipientAddress,Subject,Status,MessageID
                $removed = 0

                Write-Host "Remove Matching Messages? [y/n]?" -ForegroundColor Yellow -NoNewline             
                if ( ($response = Read-Host ) -notmatch "[yY]" ) 
                {
                    if($response -eq "s")
                    {
                        Write-Warning "Skipping removal $internetMessageID - previously removed by analyst."
                    }
                    else
                    {
                        Write-Warning "Skipping $internetMessageID - process manually"
                        continue submission
                    }
                }
                else
                {
                    # Otherwise,  run the remove messages from mailboxes and quarantine

                    foreach($group in $msgGroupsByID)
                    {
                        $quarantines = @(Get-O365QuarantineMessage -MessageId $group.Name)
                        $quarantines | Delete-O365QuarantineMessage -Confirm:$false
                        $removed += $quarantines.count
                    }

                    foreach($msgGroup in $msgGroups)
                    {
                        # https://docs.microsoft.com/en-us/exchange/security-and-compliance/in-place-ediscovery/message-properties-and-search-operators
                        $query = "subject:`"{0}`" from:{1} received:{2:d}"  -f $msgGroup.group[0].subject,$msgGroup.group[0].SenderAddress,[DateTime]::Parse($msgGroup.group[0].Received)  # 
                        $result = Search-OnPremMailbox -Identity  $msgGroup.Name -SearchQuery $query  -LogLevel:Full -TargetMailbox $mailboxName -TargetFolder $logtoFolder -DeleteContent -Force 
                        $removed += $result.ResultItemsCount
                    }
                }

                $fromIPs = @($trace | Select-Object -Unique -ExpandProperty FromIP )
            }
        }


        $summary = [pscustomobject] @{
            Subject=$reportedMessage.Subject
            FromAddr=$reportedMessage.From #-replace '[^\p{L}\p{Nd}/(/}/_]', ''  # https://lazywinadmin.com/2015/08/powershell-remove-special-characters.html
            ReplyToAddr=$reportedMessage.ReplyTo
            SenderAddr=$senderAddr
            Attachments=$attachments -join "LINE_BREAK"
            URLs=($urls -join "LINE_BREAK" ) -replace "http","hXXp"
            ReceivedUTC="{0:o}" -f $reportedMessage.ReceivedTime.ToUniversalTime() # why did I use this ? -->> [DateTime]::Parse($msgGroups[0].Group.Received)
            ReportedUTC="{0:o}" -f $submission.DateTimeSent.ToUniversalTime()
            RecipientCount=$msgGroups.Count
            MessagesRemoved=$removed
            ReportedMessageID=$internetMessageId
            ReceivedSPF=$receivedSPF
            AuthenticationResults=$authResults
            OriginalReporter=$submission.From.Address
        }

        $replyEmail = New-Object  Microsoft.Exchange.WebServices.Data.EmailMessage($Service)
        $replyEmail.Subject = "[Suspicious Message Review]: `"{0}`"" -f $reportedMessage.subject
        $replyEmail.ToRecipients.Add($submission.From)
        $summaryTable = ConvertTo-Html -InputObject  $summary  -as List -Head $Header
        $summaryTable = $summaryTable.replace("LINE_BREAK","<p/>")
        $replyBody  = "Thank you for reporting this suspicious message.   We have analyzed the message and attempted to remove it from all recipients' mailboxes.<p/>If you did not click on any links or open any attachments then no further action is required.<p/>"  + ( $summaryTable)
        $replyEmail.Body = $replyBody -replace "`0", ""
        $replyEmail.Save($DraftsFolder )


        # reformat the attachments and URLs properties (now that we're finished with HTML display for the user)
        $summary.Attachments= $attachments
        $summary.URLs = @($urls)

        # Display for the analyst
        $summary | fl *

        # We don't show the full recipient list to the reporting user but do show it to the analyst and add it to the Splunk event
        Write-Host "Full Recipient List"
        $msgGroups.Name -join ";"
        $summary | Add-Member -NotePropertyName "Recipients" -NotePropertyValue @($msgGroups.Name)  # force to array
        $summary | Add-Member -NotePropertyName "MessageIDs" -NotePropertyValue @($msgGroupsByID.Name)
        $summary | Add-Member -NotePropertyName "FromIPs" -NotePropertyValue $fromIPs
        $summary | Add-Member -NotePropertyName "Analyst" -NotePropertyValue $(whoami.exe /upn)



    }
    catch 
    {        

        $submission.IsRead = $false
        $submission.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)

        [VOID] $submission.Move($folderPickup.ID)        

        $Error[0]
        $Error[0].Exception
        $Error[0].Exception.InnerException

        continue submission
    }   


    $submission.IsRead = $true
    $submission.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
    [VOID] $submission.Move($FolderFinal.ID)

    $splunkHeaders = @{Authorization = "Splunk $splunkHECToken"}
    $splunkHecUrl = "https://$splunkHEC/services/collector/event"
    $splunkSource = $($MyInvocation.MyCommand).ToString()


    # the Splunk payload
    $splunkData = [pscustomobject] @{
        index = $splunkIndex
        sourcetype = $splunkSourcetype
        host = $env:COMPUTERNAME
        source = $splunkSource
        time = (New-TimeSpan -Start (Get-Date "1970-01-01T00:00:00Z").ToUniversalTime() -End (Get-Date ).ToUniversalTime()).TotalSeconds
        event = $summary 
    } 

    $payload = ($splunkData | ConvertTo-Json -Depth 5) -replace "\\u003c","<" -replace "\\u003e",">"  # un-escape some JSON escaping for human clarity
            
    $result = Invoke-RestMethod -Uri $splunkHecUrl -Method Post -Headers $splunkHeaders -Body $payload
    if($result.text -eq "Success")
    {
        Write-Host "logged to Splunk"
    }
}
