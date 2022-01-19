Import-Module msonline                 # for connect-msolservice
Import-Module exchangeonlinemanagement # for connect-exchangeonline

## set TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 

Connect-MsolService
Connect-ExchangeOnline

$startTime = get-date
$mailboxes = get-mailbox -ResultSize unlimited
$reportOut = "c:\temp\exolForwarding_$((get-date).ToString('yyyyMMdd_hhmmss')).csv"
$mailboxCounter = 0
$domains = Get-AcceptedDomain

Write-Host "Mailboxes to evaluate  :" $mailboxes.count
Write-Host "Report will be saved as:" $reportOut

foreach ($mailbox in $mailboxes)

    {
        $mailboxCounter = $mailboxCounter + 1
        Write-Host $mailboxCounter -nonewline
        $rules = get-inboxrule -mailbox $mailbox.PrimarySmtpAddress
        $recipients = @()
        
        ## if the mailbox has one of these 3 attributes, we add them to our array
        if($mailbox.DeliverToMailboxAndForward -or $mailbox.ForwardingAddress -or $mailbox.ForwardingSmtpAddress)
            {
                $recipients += $mailbox.ForwardingAddress     | Where-Object {$_ -match "SMTP"}
                $recipients += $mailbox.ForwardingSmtpAddress | Where-Object {$_ -match "SMTP"}

                ## evaluate each recipient to determine if it is external
                foreach ($recipient in $recipients)
                {
                    $email = ($recipient -split "SMTP:")[1].Trim("]")
                    $domain = ($email -split "@")[1]

                    ## if it is external then we gather information about it and insert it into the report (rules are added empty so we can see them all in one report)
                    if ($domains.domainname -notcontains $domain)
                        {
                            $ruleReport = new-object PSCustomObject
                            Write-Host $mailbox.PrimarySmtpAddress
                            $ruleReport | Add-Member -MemberType NoteProperty -name PrimarySmtpAddress          -Value $mailbox.PrimarySmtpAddress
                            $ruleReport | Add-Member -MemberType NoteProperty -name DisplayName                 -Value $mailbox.DisplayName
                            $ruleReport | Add-Member -MemberType NoteProperty -name RulePriority                -Value ""
                            $ruleReport | Add-Member -MemberType NoteProperty -name RuleDescription             -Value ""
                            $ruleReport | Add-Member -MemberType NoteProperty -name RuleId                      -Value ""
                            $ruleReport | Add-Member -MemberType NoteProperty -name RuleEnabled                 -Value ""
                            $ruleReport | Add-Member -MemberType NoteProperty -name RuleName                    -Value ""
                            $ruleReport | Add-Member -MemberType NoteProperty -name ForwardTo                   -Value ""
                            $ruleReport | Add-Member -MemberType NoteProperty -name RedirectTo                  -Value ""
                            $ruleReport | Add-Member -MemberType NoteProperty -name ForwardAsAttachmentTo       -Value ""
                            $ruleReport | Add-Member -MemberType NoteProperty -name DeliverToMailboxAndForward  -Value $mailbox.DeliverToMailboxAndForward
                            $ruleReport | Add-Member -MemberType NoteProperty -name ForwardingSmtpAddress       -Value $mailbox.ForwardingSmtpAddress
                            $ruleReport | Add-Member -MemberType NoteProperty -name ForwardingAddress           -Value $mailbox.ForwardingAddress
                        
                            $rulereport | Export-Csv -Path $reportOut -Append -NoTypeInformation
                        }
                }
            }
        
        if ($rules)
            {
                ## if the rule has one of these forward/redirect attributes add it to the array 
                foreach ($rule in $rules)
                    {
                        $recipients = @()
                        $recipients += $rule.ForwardTo             | Where-Object {$_ -match "SMTP"}
                        $recipients += $rule.ForwardAsAttachmentTo | Where-Object {$_ -match "SMTP"}
                        $recipients += $rule.RedirectTo            | Where-Object {$_ -match "SMTP"}

                        ## evaluate each recipient to see if is external
                        foreach ($recipient in $recipients)
                        {
                            $email  = ($recipient -split "SMTP:")[1].Trim("]")
                            $domain = ($email -split "@")[1]
                            
                            ## if it is external then we gather information and insert it into the report (mailbox attributes are added as null so we can see them as separate line items in the report)
                            if ($domains.domainname -notcontains $domain)
                                {
                                    $ruleReport = new-object PSCustomObject
                                    Write-Host $mailbox.PrimarySmtpAddress
                                    $ruleReport | Add-Member -MemberType NoteProperty -name PrimarySmtpAddress          -Value $mailbox.PrimarySmtpAddress
                                    $ruleReport | Add-Member -MemberType NoteProperty -name DisplayName                 -Value $mailbox.DisplayName
                                    $ruleReport | Add-Member -MemberType NoteProperty -name RulePriority                -Value $rule.Priority
                                    $ruleReport | Add-Member -MemberType NoteProperty -name RuleDescription             -Value $rule.Description
                                    $ruleReport | Add-Member -MemberType NoteProperty -name RuleId                      -Value $rule.identity
                                    $ruleReport | Add-Member -MemberType NoteProperty -name RuleEnabled                 -Value $rule.enabled
                                    $ruleReport | Add-Member -MemberType NoteProperty -name RuleName                    -Value $rule.name
                                    $ruleReport | Add-Member -MemberType NoteProperty -name ForwardTo                   -Value $rule.ForwardTo
                                    $ruleReport | Add-Member -MemberType NoteProperty -name RedirectTo                  -Value $rule.RedirectTo
                                    $ruleReport | Add-Member -MemberType NoteProperty -name ForwardAsAttachmentTo       -Value $rule.ForwardAsAttachmentTo
                                    $ruleReport | Add-Member -MemberType NoteProperty -name DeliverToMailboxAndForward  -Value ""
                                    $ruleReport | Add-Member -MemberType NoteProperty -name ForwardingSmtpAddress       -Value ""
                                    $ruleReport | Add-Member -MemberType NoteProperty -name ForwardingAddress           -Value ""
                                
                                    $rulereport | Export-Csv -Path $reportOut -Append -NoTypeInformation
                                }
                        }
                    }                
            }
        

    }
    
    $endTime = get-date
    Write-Host "`nResults   : " $reportOut
    Write-Host "Start Time: " $startTime
    Write-Host "End   Time: " $endTime
