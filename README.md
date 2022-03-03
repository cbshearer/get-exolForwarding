# get-exolForwarding
There are 5 types of forwarding in MS Exchange Online. Two are properties of the mailbox and 3 are types of rules. 
Note: This script can take a long time to run in large organizations.

## Prerequisites
- MSOnline PowerShell module
  ```PowerShell
  Install-Module -Name MSOnline -Scope AllUsers
  ```
- ExchangeOnlineManagement PowerShell module
  ```PowerShell
  Install-Module -Name ExchangeOnlineManagement -Scope AllUsers
  ```


## This script will

- Make folder ```c:\temp``` if it doesn't exist
- Check if the MSOnline and ExchangeOnlineManagement modules are installed
- Get all of the mailboxes in your tenant
- Get every domain that is a part of your tenant (acceptedDomains)
- For every mailbox:
  - Check to see if it has a forwarding property (DeliverToMailboxAndForward, ForwardingSmtpAddress, ForwardingAddress)
  - If it does AND is going to a domain outside of your tenant accepted domains THEN export to CSV
  - Get every rule
    - For every rule:
      - See if it is a forwarding rule (ForwardTo , or ForwardAsAttachmentTo, or RedirectTo)
      - If does AND is going to a domain outside of your tenant accepted domains THEN export to CSV
- Results will be saved to the following CSV file formatted with date and time the script was run: 
```powershell
c:\temp\exolForwarding_DATE_TIME.csv
```
- Once finished running, the script will tell you:
  - The location of the output
  - The time the script started
  - The time the script finished

## The following properties are exported to the generated CSV file
- PrimarySMTPAddress of the mailbox
- DisplayName of the mailbox
- Rule priority (for a rule)
- Rule description (for a rule)
- Rule ID (for a rule)
- Rule enabled status (for a rule)
- Rule name (for a rule)
- Rule ForwardTo address (for a rule, if exists)
- Rule RedirectTo address (for a rule, if exists)
- Rule ForwardAsAttachmentTo address (for a rule, if exists)
- Mailbox DeliverToMailboxAndForward status (mailbox property, if exists)
- Mailbox ForwardingSmtpAddress (mailbox property, if exists)
- Mailbox ForwardingAddress (mailbox property, if exists)
