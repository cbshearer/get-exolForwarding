# get-exolForwarding
There are 6 types of forwarding in MS Exchange Online. Three are properties of the mailbox and 3 are types of rules. 

## This script will

- Get all of the mailboxes in your tenant
- Get every rule for every mailbox
- Get every domain that is a part of your tenant (acceptedDomains)
- For every mailbox:
  - Check to see if it has a forwarding property (DeliverToMailboxAndForward, ForwardingSmtpAddress, ForwardingAddress)
  - If it does AND is going to a domain outside of your tenant accepted domains THEN log it
- For every rule:
  - See if it is a forwarding rule (DeliverToMailboxAndForward , or ForwardingAddress, or ForwardingSmtpAddress)
  - If does AND is going to a domain outside of your tenant accepted domains
  - If it is then log it

## The following properties are logged
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
- Mailbox DeliverToMailboxAndForward status (for mailbox property, if exists)
- Mailbox ForwardingSmtpAddress (for mailbox property, if exists)
- Mailbox ForwardingAddress (for mailbox property, if exists)
