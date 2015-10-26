IMAP Commandline Client Automation
----------------------------------

IMAP Automation Script

Usage Examples:

% imap server=imap.gmail.com user=... folder=Inbox search="FROM mowli SUBJECT urgent"

% imap server=imap.gmail.com user=... folder=Inbox search="FROM mowli SUBJECT urgent" action=move_to_mail_folder:Trash

Works with gmail as well as MS Exchange

See link for more IMAP keywords

https://www.fastmail.com/help/receive/search.html

P.S: 
1) Make sure IO::Socket:SSL is installed!!!
  This is required for some emails services that make secure authentication, 
  but a failing script does not tell you the root cause!
  
2) IMAP servers for popular email providers
      Gmail : imap.gmail.com
      
      Yahoo: imap.mail.yahoo.com
      
      Hotmail: imap-mail.outlook.com
