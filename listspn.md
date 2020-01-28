Full article and background here: List all SPNs Used in your Active Directory
URL:http://social.technet.microsoft.com/wiki/contents/articles/18996.list-all-spns-used-in-your-active-directory.aspx
 

There are a lot of hints & tips out there for troubleshooting SPNs (service principal names).

Listing duplicate SPNs is fairly easy, use the “setspn -X” command and you’ll find out.

But how do you find out which SPNs are used for which users and computers are used for this?

So you need a general script to list all SPNs, for all users and all computers…

Nice fact to know, SPNs are set as an attribute on the user or computer accounts.
So that makes it fairly ease to query for that attribute.

And modern admins do PowerShell, right?

 

How to use the script: to avoid certificate issues, download the txt file and remove the txt extension. Then use the PS1 file to list the SPNs.
