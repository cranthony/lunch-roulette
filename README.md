# Lunch roulette

This is a script that sends emails to my EMBA classmates to suggest that they have lunch together.  The emails are sent to one person at a time, to encourage people to reach out to their match without me on the thread.

The work to send the email is done in Powershell, via Outlook.  Outlook is set up with my EMBA email account, so it made sense to me to send from there.  It may make more sense to send from a different domain, but we can work that out later, I guess.  It's not great that I wasted time figuring out how to get Outlook to work when I could have spent the same amount of time figuring out how to get a Cloud email service to work, but whatever.

The meat of the program will be the matching anyway.  Sending emails will hopefully always be straightforward.  And the sending isn't implemented yet.

## Running the program

### Sending an email by script

The Powershell script sends emails via Outlook.  You'll need to enable unsigned script execution to run this, but I'll let you Google that to figure it out.

You can send an email with a command line like this, but substitute a real email for the `nobody@` address below:

```Powershell
.\lunch-roulette-email.ps1 -email 'nobody@dontspamme.com' -friendlyName Chris -lunchDate 'Saturday, October 8, 2022' -otherEmail 'nobody@dontspamme.com' -otherFriendlyName 'NotChris' -otherFullName 'NotChris NotAnthony' -otherGender 'male'
```

I tested with Outlook already opened.  I borrowed some code from a blog that should theoretically open Outlook if it's not already open, but I didn't test that.