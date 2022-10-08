# Lunch roulette

This is a script that sends emails to my EMBA classmates to suggest that they
have lunch together.  The emails are sent to one person at a time, to encourage
people to reach out to their match without me on the thread.

## Technologies

User data is stored within an XLSX file because I used Excel to track my
classmates and fill in information like their correct email addresses.  I
considered using a local database -- specifically SQLite -- to store the user
data, but decided against it because it's harder to use, and because I don't
expect any Cloud application to use SQL.  Google Cloud, for example, only has a
free tier in Firestore, a NoSQL database.

The work to send the email is done in Outlook, via Powershell.  Outlook is set
up with my EMBA email account, so it made sense to me to send from there.  It
may make more sense to send from a different domain, but we can work that out
later, I guess.  It's not great that I wasted time figuring out how to get
Outlook to work when I could have spent the same amount of time figuring out how
to get a Cloud email service to work, but whatever.

The meat of the program is implemented in Python.  There was no particular
reason for this.  Python is well supported and is easy enough to write.

## Running the program

The entry point for the program is `lunch-roulette.py`.  It has usage.  It uses
Python 3, and was built using `venv` to manage its dependencies.  To initialize
the virtual environment:

In Powershell:
```Powershell
python3 -m venv venv
.\venv\Scripts\activate.bat
python3 -m pip install -r requirements.txt
```

In Bash:
```Bash
python3 -m venv venv
source ./venv/bin/activate
python3 -m pip install -r requirements.txt
```

### Initializing your user data

The program expects your XLSX file to have columns like the following:

|email|friendly_name|full_name|gender|cluster|year|
| --- | ---         | ---     | ---  | ---   | -- |
| me@example.com | Me! | Meschievous R Cunningham | male | E | 2024 |
| you@example.com | You! | Youstice M Tallybottom | female | M | 2024 |

A `frequency` column is optional, and will be interpreted to contain the
frequency, measured in `1/class weeks`, that the person wants to be in the
roulette.  For example, a frequency of 1 means they'll be matched every class
week, while a frequency of 1/4 means they'll be matched once every 4 class
weeks.

### Running the roulette

To run the roulette, run the script with the `--roulette` option:

```Powershell
python3 .\lunch-roulette.py --xlsx my.xlsx --lunch-date 20221008 --roulette
```

The XLSX file will be updated with a column named `match_20221008` that contains
each person's match for that day.  Note that the matches are not sent out, so
that you can review and edit the matches as you'd like.  If a `match_20221008`
column already exists in the XLSX, it will be overwritten.

### Sending emails for a lunch roulette

To send the emails for a lunch roulette, use the `--send-emails` option.  This
should only be used after the XLSX has already been filled with the matches for
the given lunch date.

```
python3 .\lunch-roulette.py --xlsx my.xlsx --lunch-date 20221008 --send-emails
```

### Development only: directly sending an email by script

This scenario is for developers to test the email generation.  The main lunch
roulette script can already call the needed Powershell script with the proper
arguments, so this shouldn't need to be called once you've proven to yourself
that it works.  This can be handy, though, if you've updated the email template
and want to send emails to yourself to validate that they look as you expect.

The Powershell script sends emails via Outlook.  You'll need to enable unsigned
script execution to run this, but I'll let you Google that to figure it out.

You can send an email with a command line like this, but substitute a real email
for the `nobody@` address below:

```Powershell
.\lunch-roulette-email.ps1 `
    -email 'nobody@dontspamme.com' `
    -friendlyName Chris `
    -lunchDate 'Saturday, October 8, 2022' `
    -otherEmail 'nobody@dontspamme.com' `
    -otherFriendlyName 'NotChris' `
    -otherFullName 'NotChris NotAnthony' `
    -otherGender 'male'
```

Outlook should already be opened, before running the script.