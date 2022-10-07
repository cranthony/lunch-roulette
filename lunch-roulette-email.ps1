# The emailing structure of this was from a Microsoft dev blog at
# https://devblogs.microsoft.com/premier-developer/outlook-email-automation-with-powershell/
#
# And then I looked up a bunch of tutorials on Powershell to get the rest.
param(
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [String]$email,
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [String]$friendlyName,
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [String]$lunchDate,
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [String]$otherEmail,
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [String]$otherFriendlyName,
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [String]$otherFullName,
  [Parameter(Mandatory=$true)] [ValidateSet('male', 'female', 'nonbinary')] [String]$otherGender
)

$emailTemplatePath = ".\"
$emailTemplateName = "emba-lunch-roulette-email-template.oft"

# Connect to Outlook.
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
add-type -assembly "System.Runtime.Interopservices"
try
{
  $outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
  $outlookWasAlreadyRunning = $true
}
catch
{
  try
  {
    $outlook = New-Object -comobject Outlook.Application
    $outlookWasAlreadyRunning = $false
  }
  catch
  {
    write-host "You must exit Outlook first."
    exit
  }
}
$namespace = $Outlook.GetNameSpace("MAPI")

# Generate the email from a template.
$template = get-childitem $emailTemplatePath -Filter "$emailTemplateName"
if ((Test-Path $template.FullName) -ne $true) {
  write-host "Couldn't find template $template"
  exit
}
$message = $outlook.CreateItemFromTemplate($template.FullName.ToString())

$message.To = $email

# Replace all of the keywords that we care about
$pronouns = @{
  'male' = @{'possessive' = 'his'; 'subject' = 'he'}
  'female' = @{'possessive' = 'her'; 'subject' = 'she'}
  'nonbinary' = @{'possessive' = 'their'; 'subject' = 'they'}
}
# Note that no replacement key is contained within another key.  For example,
# if FriendlyName and OtherFriendlyName were keys, then it's possible that,
# if the FriendlyName replace runs first, the final email would contain text
# like "OtherChris".
$replacements = @{
  'VarFriendlyName' = $friendlyName
  'VarLunchDate' = $lunchDate
  'VarOtherEmail' = $otherEmail
  'VarOtherFriendlyName' = $otherFriendlyName
  'VarOtherFullName' = $otherFullName
  'VarOtherPossessivePronoun' = ($pronouns.$otherGender).possessive
  'VarOtherSubjectPronoun' = ($pronouns.$otherGender).subject
}
$replacements.GetEnumerator() | ForEach-Object {
  $message.Subject = $message.Subject.Replace($_.Key, $_.Value)
  $message.HTMLBody = $message.HTMLBody.Replace($_.Key, $_.Value)
}

$message.Send()

# Close outlook if it wasn't opened before running this script
if ($outlookWasAlreadyRunning -eq $false)
{
    Get-Process "*outlook*" | Stop-Process –force
}