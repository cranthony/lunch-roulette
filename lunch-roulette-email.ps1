# The emailing structure of this was from a Microsoft dev blog at
# https://devblogs.microsoft.com/premier-developer/outlook-email-automation-with-powershell/
#
# And then I looked up a bunch of tutorials on Powershell to get the rest.
param(
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [String]$email,
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [String]$template,
  # Every key in this hash table should match the pattern given by
  # $variablePattern.
  [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [hashtable]$replacements,
  # Note that we don't URLEncode this string or anything like that when doing
  # replacements, so prefer patterns that aren't affected by such encodings.
  #
  # For example, URL encoding may be used if the variable is included in a link
  # within the template.
  [ValidateNotNullOrEmpty()] [regex]$variablePattern = "Var[A-Z]\w+",
  # The variations of pronouns needed to account for gender are painful to
  # enumerate, so this script can do that.  The idea is to add variations of
  # each $replacements key that matches this pattern for each type of pronoun we
  # need.
  #
  # For example, if this pattern is "Gender$", and there is a replacement for:
  #   'VarOtherGender' = 'male'
  # in $replacements, then this will add replacements for:
  #   'VarOtherSubjectPronoun' = 'he'
  #   'VarOtherObjectPronoun' = 'him'
  #   'VarOtherPossessivePronoun' = 'his'
  #
  # The script doesn't currently support capitalized versions of these, because
  # they weren't needed.
  [ValidateNotNullOrEmpty()] [regex]$genderPattern = "Gender$"
)

# Stop the script whenever there's an error.  By default, Powershell continues
# past many types of errors.
$ErrorActionPreference = "Stop"

# Connect to Outlook.
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
add-type -assembly "System.Runtime.Interopservices"
try
{
  $outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
}
catch
{
  # The example that I followed would start Outlook if it wasn't running, but
  # I don't really want that.  I don't want to accidentally lock up my computer
  # by starting a script that sends 100 emails, and watching each start and
  # stop Outlook.
  Write-Host "Is Outlook running?  We expect it to be running already so that"
  Write-Host "it's easy to call this script many times in a row."
  throw
}

# Generate the email from a template.
$templatePath = Resolve-Path $template
if (-Not $templatePath.Path) {
  throw "Can't open template '$template'"
}
$message = $outlook.CreateItemFromTemplate($templatePath.Path)

$message.To = $email

# Expand gender variables, because these can be a pain to specify by command
# line.  There are too many variations.
$pronounTypes = @("Subject", "Object", "Possessive")
$pronouns = @{
  # The order of each of the value arrays matches the order of pronoun types in
  # $pronounTypes.
  'male' = @('he', 'him', 'his')
  'female' = @('she', 'her', 'her')
  'nonbinary' = @('they', 'them', 'their')
}
$genderReplacements = @{}
$replacements.GetEnumerator() | ForEach-Object {
  if ($genderPattern.Match($_.Key).Success) {
    if (-Not $pronouns.ContainsKey($_.Value)) {
      throw "Gender key '$($_.Key)' expected to have value in: $($pronouns.Keys)"
    }
    $genderKey = $_.Key
    $genderPronouns = $pronouns[$_.Value]
    if ($genderPronouns.Count -ne $pronounTypes.Count) {
      throw "Unexpected number of gender pronouns for ${genderKey}: $genderPronouns"
    }
    for ($i = 0; $i -lt $pronounTypes.Count; $i++) {
      $genderReplacements.Add(
        "$($genderKey -Replace $genderPattern,$pronounTypes[$i])Pronoun",
        $genderPronouns[$i]
      )
    }
  }
}
$genderReplacements.GetEnumerator() | ForEach-Object {
  $replacements.Add($_.Key, $_.Value)
}

# Do the replacements within the message subject and body.
$replacements.GetEnumerator() | ForEach-Object {
  if (-Not $variablePattern.Match($_.Key).Success) {
    throw "All replacement keys must match pattern [Regex]::new('$variablePattern'), but found key '$($_.Key)'"
  }

  $message.Subject = $message.Subject.Replace($_.Key, $_.Value)
  $message.HTMLBody = $message.HTMLBody.Replace($_.Key, $_.Value)
}

# Validate that the caller supplied all of the variables that we'd expect.
@{
  "Message subject" = $message.Subject
  "Message body" = $message.HTMLBody
}.GetEnumerator() | ForEach-Object {
  $unmatchedVariables = $variablePattern.Matches($_.Value)
  if ($unmatchedVariables.Count -gt 0) {
    throw "$($_.Key) contains unmatched variables: $unmatchedVariables"
  }
}

$message.Send()