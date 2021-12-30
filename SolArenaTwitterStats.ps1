#Author - Rudy Corradetti
#Purpose - Grab key metrics from Twitter, Instagram, and Discord, create report, email report to $recipients
#Requirements - SMTP2GO free account, advanced API access to Twitter account, 

#Set date
$Date = get-date -f yyyy-MM-dd

#Import PSTwitterAPI - https://github.com/mkellerman/PSTwitterAPI
Import-Module -Name PSTwitterAPI

#################################################################################################OAUTH SETTINGS START HERE#############################################################################################
$OAuthSettings = @{
  ApiKey = ""
  ApiSecret = ""
  AccessToken = ""
  AccessTokenSecret = ""
}
Set-TwitterOAuthSettings @OAuthSettings

#################################################################################################PSTWITTERAPI CODE STARTS HERE#############################################################################################
#needed to ensure all followers are captured when Get-TwitterFollowers_list runs
$FormatEnumerationLimit = -1
$AmtOfFollowers = Get-TwitterFollowers_list | Out-String 
$AmtOfFollowers -match [regex]::matches($AmtOfFollowers,"screen_name=").count
$matches.Values | Out-File ".\TwitterFollowers\$date.txt"
$AmtFollowsFormatted = [IO.File]::ReadAllText(".\TwitterFollowers\$date.txt")


#################################################################################################FORMAT CSV FILE BASED ON PSTWITTERAPI DATA STARTS HERE#############################################################################################


try {$outfile = "C:\Users\RJC\SolArenaTwitterStats\SolArenaTwitterStats.csv"}
catch {$newcsv = {} | Select-Object Date,Followers,PlusOrMinus | Export-Csv $outfile -NoTypeInformation}
$csvfile = Import-Csv $outfile
$csvfile.Date = "$Date"
$csvfile.Followers = $AmtFollowsFormatted
#$csvfile.PlusOrMinus = 
##Add a new row of data to the csv each time script runs
"{0},{1},{2}"-f $csvfile.Date,$csvfile.Followers, $csvfile.PlusOrMinus |Add-Content $outfile

#$outfile = "C:\temp\Outfile.csv"
#$newcsv = {} | Select "EMP_Name","EMP_ID","CITY" | Export-Csv $outfile
#$csvfile = Import-Csv $outfile
#$csvfile.Emp_Name = "Charles"
#$csvfile.EMP_ID = "2000"
#$csvfile.CITY = "New York"
#$csvfile | Export-CSV
#$outfile Import-Csv $outfile

##################################################################################################SMTP SECTION STARTS HERE#############################################################################################
#variables, recipients, logo, body for email
$to = "socials@solarena.io"
$from = "SolArenaTwitterStats@solarena.io"
$subject = "SolArena Social Stats - $Date"
#Build table that socials data will sit in
$a = "<style>"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:azure}"
$a = $a + "</style>"
#Imports the CSV saved to $output.
$csvforbody = Import-Csv $outfile | Select-Object Date,Followers | ConvertTo-Html -head $a -Body "<H2>SolArena Twitter Report</H2> <H3> Follower Goal $AmtFollowsFormatted/1000 </H3>"
#Selects File Name, Device Name, and File Path columns, then converts to HTML using the Head +Body Information
$emailBody = "$csvforbody"
#Create credential object 
$server = "mail.smtp2go.com"
$port = "25"
#Send send the Mail

Send-MailMessage -to $to -From $from -SmtpServer $server -Port $port -Subject $subject -BodyAsHtml "<br /><img src=solarenalogoNoEp.png> $emailBody" -Attachments "C:\Users\RJC\SolArenaTwitterStats\solarenalogoNoEp.png"