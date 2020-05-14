#----------------------------------------------------------------------------------------------------
# Author: Peter Geelen
# e-mail: peter@fim2010.be
# Web: blog.identityunderground.be
#
# Core references:
# http://social.technet.microsoft.com/wiki/contents/articles/5283.wiki-ninjas-blog-authoring-schedule.aspx
#----------------------------------------------------------------------------------------------------

#for every quarter the scripts generates a table with 13 weeks
#starting the first monday of the quarter
#to update the new schedule, change the year and quarter required, then the script will copy the HTML table to memory
#next you can paste it in the article, mentioned above

cls
[int]$year = 2014
[int]$quarter = 3
[int]$startmonth = ($quarter -1) * 3 + 1  

#get the first monday of the quarter
$startdate = get-date -Year $year -Month $startmonth -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
$firstmonday = $startdate

if ([int]$startdate.DayOfWeek -ne 1 ) 
{$firstmonday = $startdate.AddDays((8 - [int]$startdate.DayOfWeek) % 7)}

#set the weekly article name tags
[string[]]$message = "Monday+(Interview)", 'Tuesday+(Article+Spotlight)', 'Wednesday+(Wiki+Life)', 'Thursday+(CouncilSpotlight)', 'Friday+(International+Update)', 'Saturday+(Top+Contributors)', 'Sunday+(W/E+Surprise)'

#build the HTML table
[string]$body = ''
[string]$prefix = '<td><a href="http://www.timeanddate.com/scripts/ics.php?type=utc&amp;iso='
[string]$mid = 'T00&amp;p1=1440&amp;ah=1&amp;msg='
[string]$suffix = '" target="_blank">.'
[string]$suffix2 = '</a> </td>'


$day = $firstmonday

for ($week = 0; $week -ne 14; $week++)
{

    for ($weekday = 0; $weekday -lt 7; $weekday++)
    {
        $delta =  ($week*7) + $weekday
        $day = $firstmonday.AddDays($delta)

       if ($weekday -eq 0) 
        {   
            $column1 = $day.ToString("MMM/dd") + " - " + $day.addDays(6).ToString("MMM/dd") + '</td>'
            $body = $body + $column1
        }

    $body = $body + $prefix
    $body = $body + $day.ToString("yyyyMMdd")
    $body = $body + $mid
    $body = $body + $message[$weekday]
    $body = $body + $suffix
    #$body = $body + ($weekday + 1)
    $body = $body + $suffix2

    }
    
    $body = $body + '</tr>
 <tr style="border-width: medium 1pt 1pt; border-style: none solid solid; border-color: currentColor windowtext windowtext; padding: 0.75pt;">
 <td>'
}


$header = '<table border="1" cellpadding="0" cellspacing="0" style="border-color: inherit; border-collapse: collapse;" width="99%">
 <tr style="font-weight:bold; background: grey; border-width: 1pt 1pt 1pt medium; border-style: solid solid solid none; border-color: windowtext windowtext windowtext currentColor; padding: 0.75pt;">
  <td>Dates:</td>
  <td>Monday (Interview) </td>
  <td>Tuesday (Article Spotlight)</td>
  <td>Wednesday (Wiki Life)</td>
  <td>Thursday (Council Spotlight)</td>
  <td>Friday (International Update)</td>
  <td>Saturday (Top Contributors)</td>
  <td>Sunday (W/E Surprise)</td>
 </tr>
 <tr style="border-width: medium 1pt 1pt; border-style: none solid solid; border-color: currentColor windowtext windowtext; padding: 0.75pt;">
 <td>'

$footer = '</tr></table>'

$body = $header + $body + $footer

#save the table to the computer memory
$body | clip
Write-Host "The HTML code of the new table template is saved to memory, now paste it to the HTML Wiki article."