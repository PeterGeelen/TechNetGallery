#----------------------------------------------------------------------------------------------------
# # Author: Peter Geelen
# e-mail: peter@ffwd2.me
# Web: blog.identityunderground.be
#----------------------------------------------------------------------------------------------------
#Set-PSDebug -Trace 2
Set-PSDebug -off
#----------------------------------------------------------------------------------------------------
#Sorting mode
$SortingMode = 1
#where
#1 (Default) = Title (Alphabetical)
#2 = Date (Ascending, recent last)
#3 = Date (Descending, recent first)
#----------------------------------------------------------------------------------------------------

Function pause
{
    PARAM($msg = "")
    END
    {
    $message = $msg  + "... (Press any key to continue) ..."
    Write-Host $message
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}
#----------------------------------------------------------------------------------------------------
Function LoadTag
{
    PARAM($sTag = "", $sTagDesc = "")
    END
    {
        $newRecord = new-object psobject
        $newRecord | Add-Member noteproperty TagName $sTag
        $newRecord | Add-Member noteproperty TagSearch $sTag.Replace(" ", "+")
        $newRecord | Add-Member noteproperty TagDescription $sTagDesc
        return $newRecord
    }
}
#----------------------------------------------------------------------------------------------------
Function AddArticle
{
    PARAM($sNr = "", $sTitle = "", $sURL = "", $sVersion = "", $sDate = "")
    END
    {
        
        $newArticle = new-object psobject
        $newArticle | add-member noteproperty "Title" $sTitle
        $newArticle | add-member noteproperty "URL" $sURL
        $newArticle | add-member noteproperty "Version" $sVersion
        $newArticle | add-member noteproperty "Date" $sDate
        return $newArticle
    }
}
#----------------------------------------------------------------------------------------------------
Function GetDownloadInfo
{
    PARAM($pURL = "")
    END
    {

        #initialise the web connection
        $wcl = New-Object Net.Webclient

        #retrieve the URL
        $content = $wcl.DownloadString("$pURL")
        cls
    
        #get download title
        $i       = $content.indexof("header-bottom-space")
        $content = $content.substring($i)
        $i       = $content.indexof("<h2>") + 5 

        $content = $content.substring($i)
        $iEnd    = $content.indexof("</h2>") 

        $Downloadtitle = $content.substring(0,$iEnd)
        $Downloadtitle = $Downloadtitle.Trim()

        if ($Downloadtitle.ToLower().Contains("management pack")) 
        {
            #get the Version
            $i = $content.indexof("Version:") 
            $content = $content.substring($i)
            $i = $content.indexof("<p>")+3 
            $content = $content.substring($i)
            $iEnd    = $content.indexof("</p>") 
        
            $Version = $content.substring(0,$iEnd).Trim()

            $content = $content.substring($iEnd)

            #get Publication Data
            $i = $content.indexof("Date Published:") 
            $content = $content.substring($i)
            $i = $content.indexof("<p>")+3 
            $content = $content.substring($i)
            $iEnd    = $content.indexof("</p>") 
        
            $PublicationDate = $content.substring(0,$iEnd).Trim()

            $culture = New-Object System.Globalization.CultureInfo("en-us") 
            $shortdate = [datetime]::Parse($PublicationDate, $culture)

            $toprint = 
                $Downloadtitle + "#" + 
                $Version + "#" + 
                $shortdate.ToString("yyyy-MM-dd")
                
            return $toPrint
        }
        else
        {    
            return ""
        }
    }
}
#----------------------------------------------------------------------------------------------------
# MAIN
#----------------------------------------------------------------------------------------------------
$a = (Get-Host).UI.RawUI
$a.WindowTitle = "Microsoft Download Data Retriever"
$b = $a.WindowSize
#----------------------------------------------------------------------------------------------------

Clear-Host
Write-Progress -activity "Retrieving Microsoft Downloads" -status "Please wait..." -id 1
#----------------------------------------------------------------------------------------------------
# create tag list to publish
#----------------------------------------------------------------------------------------------------
$tagList = @()

$tagList += LoadTag -sTag "Management Pack" -sTagDesc "These downloads contain Management Packs."

#----------------------------------------------------------------------------------------------------
# initialise files
#----------------------------------------------------------------------------------------------------
$csvFile = $MyInvocation.MyCommand.Definition -replace 'ps1','csv'
$htmlFile = $MyInvocation.MyCommand.Definition -replace 'ps1','html'

#----------------------------------------------------------------------------------------------------
# process tags
#----------------------------------------------------------------------------------------------------

$table  = @()
$iTagCount = 0 

foreach ($tag in $taglist)
{
    $iTagCount += 1
    $iTagTotalCount = $Taglist.length
    #set tag URL
    $tagCode = $tag.TagSearch
    #https://www.microsoft.com/en-us/search/downloadresults?q=%22Management+Pack%22&first=1
       
    
    $pctComplete = ([int]($iTagCount/$tagList.length * 100))
    $activity = "Reading tag: "+ $tag.TagName 
    Write-Progress -activity $activity -status "Please wait" -percentcomplete $pctComplete -currentoperation "now processing tag ... ($iTagCount/$iTagTotalCount)" -id 1
    
    #limit article download to max 1000
    $maxlimit = 500
    
    $sArticleInfo = ""
    $iStartCount = 1

    for ($i = $iStartCount; $i -le $maxlimit; $i++)
    {
        $pctComplete = ([int]($i/$maxlimit * 100))
        $subactivity = "Collecting download pages for tag " + $tag.TagName 
        Write-Progress -activity $subactivity -status "Please wait" -percentcomplete $pctComplete -currentoperation "processing download page $i of $maxlimit..." -id 2


        #open webconnection to get first download of list
        $URL = "https://www.microsoft.com/en-us/search/downloadresults?q=%22$tagCode%22&first=$i"
        $URLlist = (Invoke-WebRequest -Uri "$URL").Links.Href -like "*download/details*"

        $webclient = New-Object Net.Webclient
        $sArticleInfo = GetDownloadInfo($URLlist[0])
        
        if($sArticleInfo -ne "")
        {
        
        $arrary = @()
        $array = $sArticleInfo.Split("#")
        $table += AddArticle -sNr $i.ToString().Padleft(5,"0") -sTitle $array[0] -sURL $URLlist[0] -sVersion $array[1] -sDate $array[2]   
        }
        else
        {
            break
        }  
    }

}
#----------------------------------------------------------------------------------------------------
#writing data to file
#-------------------------------------------------------------------------------------------------------------------

$TotalCount = $table.Length
$TotalCount
$table = $table | sort-object -Unique -property Title
$TotalCount = $table.Length
$TotalCount


switch ($SortingMode)
    {
        2 
            { 
            $table = $table | sort-object -property Date
            $htmlFile = $htmlFile.Replace(".html","_dateAS.html")
            }
        3 
            {
            $table = $table | sort-object -property Date -descending
            $htmlFile = $htmlFile.Replace(".html","_dateDS.html")

            }
        Default 
        {
            $table = $table | sort-object -property Title
        }
    }

$Count = 0
$pctComplete = ([int]($Count/$TotalCount * 100))

$output = "<table>"
$output += "<tr>"
$output  += "<td>Management Pack</td>"
$output  += "<td>Version</td>"
$output  += "<td>yyyy-mm-dd</td>"
$output += "</tr>"

foreach ($article in $table)
{
    $Count += 1
    $activity = "Printing article: "+ $article.Title
    Write-Progress -activity $activity -status "In progress" -percentcomplete $pctComplete -currentoperation "now processing article $Count of $TotalCount"

    $output += "<tr>"
    $output  += "<td><a href=`"" + $article.URL +"`">" + $article.Title + "</a></td>"
    $output  += "<td>"+$article.Version+"</td>"
    $output  += "<td>"+$article.Date+"</td>"
    $output += "</tr>"
            
}
$output += "</table>"

$output | out-file $htmlfile -Append
    

Write-Host "Command completed successfully"
#-------------------------------------------------------------------------------------------------------------------
Trap 
 { 
    $table | sort-object -property SeqNr,Title | Export-Csv $csvFile -encoding "unicode" -NoTypeInformation
    Write-Host "`nError: $($_.Exception.Message)`n" -foregroundcolor white -backgroundcolor darkred
    Exit 1
}
#----------------------------------------------------------------------------------------------------------
    