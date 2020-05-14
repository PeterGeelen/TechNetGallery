#----------------------------------------------------------------------------------------------------
# Author: Peter Geelen
# e-mail: peter@ffwd2.me
# Web: blog.identityunderground.be
#----------------------------------------------------------------------------------------------------
# PREREQUISITE:
# You need to install the PowerShell Excel module as mentioned below 
# Introducing the PowerShell Excel Module
# https://blogs.technet.microsoft.com/heyscriptingguy/2015/11/25/introducing-the-powershell-excel-module-2/
#----------------------------------------------------------------------------------------------------
#
# Start execution time tracking
$StarTime = Get-Date

#Set-PSDebug -Trace 2
Set-PSDebug -off
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
Function AddTagToCollection
{
    PARAM($sTagCount = "", $sTagName = "", $sTagLength = "",$sTagURL = "")
    END
    {
        $newArticle = new-object psobject
        $newArticle | add-member NoteProperty TagCount $sTagCount
        $newArticle | add-member NoteProperty TagName $sTagName
        $newArticle | add-member NoteProperty TagLength $sTagLength
        $iTagParts = ($sTagName.Split(' ')).count
        $newArticle | add-member NoteProperty TagPart $iTagParts
        $newArticle | add-member NoteProperty TagURL $sTagURL

        return $newArticle
    }
}
#----------------------------------------------------------------------------------------------------
Function ExtractTagInfo
{
    PARAM($strURL = "")
    END
    {
        #initialise the web connection
        $wcl = New-Object Net.Webclient

        Try
        {
            #retrieve the URL
            $webpage = ""
            $webpage = $wcl.DownloadString("$strURL") 
        }
        Catch [Exception]
        {
                #PARAM($sTagCount = 0, $sTagName = "", $sTagLength = "",$sTagURL = "")
                #return (@($iTagCount,$sTagName,$sTagName.Length))        
                $iTagCount = "X"
                $sTagName = "<Load Failure>"
                return (@($iTagCount,$sTagName,$sTagName.Length, $strURL))        
                break
        }
        

        if ($webpage.Contains("<span class=""summary"">")) 
        {
            #find the tagname
            $i = $webpage.indexof("All Tags")
            $Text  = $webpage.substring($i)

      
            $i = $Text.indexof("default.aspx")
            $Text  = $Text.substring($i)

            $i = $Text.indexof(">") + 1
            $Text  = $Text.substring($i)

            $iEndpos = $Text.indexof(">") + 1
            $sTagName  = $Text.substring(0, $iEndpos-4)
            
            #get the Count
            $i = $webpage.indexof("<span class=""summary"">") 
            $Text  = $webpage.substring($i)

            $i = $Text.indexof("(") 
            $Text  = $Text.substring($i+1)

            $iEndpos  = $Text.indexof(" ")
            $iTagCount =  $Text.substring(0, $iEndpos).Trim()
     
                
        }
        else
        {
            #PARAM($sTagCount = 0, $sTagName = "", $sTagLength = "",$sTagURL = "")
            #return (@($iTagCount,$sTagName,$sTagName.Length))        
            $iTagCount = "0"
            $sTagName = "<No Summary>"            }
        
        #PARAM($sTagCount = 0, $sTagName = "", $sTagLength = "",$sTagURL = "")
        return (@($iTagCount,$sTagName,$sTagName.Length, $strURL))        
    }
}
#----------------------------------------------------------------------------------------------------
# MAIN
#----------------------------------------------------------------------------------------------------

$a = (Get-Host).UI.RawUI
$a.WindowTitle = "Wiki Data Retriever"
$b = $a.WindowSize
#----------------------------------------------------------------------------------------------------

Clear-Host
Write-Progress -activity "Retrieving Wiki pages" -status "Please wait..." -id 1
#----------------------------------------------------------------------------------------------------
# initialise file
#----------------------------------------------------------------------------------------------------
$csvFile = $MyInvocation.MyCommand.Definition -replace 'ps1','csv'
$xlsFile = $MyInvocation.MyCommand.Definition -replace 'ps1','xlsx'
$errorlog = $MyInvocation.MyCommand.Definition -replace 'ps1','err'
$htmlFile = $MyInvocation.MyCommand.Definition -replace 'ps1','html'

#----------------------------------------------------------------------------------------------------
# process tag list
#----------------------------------------------------------------------------------------------------
#1. Load content from web
#set URL
$SourceURL = "https://social.technet.microsoft.com/wiki/SiteMap.ashx?apptype=wiki&GroupID=1&WikiApp=articles"
$webclient = New-Object Net.Webclient
[xml]$sitemap  = $webclient.DownloadString("$SourceURL") 



#2.load content from local file
#$localfile = "D:\My Documents\SkyDrive\_MS TechNet\ps_collect_tags\SiteMap.xml"
#[xml]$sitemap = Get-Content $localfile 


#check if $csvfile is already available
$controlset = @()
if (Test-Path $csvFile)
{
    $controlset = Import-Csv $csvFile  
}

$TagCollection = @()
$URLCount = 0
$TotalCount = $sitemap.urlset.url.loc.length
$activity = "Checking articles"
   
foreach ($url in $sitemap.urlset.url.loc)
{
    $URLCount += 1
    $pctComplete = ([int]($URLCount/$TotalCount * 100))
    
    Write-Progress -activity $activity -status "Please wait" -percentcomplete $pctComplete -currentoperation "Now processing article $URLCount of $TotalCount ($url)."

    #$URL
    if ($url.Contains("/tags/"))
    {
        
        $searchstring = "*"+$url+"*"
        $URLMissing = (($controlset -like $searchstring).Count -lt 1)

        if ($URLMissing)

        {
            Write-Host "MISSING:" $URL
            $articleInfo = ExtractTagInfo -strURL $URL
    
            #PARAM($sTagCount = 0, $sTagName = "", $sTagLength = "",$sTagURL = "")
            #return (@($iTagCount,$sTagName,$iTagLength))
            $articleInfo = AddTagToCollection -sTagCount $articleInfo[0] -sTagName $articleInfo[1] -sTagLength $articleInfo[2] -sTagURL $articleInfo[3] 
            $articleInfo| Format-Table

            #save article info to file
            $articleInfo | Export-Csv $csvFile -encoding "unicode" -NoTypeInformation -Append
            #$TagCollection += $articleInfo 
            Write-host $URLCount $articleInfo[1]
        }
        else
        {
        Write-Host "OK:" $URL
        }
    }
}

#----------------------------------------------------------------------------------------------------
#writing data to file
#-------------------------------------------------------------------------------------------------------------------

$CSVEndTime = Get-Date

"Importing csv"
$TagCollection = Import-Csv $csvFile
"Sorting csv"
$TagCollection = $TagCollection | sort-object -property TagCount,TagName
$TagCollection | Export-Excel $xlsFile

# Stop execution time tracking
$EndTime = Get-Date


$TagTotalCount = $TagCollection.Length

Write-host "Total tags count: " $TagTotalCount

# Report execution time tracking
#format dd:hh:mm:ss.tttt
"The script took {0:g} to run." -f (NEW-TIMESPAN –Start $StartDate –End $EndDate)

Write-Host "Command completed successfully"


#-------------------------------------------------------------------------------------------------------------------
Trap 
{ 
    #TagCount,"TagName","TagLength""TagURL"
    if (Test-Path $csvFile)
    {
        $TagCollection | sort-object -property TagCount,TagName| Export-Csv $errorlog -encoding "unicode" -NoTypeInformation -Append
    }
    else
    {
        $TagCollection | sort-object -property TagCount,TagName| Export-Csv $errorlog -encoding "unicode" -NoTypeInformation
    }
    Write-host "Data saved to errorlog..."
    Write-Host "`nError: $($_.Exception.Message)`n" -foregroundcolor white -backgroundcolor darkred
    Exit 1
}
#----------------------------------------------------------------------------------------------------------
