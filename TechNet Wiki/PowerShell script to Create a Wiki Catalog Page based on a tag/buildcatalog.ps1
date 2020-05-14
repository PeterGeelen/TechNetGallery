#----------------------------------------------------------------------------------------------------
# Basic retrieval logic by Markus Vilcinskas, Microsft
# Author: Peter Geelen
# ----------------------------------------------------------------------------------------------------
#Set-PSDebug -Trace 2
Set-PSDebug -off
#----------------------------------------------------------------------------------------------------
Function LoadTag
{
    PARAM($sTag = "", $sTagDesc = "")
    END
    {
        $newRecord = new-object psobject
        $newRecord | Add-Member NoteProperty TagName $sTag
        $newRecord | Add-Member NoteProperty TagSearch $sTag.Replace(" ", "+")
        $newRecord | Add-Member NoteProperty TagDescription $sTagDesc
        return $newRecord
    }
}
#----------------------------------------------------------------------------------------------------
Function AddArticle
{
    PARAM($sNr = "", $sTitle = "", $sUrl = "", $sAuthor = "")
    END
    {
        $newArticle = new-object psobject
        $newArticle | add-member noteproperty "SeqNr" $sNr
        $newArticle | add-member noteproperty "Title" $sTitle
        $newArticle | add-member noteproperty "URL" $sUrl
        $newArticle | add-member noteproperty "Author" $sAuthor
        return $newArticle
    }
}
#----------------------------------------------------------------------------------------------------
Function GetArticleAuthor
{
    PARAM($strURL = "")
    END
    {
    #initialise the web connection
    $wcl = New-Object Net.Webclient

    #retrieve the URL
    $source  = $wcl.DownloadString("$strURL") 
    
    if ($source.Contains("First published by")) 
        {
        $i     = $source.indexof("First published by") 
    
        $Text  = $source.substring($i)
    
        $i     = $Text.indexof("<span class=""user-name"">") 
        $Text  = $Text.substring($i+25)

        $i     = $Text.indexof(">")+1
        $Text  = $Text.substring($i)

        #get the author
        $endpos  = $Text.indexof("<")
        $Author =  $Text.substring(0, $endpos).Trim()

        $i     = $Text.indexof("attribute-item page-created")
        $Text  = $Text.substring($i)
        $i     = $Text.indexof("attribute-value")
        $Text  = $Text.substring($i)

        $i     = $Text.indexof(">")+1
        $Text  = $Text.substring($i)

        $endpos  = $Text.indexof("<")

        #get creation date
        $CreationDate =  $Text.substring(0, $endpos).Trim()
        $culture = New-Object System.Globalization.CultureInfo("en-us") 
        $shortdate = [datetime]::Parse($CreationDate, $culture)
        $toprint = "(" + $Author + ", " + $shortdate.ToString("d MMM yyyy")  + ")"
        return $toprint
        }
    else
        {    
        return "(---)"
        }
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
# create tag list to publish
#----------------------------------------------------------------------------------------------------
$tagList = @()

$tagList += LoadTag -sTag "FIM Article Development" -sTagDesc "These articles contain information about developing FIM Wiki articles."
$tagList += LoadTag -sTag "FIM Reference Article" -sTagDesc "These article contain reference information." 
$tagList += LoadTag -sTag "FIM Understanding Article" -sTagDesc "These articles contains a description of a technical concept." 
$tagList += LoadTag -sTag "FIM Troubleshooting Article" -sTagDesc "These articles describe how to troubleshoot an issue." 
$tagList += LoadTag -sTag "FIM How To Article" -sTagDesc "These articles describe how to configure something." 
$tagList += LoadTag -sTag "FIM Tool Script" -sTagDesc "These articles contain links to the script code." 
$tagList += LoadTag -sTag "FIM How To Script" -sTagDesc "These articles contain script code." 
$tagList += LoadTag -sTag "FIM ScriptBox Item" -sTagDesc "These articles are about a FIM related script." 
$tagList += LoadTag -sTag "FIM Technical Article" -sTagDesc "These articles are about a technical description." 
$tagList += LoadTag -sTag "FIM" -sTagDesc "These articles are about FIM." 
$tagTotalCount = $tagList.length

#----------------------------------------------------------------------------------------------------
# initialise file
#----------------------------------------------------------------------------------------------------
$csvFile = $MyInvocation.MyCommand.Definition -replace 'ps1','csv'
$htmlFile = "FIMcatalog.html"

#----------------------------------------------------------------------------------------------------
# process tags
#----------------------------------------------------------------------------------------------------

$dataList  = @()
$iSeqNo = 0
$tagCount = 0 
foreach ($tag in $taglist)
{
    $tagCount += 1
    $iSeqNo = $tagCount * 1000
    $tagname = $tag.TagName
    $tagdesc = $tag.TagDescription
    
    
    #set tag URL
    $tagCode = $tag.TagSearch
    $url = "http://social.technet.microsoft.com/wiki/contents/articles/tags/$tagCode/default.aspx"

    #PARAM($sNr = "", $sTitle = "", $sUrl = "", $sAuthor = "")
    #add H2 tag header to article list
    $dataList += AddArticle -sNr $iSeqNo.ToString().Padleft(5,"0") -sTitle $tagName -sUrl $url -sAuthor $tagdesc
    
    $iSeqNo++
    $pctComplete = ([int]($tagCount/$tagTotalCount * 100))
    $activity = "Reading tag: "+ $tag.TagName 
    Write-Progress -activity $activity -status "Please wait" -percentcomplete $pctComplete -currentoperation "now processing tag $tagCount of $tagTotalCount" -id 1

    $webclient = New-Object Net.Webclient
    $htmlData  = $webclient.DownloadString("$url") 
             
    $index     = $htmlData.indexof("<span class=""summary"">") 
    If($index -ge 0) 
    {
        #$iSeqNo++
        $index += 32
        $pageCount = $htmlData.substring($index)
        $index     = $pageCount.indexof(" ")
        $pageCount = $pageCount.subString(0, $index)
        for ($i = 0; $i -lt $pageCount; $i++)
        {
            $x = $i + 1
            $pctComplete = ([int]($x/$pageCount * 100))
            $subactivity = "Processing HTML Pages for tag " + $tag.TagName 
            Write-Progress -activity $subactivity -status "Please wait" -percentcomplete $pctComplete -currentoperation "processing results page $x of $pageCount" -id 2

            $htmlData  = $webclient.DownloadString("$($url)?PageIndex=$($x)") 
            $index     = $htmlData.indexof("<ul class=""content-list standard"">") + 26
            $dataText  = $htmlData.substring($index)
            $index     = $dataText.indexof("</ul>")
            $dataText  = $dataText.substring(0, $index)
            $dataText  = $dataText.Trim()

            while($dataText.indexof("<li class=""content-item"">") -gt -1)
            {
                
                $dataText = $dataText.substring(25)
                $index    = $dataText.indexof("</li>")
   
                $newItem  = ($dataText.subString(0, $index)).Trim()
                $newItem  = ($newItem.subString($newItem.indexof("<h4 class=""post-name"">")+ 22)).Trim()
                $newItem  = ($newItem.subString(10)).Trim()
                $newItem  = $newItem.subString($newItem.indexof("""") + 1)

                $link     = ($newItem.subString(0, $newItem.indexof(""""))).Trim()
                $newItem  = $newItem.subString($newItem.indexof("""") + 2)
                $title    = ($newItem.subString(0, $newItem.indexof("</a>"))).Trim() 
                $title = $title.replace("&#39;","'")
                $title = $title.replace("&quot;","'")

                $tosearch = $link
                $found = ""
                $found = $dataList | Where-Object {$_.URL -eq $tosearch}
                             
                #The if statement below check if an article arleady exst.
                #If the article is listed already, it's skipped..
                #Disable this if statement if you want to put all articles to each tag they have;

                 if ($found -eq $null )
                {                
                    #PARAM($strURL = "")
                    $articleAuthor = GetArticleAuthor -strURL $link

                    # PARAM($sNr = "", $sTitle = "", $sUrl = "", $sAuthor = "")
                 
                 $dataList += AddArticle -sNr $iSeqNo.ToString().Padleft(5,"0") -sTitle $title -sUrl $link -sAuthor $articleAuthor
                }
                $dataText = ($dataText.substring($index + 5)).Trim()
            }
        }
    }
}
#----------------------------------------------------------------------------------------------------
#writing data to file
#-------------------------------------------------------------------------------------------------------------------
$dataList = $dataList | sort-object -property SeqNr,Title 

$Count = 0
$TotalCount = $dataList.Length
$pctComplete = ([int]($articleCount/$tagTotalCount * 100))
             
#put next line in comment if you wish to skip the header 
get-content .\catalogheader.txt | out-file $htmlfile 

foreach ($article in $dataList)
{
    $Count += 1
    $activity = "Printing article: "+ $article.Title
    Write-Progress -activity $activity -status "Please wait" -percentcomplete $pctComplete -currentoperation "now processing article $Count of $TotalCount"

    if ($article.Author -ne "(---)")
    {
        if ($article.SeqNr.EndsWith("0"))
            {
            $txt = "</p><h2><a href=`"" + $article.URL + "`" target=_blank>"
            $txt | out-file $htmlfile -append

            $article.Title | out-file $htmlfile -append

            $txt = "</a></h2><p>" + $article.Author + "</p><p>"
            $txt | out-file $htmlfile -append
            }
        else
            {
            $txt = "<a href=`"" + $article.URL +"`">" + $article.Title + " " + $article.Author + "</a><br>"
            $txt | out-file $htmlfile -append
            }
    }
}
             

#put next line in comment if you wish to skip the footer
get-content catalogfooter.txt | out-file $htmlfile -append
Write-Host "Command completed successfully"
#-------------------------------------------------------------------------------------------------------------------
Trap 
 { 
    $dataList | sort-object -property SeqNr,Title | Export-Csv $csvFile -encoding "unicode" -NoTypeInformation
    Write-Host "`nError: $($_.Exception.Message)`n" -foregroundcolor white -backgroundcolor darkred
    Exit 1
}
#---------------------------------------------------------------------------------------------------------- 