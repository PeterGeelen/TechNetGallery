#----------------------------------------------------------------------------------------------------
# Author: Peter Geelen
# e-mail: peter@fim2010.be
# Web: blog.identityunderground.be
# Credits: http://brianreiter.org/2010/09/03/copy-and-paste-with-clipboard-from-powershell/
#
# Core references:
# http://social.technet.microsoft.com/wiki/contents/articles/16870.wiki-fix-color-issues-in-wiki-articles.aspx
#----------------------------------------------------------------------------------------------------
#Set-PSDebug -Trace 2
Set-PSDebug -off
#----------------------------------------------------------------------------------------------------
Function pause
{
 PARAM($msg ="")
 END
 {
 $message = $msg +"... (Press any key to continue) ..."
 Write-Host $message
 $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
 }
}
#----------------------------------------------------------------------------------------------------
Function LoadColor
{
 PARAM($sRGBCode ="", $sName ="")
 END
 {
 $newRecord = new-object psobject
 $newRecord | Add-Member NoteProperty RGBCode $sRGBCode
 $newRecord | Add-Member NoteProperty Name $sName
 return $newRecord
 }
}
#----------------------------------------------------------------------------------------------------
Function LoadColorList
{
 $sRGBList = @() 
 $sRGBList += LoadColor -sRGBCode "rgb(0, 0, 0)" -sName "Black"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 0, 128)" -sName "Navy"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 0, 139)" -sName "DarkBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 0, 205)" -sName "MediumBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 0, 255)" -sName "Blue"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 100, 0)" -sName "DarkGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 128, 0)" -sName "Green"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 128, 128)" -sName "Teal"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 139, 139)" -sName "DarkCyan"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 191, 255)" -sName "DeepSkyBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 206, 209)" -sName "DarkTurquoise"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 250, 154)" -sName "MediumSpringGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 255, 0)" -sName "Lime"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 255, 127)" -sName "SpringGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 255, 255)" -sName "Aqua"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 255, 255)" -sName "Cyan"
 $sRGBList += LoadColor -sRGBCode "rgb(25, 25, 112)" -sName "MidnightBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(30, 144, 255)" -sName "DodgerBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(32, 178, 170)" -sName "LightSeaGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(34, 139, 34)" -sName "ForestGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(46, 139, 87)" -sName "SeaGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(47, 79, 79)" -sName "DarkSlateGray"
 $sRGBList += LoadColor -sRGBCode "rgb(47, 79, 79)" -sName "DarkSlateGrey"
 $sRGBList += LoadColor -sRGBCode "rgb(50, 205, 50)" -sName "LimeGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(60, 179, 113)" -sName "MediumSeaGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(64, 224, 208)" -sName "Turquoise"
 $sRGBList += LoadColor -sRGBCode "rgb(65, 105, 225)" -sName "RoyalBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(70, 130, 180)" -sName "SteelBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(72, 61, 139)" -sName "DarkSlateBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(72, 209, 204)" -sName "MediumTurquoise"
 $sRGBList += LoadColor -sRGBCode "rgb(75, 0, 130)" -sName "Indigo"
 $sRGBList += LoadColor -sRGBCode "rgb(85, 107, 47)" -sName "DarkOliveGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(95, 158, 160)" -sName "CadetBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(100, 149, 237)" -sName "CornflowerBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(102, 205, 170)" -sName "MediumAquamarine"
 $sRGBList += LoadColor -sRGBCode "rgb(105, 105, 105)" -sName "DimGray"
 $sRGBList += LoadColor -sRGBCode "rgb(105, 105, 105)" -sName "DimGrey"
 $sRGBList += LoadColor -sRGBCode "rgb(106, 90, 205)" -sName "SlateBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(107, 142, 35)" -sName "OliveDrab"
 $sRGBList += LoadColor -sRGBCode "rgb(112, 128, 144)" -sName "SlateGray"
 $sRGBList += LoadColor -sRGBCode "rgb(112, 128, 144)" -sName "SlateGrey"
 $sRGBList += LoadColor -sRGBCode "rgb(119, 136, 153)" -sName "LightSlateGray"
 $sRGBList += LoadColor -sRGBCode "rgb(119, 136, 153)" -sName "LightSlateGrey"
 $sRGBList += LoadColor -sRGBCode "rgb(123, 104, 238)" -sName "MediumSlateBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(124, 252, 0)" -sName "LawnGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(127, 255, 0)" -sName "Chartreuse"
 $sRGBList += LoadColor -sRGBCode "rgb(127, 255, 212)" -sName "Aquamarine"
 $sRGBList += LoadColor -sRGBCode "rgb(128, 0, 0)" -sName "Maroon"
 $sRGBList += LoadColor -sRGBCode "rgb(128, 0, 128)" -sName "Purple"
 $sRGBList += LoadColor -sRGBCode "rgb(128, 128, 0)" -sName "Olive"
 $sRGBList += LoadColor -sRGBCode "rgb(92, 92, 92)" -sName "Grey"
 $sRGBList += LoadColor -sRGBCode "rgb(128, 128, 128)" -sName "Gray"
 $sRGBList += LoadColor -sRGBCode "rgb(128, 128, 128)" -sName "Grey"
 $sRGBList += LoadColor -sRGBCode "rgb(135, 206, 235)" -sName "SkyBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(135, 206, 250)" -sName "LightSkyBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(138, 43, 226)" -sName "BlueViolet"
 $sRGBList += LoadColor -sRGBCode "rgb(139, 0, 0)" -sName "DarkRed"
 $sRGBList += LoadColor -sRGBCode "rgb(139, 0, 139)" -sName "DarkMagenta"
 $sRGBList += LoadColor -sRGBCode "rgb(139, 69, 19)" -sName "SaddleBrown"
 $sRGBList += LoadColor -sRGBCode "rgb(143, 188, 143)" -sName "DarkSeaGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(144, 238, 144)" -sName "LightGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(147, 112, 219)" -sName "MediumPurple"
 $sRGBList += LoadColor -sRGBCode "rgb(148, 0, 211)" -sName "DarkViolet"
 $sRGBList += LoadColor -sRGBCode "rgb(152, 251, 152)" -sName "PaleGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(153, 50, 204)" -sName "DarkOrchid"
 $sRGBList += LoadColor -sRGBCode "rgb(154, 205, 50)" -sName "YellowGreen"
 $sRGBList += LoadColor -sRGBCode "rgb(160, 82, 45)" -sName "Sienna"
 $sRGBList += LoadColor -sRGBCode "rgb(165, 42, 42)" -sName "Brown"
 $sRGBList += LoadColor -sRGBCode "rgb(169, 169, 169)" -sName "DarkGray"
 $sRGBList += LoadColor -sRGBCode "rgb(169, 169, 169)" -sName "DarkGrey"
 $sRGBList += LoadColor -sRGBCode "rgb(173, 216, 230)" -sName "LightBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(173, 255, 47)" -sName "GreenYellow"
 $sRGBList += LoadColor -sRGBCode "rgb(175, 238, 238)" -sName "PaleTurquoise"
 $sRGBList += LoadColor -sRGBCode "rgb(176, 196, 222)" -sName "LightSteelBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(176, 224, 230)" -sName "PowderBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(178, 34, 34)" -sName "FireBrick"
 $sRGBList += LoadColor -sRGBCode "rgb(184, 134, 11)" -sName "DarkGoldenrod"
 $sRGBList += LoadColor -sRGBCode "rgb(186, 85, 211)" -sName "MediumOrchid"
 $sRGBList += LoadColor -sRGBCode "rgb(188, 143, 143)" -sName "RosyBrown"
 $sRGBList += LoadColor -sRGBCode "rgb(189, 183, 107)" -sName "DarkKhaki"
 $sRGBList += LoadColor -sRGBCode "rgb(192, 192, 192)" -sName "Silver"
 $sRGBList += LoadColor -sRGBCode "rgb(199, 21, 133)" -sName "MediumVioletRed"
 $sRGBList += LoadColor -sRGBCode "rgb(205, 92, 92)" -sName "IndianRed"
 $sRGBList += LoadColor -sRGBCode "rgb(205, 133, 63)" -sName "Peru"
 $sRGBList += LoadColor -sRGBCode "rgb(210, 105, 30)" -sName "Chocolate"
 $sRGBList += LoadColor -sRGBCode "rgb(210, 180, 140)" -sName "Tan"
 $sRGBList += LoadColor -sRGBCode "rgb(211, 211, 211)" -sName "LightGray"
 $sRGBList += LoadColor -sRGBCode "rgb(211, 211, 211)" -sName "LightGrey"
 $sRGBList += LoadColor -sRGBCode "rgb(216, 191, 216)" -sName "Thistle"
 $sRGBList += LoadColor -sRGBCode "rgb(218, 112, 214)" -sName "Orchid"
 $sRGBList += LoadColor -sRGBCode "rgb(218, 165, 32)" -sName "Goldenrod"
 $sRGBList += LoadColor -sRGBCode "rgb(219, 112, 147)" -sName "PaleVioletRed"
 $sRGBList += LoadColor -sRGBCode "rgb(220, 20, 60)" -sName "Crimson"
 $sRGBList += LoadColor -sRGBCode "rgb(220, 220, 220)" -sName "Gainsboro"
 $sRGBList += LoadColor -sRGBCode "rgb(221, 160, 221)" -sName "Plum"
 $sRGBList += LoadColor -sRGBCode "rgb(222, 184, 135)" -sName "BurlyWood"
 $sRGBList += LoadColor -sRGBCode "rgb(224, 255, 255)" -sName "LightCyan"
 $sRGBList += LoadColor -sRGBCode "rgb(230, 230, 250)" -sName "Lavender"
 $sRGBList += LoadColor -sRGBCode "rgb(233, 150, 122)" -sName "DarkSalmon"
 $sRGBList += LoadColor -sRGBCode "rgb(238, 130, 238)" -sName "Violet"
 $sRGBList += LoadColor -sRGBCode "rgb(238, 232, 170)" -sName "PaleGoldenrod"
 $sRGBList += LoadColor -sRGBCode "rgb(240, 128, 128)" -sName "LightCoral"
 $sRGBList += LoadColor -sRGBCode "rgb(240, 230, 140)" -sName "Khaki"
 $sRGBList += LoadColor -sRGBCode "rgb(240, 248, 255)" -sName "AliceBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(240, 255, 240)" -sName "Honeydew"
 $sRGBList += LoadColor -sRGBCode "rgb(240, 255, 255)" -sName "Azure"
 $sRGBList += LoadColor -sRGBCode "rgb(244, 164, 96)" -sName "SandyBrown"
 $sRGBList += LoadColor -sRGBCode "rgb(245, 222, 179)" -sName "Wheat"
 $sRGBList += LoadColor -sRGBCode "rgb(245, 245, 220)" -sName "Beige"
 $sRGBList += LoadColor -sRGBCode "rgb(245, 245, 245)" -sName "WhiteSmoke"
 $sRGBList += LoadColor -sRGBCode "rgb(245, 255, 250)" -sName "MintCream"
 $sRGBList += LoadColor -sRGBCode "rgb(248, 248, 255)" -sName "GhostWhite"
 $sRGBList += LoadColor -sRGBCode "rgb(250, 128, 114)" -sName "Salmon"
 $sRGBList += LoadColor -sRGBCode "rgb(250, 235, 215)" -sName "AntiqueWhite"
 $sRGBList += LoadColor -sRGBCode "rgb(250, 240, 230)" -sName "Linen"
 $sRGBList += LoadColor -sRGBCode "rgb(250, 250, 210)" -sName "LightGoldenrodYellow"
 $sRGBList += LoadColor -sRGBCode "rgb(253, 245, 230)" -sName "OldLace"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 0, 0)" -sName "Red"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 0, 255)" -sName "Fuchsia"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 0, 255)" -sName "Magenta"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 20, 147)" -sName "DeepPink"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 69, 0)" -sName "OrangeRed"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 99, 71)" -sName "Tomato"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 105, 180)" -sName "HotPink"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 127, 80)" -sName "Coral"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 140, 0)" -sName "DarkOrange"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 160, 122)" -sName "LightSalmon"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 165, 0)" -sName "Orange"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 182, 193)" -sName "LightPink"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 192, 203)" -sName "Pink"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 215, 0)" -sName "Gold"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 218, 185)" -sName "PeachPuff"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 222, 173)" -sName "NavajoWhite"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 228, 181)" -sName "Moccasin"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 228, 196)" -sName "Bisque"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 228, 225)" -sName "MistyRose"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 235, 205)" -sName "BlanchedAlmond"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 239, 213)" -sName "PapayaWhip"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 240, 245)" -sName "LavenderBlush"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 245, 238)" -sName "Seashell"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 248, 220)" -sName "Cornsilk"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 250, 205)" -sName "LemonChiffon"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 250, 240)" -sName "FloralWhite"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 250, 250)" -sName "Snow"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 255, 0)" -sName "Yellow"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 255, 224)" -sName "LightYellow"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 255, 240)" -sName "Ivory"
 $sRGBList += LoadColor -sRGBCode "rgb(255, 255, 255)" -sName "White"

## non mathching codes with closest match
 $sRGBList += LoadColor -sRGBCode "rgb(42, 42, 42)" -sName "darkSlateGray"
 $sRGBList += LoadColor -sRGBCode "rgb(163, 163, 163)" -sName "darkGray"
 $sRGBList += LoadColor -sRGBCode "rgb(240, 240, 240)" -sName "whiteSmoke"
 $sRGBList += LoadColor -sRGBCode "rgb(242, 242, 242)" -sName "whiteSmoke"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 102, 221)" -sName "DeepSkyBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(38, 38, 38)" -sName "DarkSlateGrey"
 $sRGBList += LoadColor -sRGBCode "rgb(51, 51, 51)" -sName "DarkSlateGrey"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 102, 153)" -sName "cornflowerblue"
 $sRGBList += LoadColor -sRGBCode "rgb(248, 248, 248)" -sName "WhiteSmoke"
 $sRGBList += LoadColor -sRGBCode "rgb(0, 130, 0)" -sName "Green"
 $sRGBList += LoadColor -sRGBCode "rgb(127, 157, 185)" -sName "SteelBlue"
 $sRGBList += LoadColor -sRGBCode "rgb(163, 21, 21)" -sName "Red"
 $sRGBList += LoadColor -sRGBCode "rgb(43, 145, 175)" -sName "cornflowerblue"
 $sRGBList += LoadColor -sRGBCode "rgb(46, 117, 181)" -sName "Royalblue"
 
 return $sRGBList
}
#----------------------------------------------------------------------------------------------------
function Set-ClipboardText()
{
 PARAM($sText ="")
 END
 {
 $Text | clip
 }
}
#----------------------------------------------------------------------------------------------------
# MAIN
#----------------------------------------------------------------------------------------------------

$a = (Get-Host).UI.RawUI
$a.WindowTitle ="RGB code replacer"
$b = $a.WindowSize
#----------------------------------------------------------------------------------------------------

Clear-Host

#----------------------------------------------------------------------------------------------------
# create RGBCode list to convert
#----------------------------------------------------------------------------------------------------
Write-host "Initialising... Please wait..."
Write-host "Initialising... Loading RGB color list..."
$RGBList = @()
$RGBList = LoadColorlist


Write-host "Starting RGB conversion... Please wait..." 
pause "Now go to HTML page, copy HTML source into memory/clipboard."
cls

$SourceText = & {powershell –sta {add-type –a system.windows.forms; [windows.forms.clipboard]::GetText()}}

Write-host "Copied clipboard" 


Write-host "Converting... Using RGB color list..."

foreach ($RGB in $RGBList)
{
$SourceText = $SourceText.replace($RGB.RGBCode, $RGB.Name)
}


if ($SourceText -match"rgb\(") 
{
 Write-Host "WARNING: the text in memory contains unmapped RGB Codes!!" -foregroundcolor white -backgroundcolor darkred
}

#----------------------------------------------------------------------------------------------------
# process RGB codes from memory
#----------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------
#writing transformed data back to memory
#-------------------------------------------------------------------------------------------------------------------

pause "Now go back to the HTML page, now paste HTML source into memory/clipboard."

$SourceText | clip

#$SourceText -match"rgb\("
#pause $SourceText
if ($SourceText -match"rgb\(") 
{
 Write-Host "WARNING: the text in memory contains unmapped RGB Codes!!" -foregroundcolor white -backgroundcolor darkred

}


pause "Command completed successfully"
#-------------------------------------------------------------------------------------------------------------------
Trap 
 { 
 Write-Host "`nError: $($_.Exception.Message)`n" -foregroundcolor white -backgroundcolor darkred
 Exit 1
}
#----------------------------------------------------------------------------------------------------------
