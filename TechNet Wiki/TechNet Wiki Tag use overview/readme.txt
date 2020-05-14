This script performs an extract of the tags used in the TechNet wiki and counts how many articles are using this tag.

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

WARNING: please realise that there are about 29500 tags in the TNWIKI.
The script takes an hour to run on average.

It has the following features:

execution time tracking
debugging switch
export to CSV for progress tracking and restarting the script after interruption
export to XLS for final reporting
DO NOT RUN this script from a onedrive-sync'ed directory, as the sync engine will interrupt the execution.

The scripts creates the csv and XLS file with the same name as the PowerShell script name.

On a regular basis results are uploaded for use here: TNWiki - tag use overview (v2019-01-20)   

https://gallery.technet.microsoft.com/TNWiki-tag-use-overview-ff5bd65b?redir=0