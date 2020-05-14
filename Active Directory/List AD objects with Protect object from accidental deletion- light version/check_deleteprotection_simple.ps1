﻿<##############################################################################
Author: Peter Geelen
Quest For Security 

October 2016
http://identityunderground.wordpress.com

This script finds all AD objects protected from accidental deletions.

Credits:
This script uses logic that has been developed by:
- Ashley McGlone, Microsoft Premier Field Engineer, March 2013, http://aka.ms/GoateePFE 
- Source: https://gallery.technet.microsoft.com/Active-Directory-OU-1d09f989

LEGAL DISCLAIMER
This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, 
provided that You agree: 
(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded;and 
(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
 
This posting is provided "AS IS" with no warranties, and confers no rights. Use of included script samples are subject to the terms specified at http://www.microsoft.com/info/cpyright.htm.
##############################################################################>
#-----------------------------------------------------------------------------
#Source references
#-----------------------------------------------------------------------------
#Preventing Unwanted/Accidental deletions and Restore deleted objects in Active Directory
#abizer_hazratJune 9, 2009
#https://blogs.technet.microsoft.com/abizerh/2009/06/09/preventing-unwantedaccidental-deletions-and-restore-deleted-objects-in-active-directory/

#Windows Server 2008 Protection from Accidental Deletion
#James ONeill, October 31, 2007
#https://blogs.technet.microsoft.com/industry_insiders/2007/10/31/windows-server-2008-protection-from-accidental-deletion/

#-----------------------------------------------------------------------------
#Prerequisites: 
#this script only runs if you can load the AD PS module
#eg. run the analysis on a DC
#-----------------------------------------------------------------------------
cls
import-module activedirectory

#-----------------------------------------------------------------------------
#initialisation
#-----------------------------------------------------------------------------

#the CSV file is saved in the same directory as the PS file
$csvFile = $MyInvocation.MyCommand.Definition -replace 'ps1','csv'
### NEED TO RECONCILE THE CONFLICTS ###
$ErrorActionPreference = 'SilentlyContinue'
Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID |
 ForEach-Object {$schemaIDGUID.add([System.GUID]$_.schemaIDGUID,$_.name)}
Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID |
 ForEach-Object {$schemaIDGUID.add([System.GUID]$_.rightsGUID,$_.name)}
$ErrorActionPreference = 'Continue'
#Functions
#-----------------------------------------------------------------------------
function CheckProtection
        @{name='objectTypeName';expression={if ($_.objectType.ToString() -eq '00000000-0000-0000-0000-000000000000') {'All'} Else {$schemaIDGUID.Item($_.objectType)}}}, `
        @{name='inheritedObjectTypeName';expression={$schemaIDGUID.Item($_.inheritedObjectType)}}, `
        #(*)
        ActiveDirectoryRights,
        AccessControlType,
        IdentityReference,
        IsInherited,
        InheritanceFlags,
        PropagationFlags
#MAIN
#-----------------------------------------------------------------------------
#add the top domain