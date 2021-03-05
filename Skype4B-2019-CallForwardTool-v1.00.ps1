########################################################################
# Name: Skype for Business 2019 Call Forwarding Tool
# Version: v1.00 (15/8/2019)
# Date: 15/8/2019
# Created By: James Cussen
# Web Site: https://www.myskypelab.com
# 
# Notes: This is a PowerShell tool. To run the tool, open it from the PowerShell command line on a Skype for Business 2019 server.
#		 For more information on the requirements for setting up and using this tool please visit https://www.myskypelab.com.
#
# Copyright: Copyright (c) 2019, James Cussen (www.myskypelab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Prerequisites:
#	- Skype for Business 2019 (CU1+) with Call Forwarding PowerShell commands available (ie. Set-CsUserCallForwardingSettings).
#	- Get more information here: https://www.myskypelab.com/
#
# Known Issues: 
#	1. There is a bug in SfB 2019 CU1 when running the following command: Set-CsUserCallForwardingSettings -Identity "sip:john.woods@domain.com" -DisableForwarding -UnansweredToOther "+61392222222" -UnansweredWaitTime 10	
#	Whilst this command sets without error the unanswered call forward is not followed when you call the user. This is due to a bug in the Set-CsUserCallForwardingSettings command which will hopefully be fixed in the next CU.
#
#	2. The Get-CsUserCallForwardingSettings command in CU1 is not capable for displaying Call Forward Immediate to Voice Mail. So if the user has set the Call Forward Immediate to Voicemail on the client it will not be displayed in the tool. The Set-CsUserCallForwardingSettings does not have a Call Forward Immediate to Voicemail option. I have implemented a work around for this issue so you can set Call Forward Immediate to Voice Mail from the tool.
#
# Release Notes:
# 1.00 Initial Release.
#
########################################################################

[cmdletbinding()]
Param()


$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------" -foreground "green"
Write-Host "PowerShell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 PowerShell installed.  This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 PowerShell installed. This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 5 PowerShell installed. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------" -foreground "green"
Write-Host ""


Function Get-MyModule 
{ 
Param([string]$name) 
	
	if(-not(Get-Module -name $name)) 
	{ 
		if(Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) 
		{ 
			Import-Module -Name $name 
			return $true 
		} #end if module available then import 
		else 
		{ 
			return $false 
		} #module not available 
	} # end if not module 
	else 
	{ 
		return $true 
	} #module already loaded 
} #end function get-MyModule 


$Script:SkypeModuleAvailable = $false

Write-Host "--------------------------------------------------------------" -foreground "green"
Write-Host "Importing Modules..." -foreground "yellow"
#Import SkypeforBusiness Module
if(Get-MyModule "SkypeforBusiness")
{
	Invoke-Expression "Import-Module SkypeforBusiness"
	Write-Host "Imported SkypeforBusiness Module..." -foreground "green"
	$Script:SkypeModuleAvailable = $true
}
else
{
	Write-Host "Unable to import SkypeforBusiness Module... (Expected on a Lync 2013 system)" -foreground "yellow"
} 
Write-Host "--------------------------------------------------------------" -foreground "green"


# Check for Skype for Business Commands
$command = "Set-CsUserCallForwardingSettings"
if(Get-Command $command -errorAction SilentlyContinue)
{
	Write-Host
	Write-Host "--------------------------------------------------------------" -foreground "green"
	Write-Host "Set-CsUserCallForwardingSettings is available." -foreground "green"
	Write-Host "--------------------------------------------------------------" -foreground "green"
	Write-Host
}
else
{
	Write-Host
	Write-Host "--------------------------------------------------------------" -foreground "red"
	Write-Host "INFO: Set-CsUserCallForwardingSettings command is not available. Exiting..." -foreground "red"
	Write-Host "--------------------------------------------------------------" -foreground "red"
	Write-Host
	exit
}

$script:users = @()
$script:users = Get-CsUser | Where {$_.EnterpriseVoiceEnabled -eq $true} | select-object SipAddress
$Script:CurrentUserForwardSettings = $null
$Script:UpdatedUserForwardSettings = $null


#STRING FUNCTIONS
function RemoveTel([string] $string)
{
	return [regex]::match($string,'([tT][eE][lL]:)?([^;]*)(;.*)?$').Groups[2].Value
}

function RemoveSip([string] $string)
{
	if($string -eq $null -or $string -eq "")
	{
		return $null
	}
	elseif($string -match "\+.*user=phone") #User phone number
	{
		return [regex]::match($string,'([sS][iI][pP]:)?([^@]*)(@.*)?$').Groups[2].Value 
	}
	else #User SIP Contact Address
	{
		return [regex]::match($string,'([sS][iI][pP]:)?([^;]*)(;.*)?$').Groups[2].Value 
	}
}

function AddSipPhone([string] $string, $domain)
{
	#FINAL FORMAT
	#sip:+61233@sfb2019lab.com;user=phone
	
	if($string -eq $null -or $string -eq "")
	{
		return $null
	}
	elseif($string -match "user=phone") #User phone number
	{
		$string = [regex]::match($string,'([sS][iI][pP]:)?([^@]*)(@.*)?$').Groups[2].Value 
	}
	else #User SIP Contact Address
	{
		$string = [regex]::match($string,'([sS][iI][pP]:)?([^;]*)(;.*)?$').Groups[2].Value 
	}
	
	return "sip:${string}@${domain};user=phone"
}

function AddSip([string] $string, $domain)
{
	#FINAL FORMAT
	#sip:+61233@sfb2019lab.com
	
	if($string -eq $null -or $string -eq "")
	{
		return $null
	}
	elseif($string -match "user=phone") #User phone number
	{
		$string = [regex]::match($string,'([sS][iI][pP]:)?([^@]*)(@.*)?$').Groups[2].Value 
	}
	else #User SIP Contact Address
	{
		$string = [regex]::match($string,'([sS][iI][pP]:)?([^;]*)(;.*)?$').Groups[2].Value 
	}
	
	return "sip:${string}"
}

function GetDomainFromSipAddress([string] $string)
{
	return [regex]::match($string,'([sS][iI][pP]:)?([^@]*)(@)([^;]*)(;.*)?$').Groups[4].Value 
}


# Set up the form  ============================================================
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Skype for Business 2019 Call Forward Tool v1.00"
$objForm.Size = New-Object System.Drawing.Size(1000,560) 
#$objForm.MinimumSize = New-Object System.Drawing.Size(1000,560)
#$objForm.MaximumSize = New-Object System.Drawing.Size(1000,560)
$objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$objForm.MaximizeBox = $false
$objForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$objForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$objForm.KeyPreview = $True
$objForm.TabStop = $false

#Setup the background colour of buttons
[System.Drawing.Color] $Script:buttonBlue = [System.Drawing.Color]::FromArgb(231, 241, 251)
[System.Drawing.Color] $Script:buttonBorderBlue = [System.Drawing.Color]::FromArgb(0, 120, 215)
[System.Drawing.Color] $Script:buttonBorderBlack = [System.Drawing.Color]::FromArgb(0, 0, 0)


# Add the listbox containing all Users ============================================================
$UsersListbox = New-Object System.Windows.Forms.Listbox 
$UsersListbox.Location = New-Object System.Drawing.Size(20,30) 
$UsersListbox.Size = New-Object System.Drawing.Size(200,395) 
$UsersListbox.Sorted = $true
$UsersListbox.tabIndex = 10
$UsersListbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$UsersListbox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$UsersListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$UsersListbox.TabStop = $false

$script:users | ForEach-Object {[void] $UsersListbox.Items.Add((RemoveSip $_.SipAddress))} #.Replace("sip:","")

$objForm.Controls.Add($UsersListbox) 


# Users SelectedIndexChange Event ============================================================
$UsersListbox.add_SelectedIndexChanged(
{
	$StatusLabel.Text = "STATUS: Getting user settings..."
	$ForwardRadioButton.Checked = $false
	$SimRingRadioButton.Checked = $false
	$OffRadioButton.Checked = $false
	ToolIsBusy
	[System.Windows.Forms.Application]::DoEvents()
	
	Write-Verbose "Users List Selected Index Changed Event"
	#$ForwardSettings = Load-TestUserSettingsObject
	$SipAddress = $UsersListbox.SelectedItem
	
	[void] $ForwardOnComboBox.Items.Clear()
	[void] $ForwardOnComboBox.Items.Add("Voice Mail")
	[void] $ForwardOnComboBox.Items.Add("My Delegates")
	[void] $ForwardOnComboBox.Items.Add("New Number or Contact")
	
	[void] $SimRingComboBox.Items.Clear()
	[void] $SimRingComboBox.Items.Add("My Delegates")
	[void] $SimRingComboBox.Items.Add("My Team-Call Group")
	[void] $SimRingComboBox.Items.Add("New Number")

	
	$UserVoicemailSettings = CheckUserVoicemail $SipAddress
	if($UserVoicemailSettings.ExUmEnabled -eq $false -and $UserVoicemailSettings.HostedVoicemail -eq $false)
	{
		if($ForwardOnComboBox.FindStringExact("Voice Mail") -gt -1)
		{
			Write-Verbose "Removing Voice mail Option"
			$ForwardOnComboBox.Items.RemoveAt($ForwardOnComboBox.FindStringExact("Voice Mail"))
		}
		$TheseSettingsWillApplyLink.Enabled = $false
	}
	else
	{
		if($ForwardOnComboBox.FindStringExact("Voice Mail") -eq -1)
		{
			Write-Verbose "Adding Voice mail Option"
			$ForwardOnComboBox.Items.Add("Voice Mail")
		}
		$TheseSettingsWillApplyLink.Enabled = $true
	}
	
	
	try{
		$ForwardSettings = Get-CsUserCallForwardingSettings -identity $SipAddress -ErrorAction Stop
		ToolIsIdle
		Get-ForwardSettings -ForwardSettings $ForwardSettings
		
		$SetUsersListbox.SelectedIndex = -1
		$SetUsersListbox.SelectedIndex = $UsersListbox.FindStringExact($SipAddress)
	}
	catch
	{
		Write-Host "Unable to get user forwarding settings." -foreground "red"
		$StatusLabel.Text = "ERROR: Unable to get user forwarding settings."
		ToolIsIdle
	}
	$StatusLabel.Text = ""

})



$UsersLabel = New-Object System.Windows.Forms.Label
$UsersLabel.Location = New-Object System.Drawing.Size(20,13) 
$UsersLabel.Size = New-Object System.Drawing.Size(200,15) 
$UsersLabel.Text = "Get User Forward Settings:"
$UsersLabel.TabStop = $False
$UsersLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($UsersLabel)


# Filter button ============================================================
$FilterButton = New-Object System.Windows.Forms.Button
$FilterButton.Location = New-Object System.Drawing.Size(173,427)
$FilterButton.Size = New-Object System.Drawing.Size(48,20)
$FilterButton.Text = "Filter"
$FilterButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$FilterButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
$FilterButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$FilterButton.Add_Click({FilterUsersList})
$objForm.Controls.Add($FilterButton)

$FilterButton.Add_MouseHover(
{
   $FilterButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
})
$FilterButton.Add_MouseLeave(
{
   $FilterButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
})

#Filter Text box ============================================================
$FilterTextBox = new-object System.Windows.Forms.textbox
$FilterTextBox.location = new-object system.drawing.size(20,427)
$FilterTextBox.size= new-object system.drawing.size(150,15)
$FilterTextBox.text = ""
$FilterTextBox.TabIndex = 1
$FilterTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$FilterTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		FilterUsersList
	}
})
$objform.controls.add($FilterTextBox)


#Get button
$GetDefaultForwardButton = New-Object System.Windows.Forms.Button
$GetDefaultForwardButton.Location = New-Object System.Drawing.Size(20,460)
$GetDefaultForwardButton.Size = New-Object System.Drawing.Size(200,23)
$GetDefaultForwardButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$GetDefaultForwardButton.Text = "Use Default Forwarding Settings"
$GetDefaultForwardButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$GetDefaultForwardButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue  #[System.Drawing.Color]::CornflowerBlue
#$GetDefaultForwardButton.FlatAppearance.BorderColor = [System.Drawing.Color]::DarkBlue
#$GetDefaultForwardButton.FlatAppearance.BorderSize = 1
$GetDefaultForwardButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Getting default settings..."	
	[System.Windows.Forms.Application]::DoEvents()
	$ForwardSettings = Load-DefaultForwardSettings   
	Get-ForwardSettings -ForwardSettings $ForwardSettings
	
	$CallsWillRingYouAtLabel.Text = "Calls will ring user directly."
	
	$SipAddress = $UsersListbox.SelectedItem
	$UserVoicemailSettings = CheckUserVoicemail $SipAddress
	if($UserVoicemailSettings.ExUmEnabled -eq $false -and $UserVoicemailSettings.HostedVoicemail -eq $false)
	{
		if($ForwardOnComboBox.FindStringExact("Voice Mail") -gt -1)
		{
			Write-Verbose "Removing Voice mail Option"
			$ForwardOnComboBox.Items.RemoveAt($ForwardOnComboBox.FindStringExact("Voice Mail"))
		}
		$TheseSettingsWillApplyLink.Enabled = $false
	}
	else
	{
		if($ForwardOnComboBox.FindStringExact("Voice Mail") -eq -1)
		{
			Write-Verbose "Adding Voice mail Option"
			$ForwardOnComboBox.Items.Add("Voice Mail")
		}
		$TheseSettingsWillApplyLink.Enabled = $true
	}
	
	
	$UnansweredCallForwardDestination = (RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination)  #.Replace("sip:","").Replace("SIP:","")
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{$UnansweredWaitTime = 30}
	else
	{$UnansweredWaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime}
	
	$UserVoicemailSettings = CheckUserVoicemail $SipAddress
	if($UnansweredCallForwardDestination -ne $null -and $UnansweredCallForwardDestination -ne "")
	{
		$UnansweredCallsWillGoToLink.Text = "$UnansweredCallForwardDestination in $UnansweredWaitTime seconds"
		Write-Host "INFO: $UnansweredCallForwardDestination in $UnansweredWaitTime seconds" -foreground "Yellow"
	}
	elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail -or ($UserVoicemailSettings.ExUmEnabled -eq $true -or $UserVoicemailSettings.HostedVoicemail -eq $true))
	{
		$UnansweredCallsWillGoToLink.Text = "Voice mail in $UnansweredWaitTime seconds"
		Write-Host "INFO: Voice mail in $UnansweredWaitTime seconds" -foreground "Yellow"
	}
	elseif($UnansweredCallForwardDestination -eq "" -or $UnansweredCallForwardDestination -eq $null)
	{
		$UnansweredCallsWillGoToLink.Text = "No voice mail. Calls will continuously ring for $UnansweredWaitTime secs."
		Write-Host "INFO: No voice mail. Calls will continuously ring for $UnansweredWaitTime seconds." -foreground "Yellow"
	}
			
	$StatusLabel.Text = ""

})
$objForm.Controls.Add($GetDefaultForwardButton)
$GetDefaultForwardButton.Enabled = $true


$GetDefaultForwardButton.Add_MouseHover(
{
   $GetDefaultForwardButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
})
$GetDefaultForwardButton.Add_MouseLeave(
{
   $GetDefaultForwardButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
})



[byte[]]$ArrowImage = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 25, 0, 0, 0, 50, 8, 6, 0, 0, 0, 169, 130, 112, 150, 0, 0, 13, 44, 122, 84, 88, 116, 82, 97, 119, 32, 112, 114, 111, 102, 105, 108, 101, 32, 116, 121, 112, 101, 32, 101, 120, 105, 102, 0, 0, 120, 218, 173, 152, 107, 114, 35, 59, 174, 132, 255, 115, 21, 179, 132, 34, 65, 16, 228, 114, 248, 140, 184, 59, 152, 229, 207, 7, 74, 118, 63, 78, 143, 187, 251, 198, 88, 97, 75, 46, 177, 248, 0, 18, 153, 137, 10, 251, 223, 255, 119, 194, 191, 248, 17, 205, 37, 100, 181, 90, 90, 41, 15, 63, 185, 229, 150, 58, 31, 234, 243, 250, 121, 189, 199, 39, 223, 191, 175, 127, 218, 251, 187, 248, 227, 245, 32, 233, 253, 69, 226, 146, 240, 46, 175, 127, 203, 126, 143, 239, 92, 215, 111, 55, 88, 126, 95, 31, 63, 94, 15, 54, 223, 243, 212, 247, 68, 239, 47, 62, 38, 20, 95, 217, 23, 123, 143, 171, 239, 137, 36, 189, 174, 199, 247, 255, 225, 99, 167, 61, 127, 119, 156, 247, 239, 220, 119, 138, 39, 190, 39, 253, 249, 255, 108, 4, 99, 41, 23, 37, 133, 180, 133, 235, 252, 189, 171, 8, 59, 144, 42, 157, 191, 145, 191, 34, 62, 40, 242, 222, 69, 239, 223, 44, 241, 215, 177, 11, 31, 113, 253, 57, 120, 159, 159, 126, 138, 221, 211, 223, 215, 229, 199, 80, 132, 167, 188, 7, 148, 159, 98, 244, 190, 30, 245, 215, 177, 187, 17, 250, 126, 71, 241, 219, 202, 63, 124, 49, 243, 71, 120, 255, 25, 187, 115, 86, 61, 103, 191, 78, 215, 115, 33, 82, 37, 188, 15, 245, 113, 148, 251, 137, 129, 131, 80, 202, 189, 173, 240, 50, 126, 149, 207, 118, 95, 141, 87, 229, 136, 147, 140, 45, 150, 27, 188, 102, 136, 45, 38, 162, 125, 98, 142, 43, 246, 120, 226, 190, 239, 51, 78, 182, 152, 211, 78, 198, 123, 74, 51, 201, 189, 86, 197, 82, 75, 243, 38, 37, 251, 43, 158, 100, 210, 100, 5, 114, 148, 100, 146, 53, 225, 114, 250, 220, 75, 188, 235, 182, 187, 222, 140, 149, 149, 87, 100, 100, 138, 76, 22, 185, 227, 31, 175, 240, 171, 139, 255, 159, 215, 231, 68, 231, 120, 108, 99, 124, 234, 43, 78, 245, 38, 56, 57, 0, 217, 134, 103, 206, 255, 50, 138, 132, 196, 243, 142, 169, 222, 248, 222, 87, 248, 14, 55, 207, 119, 137, 21, 50, 168, 55, 204, 149, 3, 246, 103, 188, 166, 24, 26, 191, 97, 75, 110, 158, 133, 113, 250, 228, 240, 188, 74, 35, 218, 122, 79, 64, 136, 88, 91, 217, 12, 232, 206, 241, 41, 81, 52, 150, 248, 88, 74, 22, 35, 113, 172, 228, 167, 179, 243, 36, 57, 13, 50, 16, 85, 211, 138, 225, 144, 27, 145, 66, 114, 106, 242, 181, 185, 199, 226, 29, 155, 52, 189, 46, 67, 45, 36, 66, 165, 136, 145, 154, 38, 157, 100, 229, 12, 7, 81, 107, 21, 12, 117, 133, 144, 130, 170, 22, 53, 173, 218, 180, 23, 41, 185, 104, 41, 197, 138, 115, 84, 55, 177, 108, 106, 197, 204, 170, 53, 235, 85, 106, 174, 90, 75, 181, 90, 107, 171, 189, 165, 38, 80, 152, 182, 210, 44, 180, 218, 90, 235, 157, 69, 59, 83, 119, 238, 238, 140, 232, 125, 164, 33, 35, 15, 29, 101, 216, 168, 163, 141, 62, 129, 207, 204, 83, 103, 153, 54, 235, 108, 179, 175, 180, 100, 81, 254, 171, 44, 11, 171, 174, 182, 250, 142, 27, 40, 237, 188, 117, 151, 109, 187, 238, 182, 251, 1, 107, 71, 78, 62, 122, 202, 177, 83, 79, 59, 253, 51, 107, 241, 93, 182, 63, 100, 45, 254, 148, 185, 175, 179, 22, 223, 89, 243, 140, 229, 59, 206, 190, 101, 141, 203, 102, 31, 83, 68, 167, 19, 245, 156, 145, 177, 148, 35, 25, 55, 207, 0, 128, 78, 158, 179, 167, 198, 156, 147, 103, 206, 115, 246, 180, 68, 81, 104, 34, 107, 81, 61, 57, 43, 122, 198, 200, 96, 222, 49, 233, 137, 159, 185, 251, 150, 185, 47, 243, 22, 52, 255, 85, 222, 210, 127, 203, 92, 240, 212, 253, 47, 50, 23, 60, 117, 239, 204, 253, 51, 111, 191, 200, 218, 234, 87, 81, 228, 38, 200, 171, 208, 99, 250, 200, 129, 216, 24, 176, 107, 79, 109, 214, 122, 210, 172, 68, 176, 212, 185, 247, 204, 169, 247, 206, 170, 41, 31, 18, 130, 132, 73, 97, 159, 123, 151, 89, 248, 118, 20, 230, 236, 171, 24, 12, 181, 173, 156, 208, 70, 108, 186, 86, 221, 163, 206, 179, 8, 160, 11, 149, 198, 172, 127, 245, 110, 37, 180, 29, 237, 60, 205, 56, 218, 179, 228, 104, 59, 195, 202, 90, 35, 213, 38, 147, 195, 13, 57, 22, 79, 247, 112, 219, 26, 187, 233, 212, 101, 125, 177, 17, 57, 59, 171, 12, 168, 5, 124, 49, 209, 57, 168, 88, 119, 74, 143, 9, 27, 224, 151, 1, 64, 60, 3, 138, 234, 122, 86, 59, 179, 46, 191, 202, 185, 143, 237, 205, 238, 153, 154, 24, 18, 204, 198, 217, 218, 96, 76, 236, 97, 228, 83, 100, 151, 211, 198, 214, 178, 136, 87, 217, 107, 110, 78, 184, 91, 147, 61, 116, 47, 209, 121, 246, 88, 59, 142, 37, 108, 113, 117, 98, 158, 87, 27, 185, 217, 44, 64, 175, 206, 190, 79, 11, 105, 70, 123, 70, 85, 215, 135, 254, 32, 187, 249, 137, 155, 240, 106, 62, 230, 0, 180, 162, 195, 246, 210, 50, 100, 198, 113, 100, 216, 92, 126, 235, 158, 6, 8, 108, 3, 26, 42, 232, 73, 129, 124, 85, 255, 240, 187, 247, 193, 214, 169, 184, 202, 172, 179, 61, 5, 40, 8, 176, 97, 157, 221, 114, 182, 93, 2, 103, 170, 103, 236, 179, 44, 130, 186, 178, 143, 58, 160, 192, 194, 70, 161, 234, 72, 185, 108, 62, 83, 19, 105, 128, 144, 232, 198, 33, 161, 41, 143, 13, 10, 88, 183, 25, 1, 101, 36, 196, 198, 228, 133, 32, 0, 201, 220, 111, 190, 60, 89, 146, 199, 108, 75, 168, 141, 200, 10, 171, 143, 49, 129, 249, 26, 182, 50, 85, 153, 54, 65, 206, 71, 82, 180, 185, 65, 52, 98, 107, 193, 99, 83, 101, 30, 56, 163, 98, 71, 108, 102, 0, 49, 60, 75, 212, 20, 199, 24, 103, 55, 59, 197, 78, 18, 243, 171, 115, 200, 58, 3, 136, 200, 154, 69, 186, 90, 43, 122, 8, 35, 196, 54, 244, 161, 140, 148, 175, 207, 108, 20, 174, 144, 96, 194, 248, 100, 240, 45, 231, 233, 36, 187, 111, 131, 142, 26, 42, 45, 214, 238, 128, 65, 137, 129, 4, 50, 72, 1, 16, 251, 30, 106, 237, 51, 239, 243, 104, 221, 91, 42, 39, 220, 231, 35, 111, 58, 22, 17, 153, 99, 174, 4, 133, 150, 103, 16, 220, 217, 161, 104, 177, 170, 139, 28, 82, 63, 177, 237, 167, 29, 164, 33, 40, 158, 239, 233, 55, 57, 127, 240, 78, 33, 79, 209, 157, 250, 0, 98, 70, 37, 64, 71, 181, 175, 185, 70, 128, 64, 252, 184, 240, 193, 64, 229, 207, 60, 210, 166, 153, 16, 95, 248, 138, 155, 101, 109, 102, 113, 84, 83, 71, 181, 150, 229, 190, 97, 20, 74, 170, 75, 91, 135, 96, 55, 160, 213, 149, 96, 199, 186, 86, 155, 59, 215, 227, 249, 108, 78, 62, 220, 211, 136, 146, 78, 160, 91, 83, 198, 158, 204, 10, 179, 82, 72, 144, 98, 121, 101, 133, 170, 124, 186, 23, 59, 243, 55, 147, 96, 146, 22, 158, 165, 207, 130, 129, 76, 169, 168, 139, 45, 54, 101, 70, 16, 69, 170, 136, 53, 81, 148, 3, 33, 215, 229, 97, 135, 8, 163, 66, 118, 16, 28, 179, 136, 123, 22, 24, 18, 159, 13, 103, 87, 67, 46, 150, 47, 4, 246, 56, 114, 132, 102, 114, 25, 117, 61, 160, 146, 226, 130, 236, 16, 7, 88, 42, 87, 156, 92, 110, 190, 75, 114, 67, 77, 105, 49, 117, 139, 172, 17, 231, 255, 250, 240, 229, 59, 132, 69, 162, 108, 42, 1, 54, 237, 206, 217, 13, 182, 80, 248, 23, 117, 57, 205, 70, 67, 178, 1, 23, 142, 115, 19, 228, 45, 227, 161, 62, 39, 224, 117, 8, 128, 124, 240, 54, 222, 83, 128, 62, 24, 1, 185, 25, 132, 78, 237, 129, 80, 132, 126, 34, 214, 166, 32, 46, 50, 81, 92, 179, 158, 12, 53, 0, 56, 16, 199, 177, 102, 26, 220, 230, 7, 210, 73, 66, 9, 62, 209, 136, 181, 206, 7, 149, 112, 147, 10, 251, 154, 7, 99, 16, 10, 8, 38, 141, 231, 4, 214, 66, 177, 102, 25, 173, 67, 126, 101, 81, 119, 209, 99, 63, 49, 144, 69, 147, 223, 166, 37, 174, 234, 177, 119, 114, 19, 120, 99, 109, 116, 198, 242, 34, 68, 157, 92, 181, 76, 93, 135, 149, 166, 130, 208, 236, 225, 67, 129, 215, 124, 184, 163, 183, 191, 34, 108, 211, 205, 209, 136, 25, 177, 92, 5, 169, 67, 209, 14, 186, 145, 97, 147, 132, 220, 66, 97, 79, 115, 148, 1, 149, 72, 31, 98, 190, 219, 207, 8, 66, 168, 199, 252, 151, 50, 105, 105, 5, 216, 39, 53, 220, 26, 251, 236, 201, 56, 72, 60, 121, 186, 60, 66, 28, 50, 29, 215, 235, 236, 25, 41, 81, 148, 242, 12, 23, 47, 72, 89, 18, 204, 5, 203, 108, 239, 210, 96, 129, 210, 130, 135, 89, 133, 227, 83, 224, 74, 241, 174, 221, 22, 176, 70, 53, 173, 151, 9, 213, 12, 250, 157, 17, 17, 114, 179, 217, 138, 179, 178, 187, 237, 42, 203, 107, 2, 200, 2, 237, 17, 199, 8, 240, 193, 56, 165, 15, 165, 140, 225, 126, 20, 7, 1, 228, 38, 79, 27, 138, 227, 85, 179, 69, 171, 108, 50, 245, 92, 53, 33, 13, 114, 32, 161, 147, 199, 65, 172, 218, 139, 189, 194, 102, 134, 67, 24, 6, 167, 150, 184, 188, 76, 89, 216, 37, 245, 80, 199, 123, 58, 252, 216, 24, 99, 144, 49, 74, 155, 97, 28, 69, 192, 54, 236, 121, 234, 167, 160, 134, 63, 76, 140, 194, 205, 228, 37, 13, 40, 157, 180, 12, 248, 4, 102, 130, 82, 70, 115, 119, 114, 176, 53, 107, 103, 142, 104, 201, 105, 13, 246, 135, 1, 183, 94, 225, 0, 115, 191, 225, 233, 239, 104, 186, 3, 72, 131, 86, 152, 115, 42, 177, 202, 85, 242, 89, 145, 35, 9, 147, 67, 72, 66, 110, 64, 183, 45, 68, 150, 42, 246, 121, 93, 176, 241, 206, 8, 9, 162, 13, 199, 224, 235, 32, 209, 224, 103, 118, 127, 235, 10, 107, 43, 197, 107, 186, 112, 99, 24, 53, 171, 9, 142, 139, 211, 241, 99, 40, 75, 89, 121, 100, 93, 184, 53, 108, 203, 128, 177, 203, 217, 94, 140, 13, 89, 151, 208, 57, 77, 196, 136, 228, 196, 17, 212, 70, 94, 229, 57, 110, 32, 174, 104, 80, 227, 11, 17, 89, 9, 44, 17, 236, 235, 12, 14, 156, 55, 14, 24, 154, 184, 175, 201, 182, 189, 114, 52, 80, 58, 6, 1, 151, 51, 98, 233, 17, 149, 30, 12, 58, 13, 61, 204, 104, 32, 75, 64, 238, 17, 246, 37, 214, 207, 87, 76, 19, 126, 186, 32, 45, 2, 230, 22, 137, 22, 4, 28, 87, 174, 88, 74, 37, 219, 16, 64, 205, 16, 36, 162, 146, 160, 5, 196, 199, 245, 168, 209, 223, 9, 149, 185, 87, 24, 158, 40, 119, 40, 134, 89, 101, 187, 219, 129, 136, 66, 68, 243, 121, 109, 244, 76, 171, 136, 26, 148, 19, 225, 127, 170, 176, 118, 205, 207, 174, 174, 125, 116, 63, 216, 0, 124, 39, 117, 24, 224, 158, 151, 186, 116, 178, 60, 137, 163, 185, 188, 192, 125, 245, 204, 177, 102, 191, 234, 210, 171, 87, 45, 218, 13, 41, 131, 77, 68, 64, 240, 85, 49, 149, 42, 212, 84, 193, 48, 21, 108, 13, 54, 3, 146, 196, 140, 61, 110, 99, 49, 183, 52, 79, 50, 102, 44, 50, 42, 155, 35, 178, 20, 31, 86, 153, 13, 251, 206, 33, 199, 241, 40, 248, 192, 68, 245, 217, 56, 41, 236, 189, 102, 32, 24, 108, 57, 159, 30, 251, 1, 102, 127, 199, 144, 31, 114, 1, 128, 195, 31, 88, 163, 218, 105, 250, 4, 59, 2, 172, 98, 23, 160, 74, 151, 192, 143, 164, 140, 110, 93, 109, 160, 205, 170, 24, 24, 12, 14, 30, 130, 248, 187, 132, 226, 234, 150, 131, 4, 204, 65, 108, 180, 72, 234, 250, 137, 7, 128, 35, 80, 188, 225, 65, 69, 63, 144, 236, 27, 84, 136, 2, 210, 1, 217, 20, 221, 153, 148, 119, 159, 87, 179, 61, 170, 5, 22, 0, 210, 30, 213, 78, 226, 49, 158, 11, 154, 246, 79, 186, 96, 166, 92, 35, 157, 10, 162, 115, 31, 111, 104, 190, 180, 92, 40, 145, 33, 158, 182, 248, 169, 134, 37, 141, 250, 77, 14, 59, 213, 233, 29, 208, 104, 136, 145, 70, 7, 189, 238, 74, 209, 56, 94, 165, 56, 203, 144, 240, 147, 67, 161, 167, 129, 189, 86, 108, 21, 40, 178, 198, 162, 147, 225, 148, 184, 8, 194, 140, 86, 225, 121, 4, 241, 171, 206, 91, 165, 80, 113, 170, 25, 81, 25, 165, 42, 68, 155, 217, 233, 227, 205, 72, 240, 205, 253, 137, 171, 113, 67, 227, 209, 233, 133, 13, 104, 98, 50, 197, 36, 51, 63, 90, 203, 142, 71, 216, 0, 132, 168, 16, 94, 157, 38, 183, 210, 93, 36, 28, 3, 199, 125, 38, 186, 26, 59, 20, 34, 78, 15, 148, 1, 125, 188, 87, 2, 162, 15, 163, 162, 19, 123, 111, 96, 184, 82, 24, 112, 86, 31, 11, 31, 70, 38, 32, 147, 218, 10, 198, 198, 245, 252, 161, 111, 0, 199, 248, 135, 72, 75, 6, 1, 111, 235, 4, 168, 220, 61, 98, 86, 188, 81, 104, 80, 82, 68, 220, 7, 29, 228, 241, 173, 227, 155, 96, 58, 184, 157, 255, 112, 135, 197, 221, 109, 165, 37, 104, 27, 94, 241, 90, 24, 235, 193, 21, 122, 45, 148, 233, 166, 211, 88, 22, 223, 206, 15, 197, 32, 242, 72, 24, 211, 25, 46, 95, 150, 29, 78, 83, 5, 129, 67, 46, 99, 31, 167, 209, 169, 193, 139, 78, 146, 125, 130, 172, 22, 41, 230, 134, 81, 237, 8, 36, 205, 36, 54, 126, 29, 26, 12, 104, 19, 199, 230, 141, 33, 246, 77, 38, 217, 237, 132, 207, 77, 9, 144, 101, 248, 137, 222, 122, 254, 62, 13, 219, 65, 26, 28, 165, 152, 120, 79, 132, 121, 34, 118, 165, 229, 185, 22, 226, 96, 43, 210, 167, 48, 16, 34, 206, 118, 93, 223, 112, 1, 38, 15, 123, 237, 43, 192, 157, 104, 91, 64, 223, 146, 74, 220, 139, 146, 192, 232, 83, 57, 240, 133, 185, 147, 112, 1, 198, 124, 186, 100, 56, 77, 1, 119, 12, 223, 84, 250, 79, 104, 1, 79, 77, 215, 30, 233, 159, 106, 84, 81, 127, 88, 247, 86, 81, 114, 0, 80, 112, 245, 87, 67, 97, 150, 108, 168, 252, 237, 222, 212, 181, 155, 141, 69, 88, 90, 188, 65, 115, 95, 217, 232, 235, 105, 157, 161, 159, 65, 79, 80, 81, 145, 117, 30, 172, 92, 162, 129, 101, 155, 130, 219, 192, 156, 83, 27, 207, 198, 254, 78, 152, 136, 110, 106, 121, 225, 226, 208, 0, 21, 149, 122, 100, 79, 164, 13, 140, 163, 119, 131, 169, 73, 192, 129, 70, 156, 40, 14, 61, 2, 112, 187, 234, 140, 213, 34, 24, 136, 195, 74, 84, 53, 180, 216, 188, 14, 228, 212, 149, 191, 210, 227, 240, 254, 112, 205, 146, 186, 115, 246, 62, 203, 159, 34, 221, 142, 168, 108, 38, 67, 66, 0, 150, 58, 184, 106, 201, 8, 178, 68, 239, 55, 128, 22, 229, 229, 205, 74, 213, 50, 131, 239, 60, 109, 36, 124, 64, 167, 160, 116, 21, 88, 164, 250, 131, 144, 202, 9, 59, 90, 85, 111, 71, 74, 76, 94, 203, 23, 24, 192, 205, 139, 187, 103, 12, 9, 233, 47, 171, 165, 20, 32, 12, 108, 64, 201, 232, 156, 251, 231, 142, 234, 129, 249, 134, 205, 254, 11, 118, 217, 217, 130, 23, 40, 137, 94, 217, 119, 111, 26, 91, 161, 57, 211, 140, 138, 227, 80, 179, 63, 119, 245, 199, 56, 155, 118, 169, 235, 139, 39, 126, 77, 202, 225, 227, 195, 160, 36, 30, 183, 229, 133, 200, 211, 56, 157, 165, 43, 57, 71, 224, 39, 11, 157, 21, 59, 20, 8, 56, 209, 31, 95, 10, 165, 236, 184, 147, 190, 11, 185, 132, 28, 115, 88, 8, 54, 130, 233, 94, 212, 242, 73, 20, 37, 208, 108, 236, 30, 88, 22, 8, 21, 177, 192, 30, 117, 4, 134, 172, 42, 126, 1, 157, 130, 5, 75, 106, 120, 199, 17, 59, 129, 237, 156, 170, 134, 216, 48, 158, 175, 62, 250, 57, 110, 186, 17, 160, 82, 205, 77, 55, 242, 74, 103, 135, 104, 100, 144, 154, 48, 50, 238, 240, 233, 0, 86, 117, 131, 47, 180, 71, 24, 98, 212, 8, 219, 210, 10, 59, 114, 16, 250, 232, 130, 59, 47, 73, 149, 46, 137, 190, 229, 221, 14, 36, 247, 169, 183, 203, 185, 121, 50, 4, 39, 109, 72, 48, 165, 108, 248, 115, 34, 158, 243, 227, 249, 13, 244, 194, 30, 72, 31, 249, 117, 99, 147, 55, 253, 125, 43, 153, 61, 8, 110, 205, 10, 162, 162, 254, 80, 139, 152, 176, 88, 72, 222, 23, 119, 115, 203, 150, 160, 159, 113, 45, 27, 60, 10, 179, 65, 187, 157, 42, 164, 218, 59, 156, 3, 44, 91, 195, 217, 31, 178, 225, 19, 123, 253, 63, 20, 177, 217, 109, 247, 3, 114, 40, 236, 134, 162, 124, 245, 251, 208, 246, 178, 219, 239, 139, 124, 244, 251, 240, 130, 123, 24, 224, 68, 175, 34, 88, 235, 8, 87, 147, 85, 186, 106, 90, 73, 56, 28, 118, 165, 57, 46, 125, 190, 173, 235, 203, 19, 26, 77, 167, 92, 27, 188, 176, 14, 164, 159, 105, 98, 243, 99, 40, 237, 32, 45, 162, 145, 83, 188, 254, 244, 157, 148, 238, 161, 166, 155, 244, 199, 62, 15, 152, 184, 6, 47, 250, 178, 148, 239, 219, 224, 65, 155, 254, 232, 6, 168, 207, 219, 150, 78, 40, 3, 106, 33, 223, 8, 45, 164, 205, 73, 218, 194, 217, 209, 210, 73, 11, 192, 183, 193, 127, 135, 8, 62, 144, 60, 230, 165, 225, 120, 58, 14, 132, 109, 177, 41, 26, 1, 154, 165, 241, 96, 69, 137, 8, 131, 39, 72, 55, 127, 148, 215, 126, 224, 128, 240, 133, 73, 255, 197, 115, 171, 247, 99, 171, 65, 115, 0, 211, 47, 60, 55, 248, 245, 231, 80, 57, 184, 161, 242, 199, 86, 254, 44, 10, 246, 166, 181, 129, 246, 182, 165, 153, 17, 210, 254, 242, 64, 203, 59, 151, 67, 13, 248, 99, 13, 0, 5, 231, 30, 133, 163, 65, 238, 62, 21, 79, 18, 163, 210, 66, 28, 252, 6, 174, 202, 91, 44, 34, 169, 112, 118, 1, 228, 8, 31, 70, 255, 96, 93, 176, 201, 169, 220, 7, 44, 84, 132, 55, 91, 152, 68, 218, 25, 230, 42, 216, 129, 113, 31, 196, 97, 107, 232, 141, 232, 189, 252, 185, 10, 4, 180, 204, 151, 191, 122, 245, 71, 252, 250, 124, 211, 171, 240, 167, 79, 69, 126, 247, 254, 155, 137, 4, 17, 95, 45, 252, 7, 122, 142, 75, 193, 238, 55, 127, 84, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 14, 195, 0, 0, 14, 195, 1, 199, 111, 168, 100, 0, 0, 0, 7, 116, 73, 77, 69, 7, 226, 11, 30, 2, 53, 42, 42, 64, 222, 154, 0, 0, 1, 75, 73, 68, 65, 84, 88, 195, 237, 215, 189, 74, 4, 49, 16, 192, 241, 255, 238, 88, 217, 216, 216, 8, 158, 133, 175, 97, 176, 177, 211, 86, 123, 31, 99, 74, 27, 33, 246, 86, 22, 62, 132, 205, 217, 169, 16, 65, 43, 31, 193, 198, 94, 16, 20, 44, 4, 155, 19, 214, 115, 247, 54, 31, 19, 176, 184, 169, 134, 36, 236, 47, 9, 73, 152, 133, 101, 228, 134, 170, 54, 53, 190, 219, 118, 128, 19, 224, 85, 85, 247, 170, 33, 192, 39, 176, 6, 92, 89, 67, 93, 196, 3, 231, 192, 170, 53, 36, 63, 73, 8, 1, 231, 220, 53, 176, 14, 236, 0, 135, 206, 185, 199, 16, 194, 179, 25, 210, 129, 166, 214, 144, 204, 55, 212, 128, 164, 175, 113, 0, 122, 200, 133, 100, 168, 163, 7, 58, 202, 133, 100, 81, 167, 21, 36, 99, 3, 44, 32, 137, 25, 84, 10, 73, 236, 108, 74, 32, 73, 217, 219, 92, 72, 82, 79, 74, 14, 36, 57, 231, 62, 21, 146, 220, 91, 156, 2, 73, 201, 155, 20, 11, 73, 233, 11, 27, 3, 21, 35, 49, 144, 9, 50, 6, 153, 23, 14, 179, 98, 228, 18, 56, 6, 222, 129, 141, 182, 66, 113, 50, 1, 118, 103, 249, 19, 240, 209, 26, 175, 98, 11, 184, 1, 182, 129, 123, 224, 192, 123, 255, 213, 84, 4, 246, 189, 247, 111, 0, 77, 109, 192, 4, 25, 3, 138, 145, 24, 160, 8, 137, 5, 178, 145, 20, 32, 11, 73, 5, 146, 17, 85, 157, 0, 183, 41, 64, 18, 146, 11, 68, 35, 37, 64, 20, 82, 10, 140, 34, 22, 192, 66, 196, 10, 24, 68, 44, 129, 94, 196, 26, 248, 131, 212, 0, 126, 33, 170, 186, 9, 220, 89, 3, 243, 127, 191, 103, 53, 0, 128, 149, 78, 126, 1, 188, 0, 167, 150, 192, 50, 254, 103, 124, 3, 191, 217, 32, 204, 12, 184, 191, 36, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)

$ArrowImage1 = new-object Windows.Forms.PictureBox
$ArrowImage1.Location = New-Object System.Drawing.Size(225,120) 
$ArrowImage1.width = 30
$ArrowImage1.height = 30
$ArrowImage1.BorderStyle = "FixedSingle"
$ArrowImage1.sizemode = "Zoom"
$ArrowImage1.Margin = 0
$ArrowImage1.WaitOnLoad = $true
$ArrowImage1.BorderStyle =  [System.Windows.Forms.BorderStyle]::None
$ArrowImage1.Image = $ArrowImage
$ArrowImage1.Add_Click(
{
	#$StatusLabel.Text = "Status: Settings..."
})
$objForm.Controls.Add($ArrowImage1)


[byte[]]$ForwardOn = @(0, 0, 1, 0, 2, 0, 24, 24, 0, 0, 1, 0, 32, 0, 136, 9, 0, 0, 38, 0, 0, 0, 16, 16, 0, 0, 1, 0, 32, 0, 104, 4, 0, 0, 174, 9, 0, 0, 40, 0, 0, 0, 24, 0, 0, 0, 48, 0, 0, 0, 1, 0, 32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 59, 59, 59, 131, 61, 61, 61, 222, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 255, 64, 64, 64, 213, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 59, 59, 59, 177, 65, 65, 65, 239, 134, 135, 135, 222, 203, 204, 204, 255, 236, 236, 236, 255, 235, 236, 236, 255, 235, 235, 235, 255, 235, 235, 234, 255, 180, 181, 181, 255, 69, 69, 69, 233, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 67, 67, 67, 137, 68, 67, 67, 242, 153, 153, 153, 238, 238, 238, 238, 255, 237, 237, 237, 255, 237, 237, 237, 255, 236, 236, 236, 255, 236, 236, 236, 255, 235, 235, 235, 255, 235, 235, 235, 255, 235, 235, 234, 255, 180, 180, 180, 253, 71, 71, 71, 221, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 11, 65, 65, 65, 213, 119, 119, 119, 241, 225, 225, 225, 255, 240, 239, 240, 255, 239, 239, 239, 255, 238, 238, 239, 255, 237, 237, 238, 255, 237, 237, 237, 255, 236, 236, 237, 255, 236, 236, 236, 255, 235, 235, 235, 255, 235, 235, 235, 255, 235, 235, 235, 255, 59, 59, 59, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 31, 61, 61, 61, 247, 184, 184, 184, 247, 241, 241, 241, 255, 240, 240, 240, 255, 240, 240, 240, 255, 240, 240, 239, 255, 238, 239, 238, 255, 238, 238, 238, 255, 238, 237, 237, 255, 237, 237, 237, 255, 237, 236, 237, 255, 236, 236, 236, 255, 235, 236, 235, 255, 235, 235, 235, 255, 59, 59, 59, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 54, 73, 73, 73, 237, 181, 182, 181, 255, 243, 243, 243, 255, 242, 242, 242, 255, 241, 242, 242, 255, 226, 225, 226, 255, 92, 92, 92, 234, 131, 131, 131, 243, 239, 239, 239, 255, 239, 238, 238, 255, 237, 237, 237, 255, 237, 237, 237, 255, 237, 236, 236, 255, 236, 236, 236, 255, 96, 96, 96, 224, 62, 62, 62, 116, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 109, 109, 109, 41, 94, 94, 94, 236, 182, 182, 182, 255, 245, 244, 244, 255, 243, 244, 244, 255, 243, 242, 243, 255, 200, 200, 200, 255, 63, 63, 63, 250, 59, 59, 59, 96, 74, 74, 74, 160, 130, 130, 130, 240, 239, 238, 239, 255, 238, 238, 238, 255, 237, 237, 238, 255, 237, 237, 237, 255, 91, 91, 91, 216, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 5, 118, 118, 118, 225, 195, 195, 195, 255, 246, 245, 246, 255, 245, 245, 245, 255, 244, 244, 244, 255, 173, 174, 173, 253, 71, 71, 71, 234, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 66, 66, 66, 184, 127, 127, 127, 236, 223, 224, 223, 255, 223, 223, 223, 249, 84, 84, 84, 206, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 115, 116, 116, 191, 184, 184, 184, 255, 247, 247, 247, 255, 246, 246, 246, 255, 245, 246, 246, 255, 160, 160, 160, 250, 76, 76, 76, 212, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 66, 66, 66, 153, 62, 62, 62, 248, 61, 61, 61, 247, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 116, 116, 116, 138, 128, 128, 128, 254, 249, 248, 248, 255, 247, 248, 248, 255, 247, 247, 247, 255, 173, 173, 173, 244, 71, 71, 71, 206, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 38, 78, 78, 78, 233, 234, 234, 234, 255, 249, 249, 249, 255, 249, 249, 249, 255, 204, 204, 204, 253, 68, 68, 68, 230, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 70, 70, 70, 188, 164, 164, 164, 252, 250, 250, 250, 255, 250, 250, 250, 255, 233, 233, 233, 255, 66, 66, 66, 239, 59, 59, 59, 41, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 153, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 73, 73, 73, 249, 251, 252, 251, 255, 251, 251, 251, 255, 251, 251, 250, 255, 82, 82, 82, 219, 59, 59, 59, 96, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 252, 61, 187, 75, 36, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 136, 61, 187, 75, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 108, 108, 108, 175, 156, 156, 155, 255, 253, 252, 252, 255, 251, 252, 252, 255, 251, 251, 251, 255, 91, 90, 90, 175, 59, 59, 59, 147, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 249, 61, 187, 75, 171, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 255, 61, 187, 75, 249, 61, 187, 75, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 77, 77, 77, 242, 217, 217, 217, 255, 252, 253, 253, 255, 252, 252, 252, 255, 252, 252, 252, 255, 251, 251, 251, 237, 89, 89, 89, 173, 59, 59, 59, 177, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 228, 61, 187, 75, 255, 61, 187, 75, 105, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 249, 61, 187, 75, 42, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 255, 253, 253, 254, 255, 253, 253, 253, 255, 252, 253, 252, 255, 252, 252, 252, 255, 252, 252, 251, 255, 251, 251, 251, 249, 97, 96, 96, 182, 59, 59, 59, 147, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 186, 61, 187, 75, 255, 61, 187, 75, 252, 61, 187, 75, 141, 61, 187, 75, 51, 61, 187, 75, 51, 61, 187, 75, 51, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 249, 61, 187, 75, 42, 0, 0, 0, 0, 59, 59, 59, 255, 254, 254, 254, 255, 253, 253, 253, 255, 253, 253, 253, 255, 253, 252, 252, 255, 252, 252, 252, 255, 252, 251, 251, 255, 235, 235, 235, 255, 60, 60, 60, 246, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 108, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 249, 61, 187, 75, 42, 59, 59, 59, 255, 255, 254, 254, 255, 254, 254, 254, 255, 254, 253, 253, 255, 253, 253, 253, 255, 252, 252, 253, 255, 252, 252, 252, 255, 235, 236, 235, 255, 60, 61, 60, 246, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 15, 61, 187, 75, 225, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 249, 59, 59, 59, 255, 255, 255, 255, 255, 254, 254, 254, 255, 254, 254, 254, 255, 254, 254, 253, 255, 253, 253, 253, 255, 252, 252, 252, 255, 92, 92, 92, 213, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 57, 61, 187, 75, 237, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 81, 81, 81, 231, 195, 195, 195, 255, 255, 254, 255, 255, 254, 254, 254, 255, 254, 253, 253, 255, 254, 254, 253, 255, 89, 89, 89, 210, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 33, 61, 187, 75, 204, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 42, 59, 59, 59, 21, 74, 74, 74, 238, 193, 193, 193, 245, 255, 254, 255, 255, 254, 254, 255, 255, 89, 89, 89, 209, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 102, 61, 187, 75, 102, 61, 187, 75, 102, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 252, 61, 187, 75, 42, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 21, 61, 61, 61, 208, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 136, 61, 187, 75, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 255, 240, 7, 0, 255, 192, 3, 0, 255, 128, 3, 0, 254, 0, 3, 0, 252, 0, 3, 0, 248, 0, 3, 0, 240, 0, 7, 0, 224, 12, 15, 0, 224, 30, 31, 0, 192, 127, 255, 0, 128, 111, 255, 0, 128, 239, 255, 0, 1, 231, 207, 0, 1, 231, 199, 0, 0, 227, 195, 0, 0, 96, 1, 0, 0, 96, 0, 0, 0, 96, 0, 0, 0, 112, 0, 0, 0, 248, 0, 0, 1, 254, 1, 0, 131, 255, 195, 0, 255, 255, 199, 0, 255, 255, 207, 0, 40, 0, 0, 0, 16, 0, 0, 0, 32, 0, 0, 0, 1, 0, 32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 96, 61, 61, 61, 209, 59, 59, 59, 255, 59, 59, 59, 255, 63, 63, 63, 244, 60, 60, 60, 192, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 61, 61, 61, 225, 87, 87, 87, 223, 190, 190, 190, 243, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 123, 123, 123, 236, 66, 66, 66, 215, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 65, 65, 65, 230, 190, 190, 190, 239, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 59, 59, 59, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 66, 66, 66, 231, 242, 242, 242, 244, 255, 255, 255, 255, 83, 83, 83, 236, 59, 59, 59, 255, 243, 243, 243, 255, 255, 255, 255, 255, 255, 255, 255, 255, 181, 181, 181, 255, 78, 78, 78, 230, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 67, 67, 67, 232, 242, 242, 242, 247, 255, 255, 255, 255, 65, 65, 65, 247, 59, 59, 59, 48, 81, 81, 81, 161, 91, 91, 91, 248, 243, 243, 243, 255, 180, 180, 180, 249, 74, 74, 74, 225, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 63, 63, 63, 244, 243, 243, 243, 249, 255, 255, 255, 255, 66, 66, 66, 248, 59, 59, 59, 48, 0, 0, 0, 0, 0, 0, 0, 0, 66, 66, 66, 148, 59, 59, 59, 255, 64, 64, 64, 213, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 64, 64, 229, 193, 193, 193, 251, 255, 255, 255, 255, 67, 67, 67, 249, 67, 67, 67, 50, 0, 0, 0, 0, 61, 187, 75, 109, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 96, 89, 89, 89, 225, 255, 255, 255, 255, 88, 88, 88, 243, 77, 77, 77, 53, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 112, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 207, 185, 185, 185, 224, 255, 255, 255, 255, 59, 59, 59, 255, 59, 59, 59, 143, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 255, 61, 187, 75, 196, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 255, 61, 187, 75, 112, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 255, 255, 255, 255, 255, 255, 255, 255, 255, 243, 243, 243, 255, 80, 80, 80, 232, 59, 59, 59, 143, 0, 0, 0, 0, 61, 187, 75, 176, 61, 187, 75, 255, 61, 187, 75, 152, 61, 187, 75, 64, 0, 0, 0, 0, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 112, 0, 0, 0, 0, 59, 59, 59, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 243, 243, 243, 252, 59, 59, 59, 255, 0, 0, 0, 0, 61, 187, 75, 62, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 112, 62, 62, 62, 243, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 177, 177, 177, 242, 65, 65, 65, 214, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 148, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 63, 63, 63, 195, 120, 120, 120, 232, 255, 255, 255, 255, 179, 179, 179, 248, 68, 68, 68, 217, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 72, 61, 187, 75, 178, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 112, 59, 59, 59, 16, 60, 60, 60, 208, 59, 59, 59, 255, 70, 70, 70, 219, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 64, 61, 187, 75, 255, 61, 187, 75, 255, 61, 187, 75, 112, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 255, 61, 187, 75, 112, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 61, 187, 75, 112, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 254, 3, 0, 0, 248, 3, 0, 0, 240, 3, 0, 0, 224, 3, 0, 0, 192, 3, 0, 0, 129, 135, 0, 0, 130, 255, 0, 0, 6, 247, 0, 0, 6, 115, 0, 0, 2, 17, 70, 188, 2, 0, 164, 252, 3, 0, 250, 255, 3, 128, 250, 255, 7, 225, 233, 255, 255, 243, 66, 239, 255, 247, 59, 41)
[byte[]]$ForwardOff = @(0, 0, 1, 0, 2, 0, 24, 24, 0, 0, 1, 0, 32, 0, 136, 9, 0, 0, 38, 0, 0, 0, 16, 16, 0, 0, 1, 0, 32, 0, 104, 4, 0, 0, 174, 9, 0, 0, 40, 0, 0, 0, 24, 0, 0, 0, 48, 0, 0, 0, 1, 0, 32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 59, 59, 59, 131, 61, 61, 61, 222, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 255, 65, 64, 64, 213, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 59, 59, 59, 177, 65, 65, 65, 239, 134, 134, 134, 222, 204, 203, 203, 255, 236, 237, 236, 255, 236, 236, 236, 255, 235, 235, 235, 255, 235, 235, 235, 255, 181, 181, 181, 255, 69, 69, 69, 233, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 67, 67, 67, 137, 68, 68, 68, 242, 153, 153, 153, 238, 239, 239, 239, 255, 238, 238, 238, 255, 238, 238, 238, 255, 236, 237, 237, 255, 237, 236, 237, 255, 236, 236, 236, 255, 236, 235, 235, 255, 235, 235, 235, 255, 180, 180, 180, 253, 71, 71, 71, 221, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 11, 65, 65, 65, 213, 119, 119, 119, 241, 225, 226, 225, 255, 240, 240, 240, 255, 239, 239, 239, 255, 238, 238, 238, 255, 238, 238, 238, 255, 237, 237, 238, 255, 237, 237, 237, 255, 236, 236, 236, 255, 236, 235, 236, 255, 236, 235, 236, 255, 235, 235, 235, 255, 59, 59, 59, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 31, 61, 61, 61, 247, 184, 185, 184, 247, 242, 242, 242, 255, 242, 241, 241, 255, 240, 240, 241, 255, 240, 239, 240, 255, 239, 239, 239, 255, 239, 238, 239, 255, 238, 238, 238, 255, 238, 237, 238, 255, 237, 236, 237, 255, 236, 236, 236, 255, 236, 236, 236, 255, 235, 235, 235, 255, 59, 59, 59, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 54, 73, 73, 73, 237, 182, 182, 182, 255, 243, 243, 243, 255, 243, 243, 242, 255, 241, 242, 242, 255, 226, 226, 226, 255, 92, 92, 92, 234, 131, 131, 131, 243, 239, 239, 239, 255, 239, 239, 239, 255, 238, 238, 238, 255, 237, 238, 238, 255, 237, 237, 237, 255, 237, 236, 236, 255, 96, 96, 96, 224, 62, 62, 62, 116, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 109, 109, 109, 41, 94, 94, 94, 236, 182, 182, 182, 255, 245, 245, 245, 255, 244, 244, 244, 255, 243, 243, 243, 255, 200, 200, 200, 255, 63, 63, 63, 250, 59, 59, 59, 96, 74, 74, 74, 160, 130, 130, 130, 240, 239, 239, 239, 255, 239, 239, 239, 255, 238, 238, 238, 255, 238, 237, 237, 255, 91, 91, 91, 216, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 5, 118, 118, 118, 225, 195, 195, 195, 255, 246, 245, 246, 255, 245, 245, 246, 255, 244, 244, 245, 255, 174, 173, 173, 253, 71, 71, 71, 234, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 66, 66, 66, 184, 128, 128, 128, 236, 224, 224, 224, 255, 223, 223, 223, 249, 84, 84, 84, 206, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 115, 115, 115, 191, 184, 184, 184, 255, 247, 248, 247, 255, 246, 247, 247, 255, 246, 246, 246, 255, 160, 160, 160, 250, 76, 77, 76, 212, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 66, 66, 66, 153, 62, 62, 62, 248, 61, 61, 61, 247, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 116, 116, 116, 138, 128, 128, 128, 254, 248, 249, 248, 255, 247, 248, 248, 255, 247, 247, 247, 255, 173, 173, 173, 244, 71, 71, 71, 206, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 38, 79, 78, 78, 233, 234, 234, 234, 255, 249, 250, 249, 255, 248, 249, 249, 255, 204, 204, 204, 253, 68, 68, 68, 230, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 70, 70, 70, 188, 164, 164, 164, 252, 250, 250, 250, 255, 250, 250, 250, 255, 233, 233, 234, 255, 66, 66, 66, 239, 59, 59, 59, 41, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 153, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 73, 73, 73, 249, 251, 252, 251, 255, 251, 251, 251, 255, 251, 251, 250, 255, 82, 82, 82, 219, 59, 59, 59, 96, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 252, 182, 182, 182, 36, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 136, 182, 182, 182, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 108, 108, 108, 175, 156, 156, 156, 255, 252, 252, 252, 255, 252, 252, 252, 255, 251, 251, 251, 255, 91, 91, 91, 175, 59, 59, 59, 147, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 249, 182, 182, 182, 171, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 255, 182, 182, 182, 249, 182, 182, 182, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 77, 77, 77, 242, 217, 217, 218, 255, 253, 253, 253, 255, 252, 252, 252, 255, 252, 252, 251, 255, 251, 251, 251, 237, 89, 89, 89, 173, 59, 59, 59, 177, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 228, 182, 182, 182, 255, 182, 182, 182, 105, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 249, 182, 182, 182, 42, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 255, 254, 253, 254, 255, 253, 253, 253, 255, 252, 252, 253, 255, 252, 252, 252, 255, 251, 252, 252, 255, 251, 251, 251, 249, 97, 96, 96, 182, 59, 59, 59, 147, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 186, 182, 182, 182, 255, 182, 182, 182, 252, 182, 182, 182, 141, 182, 182, 182, 51, 182, 182, 182, 51, 182, 182, 182, 51, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 249, 182, 182, 182, 42, 0, 0, 0, 0, 59, 59, 59, 255, 254, 254, 254, 255, 253, 253, 253, 255, 253, 253, 253, 255, 252, 253, 252, 255, 252, 253, 252, 255, 252, 251, 252, 255, 235, 235, 235, 255, 61, 61, 61, 246, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 108, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 249, 182, 182, 182, 42, 59, 59, 59, 255, 254, 254, 255, 255, 254, 254, 254, 255, 254, 253, 254, 255, 254, 253, 253, 255, 253, 252, 253, 255, 252, 252, 252, 255, 236, 236, 235, 255, 61, 61, 61, 246, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 15, 182, 182, 182, 225, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 249, 59, 59, 59, 255, 255, 255, 255, 255, 255, 254, 255, 255, 254, 254, 254, 255, 253, 254, 254, 255, 253, 254, 253, 255, 253, 253, 252, 255, 92, 92, 92, 213, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 57, 182, 182, 182, 237, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 81, 81, 81, 231, 195, 195, 195, 255, 255, 254, 255, 255, 254, 254, 254, 255, 254, 254, 254, 255, 254, 254, 253, 255, 89, 89, 89, 210, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 33, 182, 182, 182, 204, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 42, 59, 59, 59, 21, 74, 74, 74, 238, 193, 193, 193, 245, 254, 255, 255, 255, 255, 254, 254, 255, 89, 89, 89, 209, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 102, 182, 182, 182, 102, 182, 182, 182, 102, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 252, 182, 182, 182, 42, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 21, 61, 61, 61, 208, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 136, 182, 182, 182, 42, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 255, 240, 7, 0, 255, 192, 3, 0, 255, 128, 3, 0, 254, 0, 3, 0, 252, 0, 3, 0, 248, 0, 3, 0, 240, 0, 7, 0, 224, 12, 15, 0, 224, 30, 31, 0, 192, 127, 255, 0, 128, 111, 255, 0, 128, 239, 255, 0, 1, 231, 207, 0, 1, 231, 199, 0, 0, 227, 195, 0, 0, 96, 1, 0, 0, 96, 0, 0, 0, 96, 0, 0, 0, 112, 0, 0, 0, 248, 0, 0, 1, 254, 1, 0, 131, 255, 195, 0, 255, 255, 199, 0, 255, 255, 207, 0, 40, 0, 0, 0, 16, 0, 0, 0, 32, 0, 0, 0, 1, 0, 32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 96, 61, 61, 61, 209, 59, 59, 59, 255, 59, 59, 59, 255, 63, 63, 63, 244, 60, 60, 60, 192, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 61, 61, 61, 225, 87, 87, 87, 223, 190, 190, 190, 243, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 123, 123, 123, 236, 66, 66, 66, 215, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 65, 65, 65, 230, 190, 190, 190, 239, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 59, 59, 59, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 66, 66, 66, 231, 242, 242, 242, 244, 255, 255, 255, 255, 83, 83, 83, 236, 59, 59, 59, 255, 243, 243, 243, 255, 255, 255, 255, 255, 255, 255, 255, 255, 181, 181, 181, 255, 78, 78, 78, 230, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 67, 67, 67, 232, 242, 242, 242, 247, 255, 255, 255, 255, 65, 65, 65, 247, 59, 59, 59, 48, 81, 81, 81, 161, 91, 91, 91, 248, 243, 243, 243, 255, 180, 180, 180, 249, 74, 74, 74, 225, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 64, 63, 63, 63, 244, 243, 243, 243, 249, 255, 255, 255, 255, 66, 66, 66, 248, 59, 59, 59, 48, 0, 0, 0, 0, 0, 0, 0, 0, 66, 66, 66, 148, 59, 59, 59, 255, 64, 64, 64, 213, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 64, 64, 229, 193, 193, 193, 251, 255, 255, 255, 255, 67, 67, 67, 249, 67, 67, 67, 50, 0, 0, 0, 0, 182, 182, 182, 109, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 96, 89, 89, 89, 225, 255, 255, 255, 255, 88, 88, 88, 243, 77, 77, 77, 53, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 112, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 207, 185, 185, 185, 224, 255, 255, 255, 255, 59, 59, 59, 255, 59, 59, 59, 143, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 255, 182, 182, 182, 196, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 255, 182, 182, 182, 112, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 255, 255, 255, 255, 255, 255, 255, 255, 255, 243, 243, 243, 255, 80, 80, 80, 232, 59, 59, 59, 143, 0, 0, 0, 0, 182, 182, 182, 176, 182, 182, 182, 255, 182, 182, 182, 152, 182, 182, 182, 64, 0, 0, 0, 0, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 112, 0, 0, 0, 0, 59, 59, 59, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 243, 243, 243, 252, 59, 59, 59, 255, 0, 0, 0, 0, 182, 182, 182, 62, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 112, 62, 62, 62, 243, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 177, 177, 177, 242, 65, 65, 65, 214, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 148, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 63, 63, 63, 195, 120, 120, 120, 232, 255, 255, 255, 255, 179, 179, 179, 248, 68, 68, 68, 217, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 72, 182, 182, 182, 178, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 112, 59, 59, 59, 16, 60, 60, 60, 208, 59, 59, 59, 255, 70, 70, 70, 219, 59, 59, 59, 16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 64, 182, 182, 182, 255, 182, 182, 182, 255, 182, 182, 182, 112, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 255, 182, 182, 182, 112, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 182, 182, 182, 112, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 254, 3, 0, 0, 248, 3, 0, 0, 240, 3, 0, 0, 224, 3, 0, 0, 192, 3, 0, 0, 129, 135, 0, 0, 130, 255, 0, 0, 6, 247, 0, 0, 6, 115, 0, 0, 2, 17, 70, 188, 2, 0, 164, 252, 3, 0, 250, 255, 3, 128, 250, 255, 7, 225, 234, 255, 255, 243, 66, 239, 255, 247, 59, 41)
[byte[]]$SimRing = @(0, 0, 1, 0, 1, 0, 24, 24, 0, 0, 1, 0, 32, 0, 136, 9, 0, 0, 22, 0, 0, 0, 40, 0, 0, 0, 24, 0, 0, 0, 48, 0, 0, 0, 1, 0, 32, 0, 0, 0, 0, 0, 96, 9, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 59, 59, 59, 131, 61, 61, 61, 222, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 255, 64, 64, 64, 213, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 59, 59, 59, 177, 65, 65, 65, 239, 134, 134, 134, 222, 204, 204, 204, 255, 236, 236, 236, 255, 236, 236, 236, 255, 235, 235, 235, 255, 234, 235, 235, 255, 181, 180, 181, 255, 69, 69, 69, 233, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 67, 67, 67, 137, 67, 68, 67, 242, 153, 153, 153, 238, 238, 238, 238, 255, 238, 237, 238, 255, 237, 237, 238, 255, 237, 237, 237, 255, 236, 236, 236, 255, 236, 236, 235, 255, 235, 235, 235, 255, 235, 235, 235, 255, 180, 180, 180, 253, 71, 71, 71, 221, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 11, 65, 65, 65, 213, 119, 119, 119, 241, 226, 225, 225, 255, 240, 240, 239, 255, 239, 239, 239, 255, 239, 239, 239, 255, 238, 237, 238, 255, 237, 237, 237, 255, 236, 237, 236, 255, 236, 236, 236, 255, 236, 235, 236, 255, 235, 235, 235, 255, 234, 235, 235, 255, 59, 59, 59, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 31, 61, 61, 61, 247, 184, 184, 185, 247, 242, 242, 242, 255, 241, 241, 241, 255, 240, 240, 240, 255, 240, 240, 240, 255, 239, 239, 239, 255, 238, 238, 238, 255, 237, 237, 238, 255, 237, 238, 237, 255, 236, 237, 236, 255, 236, 236, 236, 255, 236, 236, 235, 255, 235, 235, 235, 255, 59, 59, 59, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 54, 73, 73, 73, 237, 182, 181, 182, 255, 243, 243, 243, 255, 242, 242, 242, 255, 242, 241, 242, 255, 226, 226, 226, 255, 92, 92, 92, 234, 131, 131, 131, 243, 239, 238, 239, 255, 238, 238, 238, 255, 238, 237, 238, 255, 237, 237, 237, 255, 237, 237, 237, 255, 236, 236, 236, 255, 96, 96, 96, 224, 62, 62, 62, 116, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 109, 109, 109, 41, 94, 94, 94, 236, 182, 182, 182, 255, 245, 245, 245, 255, 244, 244, 244, 255, 243, 244, 243, 255, 200, 200, 200, 255, 63, 63, 63, 250, 59, 59, 59, 96, 74, 74, 74, 160, 130, 130, 130, 240, 239, 239, 239, 255, 238, 238, 239, 255, 237, 237, 238, 255, 237, 237, 237, 255, 91, 91, 91, 216, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 5, 118, 118, 118, 225, 195, 195, 195, 255, 246, 246, 246, 255, 245, 245, 245, 255, 244, 244, 244, 255, 174, 174, 174, 253, 71, 71, 71, 234, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 66, 66, 66, 184, 128, 127, 128, 236, 224, 224, 224, 255, 223, 223, 223, 249, 84, 84, 84, 206, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 115, 115, 116, 191, 184, 184, 184, 255, 247, 247, 247, 255, 246, 247, 247, 255, 246, 246, 246, 255, 160, 160, 160, 250, 76, 77, 77, 212, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 66, 66, 66, 153, 62, 62, 62, 248, 61, 61, 61, 247, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 116, 116, 116, 138, 128, 128, 128, 254, 248, 249, 249, 255, 248, 248, 248, 255, 248, 247, 247, 255, 173, 173, 173, 244, 71, 71, 71, 206, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 38, 79, 79, 79, 233, 234, 234, 234, 255, 249, 250, 249, 255, 249, 249, 248, 255, 204, 204, 204, 253, 68, 68, 68, 230, 59, 59, 59, 21, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 70, 70, 70, 188, 164, 164, 164, 252, 251, 250, 251, 255, 251, 250, 250, 255, 233, 234, 234, 255, 66, 66, 66, 239, 59, 59, 59, 41, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 41, 73, 73, 73, 249, 252, 252, 252, 255, 251, 251, 252, 255, 251, 251, 251, 255, 82, 82, 82, 219, 59, 59, 59, 96, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 62, 187, 76, 255, 62, 187, 76, 255, 62, 187, 76, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 108, 108, 108, 175, 156, 156, 156, 255, 253, 253, 253, 255, 252, 252, 252, 255, 252, 251, 252, 255, 91, 90, 91, 175, 59, 59, 59, 147, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 63, 187, 77, 255, 62, 187, 76, 255, 62, 187, 76, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 77, 77, 77, 242, 217, 217, 218, 255, 253, 253, 253, 255, 253, 252, 252, 255, 252, 252, 252, 255, 251, 251, 252, 237, 89, 89, 89, 173, 59, 59, 59, 177, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 63, 187, 77, 255, 63, 187, 77, 255, 62, 187, 76, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 255, 254, 254, 254, 255, 253, 253, 253, 255, 253, 253, 253, 255, 252, 253, 253, 255, 252, 252, 252, 255, 252, 252, 251, 249, 97, 96, 96, 182, 59, 59, 59, 147, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 63, 187, 77, 255, 63, 187, 77, 255, 63, 187, 77, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 255, 254, 255, 254, 255, 254, 254, 254, 255, 254, 253, 253, 255, 253, 253, 253, 255, 252, 252, 252, 255, 252, 252, 252, 255, 235, 235, 235, 255, 60, 61, 61, 246, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 63, 188, 77, 255, 63, 188, 77, 255, 63, 187, 77, 255, 63, 187, 77, 255, 63, 187, 77, 255, 62, 187, 76, 255, 62, 187, 76, 255, 62, 187, 76, 255, 0, 0, 0, 0, 59, 59, 59, 255, 255, 255, 255, 255, 255, 254, 254, 255, 254, 254, 254, 255, 253, 253, 253, 255, 253, 253, 253, 255, 252, 252, 253, 255, 236, 236, 236, 255, 61, 61, 61, 246, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 63, 188, 77, 255, 63, 188, 77, 255, 63, 187, 77, 255, 63, 187, 77, 255, 63, 187, 77, 255, 62, 187, 76, 255, 62, 187, 76, 255, 0, 0, 0, 0, 59, 59, 59, 255, 255, 255, 255, 255, 254, 254, 255, 255, 255, 254, 254, 255, 254, 254, 254, 255, 253, 253, 253, 255, 253, 253, 253, 255, 92, 92, 92, 213, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 63, 188, 77, 255, 63, 188, 77, 255, 63, 187, 77, 255, 63, 187, 77, 255, 63, 187, 77, 255, 62, 187, 76, 255, 0, 0, 0, 0, 81, 81, 81, 231, 195, 195, 195, 255, 255, 255, 255, 255, 255, 255, 254, 255, 254, 254, 255, 255, 254, 254, 254, 255, 90, 89, 90, 210, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 188, 78, 255, 64, 188, 78, 255, 63, 188, 77, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 21, 74, 74, 74, 238, 193, 193, 193, 245, 255, 255, 255, 255, 255, 254, 255, 255, 89, 89, 89, 209, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 59, 59, 59, 21, 61, 61, 61, 208, 59, 59, 59, 255, 59, 59, 59, 255, 59, 59, 59, 114, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 64, 188, 78, 255, 64, 188, 78, 255, 64, 188, 78, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 255, 240, 7, 0, 255, 192, 3, 0, 255, 128, 3, 0, 254, 0, 3, 0, 252, 0, 3, 0, 248, 0, 3, 0, 240, 0, 7, 0, 224, 12, 15, 0, 224, 30, 31, 0, 192, 127, 255, 0, 128, 127, 255, 0, 128, 255, 255, 0, 1, 255, 31, 0, 1, 255, 31, 0, 0, 255, 31, 0, 0, 127, 31, 0, 0, 112, 1, 0, 0, 112, 1, 0, 0, 112, 1, 0, 0, 255, 31, 0, 1, 255, 31, 0, 131, 255, 31, 0, 255, 255, 31, 0, 255, 255, 255, 0)


$ForwardOffImage = new-object Windows.Forms.PictureBox
$ForwardOffImage.Location = New-Object System.Drawing.Size(5,10) 
$ForwardOffImage.width = 30
$ForwardOffImage.height = 30
$ForwardOffImage.BorderStyle = "FixedSingle"
$ForwardOffImage.sizemode = "Zoom"
$ForwardOffImage.Margin = 0
$ForwardOffImage.WaitOnLoad = $true
$ForwardOffImage.BorderStyle =  [System.Windows.Forms.BorderStyle]::None
$ForwardOffImage.Image = $ForwardOff
$ForwardOffImage.Add_Click(
{
	#$StatusLabel.Text = "Status: Settings..."
})


$Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Bold)
#Radio Buttons
$OffRadioButton = New-Object System.Windows.Forms.RadioButton
$OffRadioButton.Location = New-Object System.Drawing.Point(45,10)
$OffRadioButton.Name = "radiobutton1"
$OffRadioButton.Size = New-Object System.Drawing.Size(300,20)
$OffRadioButton.TabStop = $false
$OffRadioButton.Text = "Turn off call forwarding"
#$OffRadioButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$OffRadioButton.Checked = $true
$OffRadioButton.Font = $Font


$OffRadioButton.Add_CheckedChanged(
{
	Write-Verbose "Fowarding OffRadioButton Checked Changed"
	if($OffRadioButton.Checked)
	{
		# Do nothing for OffRadioButton.
		$SimRingComboBox.Enabled = $false
		$ForwardOnComboBox.Enabled = $false
	}
})

$OffRadioButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Getting settings..."
	[System.Windows.Forms.Application]::DoEvents()
	Write-Host "INFO: Off Button clicked - Radio Button - " $OffRadioButton.Checked -foreground "yellow"
	if($OffRadioButton.Checked)
	{
		$UnansweredCallsWillGoToLabel.Visible = $true
		$UnansweredCallsWillGoToLink.Visible = $true

		Write-Verbose "Off Radio Button Checked Changed"
		if($OffRadioButton.Checked)
		{
			$SipAddress = $UsersListbox.SelectedItem
			$UserDetails = Get-CsUser -identity "sip:${SipAddress}" | Select-Object DisplayName, LineURI
			$displayname = $UserDetails.DisplayName
			$lineuri = RemoveTel ([string]$UserDetails.LineURI) 
			#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri"
			$CallsWillRingYouAtLabel.Text = "Calls will ring the user directly."
			
			if($lineuri -eq "" -or $lineuri -eq $null)
			{
				$lineuri = "<not set>"
			}
			
			#UNANSWERED CALL DETAILS
			$UnansweredCallForwardDestination = (RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination)  #.Replace("sip:","").Replace("SIP:","")
			if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
			{$UnansweredWaitTime = 30}
			else
			{$UnansweredWaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime}
			
			$UserVoicemailSettings = CheckUserVoicemail $SipAddress
			if($UnansweredCallForwardDestination -ne $null -and $UnansweredCallForwardDestination -ne "")
			{
				$UnansweredCallsWillGoToLink.Text = "$UnansweredCallForwardDestination in $UnansweredWaitTime seconds"
				Write-Host "INFO: $UnansweredCallForwardDestination in $UnansweredWaitTime seconds" -foreground "Yellow"
			}
			elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail -or ($UserVoicemailSettings.ExUmEnabled -eq $true -or $UserVoicemailSettings.HostedVoicemail -eq $true))
			{
				$UnansweredCallsWillGoToLink.Text = "Voice mail in $UnansweredWaitTime seconds"
				Write-Host "INFO: Voice mail in $UnansweredWaitTime seconds" -foreground "Yellow"
			}
			elseif($UnansweredCallForwardDestination -eq "" -or $UnansweredCallForwardDestination -eq $null)
			{
				$UnansweredCallsWillGoToLink.Text = "No voice mail. Calls will continuously ring for $UnansweredWaitTime secs."
				Write-Host "INFO: No voice mail. Calls will continuously ring for $UnansweredWaitTime seconds." -foreground "Yellow"
			}
			
			
			
			#Update the current setting
			$Script:UpdatedUserForwardSettings.CallForwardingEnabled = $False
			$Script:UpdatedUserForwardSettings.ForwardDestination = $Null
			$Script:UpdatedUserForwardSettings.ForwardImmediateEnabled = $False
			$Script:UpdatedUserForwardSettings.SimultaneousRingEnabled = $False
			$Script:UpdatedUserForwardSettings.SimultaneousRingDestination = $Null
			$Script:UpdatedUserForwardSettings.ForwardToDelegates = $False
			$Script:UpdatedUserForwardSettings.TeamRingEnabled = $False
						
		}
		else
		{
			#Update the current setting
			#$Script:UpdatedUserForwardSettings.CallForwardingEnabled = $True
		}
	}
	PrintUpdatedUserSettings
	$StatusLabel.Text = ""
})

$ForwardOffLabel = New-Object System.Windows.Forms.Label
$ForwardOffLabel.Location = New-Object System.Drawing.Size(45,30) 
$ForwardOffLabel.Size = New-Object System.Drawing.Size(300,15) 
#$ForwardOffLabel.Text = "Calls will ring you at work and not be forwarded."
$ForwardOffLabel.Text = "Calls will ring the user directly and not be forwarded."
$ForwardOffLabel.TabStop = $False
#$ForwardOffLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left


$ForwardOnImage = new-object Windows.Forms.PictureBox
$ForwardOnImage.Location = New-Object System.Drawing.Size(5,60) 
$ForwardOnImage.width = 30
$ForwardOnImage.height = 30
$ForwardOnImage.BorderStyle = "FixedSingle"
$ForwardOnImage.sizemode = "Zoom"
$ForwardOnImage.Margin = 0
$ForwardOnImage.WaitOnLoad = $true
$ForwardOnImage.BorderStyle =  [System.Windows.Forms.BorderStyle]::None
$ForwardOnImage.Image = $ForwardOn
$ForwardOnImage.Add_Click(
{
	#$StatusLabel.Text = "Status: Settings..."
})

$ForwardRadioButton = New-Object System.Windows.Forms.RadioButton
$ForwardRadioButton.Location = New-Object System.Drawing.Point(45,60)
$ForwardRadioButton.Name = "radiobutton2"
$ForwardRadioButton.Size = New-Object System.Drawing.Size(146,20)
$ForwardRadioButton.TabStop = $false
$ForwardRadioButton.Text = "Forward my calls to:"
#$ForwardRadioButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$ForwardRadioButton.Font = $Font


$ForwardRadioButton.Add_CheckedChanged(
{
	$StatusLabel.Text = "STATUS: Getting settings..."
	[System.Windows.Forms.Application]::DoEvents()
	Write-Verbose "Off Checked Changed"
	if($ForwardRadioButton.Checked)
	{
		# Do nothing for OffRadioButton.
		$SimRingComboBox.Enabled = $false
		$ForwardOnComboBox.Enabled = $true
	}

	$StatusLabel.Text = ""
})

$ForwardRadioButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Getting settings..."
	[System.Windows.Forms.Application]::DoEvents()
	Write-Host "INFO: Forward Button clicked - Radio Button - " $ForwardRadioButton.Checked -foreground "yellow"
	if($ForwardRadioButton.Checked)
	{
		$UnansweredCallsWillGoToLabel.Visible = $false
		$UnansweredCallsWillGoToLink.Visible = $false

		Write-Verbose "Forward Radio Button Checked Changed"
		if($ForwardRadioButton.Checked)
		{
			if($ForwardOnComboBox.SelectedItem -eq "" -or $ForwardOnComboBox.SelectedItem -eq $null)
			{
				if($ForwardOnComboBox.FindStringExact("Select from this list") -eq -1)
				{
					Write-Verbose "Adding Select from this list Option"
					$ForwardOnComboBox.Items.Insert(0,"Select from this list")
					$ForwardOnComboBox.SelectedIndex = $ForwardOnComboBox.FindStringExact("Select from this list")
				}
			}
			elseif($ForwardOnComboBox.SelectedItem -ne "Select from this list") #If Select from this list is there then remove it.
			{
				if($ForwardOnComboBox.FindStringExact("Select from this list") -gt -1)
				{
					Write-Verbose "Removing Select from this list Option"
					$ForwardOnComboBox.Items.RemoveAt($ForwardOnComboBox.FindStringExact("Select from this list"))
				}
			}
			
			$ForwardOnComboBox.Enabled = $true
			$ForwardLocation = $ForwardOnComboBox.SelectedItem
			Write-Verbose "Currently selected forward location: $ForwardLocation"
			if($ForwardLocation -eq "Select from this list")
			{
				$CallsWillRingYouAtLabel.Text = "Select a forward location."
				$ForwardLocation = $null
			}
			elseif($ForwardLocation -eq "My Delegates")
			{
				$ForwardLocation = "Delegates"
				$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to Delegates."
			}
			else
			{
				$ForwardLocation = RemoveSip $ForwardLocation
				$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to ${ForwardLocation}."
			}
			
			
			#Update the current setting
			$Script:UpdatedUserForwardSettings.CallForwardingEnabled = $True #TRUE
			$Script:UpdatedUserForwardSettings.ForwardDestination = $ForwardLocation #VALUE
			$Script:UpdatedUserForwardSettings.ForwardImmediateEnabled = $True #TRUE
			$Script:UpdatedUserForwardSettings.SimultaneousRingEnabled = $False
			$Script:UpdatedUserForwardSettings.SimultaneousRingDestination = $Null
			if($ForwardLocation -eq "Delegates")
			{
				$Script:UpdatedUserForwardSettings.ForwardToDelegates = $True
			}
			else
			{
				$Script:UpdatedUserForwardSettings.ForwardToDelegates = $False
			}
			$Script:UpdatedUserForwardSettings.TeamRingEnabled = $False
		}
		else
		{
			$ForwardOnComboBox.Enabled = $false
		}
	}
	
	PrintUpdatedUserSettings
	$StatusLabel.Text = ""
})

$ForwardOnLabel = New-Object System.Windows.Forms.Label
$ForwardOnLabel.Location = New-Object System.Drawing.Size(45,80) 
$ForwardOnLabel.Size = New-Object System.Drawing.Size(380,15) 
$ForwardOnLabel.Text = "Calls will be forwarded immediately and not ring the user's work number."
$ForwardOnLabel.TabStop = $False
#$ForwardOffLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left


$ForwardOnComboBox = New-Object System.Windows.Forms.ComboBox 
$ForwardOnComboBox.Location = New-Object System.Drawing.Size(195,57) 
$ForwardOnComboBox.Size = New-Object System.Drawing.Size(200,20) 
$ForwardOnComboBox.DropDownHeight = 100 
$ForwardOnComboBox.tabIndex = 4
$ForwardOnComboBox.Sorted = $false
$ForwardOnComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$ForwardOnComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::PopUp
#$ForwardOnComboBox.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
#$ForwardOnComboBox.BorderColor = DarkGray
#$ForwardOnComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$ForwardOnComboBox.Enabled = $false

[void] $ForwardOnComboBox.Items.Add("Voice Mail")
[void] $ForwardOnComboBox.Items.Add("My Delegates")
[void] $ForwardOnComboBox.Items.Add("New Number or Contact")

$numberOfItems = $ForwardOnComboBox.Items.count
if($numberOfItems -gt 0)
{
	$ForwardOnComboBox.SelectedIndex = 0
}

$ForwardOnComboBox.add_SelectionChangeCommitted(
{
	$StatusLabel.Text = "STATUS: Getting settings..."
	[System.Windows.Forms.Application]::DoEvents()
	Write-Verbose "ForwardOnComboBox add_SelectedIndexChanged Changed"
    if($ForwardOnComboBox.SelectedItem -eq "New Number or Contact")
	{
		$StatusLabel.Text = "STATUS: Opening New Number Dialog..."
		[System.Windows.Forms.Application]::DoEvents()
		$DialogReturn = NewNumberDialog
		$Result = RemoveSip ($DialogReturn.Number) #$Result.Number
		if($Result -ne $null -and $Result -ne "")
		{
			if($ForwardOnComboBox.FindStringExact($Result) -eq -1)
			{
				[void] $ForwardOnComboBox.Items.Add($Result)
			}
			$ForwardOnComboBox.SelectedIndex = $ForwardOnComboBox.FindStringExact($Result)
		}

	}
	elseif($ForwardOnComboBox.SelectedItem -eq "My Delegates")
	{
		Write-Verbose "Delegates $(($Script:UpdatedUserForwardSettings.Delegates).count)" 
		if(([array]$Script:UpdatedUserForwardSettings.Delegates).count -eq 0)
		{
			DelegateDialog
			$DelegateCount = ($Script:UpdatedUserForwardSettings.Delegates).count
			$DelegateGroupLabel.Text = "Edit delegate members (${DelegateCount})"
			$DelegateGroupLabel.AutoSize = $true
		}
	}
	
	$ForwardOnComboBox.Enabled = $true
	$ForwardLocation = $ForwardOnComboBox.SelectedItem
	Write-Verbose "Currently selected forward location: $ForwardLocation"
	if($ForwardLocation -eq "Select from this list")
	{
		$finalNumber = $null
		$CallsWillRingYouAtLabel.Text = "Select a forward location."
	}
	elseif($ForwardLocation -eq "My Delegates")
	{
		$finalNumber = "Delegates"
		$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to Delegates."
	}
	elseif($ForwardLocation -eq "New Number or Contact" -or $ForwardLocation -eq "")
	{
		$finalNumber = $null
		$CallsWillRingYouAtLabel.Text = "Select a forward location."
	}
	elseif($DialogReturn.ResponseType -eq "Number")
	{
		$domain = (GetDomainFromSipAddress ($UsersListbox.SelectedItem))
		$finalNumber = (AddSipPhone ($ForwardLocation) $domain)
		$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to ${ForwardLocation}."
	}
	elseif($DialogReturn.ResponseType -eq "Contact")
	{
		$finalNumber = (AddSip ($ForwardLocation))
		$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to ${ForwardLocation}."
	}
	else
	{
		$finalNumber = $ForwardLocation
		$ForwardLocation = RemoveSip $ForwardLocation
		$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to ${ForwardLocation}."
	}
	
	
	#Update the current setting
	$Script:UpdatedUserForwardSettings.CallForwardingEnabled = $True #TRUE
	$Script:UpdatedUserForwardSettings.ForwardDestination = $finalNumber #VALUE
	$Script:UpdatedUserForwardSettings.ForwardImmediateEnabled = $True #TRUE
	$Script:UpdatedUserForwardSettings.SimultaneousRingEnabled = $False
	$Script:UpdatedUserForwardSettings.SimultaneousRingDestination = $Null
	if($ForwardLocation -eq "Delegates")
	{
		$Script:UpdatedUserForwardSettings.ForwardToDelegates = $True
	}
	else
	{
		$Script:UpdatedUserForwardSettings.ForwardToDelegates = $False
	}
	$Script:UpdatedUserForwardSettings.TeamRingEnabled = $False
	
	PrintUpdatedUserSettings
	$StatusLabel.Text = ""
})


$SimRingImage = new-object Windows.Forms.PictureBox
$SimRingImage.Location = New-Object System.Drawing.Size(5,110) 
$SimRingImage.width = 30
$SimRingImage.height = 30
$SimRingImage.BorderStyle = "FixedSingle"
$SimRingImage.sizemode = "Zoom"
$SimRingImage.Margin = 0
$SimRingImage.WaitOnLoad = $true
$SimRingImage.BorderStyle =  [System.Windows.Forms.BorderStyle]::None
$SimRingImage.Image = $SimRing
$SimRingImage.Add_Click(
{
	#$StatusLabel.Text = "Status: Settings..."
})

$SimRingRadioButton = New-Object System.Windows.Forms.RadioButton
$SimRingRadioButton.Location = New-Object System.Drawing.Point(45,110)
$SimRingRadioButton.Name = "radiobutton3"
$SimRingRadioButton.Size = New-Object System.Drawing.Size(146,20)
$SimRingRadioButton.TabStop = $false
$SimRingRadioButton.Text = "Simultaneously ring:"
#$SimRingRadioButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$SimRingRadioButton.Font = $Font


$SimRingRadioButton.Add_CheckedChanged(
{
	$StatusLabel.Text = "STATUS: Getting settings..."
	[System.Windows.Forms.Application]::DoEvents()
	Write-Verbose "Add_CheckedChanged - Sim Ring Radio Button Checked Changed" 
	if($SimRingRadioButton.Checked)
	{
		# Do nothing for OffRadioButton.
		$SimRingComboBox.Enabled = $true
		$ForwardOnComboBox.Enabled = $false
	}
	$StatusLabel.Text = ""
})

$SimRingRadioButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Getting settings..."
	[System.Windows.Forms.Application]::DoEvents()
	Write-Host "INFO: SimRing Button clicked - Radio Button - " $SimRingRadioButton.Checked -foreground "yellow"
	if($SimRingRadioButton.Checked)
	{
		$UnansweredCallsWillGoToLabel.Visible = $true
		$UnansweredCallsWillGoToLink.Visible = $true

		Write-Verbose "Add_Click - Sim Ring Radio Button Checked Changed"
		if($SimRingRadioButton.Checked)
		{
			
			if($SimRingComboBox.SelectedItem -eq "" -or $SimRingComboBox.SelectedItem -eq $null)
			{
				if($SimRingComboBox.FindStringExact("Select from this list") -eq -1)
				{
					Write-Verbose "Adding Select from this list Option"
					$SimRingComboBox.Items.Insert(0,"Select from this list")
					$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact("Select from this list")
				}
			}
			elseif($SimRingComboBox.SelectedItem -ne "Select from this list") #If Select from this list is there then remove it.
			{
				if($SimRingComboBox.FindStringExact("Select from this list") -gt -1)
				{
					Write-Verbose "Removing Select from this list Option"
					$SimRingComboBox.Items.RemoveAt($SimRingComboBox.FindStringExact("Select from this list"))
				}
			}			
			
			$SipAddress = $UsersListbox.SelectedItem
			$UserDetails = Get-CsUser -identity "sip:${SipAddress}" | Select-Object DisplayName, LineURI
			$displayname = $UserDetails.DisplayName
			$lineuri = RemoveTel ([string]$UserDetails.LineURI) #[regex]::match(([string]$UserDetails.LineURI),'(tel:)?([^;]*)(;.*)?$').Groups[2].Value #([string]$UserDetails.LineURI) -replace "tel:(.*);.*","$1"
			$SimRingLocation = RemoveTel ([string]$SimRingComboBox.SelectedItem) #[regex]::match(([string]$SimRingComboBox.Text),'(tel:)?([^;]*)(;.*)?$').Groups[2].Value #$SimRingComboBox.Text
			
			if($lineuri -eq "" -or $lineuri -eq $null)
			{
				$lineuri = "<not set>"
			}
			
			Write-Verbose "Currently selected SimRing location: $SimRingLocation"
			if($SimRingLocation -eq "Select from this list")
			{
				$CallsWillRingYouAtLabel.Text = "Select a Sim-Ring location."
				$SimRingLocation = $null
			}
			elseif($SimRingLocation -eq "My Delegates")
			{
				$SimRingLocation = "Delegates"
				#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring My Delegates."
				$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring their Delegates."
			}
			elseif($SimRingLocation -eq "My Team-Call Group")
			{
				$SimRingLocation = "Team"
				#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring My Team-Call Group."
				$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring their Team-Call Group."
			}
			elseif($SimRingLocation -eq "New Number" -or $SimRingLocation -eq "")
			{
				$SimRingLocation = $null
				#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring My Team-Call Group."
				$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring their Team-Call Group."
			}
			else
			{
				#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring ${SimRingLocation}."
				$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring ${SimRingLocation}."
			}
			
			$SimRingComboBox.Enabled = $true
			
			#UNANSWERED CALL DETAILS
			$UnansweredCallForwardDestination = (RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination)  #.Replace("sip:","").Replace("SIP:","")
			if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
			{$UnansweredWaitTime = 30}
			else
			{$UnansweredWaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime}
			
			$UserVoicemailSettings = CheckUserVoicemail $SipAddress
			if($UnansweredCallForwardDestination -ne $null -and $UnansweredCallForwardDestination -ne "")
			{
				$UnansweredCallsWillGoToLink.Text = "$UnansweredCallForwardDestination in $UnansweredWaitTime seconds"
				Write-Host "INFO: $UnansweredCallForwardDestination in $UnansweredWaitTime seconds" -foreground "Yellow"
			}
			elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail -or ($UserVoicemailSettings.ExUmEnabled -eq $true -or $UserVoicemailSettings.HostedVoicemail -eq $true))
			{
				$UnansweredCallsWillGoToLink.Text = "Voice mail in $UnansweredWaitTime seconds"
				Write-Host "INFO: Voice mail in $UnansweredWaitTime seconds" -foreground "Yellow"
			}
			elseif($UnansweredCallForwardDestination -eq "" -or $UnansweredCallForwardDestination -eq $null)
			{
				$UnansweredCallsWillGoToLink.Text = "No voice mail. Calls will continuously ring for $UnansweredWaitTime secs."
				Write-Host "INFO: No voice mail. Calls will continuously ring for $UnansweredWaitTime seconds." -foreground "Yellow"
			}
						
			
			#Update the current setting
			$Script:UpdatedUserForwardSettings.CallForwardingEnabled = $True #TRUE
			$Script:UpdatedUserForwardSettings.ForwardDestination = $Null #VALUE
			$Script:UpdatedUserForwardSettings.ForwardImmediateEnabled = $False #FALSE
			$Script:UpdatedUserForwardSettings.SimultaneousRingEnabled = $True  #TRUE
			$Script:UpdatedUserForwardSettings.SimultaneousRingDestination = $SimRingLocation
			$Script:UpdatedUserForwardSettings.ForwardToDelegates = $False
			if($SimRingLocation -eq "Team")
			{
				$Script:UpdatedUserForwardSettings.TeamRingEnabled = $True
			}
			else
			{
				$Script:UpdatedUserForwardSettings.TeamRingEnabled = $False
			}
		}
		else
		{
			$SimRingComboBox.Enabled = $false
		}
	}
	PrintUpdatedUserSettings
	$StatusLabel.Text = ""
})


$SimRingComboBox = New-Object System.Windows.Forms.ComboBox 
$SimRingComboBox.Location = New-Object System.Drawing.Size(195,109) 
$SimRingComboBox.Size = New-Object System.Drawing.Size(200,20) 
$SimRingComboBox.DropDownHeight = 100 
$SimRingComboBox.tabIndex = 5
$SimRingComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$SimRingComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
#$ForwardOnComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$SimRingComboBox.Enabled = $false

[void] $SimRingComboBox.Items.Add("My Delegates")
[void] $SimRingComboBox.Items.Add("My Team-Call Group")
[void] $SimRingComboBox.Items.Add("New Number")


$numberOfItems = $SimRingComboBox.Items.count
if($numberOfItems -gt 0)
{
	$SimRingComboBox.SelectedIndex = 0
}

#SelectedIndexChanged
$SimRingComboBox.add_SelectionChangeCommitted(
{
	$StatusLabel.Text = "STATUS: Updating settings..."
	[System.Windows.Forms.Application]::DoEvents()
	Write-Verbose "add_SelectedIndexChanged - SimRing Combobox Changed"
	
	$SimRingLocation = $SimRingComboBox.SelectedItem
	
    if($SimRingComboBox.SelectedItem -eq "New Number")
	{
		$StatusLabel.Text = "STATUS: Opening New Number Dialog..."
		[System.Windows.Forms.Application]::DoEvents()
		$Result = (NewSimRingNumberDialog).Number #$Result.Number
		#$SimRingComboBox.BeginInvoke([Action[string]] {param($i); $SimRingComboBox.Text = $i}, $Result)
		if($Result -ne $null -and $Result -ne "")
		{
			if($SimRingComboBox.FindStringExact($Result) -eq -1)
			{
				[void] $SimRingComboBox.Items.Add($Result)
			}
			$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact($Result)
		}
		
		$SipAddress = $UsersListbox.SelectedItem
		$UserDetails = Get-CsUser -identity "sip:${SipAddress}" | Select-Object DisplayName, LineURI
		$displayname = $UserDetails.DisplayName
		$lineuri = RemoveTel ([string]$UserDetails.LineURI) 
		$SimRingLocation = RemoveTel ([string]$SimRingComboBox.SelectedItem) 
		
		if($lineuri -eq "" -or $lineuri -eq $null)
		{
			$lineuri = "<not set>"
		}
		
		if($SimRingLocation -eq "New Number" -or $SimRingLocation -eq "")
		{
			$CallsWillRingYouAtLabel.Text = "Select a Sim-Ring location."
		}
		else
		{
			#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring ${SimRingLocation}."
			$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring ${SimRingLocation}."
		}
	}
	elseif($SimRingComboBox.SelectedItem -eq "My Delegates")
	{
		if(([array]$Script:UpdatedUserForwardSettings.Delegates).count -eq 0)
		{
			DelegateDialog
			$DelegateCount = ($Script:UpdatedUserForwardSettings.Delegates).count
			$DelegateGroupLabel.Text = "Edit delegate members (${DelegateCount})"
			$DelegateGroupLabel.AutoSize = $true
		}
		
		$SipAddress = $UsersListbox.SelectedItem
		$UserDetails = Get-CsUser -identity "sip:${SipAddress}" | Select-Object DisplayName, LineURI
		$displayname = $UserDetails.DisplayName
		$lineuri = RemoveTel ([string]$UserDetails.LineURI) #[regex]::match(([string]$UserDetails.LineURI),'(tel:)?([^;]*)(;.*)?$').Groups[2].Value #([string]$UserDetails.LineURI) -replace "tel:(.*);.*","$1"
		
		if($lineuri -eq "" -or $lineuri -eq $null)
		{
			$lineuri = "<not set>"
		}
		
		#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring My Delegates."
		$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring their Delegates."
	}
	elseif($SimRingComboBox.SelectedItem -eq "My Team-Call Group")
	{
		if(([array]$Script:UpdatedUserForwardSettings.Team).count -eq 0)
		{
			TeamCallDialog
			$TeamCount = ($Script:UpdatedUserForwardSettings.Team).count
			$TeamCallGroupLabel.Text = "Edit team-call group members (${TeamCount})"
			$TeamCallGroupLabel.AutoSize = $true
		}
		
		$SipAddress = $UsersListbox.SelectedItem
		$UserDetails = Get-CsUser -identity "sip:${SipAddress}" | Select-Object DisplayName, LineURI
		$displayname = $UserDetails.DisplayName
		$lineuri = RemoveTel ([string]$UserDetails.LineURI) #[regex]::match(([string]$UserDetails.LineURI),'(tel:)?([^;]*)(;.*)?$').Groups[2].Value #([string]$UserDetails.LineURI) -replace "tel:(.*);.*","$1"
		
		if($lineuri -eq "" -or $lineuri -eq $null)
		{
			$lineuri = "<not set>"
		}
		
		#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring Team-Call Group."
		$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring Team-Call Group."
	}
	else
	{
		$SipAddress = $UsersListbox.SelectedItem
		$UserDetails = Get-CsUser -identity "sip:${SipAddress}" | Select-Object DisplayName, LineURI
		$displayname = $UserDetails.DisplayName
		$lineuri = RemoveTel ([string]$UserDetails.LineURI) #[regex]::match(([string]$UserDetails.LineURI),'(tel:)?([^;]*)(;.*)?$').Groups[2].Value #([string]$UserDetails.LineURI) -replace "tel:(.*);.*","$1"
		$SimRingLocation = RemoveTel ([string]$SimRingComboBox.SelectedItem) #[regex]::match(([string]$SimRingComboBox.Text),'(tel:)?([^;]*)(;.*)?$').Groups[2].Value #$SimRingComboBox.Text
		
		if($lineuri -eq "" -or $lineuri -eq $null)
		{
			$lineuri = "<not set>"
		}
		
		#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring ${SimRingLocation}."
		$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring ${SimRingLocation}."
	}
	
	Write-Verbose "Currently selected SimRing location: $SimRingLocation"
	if($SimRingLocation -eq "Select from this list")
	{
		$CallsWillRingYouAtLabel.Text = "Select a Sim-Ring location."
		$SimRingLocation = $null
		$finalNumber = $null
	}
	elseif($SimRingLocation -eq "My Delegates")
	{
		$SimRingLocation = "Delegates"
		$finalNumber = "Delegates"
	}
	elseif($SimRingLocation -eq "My Team-Call Group")
	{
		$SimRingLocation = "Team"
		$finalNumber = "Team"
	}
	elseif($SimRingLocation -eq "New Number")
	{
		$CallsWillRingYouAtLabel.Text = "Select a Sim-Ring location."
		$SimRingLocation = $null
		$finalNumber = $null
	}
	else
	{
		#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring ${SimRingLocation}."
		$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring ${SimRingLocation}."
		$domain = (GetDomainFromSipAddress ($UsersListbox.SelectedItem))
		$finalNumber = (AddSipPhone ($SimRingLocation) $domain)
	}
	
	
	
	#Update the current setting
	$Script:UpdatedUserForwardSettings.CallForwardingEnabled = $True #TRUE
	$Script:UpdatedUserForwardSettings.ForwardDestination = $Null #VALUE
	$Script:UpdatedUserForwardSettings.ForwardImmediateEnabled = $False #FALSE
	$Script:UpdatedUserForwardSettings.SimultaneousRingEnabled = $True  #TRUE
	$Script:UpdatedUserForwardSettings.SimultaneousRingDestination = $finalNumber
	$Script:UpdatedUserForwardSettings.ForwardToDelegates = $False
	if($SimRingLocation -eq "Team")
	{
		$Script:UpdatedUserForwardSettings.TeamRingEnabled = $True
	}
	else
	{
		$Script:UpdatedUserForwardSettings.TeamRingEnabled = $False
	}
	
	PrintUpdatedUserSettings
	$StatusLabel.Text = ""
})


$SimRingLabel = New-Object System.Windows.Forms.Label
$SimRingLabel.Location = New-Object System.Drawing.Size(45,130) 
$SimRingLabel.Size = New-Object System.Drawing.Size(380,15) 
$SimRingLabel.Text = "Calls will ring the user at work and also ring another number or person."
$SimRingLabel.TabStop = $False
#$ForwardOffLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left


$groupPanel = New-Object System.Windows.Forms.Panel
$groupPanel.Controls.Add($OffRadioButton)
$groupPanel.Controls.Add($ForwardRadioButton)
$groupPanel.Controls.Add($ForwardOnComboBox)
$groupPanel.Controls.Add($SimRingRadioButton)
$groupPanel.Controls.Add($SimRingComboBox)
$groupPanel.Controls.Add($ForwardOffImage)
$groupPanel.Controls.Add($ForwardOnImage)
$groupPanel.Controls.Add($SimRingImage)
$groupPanel.Controls.Add($ForwardOffLabel)
$groupPanel.Controls.Add($ForwardOnLabel)
$groupPanel.Controls.Add($SimRingLabel)

$groupPanel.Location = New-Object System.Drawing.Size(275,50)
#$groupPanel.Name = "groupbox1"
$groupPanel.Size = New-Object System.Drawing.Size(400,170)
$groupPanel.TabStop = $False
$groupPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top 
#$groupPanel.Text = ""
$objForm.Controls.Add($groupPanel) 

<#
$groupbox1 = New-Object System.Windows.Forms.GroupBox
$groupbox1.Controls.Add($OffRadioButton)
$groupbox1.Controls.Add($ForwardRadioButton)
$groupbox1.Controls.Add($SimRingRadioButton)
$groupbox1.Location = New-Object System.Drawing.Size(190,50)
$groupbox1.Name = "groupbox1"
$groupbox1.Size = New-Object System.Drawing.Size(300,200)
$groupbox1.TabStop = $False
$groupbox1.Text = ""
$objForm.Controls.Add($groupbox1) 
#>

# Add a groupbox ============================================================
$GroupsBox = New-Object System.Windows.Forms.Groupbox
$GroupsBox.Location = New-Object System.Drawing.Size(263,24) 
$GroupsBox.Size = New-Object System.Drawing.Size(460,200) 
$GroupsBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top #-bor [System.Windows.Forms.AnchorStyles]::Left
$GroupsBox.TabStop = $False
$GroupsBox.Text = "Call Forwarding"
#$GroupsBox.BackColor = [System.Drawing.Color]::DarkGray
$GroupsBox.ForeColor = [System.Drawing.Color]::Black
#$GroupsBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$GroupsBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($GroupsBox)


$CurrentSettingsLabel = New-Object System.Windows.Forms.Label
$CurrentSettingsLabel.Location = New-Object System.Drawing.Size(263,240) 
$CurrentSettingsLabel.Size = New-Object System.Drawing.Size(250,15) 
$CurrentSettingsLabel.Text = "User's current call forwarding settings:"
$CurrentSettingsLabel.TabStop = $False
$objForm.Controls.Add($CurrentSettingsLabel)


$CallsWillRingYouAtLabel = New-Object System.Windows.Forms.Label
$CallsWillRingYouAtLabel.Location = New-Object System.Drawing.Size(8,10) 
$CallsWillRingYouAtLabel.Size = New-Object System.Drawing.Size(420,30) 
$CallsWillRingYouAtLabel.Text = "Calls will ring you at work"
$CallsWillRingYouAtLabel.TabStop = $False

$UnansweredCallsWillGoToLabel = New-Object System.Windows.Forms.Label
$UnansweredCallsWillGoToLabel.Location = New-Object System.Drawing.Size(8,38) 
$UnansweredCallsWillGoToLabel.Size = New-Object System.Drawing.Size(150,30) 
$UnansweredCallsWillGoToLabel.Text = "Unanswered calls will go to:"
$UnansweredCallsWillGoToLabel.TabStop = $False


$UnansweredCallsWillGoToLink = New-Object System.Windows.Forms.LinkLabel
$UnansweredCallsWillGoToLink.Location = New-Object System.Drawing.Size(160,40)
$UnansweredCallsWillGoToLink.Size = New-Object System.Drawing.Size(295,30)
#$UnansweredCallsWillGoToLink.AutoSize = $true
$UnansweredCallsWillGoToLink.DisabledLinkColor = $buttonBorderBlue
$UnansweredCallsWillGoToLink.VisitedLinkColor = $buttonBorderBlue
$UnansweredCallsWillGoToLink.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$UnansweredCallsWillGoToLink.LinkColor = $buttonBorderBlue
$UnansweredCallsWillGoToLink.TabStop = $False
$UnansweredCallsWillGoToLink.Text = "Voice mail in 30 seconds"
$UnansweredCallsWillGoToLink.add_click(
{
	$StatusLabel.Text = "STATUS: Opening Forward Unanswered Dialog..."
	[System.Windows.Forms.Application]::DoEvents()
	 ForwardUnansweredDialog
	 $StatusLabel.Text = ""
})


$TheseSettingsWillApplyLabel = New-Object System.Windows.Forms.Label
$TheseSettingsWillApplyLabel.Location = New-Object System.Drawing.Size(8,68) 
$TheseSettingsWillApplyLabel.Size = New-Object System.Drawing.Size(150,30) 
$TheseSettingsWillApplyLabel.Text = "These settings will apply:"
$TheseSettingsWillApplyLabel.TabStop = $False


$TheseSettingsWillApplyLink = New-Object System.Windows.Forms.LinkLabel
$TheseSettingsWillApplyLink.Location = New-Object System.Drawing.Size(160,70)
$TheseSettingsWillApplyLink.Size = New-Object System.Drawing.Size(210,15)
$TheseSettingsWillApplyLink.DisabledLinkColor = $buttonBorderBlue
$TheseSettingsWillApplyLink.VisitedLinkColor = $buttonBorderBlue
$TheseSettingsWillApplyLink.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$TheseSettingsWillApplyLink.LinkColor = $buttonBorderBlue
$TheseSettingsWillApplyLink.TabStop = $False
$TheseSettingsWillApplyLink.Text = "All of the time"
$TheseSettingsWillApplyLink.add_click(
{
	$StatusLabel.Text = "STATUS: Opening Time Applied Dialog..."
	[System.Windows.Forms.Application]::DoEvents()
	 TheseSettingsWillApplyLinkDialog
	 $StatusLabel.Text = ""
})



$GroupBoxCurrent = New-Object System.Windows.Forms.Panel
$GroupBoxCurrent.Location = New-Object System.Drawing.Size(265,260) 
$GroupBoxCurrent.Size = New-Object System.Drawing.Size(460,100) 
#$GroupBoxCurrent.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$GroupBoxCurrent.TabStop = $False
$GroupBoxCurrent.Controls.Add($CallsWillRingYouAtLabel)
$GroupBoxCurrent.Controls.Add($UnansweredCallsWillGoToLabel)
$GroupBoxCurrent.Controls.Add($TheseSettingsWillApplyLabel)
$GroupBoxCurrent.Controls.Add($UnansweredCallsWillGoToLink)
$GroupBoxCurrent.Controls.Add($TheseSettingsWillApplyLink)
#$GroupsBox.Text = "Call Forwarding"
$GroupBoxCurrent.BackColor = [System.Drawing.Color]::White
$GroupBoxCurrent.ForeColor = [System.Drawing.Color]::Black
$GroupBoxCurrent.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$objForm.Controls.Add($GroupBoxCurrent)




$TeamCallGroupLabel = New-Object System.Windows.Forms.LinkLabel
$TeamCallGroupLabel.Location = New-Object System.Drawing.Size(265,380)
#$TeamCallGroupLabel.Size = New-Object System.Drawing.Size(400,15)
$TeamCallGroupLabel.AutoSize = $true
$TeamCallGroupLabel.DisabledLinkColor = $buttonBorderBlue
$TeamCallGroupLabel.VisitedLinkColor = $buttonBorderBlue
$TeamCallGroupLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$TeamCallGroupLabel.LinkColor = $buttonBorderBlue
$TeamCallGroupLabel.TabStop = $False
$TeamCallGroupLabel.Text = "Edit team-call group members"
$TeamCallGroupLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$TeamCallGroupLabel.add_click(
{
	$StatusLabel.Text = "STATUS: Opening Team-Call Dialog..."
	[System.Windows.Forms.Application]::DoEvents()
	TeamCallDialog
	 
	$TeamCount = ($Script:UpdatedUserForwardSettings.Team).count
	$TeamCallGroupLabel.Text = "Edit team-call group members (${TeamCount})"
	$TeamCallGroupLabel.AutoSize = $true
	$StatusLabel.Text = ""
	
})
$objForm.Controls.Add($TeamCallGroupLabel)


$DelegateGroupLabel = New-Object System.Windows.Forms.LinkLabel
$DelegateGroupLabel.Location = New-Object System.Drawing.Size(265,410)
#$DelegateGroupLabel.Size = New-Object System.Drawing.Size(400,15)
$DelegateGroupLabel.AutoSize = $true
$DelegateGroupLabel.DisabledLinkColor = $buttonBorderBlue
$DelegateGroupLabel.VisitedLinkColor = $buttonBorderBlue
$DelegateGroupLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$DelegateGroupLabel.LinkColor = $buttonBorderBlue
$DelegateGroupLabel.TabStop = $False
$DelegateGroupLabel.Text = "Edit delegate members"
$DelegateGroupLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$DelegateGroupLabel.add_click(
{
	$StatusLabel.Text = "STATUS: Opening Team-Call Dialog..."
	[System.Windows.Forms.Application]::DoEvents()
	DelegateDialog
	 
	$DelegateCount = ($Script:UpdatedUserForwardSettings.Delegates).count
	$DelegateGroupLabel.Text = "Edit delegate members (${DelegateCount})"
	$StatusLabel.Text = ""
})
$objForm.Controls.Add($DelegateGroupLabel)

<#
$DirectionsLabel = New-Object System.Windows.Forms.Label
$DirectionsLabel.Location = New-Object System.Drawing.Size(265,440) 
$DirectionsLabel.Size = New-Object System.Drawing.Size(450,30) 
$DirectionsLabel.Text = "Configure the desired state and then click the `"Set User(s) Forward Settings`" button."
$DirectionsLabel.TabStop = $False
$DirectionsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($DirectionsLabel)
#>

<#
# Add the Status Label ============================================================
$SettingsUpdatedLabel = New-Object System.Windows.Forms.Label
$SettingsUpdatedLabel.Location = New-Object System.Drawing.Size(265,450) 
$SettingsUpdatedLabel.Size = New-Object System.Drawing.Size(480,40) 
$SettingsUpdatedLabel.Text = "After making changes to settings. Use the `"Set User(s) Forward Settings`" button to update apply to user(s)."
$SettingsUpdatedLabel.forecolor = [System.Drawing.ColorTranslator]::FromHtml("#696969")
$SettingsUpdatedLabel.TabStop = $false
$SettingsUpdatedLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($SettingsUpdatedLabel)
#>

$ArrowImage2 = new-object Windows.Forms.PictureBox
$ArrowImage2.Location = New-Object System.Drawing.Size(730,120) 
$ArrowImage2.width = 30
$ArrowImage2.height = 30
$ArrowImage2.BorderStyle = "FixedSingle"
$ArrowImage2.sizemode = "Zoom"
$ArrowImage2.Margin = 0
$ArrowImage2.WaitOnLoad = $true
$ArrowImage2.BorderStyle =  [System.Windows.Forms.BorderStyle]::None
$ArrowImage2.Image = $ArrowImage
$ArrowImage2.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top 
$ArrowImage2.Add_Click(
{
	#$StatusLabel.Text = "Status: Settings..."
})
$objForm.Controls.Add($ArrowImage2)



# Add the listbox containing all Users ============================================================
$SetUsersListbox = New-Object System.Windows.Forms.Listbox 
$SetUsersListbox.Location = New-Object System.Drawing.Size(765,60) 
$SetUsersListbox.Size = New-Object System.Drawing.Size(200,435) 
$SetUsersListbox.Sorted = $true
$SetUsersListbox.tabIndex = 10
$SetUsersListbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$SetUsersListbox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$SetUsersListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$SetUsersListbox.TabStop = $false

# Add Users ============================================================
$script:users | ForEach-Object {[void] $SetUsersListbox.Items.Add((RemoveSip $_.SipAddress))}


$objForm.Controls.Add($SetUsersListbox) 

# SetUsersListbox Click Event ============================================================
$SetUsersListbox.add_Click(
{
	#Do Nothing
})

# SetUsersListbox Key Event ============================================================
$SetUsersListbox.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		#Do Nothing
	}
})

$SetUsersLabel = New-Object System.Windows.Forms.Label
$SetUsersLabel.Location = New-Object System.Drawing.Size(765,13) 
$SetUsersLabel.Size = New-Object System.Drawing.Size(200,15) 
$SetUsersLabel.Text = "Set user(s) Forward Settings:"
$SetUsersLabel.TabStop = $False
$SetUsersLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$objForm.Controls.Add($SetUsersLabel)

#Set button
$SetUserForwardButton = New-Object System.Windows.Forms.Button
$SetUserForwardButton.Location = New-Object System.Drawing.Size(765,30)
$SetUserForwardButton.Size = New-Object System.Drawing.Size(200,23)
$SetUserForwardButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$SetUserForwardButton.Text = "  Set User(s) Forward Settings"
$SetUserForwardButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$SetUserForwardButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
$SetUserForwardButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Setting user's forward settings..."
	ToolIsBusy
	[System.Windows.Forms.Application]::DoEvents()
	Set-ForwardSettings
	ToolIsIdle
	$StatusLabel.Text = ""
})
$objForm.Controls.Add($SetUserForwardButton)
$SetUserForwardButton.Enabled = $true

$SetUserForwardButton.Add_MouseHover(
{
   $SetUserForwardButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
})
$SetUserForwardButton.Add_MouseLeave(
{
   $SetUserForwardButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
})


# Add the Status Label ============================================================
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Size(20,503) 
$StatusLabel.Size = New-Object System.Drawing.Size(420,15) 
$StatusLabel.Text = ""
$StatusLabel.forecolor = "DarkBlue"
$StatusLabel.TabStop = $false
$StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($StatusLabel)


$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(600,3)
#$MyLinkLabel.Size = New-Object System.Drawing.Size(170,15)
$MyLinkLabel.AutoSize = $true
$MyLinkLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$MyLinkLabel.DisabledLinkColor = $buttonBorderBlue
$MyLinkLabel.VisitedLinkColor = $buttonBorderBlue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = $buttonBorderBlue
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Text = "  www.myskypelab.com"
$MyLinkLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("https://www.myskypelab.com")
})
$objForm.Controls.Add($MyLinkLabel)


$ToolTip = New-Object System.Windows.Forms.ToolTip 
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
$ToolTip.IsBalloon = $true 
$ToolTip.InitialDelay = 2000 
$ToolTip.ReshowDelay = 1000 
$ToolTip.AutoPopDelay = 10000
#$ToolTip.ToolTipTitle = "Help:"
$ToolTip.SetToolTip($SetUserForwardButton, "After making changes to settings. Use the `"Set User(s) Forward Settings`"`r`nbutton to update apply to user(s).") 
$ToolTip.SetToolTip($GroupsBox, "Configure the desired state and then click the `"Set User(s) Forward Settings`" button.") 


function NewNumberDialog()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    
	$NewNumberLabel = New-Object System.Windows.Forms.Label
	$NewNumberLabel.Location = New-Object System.Drawing.Size(10,10) 
	$NewNumberLabel.Size = New-Object System.Drawing.Size(380,20)
	$NewNumberLabel.Text = "Enter the number you would like to forward to below (E.164 format):"
	$NewNumberLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
		
	$PhoneNumberLabel = New-Object System.Windows.Forms.Label
	$PhoneNumberLabel.Location = New-Object System.Drawing.Size(20,50) 
	$PhoneNumberLabel.Size = New-Object System.Drawing.Size(90,20)
	$PhoneNumberLabel.Text = "Phone number:"
	$PhoneNumberLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$PhoneNumberLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	
	$PhoneNumberTextBox = new-object System.Windows.Forms.textbox
	$PhoneNumberTextBox.location = new-object system.drawing.size(115,50)
	$PhoneNumberTextBox.size= new-object system.drawing.size(200,15)
	$PhoneNumberTextBox.text = ""
	$PhoneNumberTextBox.TabIndex = 1
	$PhoneNumberTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	
	$ContactLabel = New-Object System.Windows.Forms.Label
	$ContactLabel.Location = New-Object System.Drawing.Size(20,80) 
	$ContactLabel.Size = New-Object System.Drawing.Size(90,20)
	$ContactLabel.Text = "Contact:"
	$ContactLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$ContactLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	
	$ContactComboBox = New-Object System.Windows.Forms.ComboBox 
	$ContactComboBox.Location = New-Object System.Drawing.Size(115,80) 
	$ContactComboBox.Size = New-Object System.Drawing.Size(200,20) 
	$ContactComboBox.DropDownHeight = 100 
	$ContactComboBox.DropDownWidth = 230
	$ContactComboBox.tabIndex = 2
	$ContactComboBox.Sorted = $true
	$ContactComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$ContactComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
	#$ForwardOnComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$ContactComboBox.Enabled = $false

	#$ContactComboBox.Items.Add("Voice Mail")
	$Users = Get-CsUser | Select-Object SipAddress
	foreach($User in $Users.SipAddress)
	{
		$User = (RemoveSip $User)    #.Replace("sip:","")
		[void] $ContactComboBox.Items.Add("$User")
	}
	
	# Add NewCheckBox ============================================================
	$NewCheckBox = New-Object System.Windows.Forms.Checkbox 
	$NewCheckBox.Location = New-Object System.Drawing.Size(320,81) 
	$NewCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$NewCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$NewCheckBox.tabIndex = 3
	$NewCheckBox.Add_Click(
	{
		if($NewCheckBox.Checked -eq $true)
		{
			$ContactComboBox.Enabled = $true
			$PhoneNumberTextBox.Enabled = $false
			$NewNumberLabel.Text = "Select the contact you would like to forward to below:"
			if($ContactComboBox.Items.Count -ge 0)
			{
				$ContactComboBox.SelectedIndex = 0
			}
		}
		else
		{
			$ContactComboBox.Enabled = $false
			$PhoneNumberTextBox.Enabled = $true
			$NewNumberLabel.Text = "Enter the number you would like to forward to below (E.164 format):"
			$ContactComboBox.SelectedIndex = -1
		}
	})

	#NewLabel Label ============================================================
	$NewLabel = New-Object System.Windows.Forms.Label
	$NewLabel.Location = New-Object System.Drawing.Size(337,83) 
	$NewLabel.Size = New-Object System.Drawing.Size(40,15) 
	$NewLabel.Text = "Select"
	#$NewLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$NewLabel.TabStop = $false
	$NewLabel.Add_Click(
	{
		if($NewCheckBox.Checked -eq $true)
		{
			$ContactComboBox.Enabled = $false
			$PhoneNumberTextBox.Enabled = $true
			$NewCheckBox.Checked = $false
			$ContactComboBox.SelectedIndex = -1
			$NewNumberLabel.Text = "Enter the number you would like to forward to below (E.164 format):"
		}
		else
		{
			$ContactComboBox.Enabled = $true
			$PhoneNumberTextBox.Enabled = $false
			$NewCheckBox.Checked = $true
			if($ContactComboBox.Items.Count -ge 0)
			{
				$ContactComboBox.SelectedIndex = 0
			}
			$NewNumberLabel.Text = "Select the contact you would like to forward to below:"			
		}
	})
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(200,140)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
	$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$okButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $okButton.Add_Click({ 
	
		Write-Host "INFO: New Number Updated." -foreground "Yellow"
		
		$returnObj = New-Object PSObject
		
		
		
		if($NewCheckBox.Checked)
		{
			Add-Member -InputObject $returnObj -MemberType NoteProperty -Name ResponseType -Value ("Contact")
			Add-Member -InputObject $returnObj -MemberType NoteProperty -Name Number -Value ("sip:" + $ContactComboBox.SelectedItem)
			$form.Tag = $returnObj
			$form.Close()
		}
		else
		{
			$theNumber = (RemoveTel (RemoveSip $PhoneNumberTextBox.text))
			if($theNumber -match '(\+)?[0123456789]+$')
			{
				Add-Member -InputObject $returnObj -MemberType NoteProperty -Name ResponseType -Value ("Number")
				Add-Member -InputObject $returnObj -MemberType NoteProperty -Name Number -Value $theNumber 
				$form.Tag = $returnObj
				$form.Close()
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("The number format is incorrect. This field should contain only +,0,1,2,3,4,5,6,7,8,9,0 characters. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			}
		}
	})
	
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(290,140)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
	$CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$CancelButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $CancelButton.Add_Click({ 
	
		Write-Host "INFO: Cancelled dialog." -foreground "Yellow"
		$form.Tag = $null
		$form.Close() 
		
	})

	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "New Number or Contact"
    $form.Size = New-Object System.Drawing.Size(400,220)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
     
	$form.Tag = $null
	
	$form.Controls.Add($NewNumberLabel)
	$form.Controls.Add($PhoneNumberLabel)
	$form.Controls.Add($PhoneNumberTextBox)
	$form.Controls.Add($ContactLabel)
	$form.Controls.Add($ContactComboBox)
	$form.Controls.Add($NewCheckBox)
	$form.Controls.Add($NewLabel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
			
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
	# Return the text that the user entered.
	return $form.Tag
	
}



function NewSimRingNumberDialog()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    
	$NewNumberLabel = New-Object System.Windows.Forms.Label
	$NewNumberLabel.Location = New-Object System.Drawing.Size(10,10) 
	$NewNumberLabel.Size = New-Object System.Drawing.Size(380,20)
	$NewNumberLabel.Text = "Enter the number you would like to forward to below (E.164 format):"
	$NewNumberLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
		
	$PhoneNumberLabel = New-Object System.Windows.Forms.Label
	$PhoneNumberLabel.Location = New-Object System.Drawing.Size(20,50) 
	$PhoneNumberLabel.Size = New-Object System.Drawing.Size(90,20)
	$PhoneNumberLabel.Text = "Phone number:"
	$PhoneNumberLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$PhoneNumberLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	
	$PhoneNumberTextBox = new-object System.Windows.Forms.textbox
	$PhoneNumberTextBox.location = new-object system.drawing.size(115,50)
	$PhoneNumberTextBox.size= new-object system.drawing.size(200,15)
	$PhoneNumberTextBox.text = ""
	$PhoneNumberTextBox.TabIndex = 1
	$PhoneNumberTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(200,110)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
	$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$okButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $okButton.Add_Click({ 
	
		Write-Host "INFO: New Number Updated." -foreground "Yellow"
		
		$theNumber = $PhoneNumberTextBox.text
		if($theNumber -match '(\+)?[0123456789]+$')
		{
			$returnObj = New-Object PSObject
			Add-Member -InputObject $returnObj -MemberType NoteProperty -Name Number -Value $theNumber 
		
			$form.Tag = $returnObj
			$form.Close()
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show("The number format is incorrect. This field should contain only +,0,1,2,3,4,5,6,7,8,9,0 characters. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		}
			
	})
	
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(290,110)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
	$CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$CancelButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $CancelButton.Add_Click({ 
	
		Write-Host "INFO: Cancelled dialog." -foreground "Yellow"
		$form.Tag = $null
		$form.Close() 
		
	})

	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "New Number"
    $form.Size = New-Object System.Drawing.Size(400,190)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
     
	$form.Tag = $null
	
	$form.Controls.Add($NewNumberLabel)
	$form.Controls.Add($PhoneNumberLabel)
	$form.Controls.Add($PhoneNumberTextBox)
	$form.Controls.Add($ContactLabel)
	$form.Controls.Add($ContactComboBox)
	$form.Controls.Add($NewCheckBox)
	$form.Controls.Add($NewLabel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	
		
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
	# Return the text that the user entered.
	return $form.Tag
	
}

function UserPickerDialog()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	function UserPickerFilterUsersList()
	{
		$UserPickerListbox.Items.Clear()
		[string] $theFilter = $UserPickerFilterTextBox.text
		# Add Lync Users ============================================================
		$Script:users | ForEach-Object {if($_.SipAddress -imatch $theFilter){[void] $UserPickerListbox.Items.Add((RemoveSip $_.SipAddress))}}  #.ToLower().Replace('sip:','').Replace("SIP:","")
	}
     
	# Add the listbox containing all Users ============================================================
	$UserPickerListbox = New-Object System.Windows.Forms.Listbox 
	$UserPickerListbox.Location = New-Object System.Drawing.Size(20,30) 
	$UserPickerListbox.Size = New-Object System.Drawing.Size(300,261) 
	$UserPickerListbox.Sorted = $true
	$UserPickerListbox.tabIndex = 10
	$UserPickerListbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$UserPickerListbox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
	$UserPickerListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$UserPickerListbox.TabStop = $false

	
	foreach($user in $script:users )
	{
		[void] $UserPickerListbox.Items.Add((RemoveSip $user.SipAddress))
	}
	
	$objForm.Controls.Add($UserPickerListbox) 

	# UserPickerListbox Click Event ============================================================
	$UserPickerListbox.add_Click(
	{
		#DO Nothing
	})

	# UserPickerListbox Key Event ============================================================
	$UserPickerListbox.add_KeyUp(
	{
		if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
		{	
			#DO Nothing
		}
	})

	$UsersLabel = New-Object System.Windows.Forms.Label
	$UsersLabel.Location = New-Object System.Drawing.Size(20,15) 
	$UsersLabel.Size = New-Object System.Drawing.Size(200,15) 
	$UsersLabel.Text = "Choose User(s):"
	$UsersLabel.TabStop = $False
	$UsersLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	$objForm.Controls.Add($UsersLabel)


	# Filter button ============================================================
	$UserPickerFilterButton = New-Object System.Windows.Forms.Button
	$UserPickerFilterButton.Location = New-Object System.Drawing.Size(273,288)
	$UserPickerFilterButton.Size = New-Object System.Drawing.Size(48,20)
	$UserPickerFilterButton.Text = "Filter"
	$UserPickerFilterButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$UserPickerFilterButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
	$UserPickerFilterButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
	$UserPickerFilterButton.Add_Click({UserPickerFilterUsersList})
	$objForm.Controls.Add($UserPickerFilterButton)
	
	$UserPickerFilterButton.Add_MouseHover(
	{
	   $UserPickerFilterButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$UserPickerFilterButton.Add_MouseLeave(
	{
	   $UserPickerFilterButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})

	#Filter Text box ============================================================
	$UserPickerFilterTextBox = new-object System.Windows.Forms.textbox
	$UserPickerFilterTextBox.location = new-object system.drawing.size(20,288)
	$UserPickerFilterTextBox.size= new-object system.drawing.size(250,15)
	$UserPickerFilterTextBox.text = ""
	$UserPickerFilterTextBox.TabIndex = 1
	$UserPickerFilterTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
	$UserPickerFilterTextBox.add_KeyUp(
	{
		if ($_.KeyCode -eq "Enter") 
		{	
			UserPickerFilterUsersList
		}
	})
	$objform.controls.add($UserPickerFilterTextBox)

	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(80,320)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
	$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$okButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $okButton.Add_Click({ 
	
		$form.Close() 
			
	})
	
	$okButton.Add_MouseHover(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$okButton.Add_MouseLeave(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(170,320)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
	$CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$CancelButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $CancelButton.Add_Click({ 
	
		$form.Close() 
		
	})
	
	$CancelButton.Add_MouseHover(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$CancelButton.Add_MouseLeave(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})

	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Choose User"
    $form.Size = New-Object System.Drawing.Size(350,400)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
     
	$form.Controls.Add($UserPickerListbox)
	$form.Controls.Add($UsersLabel)
	$form.Controls.Add($UserPickerFilterButton)
	$form.Controls.Add($UserPickerFilterTextBox)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
			
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
    
	$usersArray = New-Object System.Collections.ArrayList
	
	foreach($item in $UserPickerListbox.SelectedItems)
	{
		Write-Verbose "Adding user: $item"
		$usersArray.Add($item) > $null 
	}
		 
	# Return the text that the user entered.
	return $usersArray
	
}




function ForwardUnansweredDialog()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	$Script:ForwardUnansweredDialogSettings = $Script:UpdatedUserForwardSettings
	$Script:ForwardUnansweredDialogResult = $null
	Write-Verbose "ForwardUnansweredDialogSettings = $($Script:ForwardUnansweredDialogSettings)"
	
	$UnansweredCallLocationLabel = New-Object System.Windows.Forms.Label
	$UnansweredCallLocationLabel.Location = New-Object System.Drawing.Size(10,22) 
	$UnansweredCallLocationLabel.Size = New-Object System.Drawing.Size(240,20)
	$UnansweredCallLocationLabel.Text = "Send unanswered calls to the following:"
	$UnansweredCallLocationLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	$UnansweredCallLocationComboBox = New-Object System.Windows.Forms.ComboBox 
	$UnansweredCallLocationComboBox.Location = New-Object System.Drawing.Size(255,20) 
	$UnansweredCallLocationComboBox.Size = New-Object System.Drawing.Size(220,20) 
	$UnansweredCallLocationComboBox.DropDownHeight = 100
	#$UnansweredCallLocationComboBox.DropDownWidth = 250	
	$UnansweredCallLocationComboBox.tabIndex = 4
	$UnansweredCallLocationComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
	$UnansweredCallLocationComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	
	#Need to test is Voice mail Exists - Get-CsHostedVoicemailPolicy? 
	#if((!$Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination -match "sip:") -and $Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination -ne $null -and $Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination -ne "")
	$SipAddress = $UsersListbox.SelectedItem
	$UserVoicemailSettings = CheckUserVoicemail $SipAddress
	Write-Verbose "ExUmEnabled = $($UserVoicemailSettings.ExUmEnabled)  HostedVoicemail = $($UserVoicemailSettings.HostedVoicemail)"
	if($UserVoicemailSettings.ExUmEnabled -eq $true -or $UserVoicemailSettings.HostedVoicemail -eq $true)
	{
		Write-Verbose "Setting Combobox to Voice Mail"
		[void]$UnansweredCallLocationComboBox.Items.Add("Voice Mail")
	}
	else
	{
		Write-Verbose "Setting Combobox to None"
		[void]$UnansweredCallLocationComboBox.Items.Add("None")
	}
	
	#If there is a number set then add it to the combo
	if($Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination -ne "" -and $Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination -ne $null)
	{
		if($Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination -match "user=phone")
		{	
			Write-Host "CURRENT SETTING = NUMBER:" (RemoveSip ($Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination)) -foreground "Yellow"
			
			$returnObj = New-Object PSObject
			Add-Member -InputObject $returnObj -MemberType NoteProperty -Name ResponseType -Value ("Number")
			Add-Member -InputObject $returnObj -MemberType NoteProperty -Name Number -Value (RemoveSip ($Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination)) 
		
			$Script:ForwardUnansweredDialogResult = $returnObj
		}
		else
		{
			Write-Host "CURRENT SETTING = CONTACT:" (RemoveSip ($Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination)) -foreground "Yellow"
			
			$returnObj = New-Object PSObject
			Add-Member -InputObject $returnObj -MemberType NoteProperty -Name ResponseType -Value ("Contact")
			Add-Member -InputObject $returnObj -MemberType NoteProperty -Name Number -Value (RemoveSip ($Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination))
			
			$Script:ForwardUnansweredDialogResult = $returnObj
		}
		
		[void]$UnansweredCallLocationComboBox.Items.Add((RemoveSip ($Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination))) #.Replace("sip:","").Replace("SIP:","")
	}
	[void]$UnansweredCallLocationComboBox.Items.Add("New Number or Contact")
	
	
	#Add this after initial combo box selection is made
	$UnansweredCallLocationComboBox.add_SelectedIndexChanged(
	{
		Write-Verbose "UnansweredCallLocationComboBox add_SelectedIndexChanged Changed"
		if($UnansweredCallLocationComboBox.SelectedItem -eq "New Number or Contact")
		{
			$Script:ForwardUnansweredDialogResult = NewNumberDialog
			
			$Result = $Script:ForwardUnansweredDialogResult.Number #$Result.Number
			if($Result -ne $null -and $Result -ne "")
			{
				if($UnansweredCallLocationComboBox.FindStringExact($Result) -eq -1)
				{
					[void] $UnansweredCallLocationComboBox.Items.Add($Result)
				}
								
				$UnansweredCallLocationComboBox.SelectedIndex = $UnansweredCallLocationComboBox.FindStringExact($Result)
				#$UnansweredCallLocationComboBox.BeginInvoke([Action[string]] {param($i); $UnansweredCallLocationComboBox.Text = $i}, $Result)
			}
		}
	})
	
	$RingSecondsLabel = New-Object System.Windows.Forms.Label
	$RingSecondsLabel.Location = New-Object System.Drawing.Size(10,62) 
	$RingSecondsLabel.Size = New-Object System.Drawing.Size(240,20)
	$RingSecondsLabel.Text = "Ring for this many seconds before redirecting:"
	$RingSecondsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	$RingSecondsComboBox = New-Object System.Windows.Forms.ComboBox 
	$RingSecondsComboBox.Location = New-Object System.Drawing.Size(255,60) 
	$RingSecondsComboBox.Size = New-Object System.Drawing.Size(100,20) 
	$RingSecondsComboBox.DropDownHeight = 200 
	$RingSecondsComboBox.tabIndex = 4
	$RingSecondsComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
	$RingSecondsComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	
	[void]$RingSecondsComboBox.Items.Add("5")
	[void]$RingSecondsComboBox.Items.Add("10")
	[void]$RingSecondsComboBox.Items.Add("15")
	[void]$RingSecondsComboBox.Items.Add("20")
	[void]$RingSecondsComboBox.Items.Add("25")
	[void]$RingSecondsComboBox.Items.Add("30")
	[void]$RingSecondsComboBox.Items.Add("35")
	[void]$RingSecondsComboBox.Items.Add("40")
	[void]$RingSecondsComboBox.Items.Add("45")
	[void]$RingSecondsComboBox.Items.Add("50")
	[void]$RingSecondsComboBox.Items.Add("55")
	[void]$RingSecondsComboBox.Items.Add("60")
	

	$numberOfItems = $RingSecondsComboBox.Items.count
	if($numberOfItems -gt 0)
	{
		if($Script:ForwardUnansweredDialogSettings.UnansweredWaitTime -eq 0)
		{$RingSecondsComboBox.SelectedIndex = $RingSecondsComboBox.FindStringExact("30")}
		else
		{$RingSecondsComboBox.SelectedIndex = $RingSecondsComboBox.FindStringExact($Script:ForwardUnansweredDialogSettings.UnansweredWaitTime)}
		
	}


	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(300,100)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
	$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$okButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $okButton.Add_Click({ 
	
		Write-Host "INFO: Unanswered Calls Updated." -foreground "Yellow"
		 
		$location = $UnansweredCallLocationComboBox.text
		$time = $RingSecondsComboBox.text
		 
		#Update the current setting
		if($location -eq "Voicemail" -or $location -eq "Voice Mail")
		{
			$Script:UpdatedUserForwardSettings.UnansweredToVoicemail = $true
			$Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination = $null
			$UnansweredCallsWillGoToLink.Text = "Voice mail in $time seconds"
		}
		else
		{
			if($location -eq "None")
			{
				$Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination = $null
				$UnansweredCallsWillGoToLink.Text = "No voice mail. Calls will continuously ring for $time secs."
			}
			else
			{
				Write-Verbose "Dialog result: $($Script:ForwardUnansweredDialogResult.ResponseType)" 
				Write-Verbose "Dialog Number: $($Script:ForwardUnansweredDialogResult.Number)" 
				Write-Verbose "Dialog Contact: $($Script:ForwardUnansweredDialogResult.Contact)" 
				
				if($Script:ForwardUnansweredDialogResult.ResponseType -eq "Number")
				{
					$domain = (GetDomainFromSipAddress ($UsersListbox.SelectedItem))
					$finalNumber = (AddSipPhone ($Script:ForwardUnansweredDialogResult.Number) $domain)
				}
				elseif($Script:ForwardUnansweredDialogResult.ResponseType -eq "Contact")
				{
					$finalNumber = (AddSip ($Script:ForwardUnansweredDialogResult.Number))
				}

				$Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination = $finalNumber
				$locationString = RemoveSip $finalNumber
				$UnansweredCallsWillGoToLink.Text = "$locationString in $time seconds"
			}
			#FORMAT: sip:+61407532919@sfb2019lab.com;user=phone
			
		}
		$Script:UpdatedUserForwardSettings.UnansweredWaitTime = $time
		
		PrintUpdatedUserSettings
		
		$form.Tag = $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
		$form.Close() 
		
			
	})
	
	$okButton.Add_MouseHover(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$okButton.Add_MouseLeave(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(390,100)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
	$CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$CancelButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $CancelButton.Add_Click({ 
	
		$form.Tag = $null
		$form.Close() 
		
	})
	
	$CancelButton.Add_MouseHover(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$CancelButton.Add_MouseLeave(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})

	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Call Forwarding - Unanswered Calls"
    $form.Size = New-Object System.Drawing.Size(500,180)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	#Myskypelab Icon
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    #$form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.Tag = $null
     
	$form.Controls.Add($UnansweredCallLocationLabel)
	$form.Controls.Add($UnansweredCallLocationComboBox)
	$form.Controls.Add($RingSecondsLabel)
	$form.Controls.Add($RingSecondsComboBox)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	
	
    # Initialize and show the form.
    $form.Add_Shown({
	$form.Activate()
	
	$numberOfItems = $UnansweredCallLocationComboBox.Items.count
	if($numberOfItems -gt 0)
	{
		#if($Script:ForwardUnansweredDialogSettings.UnansweredToVoicemail)
		#{
			
		#}
		Write-Verbose ("UnansweredCallForwardDestination {0}" -f $Script:UpdatedUserForwardSettings.UnansweredToVoicemail)
		Write-Verbose ("UnansweredCallForwardDestination {0}" -f $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination)
		
		#$SipAddress = $UsersListbox.SelectedItem
		$UserVoicemailSettings = CheckUserVoicemail $SipAddress
		Write-Verbose "ExUmEnabled = $($UserVoicemailSettings.ExUmEnabled)  HostedVoicemail = $($UserVoicemailSettings.HostedVoicemail)"
		
		if($Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne $null -and $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne "")
		{
			Write-Verbose "Selecting number from combobox"
			$UnansweredCallLocationComboBox.SelectedIndex = $UnansweredCallLocationComboBox.FindStringExact((RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination)) #.Replace("sip:","").Replace("SIP:","")
		}
		elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail -eq $true)
		{
			Write-Verbose "Selecting Voice Mail from Combobox"
			$UnansweredCallLocationComboBox.SelectedIndex = $UnansweredCallLocationComboBox.FindStringExact("Voice Mail")
		}
		elseif($UserVoicemailSettings.ExUmEnabled -eq $true -or $UserVoicemailSettings.HostedVoicemail -eq $true)
		{
			Write-Verbose "Selecting Voice Mail from Combobox"
			$UnansweredCallLocationComboBox.SelectedIndex = $UnansweredCallLocationComboBox.FindStringExact("Voice Mail")
		}
		elseif($Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -eq $null -or $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -eq "")
		{
			Write-Verbose "Selecting None from Combobox"
			$UnansweredCallLocationComboBox.SelectedIndex = $UnansweredCallLocationComboBox.FindStringExact("None")
		}
		else
		{
			Write-Verbose "Setting combobox to item 0"
			$UnansweredCallLocationComboBox.SelectedIndex = 0
		}
	}
	
	})
	
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
}


function TheseSettingsWillApplyLinkDialog()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    
	
	$TimeLabel = New-Object System.Windows.Forms.Label
	$TimeLabel.Location = New-Object System.Drawing.Size(10,10) 
	$TimeLabel.Size = New-Object System.Drawing.Size(250,20)
	$TimeLabel.Text = "When do you want the settings to be applied?"
	$TimeLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	$AllTheTimeRadioButton = New-Object System.Windows.Forms.RadioButton
	$AllTheTimeRadioButton.Location = New-Object System.Drawing.Point(10,10)
	$AllTheTimeRadioButton.Name = "radiobutton1"
	$AllTheTimeRadioButton.Size = New-Object System.Drawing.Size(200,20)
	$AllTheTimeRadioButton.TabStop = $false
	$AllTheTimeRadioButton.Text = "All of the time"
	
	if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
	{
		$AllTheTimeRadioButton.Checked = $false
	}
	else
	{
		$AllTheTimeRadioButton.Checked = $true
	}
			
	$OutlookHoursRadioButton = New-Object System.Windows.Forms.RadioButton
	$OutlookHoursRadioButton.Location = New-Object System.Drawing.Point(10,40)
	$OutlookHoursRadioButton.Name = "radiobutton2"
	$OutlookHoursRadioButton.Size = New-Object System.Drawing.Size(200,20)
	$OutlookHoursRadioButton.TabStop = $false
	$OutlookHoursRadioButton.Text = "During work hours set in Outlook"

	if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
	{
		$OutlookHoursRadioButton.Checked = $true
	}
	else
	{
		$OutlookHoursRadioButton.Checked = $false
	}
	
	#if((!$Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -match "sip:") -and $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne $null -and $Script:ForwardUnansweredDialogSettings.UnansweredCallForwardDestination -ne "")
	#{
		#Do Nothing
	#}
	#else
	$SipAddress = $UsersListbox.SelectedItem
	$UserVoicemailSettings = CheckUserVoicemail $SipAddress
	if($UserVoicemailSettings.ExUmEnabled -eq $false -and $UserVoicemailSettings.HostedVoicemail -eq $false)
	{
		Write-Verbose "OutlookHoursRadioButton Disabled"
		$OutlookHoursRadioButton.Enabled = $false
	}

	$groupPanel = New-Object System.Windows.Forms.Panel
	$groupPanel.Controls.Add($AllTheTimeRadioButton)
	$groupPanel.Controls.Add($OutlookHoursRadioButton)
	$groupPanel.Location = New-Object System.Drawing.Size(15,30)
	#$groupPanel.Name = "groupbox1"
	$groupPanel.Size = New-Object System.Drawing.Size(300,70)
	$groupPanel.TabStop = $False
	#$groupPanel.Text = ""
	
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(300,100)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
	$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$okButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $okButton.Add_Click({ 
	
				
		if($AllTheTimeRadioButton.Checked)
		{
			$TheseSettingsWillApplyLink.Text = $AllTheTimeRadioButton.Text
			$Script:UpdatedUserForwardSettings.SettingsActiveWorkhours = $false
		}
		else
		{
			$TheseSettingsWillApplyLink.Text = $OutlookHoursRadioButton.Text
			$Script:UpdatedUserForwardSettings.SettingsActiveWorkhours = $true
		}
		 		
		Write-Host "INFO: Settings Updated." -foreground "Yellow"
		$form.Close() 
			
	})
	
	$okButton.Add_MouseHover(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$okButton.Add_MouseLeave(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(390,100)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
	$CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$CancelButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $CancelButton.Add_Click({ 
	
		$form.Close() 
		
	})
	
	$CancelButton.Add_MouseHover(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$CancelButton.Add_MouseLeave(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})

	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Call Forwarding - Time Setting"
    $form.Size = New-Object System.Drawing.Size(500,180)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	#Myskypelab Icon
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
     
	$form.Controls.Add($TimeLabel)
	$form.Controls.Add($groupPanel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	
	
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.

}


function TeamCallDialog()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
     
	[array]$Script:editDialogTeamMembers = $Script:UpdatedUserForwardSettings.Team
	
	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size(8,8) 
	$TitleLabel.Size = New-Object System.Drawing.Size(300,20)
	$TitleLabel.Text = "A team-call group can answer calls on behalf of the user."
	$TitleLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	$TitleLabel2 = New-Object System.Windows.Forms.Label
	$TitleLabel2.Location = New-Object System.Drawing.Size(8,28) 
	$TitleLabel2.Size = New-Object System.Drawing.Size(300,20)
	$TitleLabel2.Text = "User's calls will be forwarded to people in this list."
	$TitleLabel2.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
		
	#Data Grid View ============================================================
	$dgv = New-Object Windows.Forms.DataGridView
	$dgv.Size = New-Object System.Drawing.Size(560,180)
	$dgv.Location = New-Object System.Drawing.Size(8,50)
	$dgv.AutoGenerateColumns = $false
	$dgv.RowHeadersVisible = $false
	$dgv.AllowUserToAddRows = $false
	$dgv.BackgroundColor = [System.Drawing.Color]::White
	$dgv.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
	$dgv.AllowUserToResizeColumns = $false
	$dgv.AllowUserToResizeRows = $false
	$dgv.ColumnHeadersHeightSizeMode  = [Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
	$dgv.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

	$titleColumn0 = New-Object Windows.Forms.DataGridViewCheckBoxColumn
	$titleColumn0.HeaderText = "Receive Calls"
	$titleColumn0.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$titleColumn0.MinimumWidth = 80
	$titleColumn0.Width = 80
	$titleColumn0.Visible = $false
	$dgv.Columns.Add($titleColumn0) | Out-Null


	$titleColumn1 = New-Object Windows.Forms.DataGridViewTextBoxColumn
	$titleColumn1.HeaderText = "Team-Call Group"
	$titleColumn1.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$titleColumn1.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
	$titleColumn1.ReadOnly = $true
	$titleColumn1.MinimumWidth = 557
	$titleColumn1.Width = 557
	$dgv.Columns.Add($titleColumn1) | Out-Null
	
	
	foreach($item in $Script:editDialogTeamMembers)
	{
		$item = RemoveSip $item
		[bool] $delegate = $true
		$dgv.Rows.Add( @($delegate,$item) )
	}
	
	# Create the OK button.
    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Size(10,235)
    $AddButton.Size = New-Object System.Drawing.Size(100,20)
    $AddButton.Text = "Add..."
	$AddButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$AddButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $AddButton.Add_Click({ 
	
		$result = UserPickerDialog
		
		foreach($item in $result)
		{
			$dgv.Rows.Add( @($false,$item) )
			[array]$Script:editDialogTeamMembers += $item
		}
				
	})
	
	$AddButton.Add_MouseHover(
	{
	   $AddButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$AddButton.Add_MouseLeave(
	{
	   $AddButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	# Create the OK button.
    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Size(125,235)
    $RemoveButton.Size = New-Object System.Drawing.Size(100,20)
    $RemoveButton.Text = "Remove"
	$RemoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$RemoveButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $RemoveButton.Add_Click({ 
	
		
	$beforeDelete = $dgv.CurrentCell.RowIndex
	[array]$itemArray = @()
	foreach($item in $dgv.SelectedRows.Cells[1].Value)
	{
		$itemArray += $item
	}
	foreach($item in $itemArray)
	{
		$dgv.Rows.RemoveAt($dgv.CurrentCell.RowIndex)
		[array]$Script:editDialogTeamMembers = [array]$Script:editDialogTeamMembers | Where-Object { $_ -ne $item }
	}
	
	if($beforeDelete -gt ($dgv.Rows.Count - 1))
	{
		$beforeDelete = $beforeDelete - 1
	}
	if($dgv.Rows.Count -gt 0)
	{
		$dgv.Rows[$beforeDelete].Selected = $true
	}
	elseif($dgv.Rows.Count -eq 0)
	{
		#Do nothing
	}
					
	})
	
	$RemoveButton.Add_MouseHover(
	{
	   $RemoveButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$RemoveButton.Add_MouseLeave(
	{
	   $RemoveButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	
	$RingAfterLabel = New-Object System.Windows.Forms.Label
	$RingAfterLabel.Location = New-Object System.Drawing.Size(20,273) 
	$RingAfterLabel.Size = New-Object System.Drawing.Size(240,20)
	$RingAfterLabel.Text = "Ring team-call group after this many seconds:"
	$RingAfterLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	$RingAfterTimeComboBox = New-Object System.Windows.Forms.ComboBox 
	$RingAfterTimeComboBox.Location = New-Object System.Drawing.Size(260,270) 
	$RingAfterTimeComboBox.Size = New-Object System.Drawing.Size(160,18) 
	$RingAfterTimeComboBox.DropDownHeight = 80
	$RingAfterTimeComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
	$RingAfterTimeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	
	[void]$RingAfterTimeComboBox.Items.Add("0 - at the same time")
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 5 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{[void]$RingAfterTimeComboBox.Items.Add("5")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 10 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{[void]$RingAfterTimeComboBox.Items.Add("10")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 15 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0) #ONLY 0-10 IS SUPPORTED VIA POWERSHELL.... WHY? I DON'T KNOW.
	{[void]$RingAfterTimeComboBox.Items.Add("15")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 20 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{[void]$RingAfterTimeComboBox.Items.Add("20")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 25 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{[void]$RingAfterTimeComboBox.Items.Add("25")}	
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 30)
	{[void]$RingAfterTimeComboBox.Items.Add("30")}	
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 35)
	{[void]$RingAfterTimeComboBox.Items.Add("35")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 40)
	{[void]$RingAfterTimeComboBox.Items.Add("40")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 45)
	{[void]$RingAfterTimeComboBox.Items.Add("45")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 50)
	{[void]$RingAfterTimeComboBox.Items.Add("50")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 55)
	{[void]$RingAfterTimeComboBox.Items.Add("55")}
	

	$numberOfItems = $RingAfterTimeComboBox.Items.count
	if($numberOfItems -gt 0)
	{
		if($Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime -eq 0 -and $Script:UpdatedUserForwardSettings.DelegateRingWaitTime -eq 0)
		{
			$RingAfterTimeComboBox.SelectedIndex = $RingAfterTimeComboBox.FindStringExact("0 - at the same time")
		}
		elseif($Script:UpdatedUserForwardSettings.DelegateRingWaitTime -gt 0)
		{
			$RingAfterTimeComboBox.SelectedIndex = $RingAfterTimeComboBox.FindStringExact($Script:UpdatedUserForwardSettings.DelegateRingWaitTime)
		}
		elseif($Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime -gt 0)
		{
			$RingAfterTimeComboBox.SelectedIndex = $RingAfterTimeComboBox.FindStringExact($Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime)
		}
		else
		{
			$RingAfterTimeComboBox.SelectedIndex = 0
		}
	}
	
	$RingAfterTimeComboBox.Add_SelectionChangeCommitted(
	{
		if($RingAfterTimeComboBox.SelectedItem -eq "15" -or $RingAfterTimeComboBox.SelectedItem -eq "20" -or $RingAfterTimeComboBox.SelectedItem -eq "25" -or $RingAfterTimeComboBox.SelectedItem -eq "30" -or $RingAfterTimeComboBox.SelectedItem -eq "35" -or $RingAfterTimeComboBox.SelectedItem -eq "40" -or $RingAfterTimeComboBox.SelectedItem -eq "45" -or $RingAfterTimeComboBox.SelectedItem -eq "50" -or $RingAfterTimeComboBox.SelectedItem -eq "55" -or $RingAfterTimeComboBox.SelectedItem -eq "60")
		{
			[System.Windows.Forms.MessageBox]::Show("The PowerShell commands currently only support settings between 0-10 seconds. Changing the selection back to 10.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			$RingAfterTimeComboBox.SelectedIndex = $RingAfterTimeComboBox.FindStringExact("10")
		}
	})
	
		
	$ToolTip = New-Object System.Windows.Forms.ToolTip 
	$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
	$ToolTip.IsBalloon = $true 
	$ToolTip.InitialDelay = 1000 
	$ToolTip.ReshowDelay = 2000 
	$ToolTip.AutoPopDelay = 10000
	#$ToolTip.ToolTipTitle = "Help:"
	$ToolTip.SetToolTip($RingAfterTimeComboBox, "Note: The PowerShell commands only support 0-15 seconds right now.") 
	

	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(340,320)
    $okButton.Size = New-Object System.Drawing.Size(100,25)
    $okButton.Text = "OK"
	$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$okButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
	$okButton.tabIndex = 1
    $okButton.Add_Click({ 
	
		$Script:UpdatedUserForwardSettings.Team = $Script:editDialogTeamMembers
		if($RingAfterTimeComboBox.Text -eq "0 - at the same time")
		{
			$Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime = 0
			$Script:UpdatedUserForwardSettings.DelegateRingWaitTime = 0
		}
		else
		{
			$Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime = $RingAfterTimeComboBox.Text
			$Script:UpdatedUserForwardSettings.DelegateRingWaitTime = $RingAfterTimeComboBox.Text
		}
		Write-Host "INFO: Settings Updated." -foreground "Yellow"
		$form.Close() 
			
	})
	
	$okButton.Add_MouseHover(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$okButton.Add_MouseLeave(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(460,320)
    $CancelButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelButton.Text = "Cancel"
	$CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$CancelButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $CancelButton.Add_Click({ 
	
		$form.Close() 
		
	})
	
	$CancelButton.Add_MouseHover(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$CancelButton.Add_MouseLeave(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})

	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Call Forwarding - Team-Call Group"
    $form.Size = New-Object System.Drawing.Size(600,400)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
     
	$form.Controls.Add($TitleLabel)
	$form.Controls.Add($TitleLabel2)
	$form.Controls.Add($dgv)
	$form.Controls.Add($AddButton)
	$form.Controls.Add($RemoveButton)
	$form.Controls.Add($RingAfterLabel)
	$form.Controls.Add($RingAfterTimeComboBox)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)

	
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}


function DelegateDialog()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	[array]$Script:editDialogDelegateMembers = $Script:UpdatedUserForwardSettings.Delegates
     
	
	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size(8,8) 
	$TitleLabel.Size = New-Object System.Drawing.Size(590,20)
	$TitleLabel.Text = "Delegates can schedule Skype Meetings, make calls, and receive calls on the user's behalf."
	$TitleLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	#Data Grid View ============================================================
	$dgv = New-Object Windows.Forms.DataGridView
	$dgv.Size = New-Object System.Drawing.Size(560,180)
	$dgv.Location = New-Object System.Drawing.Size(8,50)
	$dgv.AutoGenerateColumns = $false
	$dgv.RowHeadersVisible = $false
	$dgv.AllowUserToAddRows = $false
	$dgv.BackgroundColor = [System.Drawing.Color]::White
	$dgv.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
	$dgv.AllowUserToResizeColumns = $false
	$dgv.AllowUserToResizeRows = $false
	$dgv.ColumnHeadersHeightSizeMode  = [Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
	$dgv.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

	$titleColumn0 = New-Object Windows.Forms.DataGridViewCheckBoxColumn
	$titleColumn0.HeaderText = "Receive Calls"
	$titleColumn0.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$titleColumn0.MinimumWidth = 80
	$titleColumn0.Width = 80
	$titleColumn0.Visible = $false
	$dgv.Columns.Add($titleColumn0) | Out-Null


	$titleColumn1 = New-Object Windows.Forms.DataGridViewTextBoxColumn
	$titleColumn1.HeaderText = "Delegate"
	$titleColumn1.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$titleColumn1.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
	$titleColumn1.ReadOnly = $true
	$titleColumn1.MinimumWidth = 557
	$titleColumn1.Width = 557
	$dgv.Columns.Add($titleColumn1) | Out-Null
	
	
	foreach($item in $Script:editDialogDelegateMembers)
	{
		$item = RemoveSip $item
		[bool] $delegate = $true
		$dgv.Rows.Add( @($delegate,$item) )
	}
	
	
	# Create the OK button.
    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Size(10,235)
    $AddButton.Size = New-Object System.Drawing.Size(100,20)
    $AddButton.Text = "Add..."
	$AddButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$AddButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $AddButton.Add_Click({ 

		$result = UserPickerDialog
		
		foreach($item in $result)
		{
			$dgv.Rows.Add( @($false,$item) )
			[array]$Script:editDialogDelegateMembers += $item
		}
		
	})
	
	$AddButton.Add_MouseHover(
	{
	   $AddButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$AddButton.Add_MouseLeave(
	{
	   $AddButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	# Create the OK button.
    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Size(125,235)
    $RemoveButton.Size = New-Object System.Drawing.Size(100,20)
    $RemoveButton.Text = "Remove"
	$RemoveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$RemoveButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $RemoveButton.Add_Click({ 
	
		$beforeDelete = $dgv.CurrentCell.RowIndex
		[array]$itemArray = @()
		foreach($item in $dgv.SelectedRows.Cells[1].Value)
		{
			$itemArray += $item
		}
		foreach($item in $itemArray)
		{
			$dgv.Rows.RemoveAt($dgv.CurrentCell.RowIndex)
			[array]$Script:editDialogDelegateMembers = [array]$Script:editDialogDelegateMembers | Where-Object { $_ -ne $item }
		}
		
		if($beforeDelete -gt ($dgv.Rows.Count - 1))
		{
			$beforeDelete = $beforeDelete - 1
		}
		if($dgv.Rows.Count -gt 0)
		{
			$dgv.Rows[$beforeDelete].Selected = $true
		}
		elseif($dgv.Rows.Count -eq 0)
		{
			#Do nothing
		}
	
	})
	
	$RemoveButton.Add_MouseHover(
	{
	   $RemoveButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$RemoveButton.Add_MouseLeave(
	{
	   $RemoveButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	
	$RingAfterLabel = New-Object System.Windows.Forms.Label
	$RingAfterLabel.Location = New-Object System.Drawing.Size(20,278) 
	$RingAfterLabel.Size = New-Object System.Drawing.Size(220,20)
	$RingAfterLabel.Text = "Ring delegates after this many seconds:"
	$RingAfterLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	$RingAfterTimeComboBox = New-Object System.Windows.Forms.ComboBox 
	$RingAfterTimeComboBox.Location = New-Object System.Drawing.Size(250,275) 
	$RingAfterTimeComboBox.Size = New-Object System.Drawing.Size(130,18) 
	$RingAfterTimeComboBox.DropDownHeight = 80
	$RingAfterTimeComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
	$RingAfterTimeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	
	[void]$RingAfterTimeComboBox.Items.Add("0 - at the same time")
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 5 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{[void]$RingAfterTimeComboBox.Items.Add("5")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 10 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{[void]$RingAfterTimeComboBox.Items.Add("10")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 15 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)  #ONLY 0-10 IS SUPPORTED VIA POWERSHELL.... WHY? I DON'T KNOW.
	{[void]$RingAfterTimeComboBox.Items.Add("15")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 20 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{[void]$RingAfterTimeComboBox.Items.Add("20")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 25 -or $Script:UpdatedUserForwardSettings.UnansweredWaitTime -eq 0)
	{[void]$RingAfterTimeComboBox.Items.Add("25")}	
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 30)
	{[void]$RingAfterTimeComboBox.Items.Add("30")}	
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 35)
	{[void]$RingAfterTimeComboBox.Items.Add("35")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 40)
	{[void]$RingAfterTimeComboBox.Items.Add("40")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 45)
	{[void]$RingAfterTimeComboBox.Items.Add("45")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 50)
	{[void]$RingAfterTimeComboBox.Items.Add("50")}
	if($Script:UpdatedUserForwardSettings.UnansweredWaitTime -gt 55)
	{[void]$RingAfterTimeComboBox.Items.Add("55")}

	$numberOfItems = $RingAfterTimeComboBox.Items.count
	if($numberOfItems -gt 0)
	{
		if($Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime -eq 0 -and $Script:UpdatedUserForwardSettings.DelegateRingWaitTime -eq 0)
		{
			$RingAfterTimeComboBox.SelectedIndex = $RingAfterTimeComboBox.FindStringExact("0 - at the same time")
		}
		elseif($Script:UpdatedUserForwardSettings.DelegateRingWaitTime -gt 0)
		{
			$RingAfterTimeComboBox.SelectedIndex = $RingAfterTimeComboBox.FindStringExact($Script:UpdatedUserForwardSettings.DelegateRingWaitTime)
		}
		elseif($Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime -gt 0)
		{
			$RingAfterTimeComboBox.SelectedIndex = $RingAfterTimeComboBox.FindStringExact($Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime)
		}
		else
		{
			$RingAfterTimeComboBox.SelectedIndex = 0
		}
	}
	
	$RingAfterTimeComboBox.Add_SelectionChangeCommitted({
		if($RingAfterTimeComboBox.SelectedItem -eq "15" -or $RingAfterTimeComboBox.SelectedItem -eq "20" -or $RingAfterTimeComboBox.SelectedItem -eq "25" -or $RingAfterTimeComboBox.SelectedItem -eq "30" -or $RingAfterTimeComboBox.SelectedItem -eq "35" -or $RingAfterTimeComboBox.SelectedItem -eq "40" -or $RingAfterTimeComboBox.SelectedItem -eq "45" -or $RingAfterTimeComboBox.SelectedItem -eq "50" -or $RingAfterTimeComboBox.SelectedItem -eq "55" -or $RingAfterTimeComboBox.SelectedItem -eq "60")
		{
			[System.Windows.Forms.MessageBox]::Show("The PowerShell commands currently only support settings between 0-10 seconds. Changing the selection back to 10.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			$RingAfterTimeComboBox.SelectedIndex = $RingAfterTimeComboBox.FindStringExact("10")
		}
	})
	
	$ToolTip = New-Object System.Windows.Forms.ToolTip 
	$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
	$ToolTip.IsBalloon = $true 
	$ToolTip.InitialDelay = 1000 
	$ToolTip.ReshowDelay = 2000 
	$ToolTip.AutoPopDelay = 10000
	#$ToolTip.ToolTipTitle = "Help:"
	$ToolTip.SetToolTip($RingAfterTimeComboBox, "Note: The PowerShell commands only support 0-15 seconds right now.") 
	

	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(340,320)
    $okButton.Size = New-Object System.Drawing.Size(100,25)
    $okButton.Text = "OK"
	$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$okButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
	$okButton.tabIndex = 1
    $okButton.Add_Click({ 
	
		$Script:UpdatedUserForwardSettings.Delegates = $Script:editDialogDelegateMembers
		if($RingAfterTimeComboBox.Text -eq "0 - at the same time")
		{
			$Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime = 0
			$Script:UpdatedUserForwardSettings.DelegateRingWaitTime = 0
		}
		else
		{
			$Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime = $RingAfterTimeComboBox.Text
			$Script:UpdatedUserForwardSettings.DelegateRingWaitTime = $RingAfterTimeComboBox.Text
		}
		
		Write-Host "INFO: Settings Updated." -foreground "Yellow"
		$form.Close() 
			
	})
	
	$okButton.Add_MouseHover(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$okButton.Add_MouseLeave(
	{
	   $okButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})
	
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(460,320)
    $CancelButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelButton.Text = "Cancel"
	$CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$CancelButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $CancelButton.Add_Click({ 
	
		$form.Close() 
		
	})
	
	$CancelButton.Add_MouseHover(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlue
	})
	$CancelButton.Add_MouseLeave(
	{
	   $CancelButton.FlatAppearance.BorderColor = $Script:buttonBorderBlack
	})

	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Call Forwarding - Delegates"
    $form.Size = New-Object System.Drawing.Size(600,400)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
     
	$form.Controls.Add($TitleLabel)
	$form.Controls.Add($TitleLabel2)
	$form.Controls.Add($dgv)
	$form.Controls.Add($AddButton)
	$form.Controls.Add($RemoveButton)
	$form.Controls.Add($RingAfterLabel)
	$form.Controls.Add($RingAfterTimeComboBox)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)

	
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}


function FilterUsersList()
{
	$UsersListbox.Items.Clear()
	[string] $theFilter = $FilterTextBox.text
	# Add Lync Users ============================================================
	$Script:users | ForEach-Object {if($_.SipAddress -imatch $theFilter){[void] $UsersListbox.Items.Add((RemoveSip $_.SipAddress))}}  #.ToLower()  .Replace('sip:','').Replace("SIP:","")
	
	if($UsersListbox.Items.Count -gt 1)
	{
		$UsersListbox.SelectedIndex = 0
	}
}


#User							: sip:name@domain.com
#CallForwardingEnabled			: False
#ForwardDestination				: 
#ForwardImmediateEnabled		: False
#SimultaneousRingEnabled		: False
#SimultaneousRingDestination 	:
#ForwardToDelegates				: False
#TeamRingEnabled				: False
#Team							: {sip:name@domain.com}
#Delegates						: {}
#DelegateRingWaitTime				: 0
#TeamDelegateRingWaitTime		: 0
#SettingsActiveWorkhours		: False
#UnansweredToVoicemail			: False
#UnansweredCallForwardDestination
#UnansweredWaitTime				: 30



function Get-ForwardSettings([PSObject] $ForwardSettings)
{
	Write-Verbose "Function Get-ForwardSettings"
	$SipAddress = $UsersListbox.SelectedItem
	$Script:CurrentUserForwardSettings = $ForwardSettings
	$Script:UpdatedUserForwardSettings = $ForwardSettings
	<#
	Write-Host "-----------------$SipAddress ForwardSettings SETTINGS --------------------------"
	Write-Host "CallForwardingEnabled           : " $ForwardSettings.CallForwardingEnabled
	Write-Host "ForwardDestination              : " $ForwardSettings.ForwardDestination
	Write-Host "ForwardImmediateEnabled         : " $ForwardSettings.ForwardImmediateEnabled
	Write-Host "SimultaneousRingEnabled         : " $ForwardSettings.SimultaneousRingEnabled
	Write-Host "SimultaneousRingDestination     : " $ForwardSettings.SimultaneousRingDestination
	Write-Host "ForwardToDelegates              : " $ForwardSettings.ForwardToDelegates
	Write-Host "TeamRingEnabled                 : " $ForwardSettings.TeamRingEnabled
	Write-Host "Team                            : " $ForwardSettings.Team
	Write-Host "Delegates                       : " $ForwardSettings.Delegates
	Write-Host "DelegateRingWaitTime            : " $ForwardSettings.DelegateRingWaitTime
	Write-Host "TeamDelegateRingWaitTime        : " $ForwardSettings.TeamDelegateRingWaitTime
	Write-Host "SettingsActiveWorkhours         : " $ForwardSettings.SettingsActiveWorkhours
	Write-Host "UnansweredToVoicemail           : " $ForwardSettings.UnansweredToVoicemail
	Write-Host "UnansweredCallForwardDestination: " $ForwardSettings.UnansweredCallForwardDestination
	Write-Host "UnansweredWaitTime              : " $ForwardSettings.UnansweredWaitTime
	#>
	Write-Host "-----------------$SipAddress CurrentUserForwardSettings SETTINGS --------------------------"
	Write-Host "CallForwardingEnabled           : " $Script:CurrentUserForwardSettings.CallForwardingEnabled
	Write-Host "ForwardDestination              : " $Script:CurrentUserForwardSettings.ForwardDestination
	Write-Host "ForwardImmediateEnabled         : " $Script:CurrentUserForwardSettings.ForwardImmediateEnabled
	Write-Host "SimultaneousRingEnabled         : " $Script:CurrentUserForwardSettings.SimultaneousRingEnabled
	Write-Host "SimultaneousRingDestination     : " $Script:CurrentUserForwardSettings.SimultaneousRingDestination
	Write-Host "ForwardToDelegates              : " $Script:CurrentUserForwardSettings.ForwardToDelegates
	Write-Host "SimultaneousRingDelegates       : " $Script:CurrentUserForwardSettings.SimultaneousRingDelegates
	Write-Host "TeamRingEnabled                 : " $Script:CurrentUserForwardSettings.TeamRingEnabled
	Write-Host "Team                            : " $Script:CurrentUserForwardSettings.Team
	Write-Host "Delegates                       : " $Script:CurrentUserForwardSettings.Delegates
	Write-Host "DelegateRingWaitTime            : " $Script:CurrentUserForwardSettings.DelegateRingWaitTime
	Write-Host "TeamDelegateRingWaitTime        : " $Script:CurrentUserForwardSettings.TeamDelegateRingWaitTime
	Write-Host "SettingsActiveWorkhours         : " $Script:CurrentUserForwardSettings.SettingsActiveWorkhours
	Write-Host "UnansweredToVoicemail           : " $Script:CurrentUserForwardSettings.UnansweredToVoicemail
	Write-Host "UnansweredCallForwardDestination: " $Script:CurrentUserForwardSettings.UnansweredCallForwardDestination
	Write-Host "UnansweredWaitTime              : " $Script:CurrentUserForwardSettings.UnansweredWaitTime
	Write-Host "-------------------------------------------------------------------------------------------"
	<#
	Write-Host "-----------------$SipAddress UpdatedUserForwardSettings SETTINGS --------------------------"
	Write-Host "CallForwardingEnabled           : " $Script:UpdatedUserForwardSettings.CallForwardingEnabled
	Write-Host "ForwardDestination              : " $Script:UpdatedUserForwardSettings.ForwardDestination
	Write-Host "ForwardImmediateEnabled         : " $Script:UpdatedUserForwardSettings.ForwardImmediateEnabled
	Write-Host "SimultaneousRingEnabled         : " $Script:UpdatedUserForwardSettings.SimultaneousRingEnabled
	Write-Host "SimultaneousRingDestination     : " $Script:UpdatedUserForwardSettings.SimultaneousRingDestination
	Write-Host "ForwardToDelegates              : " $Script:UpdatedUserForwardSettings.ForwardToDelegates
	Write-Host "TeamRingEnabled                 : " $Script:UpdatedUserForwardSettings.TeamRingEnabled
	Write-Host "Team                            : " $Script:UpdatedUserForwardSettings.Team
	Write-Host "Delegates                       : " $Script:UpdatedUserForwardSettings.Delegates
	Write-Host "DelegateRingWaitTime            : " $Script:UpdatedUserForwardSettings.DelegateRingWaitTime
	Write-Host "TeamDelegateRingWaitTime        : " $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
	Write-Host "SettingsActiveWorkhours         : " $Script:UpdatedUserForwardSettings.SettingsActiveWorkhours
	Write-Host "UnansweredToVoicemail           : " $Script:UpdatedUserForwardSettings.UnansweredToVoicemail
	Write-Host "UnansweredCallForwardDestination: " $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
	Write-Host "UnansweredWaitTime              : " $Script:UpdatedUserForwardSettings.UnansweredWaitTime
	#>
	
	
	Write-Host
	
	if($Script:CurrentUserForwardSettings.SimultaneousRingEnabled -eq $true) #SIM RING
	{
		Write-Host "INFO: SimultaneousRingEnabled = True" -foreground "Yellow"
		$SimRingRadioButton.Checked = $true
		
		$UnansweredCallsWillGoToLabel.Visible = $true
		$UnansweredCallsWillGoToLink.Visible = $true
		
		#SIM RING
		Write-Host "INFO: SimultaneousRingDestination = " $Script:CurrentUserForwardSettings.SimultaneousRingDestination -foreground "Yellow"
		$SimRingDestination = RemoveSip ($Script:CurrentUserForwardSettings.SimultaneousRingDestination) #[regex]::match($Script:CurrentUserForwardSettings.SimultaneousRingDestination,'(sip:)?([^@]*)(@.*)?$').Groups[2].Value 

		$SipAddress = $UsersListbox.SelectedItem
		$UserDetails = Get-CsUser -identity "sip:${SipAddress}" | Select-Object DisplayName, LineURI
		$displayname = $UserDetails.DisplayName
		$lineuri = RemoveTel ([string]$UserDetails.LineURI)
		$SimRingLocation = $SimRingComboBox.SelectedItem
		
		if($SimRingDestination -ne $null -and $SimRingDestination -ne "")
		{
			if($SimRingDestination -eq "Team")
			{
				if($SimRingComboBox.FindStringExact("My Team-Call Group") -eq -1)
				{
					$SimRingComboBox.Items.Add("My Team-Call Group")
				}
				$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact($SimRingDestination)
				#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring My Team-Call Group."
				$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring their Team-Call Group."
			}
			elseif($SimRingDestination -eq "Delegates")
			{
				if($SimRingComboBox.FindStringExact("My Delegates") -eq -1)
				{
					$SimRingComboBox.Items.Add("My Delegates")
				}
				$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact($SimRingDestination)
				#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring My Delegates."
				$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring their Delegates."
			}
			elseif($SimRingDestination -eq "Select from this list")
			{
				$CallsWillRingYouAtLabel.Text = "Select a Sim-Ring location."
			}
			else
			{
				$SimRingComboBox.Items.Add($SimRingDestination)
				$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact($SimRingDestination)
				#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri and also ring ${SimRingDestination}."
				$CallsWillRingYouAtLabel.Text = "Calls will ring the user and also ring ${SimRingDestination}."
			}
		}
		else
		{
			if($SimRingComboBox.FindStringExact("Select from this list") -eq -1)
			{
				Write-Host "Adding SimRingComboBox Select from this list Option"
				$SimRingComboBox.Items.Insert(0,"Select from this list")
				$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact("Select from this list")
				
				#ADDED
				$CallsWillRingYouAtLabel.Text = "Select a Sim Ring location."
			}
		}
		
		if($Script:CurrentUserForwardSettings.SimultaneousRingDelegates)
		{
			$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact("My Delegates")
			Write-Host "INFO: SimRingDestination = My Team-Call Group" -foreground "Yellow"
		}
		elseif($Script:CurrentUserForwardSettings.TeamRingEnabled)
		{
			$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact("My Team-Call Group")
			Write-Host "INFO: SimRingDestination = My Team-Call Group" -foreground "Yellow"
		}
		
		
	}
	elseif($Script:CurrentUserForwardSettings.ForwardImmediateEnabled -eq $true) #FORWARDING
	{
		Write-Host "INFO: ForwardImmediateEnabled = True" -foreground "Yellow"
		$ForwardRadioButton.Checked = $true
		
		$UnansweredCallsWillGoToLabel.Visible = $false
		$UnansweredCallsWillGoToLink.Visible = $false
		
		#FORWARD DEST
		Write-Host "INFO: ForwardDestination = " $Script:CurrentUserForwardSettings.ForwardDestination -foreground "Yellow"
		$ForwardDestination = RemoveSip ($Script:CurrentUserForwardSettings.ForwardDestination) 
		Write-Host "INFO: ForwardDestination After Regex = $ForwardDestination" -foreground "Yellow"
		if($ForwardDestination -ne $null -and $ForwardDestination -ne "")
		{
			
			if($ForwardDestination -eq "Delegates")
			{
				$ForwardOnComboBox.SelectedIndex = $ForwardOnComboBox.FindStringExact("My Delegates")
			}
			else
			{
				$ForwardOnComboBox.Items.Add($ForwardDestination)
				$ForwardOnComboBox.SelectedIndex = $ForwardOnComboBox.FindStringExact($ForwardDestination)
			}
						
			#ADDED
			Write-Verbose "Currently selected forward location: $ForwardLocation"
			if($ForwardDestination -eq "Select from this list")
			{
				$CallsWillRingYouAtLabel.Text = "Select a forward location."
			}
			elseif($ForwardDestination -eq "Delegates")
			{
				#$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to My Delegates."
				$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to the user's Delegates."
			}
			elseif($Script:CurrentUserForwardSettings.ForwardDestination -match ";opaque=app:voicemail")
			{
				$ForwardOnComboBox.SelectedIndex = $ForwardOnComboBox.FindStringExact("Voice Mail")
				$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to Voice Mail."
			}
			else
			{
				$ForwardDestination = RemoveSip $ForwardDestination
				$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to ${ForwardDestination}."
			}
		}
		else
		{
			if($ForwardOnComboBox.FindStringExact("Select from this list") -eq -1)
			{
				Write-Host "Adding ForwardOnComboBox Select from this list Option"
				$ForwardOnComboBox.Items.Insert(0,"Select from this list")
				$ForwardOnComboBox.SelectedIndex = $ForwardOnComboBox.FindStringExact("Select from this list")
				
				$CallsWillRingYouAtLabel.Text = "Select a forward location."
			}
		}

		if($Script:CurrentUserForwardSettings.ForwardToDelegates)
		{
			$ForwardOnComboBox.SelectedIndex = $ForwardOnComboBox.FindStringExact("My Delegates")
			Write-Host "INFO: ForwardDestination = My Delegates" -foreground "Yellow"
			
			#$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to My Delegates."
			$CallsWillRingYouAtLabel.Text = "Calls will be forwarded directly to the user's Delegates."
		}
		
	}
	elseif($Script:CurrentUserForwardSettings.CallForwardingEnabled -eq $false) #FORWARD DISABLED
	{
		Write-Host "INFO: Call forwarding off" -foreground "Yellow"
		$OffRadioButton.Checked = $true
		
		$UnansweredCallsWillGoToLabel.Visible = $true
		$UnansweredCallsWillGoToLink.Visible = $true
		
		$SipAddress = $UsersListbox.SelectedItem
		$UserDetails = Get-CsUser -identity "sip:${SipAddress}" | Select-Object DisplayName, LineURI
		$displayname = $UserDetails.DisplayName
		$lineuri = RemoveTel ([string]$UserDetails.LineURI)
		if($lineuri -eq "" -or $lineuri -eq $null)
		{
			$lineuri = "<not set>"
		}
		#$CallsWillRingYouAtLabel.Text = "Calls will ring $displayname at $lineuri"
		$CallsWillRingYouAtLabel.Text = "Calls will ring the user directly."
	}
	
	
	#UNANSWERED CALL SETTINGS
	$UnansweredToVoicemail = $Script:CurrentUserForwardSettings.UnansweredToVoicemail
	$UnansweredCallForwardDestination = (RemoveSip $Script:CurrentUserForwardSettings.UnansweredCallForwardDestination)  #.Replace("sip:","").Replace("SIP:","")
	$UnansweredWaitTime = $Script:CurrentUserForwardSettings.UnansweredWaitTime
	$UserVoicemailSettings = CheckUserVoicemail $SipAddress
	if($UnansweredCallForwardDestination -ne $null -and $UnansweredCallForwardDestination -ne "")
	{
		$UnansweredCallsWillGoToLink.Text = "$UnansweredCallForwardDestination in $UnansweredWaitTime seconds"
		Write-Host "INFO: $UnansweredCallForwardDestination in $UnansweredWaitTime seconds" -foreground "Yellow"
	}
	elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail -or ($UserVoicemailSettings.ExUmEnabled -eq $true -or $UserVoicemailSettings.HostedVoicemail -eq $true))
	{
		$UnansweredCallsWillGoToLink.Text = "Voice mail in $UnansweredWaitTime seconds"
		Write-Host "INFO: Voice mail in $UnansweredWaitTime seconds" -foreground "Yellow"
	}
	elseif($UnansweredCallForwardDestination -eq "" -or $UnansweredCallForwardDestination -eq $null)
	{
		$UnansweredCallsWillGoToLink.Text = "No voice mail. Calls will continuously ring for $UnansweredWaitTime secs."
		Write-Host "INFO: No voice mail. Calls will continuously ring for $UnansweredWaitTime seconds." -foreground "Yellow"
	}
	
	
	if($Script:CurrentUserForwardSettings.SettingsActiveWorkhours)
	{
		$TheseSettingsWillApplyLink.Text = "During work hours set in Outlook"
		Write-Host "INFO: During work hours set in Outlook" -foreground "Yellow"
	}
	else
	{
		$TheseSettingsWillApplyLink.Text = "All of the time"
		Write-Host "INFO: All of the time" -foreground "Yellow"
	}
	
	#DELGATES AND TEAM CALL MEMBERS
	$TeamCount = ($Script:UpdatedUserForwardSettings.Team).count
	$TeamCallGroupLabel.Text = "Edit team-call group members (${TeamCount})"
	$TeamCallGroupLabel.AutoSize = $true
	$DelegateCount = ($Script:UpdatedUserForwardSettings.Delegates).count
	$DelegateGroupLabel.Text = "Edit delegate members (${DelegateCount})"
	$DelegateGroupLabel.AutoSize = $true
	
	if($ForwardRadioButton.Checked -eq $false)
	{
		if($ForwardOnComboBox.FindStringExact("Select from this list") -eq -1)
		{
			Write-Verbose "Adding ForwardOnComboBox Select from this list Option"
			$ForwardOnComboBox.Items.Insert(0,"Select from this list")
			$ForwardOnComboBox.SelectedIndex = $ForwardOnComboBox.FindStringExact("Select from this list")
		}
	}
	if($SimRingRadioButton.Checked -eq $false)
	{
		if($SimRingComboBox.FindStringExact("Select from this list") -eq -1)
		{
			Write-Verbose "Adding Select from this list Option"
			$SimRingComboBox.Items.Insert(0,"Select from this list")
			$SimRingComboBox.SelectedIndex = $SimRingComboBox.FindStringExact("Select from this list")
		}
	}
	
	PrintUpdatedUserSettings
}

#User							: sip:name@domain.com
#CallForwardingEnabled			: False
#ForwardDestination				: 
#ForwardImmediateEnabled		: False
#SimultaneousRingEnabled		: False
#SimultaneousRingDestination 	:
#ForwardToDelegates				: False
#TeamRingEnabled				: False
#Team							: {sip:name@domain.com}
#Delegates						: {}
#DelegateRingWaitTime				: 0
#TeamDelegateRingWaitTime		: 0
#SettingsActiveWorkhours		: False
#UnansweredToVoicemail			: False
#UnansweredCallForwardDestination
#UnansweredWaitTime				: 30

function Load-TestUserSettingsObject
{
	$SipAddress = $UsersListbox.SelectedItem
	
	$UserSettings = New-Object PSObject
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name User -Value $SipAddress 
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name CallForwardingEnabled -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name ForwardDestination -Value "+61399991111"
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name ForwardImmediateEnabled -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name SimultaneousRingEnabled -Value $True
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name SimultaneousRingDestination -Value "+61399992222"
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name ForwardToDelegates -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name SimultaneousRingDelegates -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name TeamRingEnabled -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name Team -Value @("sip:name1@myteamslab.com", "sip:name2@myteamslab.com")
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name Delegates -Value @("sip:name1@myteamslab.com")
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name DelegateRingWaitTime -Value "5"
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name TeamDelegateRingWaitTime -Value "10"
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name SettingsActiveWorkhours -Value $false
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name UnansweredToVoicemail -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name UnansweredCallForwardDestination -Value "+61399993333"
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name UnansweredWaitTime -Value "35"
	
	Write-Host $UserSettings
	return $UserSettings
}

function Load-DefaultForwardSettings
{
		$SipAddress = $UsersListbox.SelectedItem
	
	$UserSettings = New-Object PSObject
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name User -Value $SipAddress 
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name CallForwardingEnabled -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name ForwardDestination -Value $Null
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name ForwardImmediateEnabled -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name SimultaneousRingEnabled -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name SimultaneousRingDestination -Value $Null
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name ForwardToDelegates -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name SimultaneousRingDelegates -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name TeamRingEnabled -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name Team -Value @()
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name Delegates -Value @()
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name DelegateRingWaitTime -Value "0"
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name TeamDelegateRingWaitTime -Value "0"
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name SettingsActiveWorkhours -Value $false
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name UnansweredToVoicemail -Value $False
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name UnansweredCallForwardDestination -Value $Null
		Add-Member -InputObject $UserSettings -MemberType NoteProperty -Name UnansweredWaitTime -Value "20"
		
	return $UserSettings
}

function PrintUpdatedUserSettings()
{
	Write-Verbose "-----------------UPDATED USER SETTINGS --------------------------"
	Write-Verbose "CallForwardingEnabled           : $($Script:UpdatedUserForwardSettings.CallForwardingEnabled)" 
	Write-Verbose "ForwardDestination              : $($Script:UpdatedUserForwardSettings.ForwardDestination)" 
	Write-Verbose "ForwardImmediateEnabled         : $($Script:UpdatedUserForwardSettings.ForwardImmediateEnabled)" 
	Write-Verbose "SimultaneousRingEnabled         : $($Script:UpdatedUserForwardSettings.SimultaneousRingEnabled)" 
	Write-Verbose "SimultaneousRingDestination     : $($Script:UpdatedUserForwardSettings.SimultaneousRingDestination)" 
	Write-Verbose "ForwardToDelegates              : $($Script:UpdatedUserForwardSettings.ForwardToDelegates)" 
	Write-Verbose "SimultaneousRingDelegates       : $($Script:UpdatedUserForwardSettings.SimultaneousRingDelegates)" 
	Write-Verbose "TeamRingEnabled                 : $($Script:UpdatedUserForwardSettings.TeamRingEnabled)" 
	Write-Verbose "Team                            : $($Script:UpdatedUserForwardSettings.Team)" 
	Write-Verbose "Delegates                       : $($Script:UpdatedUserForwardSettings.Delegates)" 
	Write-Verbose "DelegateRingWaitTime            : $($Script:UpdatedUserForwardSettings.DelegateRingWaitTime)" 
	Write-Verbose "TeamDelegateRingWaitTime        : $($Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime)" 
	Write-Verbose "SettingsActiveWorkhours         : $($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)" 
	Write-Verbose "UnansweredToVoicemail           : $($Script:UpdatedUserForwardSettings.UnansweredToVoicemail)" 
	Write-Verbose "UnansweredCallForwardDestination: $($Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination)" 
	Write-Verbose "UnansweredWaitTime              : $($Script:UpdatedUserForwardSettings.UnansweredWaitTime)" 
	Write-Verbose "-----------------------------------------------------------------"
}

function Set-ForwardSettings
{
	Write-Verbose "Function Set-ForwardSettings"
	$SipAddresses = $SetUsersListbox.SelectedItems
	
	if($SetUsersListbox.SelectedItems.Count -eq 0)
	{
		Write-Host "INFO: No users selected." -foreground "yellow"
	}
	
	$DelegateTeamCallCheck = $true
	foreach($SipAddress in $SipAddresses)
	{
		Write-Verbose "Checking if $SipAddress exists in delegates or team-call settings"
		[array]$UpdatedDelegates = $Script:UpdatedUserForwardSettings.Delegates
		foreach($User in $UpdatedDelegates)
		{
			$User = RemoveSip($User)
			Write-Verbose "Delegate check: is $User == $SipAddress"
			if($User -eq $SipAddress)
			{
				$DelegateTeamCallCheck = $false
				[System.Windows.Forms.MessageBox]::Show("The user $SipAddress is configured as a delegate. You cannot set a user as their own delegate. Please remove them and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				break
			}
		}
		if(!$DelegateTeamCallCheck)
		{break}
		
		[array]$UpdatedTeam = $Script:UpdatedUserForwardSettings.Team
		foreach($User in $UpdatedTeam)
		{
			$User = RemoveSip($User)
			Write-Verbose "Team check: is $User == $SipAddress"
			if($User -eq $SipAddress)
			{
				$DelegateTeamCallCheck = $false
				[System.Windows.Forms.MessageBox]::Show("The user $SipAddress is configured in the team-call group. You cannot configure a user in their own team-call group. Please remove them and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				break
			}
		}
		if(!$DelegateTeamCallCheck)
		{break}
	}
		
	foreach($SipAddress in $SipAddresses)
	{
		$ForwardDestination = $Script:UpdatedUserForwardSettings.ForwardDestination
		$ForwardDestination = RemoveSip($ForwardDestination)
		Write-Verbose "Forward Destination check: is $SipAddress == $ForwardDestination"
		if($SipAddress -eq $ForwardDestination)
		{
			$DelegateTeamCallCheck = $false
			[System.Windows.Forms.MessageBox]::Show("The user $SipAddress is configured as the forwarding location. You cannot configure a user to forward to themself. Please change the forward location and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		}
		if(!$DelegateTeamCallCheck)
		{break}
		<#
		$SimRingDestination = $Script:UpdatedUserForwardSettings.SimultaneousRingDestination
		$SimRingDestination = RemoveSip($SimRingDestination)
		Write-Verbose "SimultaneousRing Destination check: is $User == $SimRingDestination"
		if($SipAddress -eq $SimRingDestination)
		{
			$DelegateTeamCallCheck = $false
			[System.Windows.Forms.MessageBox]::Show("The user $SipAddress is configured as the sim ring location. You cannot configure a user to sim ring to themself. Please change the sim ring location and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		}
		if(!$DelegateTeamCallCheck)
		{break}
		#>
		$ForwardUnanswered = $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
		$ForwardUnanswered = RemoveSip($ForwardUnanswered)
		Write-Verbose "UpdatedUserForwardSettings check: is $User == $ForwardUnanswered"
		if($SipAddress -eq $ForwardUnanswered)
		{
			$DelegateTeamCallCheck = $false
			[System.Windows.Forms.MessageBox]::Show("The user $SipAddress is configured as the unanswered forward location. You cannot configure a user to forward to themself. Please change the unanswered forward location and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		}
		if(!$DelegateTeamCallCheck)
		{break}
	}
	
	if($DelegateTeamCallCheck)
	{
		#DO FOR EACH SELECTED USER
		foreach($SipAddress in $SipAddresses)
		{
			$UserSettings = Get-CsUserCallForwardingSettings -identity $SipAddress
			
			PrintUpdatedUserSettings
			
			#CHECK SETTINGS:
			Write-Verbose "Checking user settings for errors"
			if($Script:UpdatedUserForwardSettings.ForwardImmediateEnabled -eq $true -and ($Script:UpdatedUserForwardSettings.ForwardDestination -eq $null -or $Script:UpdatedUserForwardSettings.ForwardDestination -eq ""))
			{
				[System.Windows.Forms.MessageBox]::Show("A number, contact or delegate must be supplied in order to set call forward immediate. Please select a location for the call forward.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			}
			elseif($Script:UpdatedUserForwardSettings.SimultaneousRingEnabled -eq $true -and ($Script:UpdatedUserForwardSettings.SimultaneousRingDestination -eq $null -or $Script:UpdatedUserForwardSettings.SimultaneousRingDestination -eq ""))
			{
				[System.Windows.Forms.MessageBox]::Show("A number, contact, delegate(s) or team-call group must be supplied in order to enable simultaneous ring. Please select a location to sim ring.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			}
			else
			{			
				
				<#
				Write-Host "-----------------CURRENT USER SETTINGS --------------------------"
				Write-Host "CallForwardingEnabled           : " $UserSettings.CallForwardingEnabled
				Write-Host "ForwardDestination              : " $UserSettings.ForwardDestination
				Write-Host "ForwardImmediateEnabled         : " $UserSettings.ForwardImmediateEnabled
				Write-Host "SimultaneousRingEnabled         : " $UserSettings.SimultaneousRingEnabled
				Write-Host "SimultaneousRingDestination     : " $UserSettings.SimultaneousRingDestination
				Write-Host "ForwardToDelegates              : " $UserSettings.ForwardToDelegates
				Write-Host "SimultaneousRingDelegates       : " $UserSettings.SimultaneousRingDelegates
				Write-Host "TeamRingEnabled                 : " $UserSettings.TeamRingEnabled
				Write-Host "Team                            : " $UserSettings.Team
				Write-Host "Delegates                       : " $UserSettings.Delegates
				Write-Host "DelegateRingWaitTime            : " $UserSettings.DelegateRingWaitTime
				Write-Host "TeamDelegateRingWaitTime        : " $UserSettings.TeamDelegateRingWaitTime
				Write-Host "SettingsActiveWorkhours         : " $UserSettings.SettingsActiveWorkhours
				Write-Host "UnansweredToVoicemail           : " $UserSettings.UnansweredToVoicemail
				Write-Host "UnansweredCallForwardDestination: " $UserSettings.UnansweredCallForwardDestination
				Write-Host "UnansweredWaitTime              : " $UserSettings.UnansweredWaitTime
				Write-Host
				#>
				<#
				Write-Host "-----------------UPDATED USER SETTINGS --------------------------"
				Write-Host "CallForwardingEnabled           : " $Script:UpdatedUserForwardSettings.CallForwardingEnabled
				Write-Host "ForwardDestination              : " $Script:UpdatedUserForwardSettings.ForwardDestination
				Write-Host "ForwardImmediateEnabled         : " $Script:UpdatedUserForwardSettings.ForwardImmediateEnabled
				Write-Host "SimultaneousRingEnabled         : " $Script:UpdatedUserForwardSettings.SimultaneousRingEnabled
				Write-Host "SimultaneousRingDestination     : " $Script:UpdatedUserForwardSettings.SimultaneousRingDestination
				Write-Host "ForwardToDelegates              : " $Script:UpdatedUserForwardSettings.ForwardToDelegates
				Write-Host "SimultaneousRingDelegates       : " $Script:UpdatedUserForwardSettings.SimultaneousRingDelegates
				Write-Host "TeamRingEnabled                 : " $Script:UpdatedUserForwardSettings.TeamRingEnabled
				Write-Host "Team                            : " $Script:UpdatedUserForwardSettings.Team
				Write-Host "Delegates                       : " $Script:UpdatedUserForwardSettings.Delegates
				Write-Host "DelegateRingWaitTime            : " $Script:UpdatedUserForwardSettings.DelegateRingWaitTime
				Write-Host "TeamDelegateRingWaitTime        : " $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
				Write-Host "SettingsActiveWorkhours         : " $Script:UpdatedUserForwardSettings.SettingsActiveWorkhours
				Write-Host "UnansweredToVoicemail           : " $Script:UpdatedUserForwardSettings.UnansweredToVoicemail
				Write-Host "UnansweredCallForwardDestination: " $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
				Write-Host "UnansweredWaitTime              : " $Script:UpdatedUserForwardSettings.UnansweredWaitTime
				Write-Host
				#>
							
				$UpdatedDelegates = $Script:UpdatedUserForwardSettings.Delegates
				$UpdatedTeam = $Script:UpdatedUserForwardSettings.Team
				
				if($UpdatedTeam.count -gt 0)
				{
					Write-Host "RUNNING: Set-CsUserTeamMembers -Identity `"sip:${SipAddress}`" -Team @{replace=$UpdatedTeam}" -foreground "green"
					Set-CsUserTeamMembers -Identity "sip:${SipAddress}" -Team @{replace=$UpdatedTeam}
				}
				else
				{
					Write-Host "RUNNING: Set-CsUserTeamMembers -Identity `"sip:${SipAddress}`" -Team `$null" -foreground "green"
					Set-CsUserTeamMembers -Identity "sip:${SipAddress}" -Team $null
				}
				
				if($UpdatedDelegates -gt 0)
				{
					Write-Host "RUNNING: Set-CsUserDelegates -Identity `"sip:${SipAddress}`" -Delegates @{replace=$UpdatedDelegates}" -foreground "green"
					Set-CsUserDelegates -Identity "sip:${SipAddress}" -Delegates @{replace=$UpdatedDelegates}
				}
				else
				{
					Write-Host "RUNNING: Set-CsUserDelegates -Identity `"sip:${SipAddress}`" -Delegates `$null" -foreground "green"
					Set-CsUserDelegates -Identity "sip:${SipAddress}" -Delegates $null
				}
				
				Write-Host 
				
				#OTHER Settings
				if($OffRadioButton.Checked -eq $true)
				{
					Write-Verbose "Off Radio Button Checked"
					if($Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne "" -and $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne $null)
					{
						if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
						{
							$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
							$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
							#-DisableForwarding
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredToOther `"${UnansweredLocation}`" -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredToOther `"${UnansweredLocation}`" -UnansweredWaitTime `"$WaitTime`" -SettingsActiveWorkHours"
						}
						else
						{
							$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
							$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredToOther `"${UnansweredLocation}`" -UnansweredWaitTime $WaitTime" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredToOther `"${UnansweredLocation}`" -UnansweredWaitTime `"$WaitTime`""
						}
					}
					elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail)
					{
						Write-Verbose "Off Radio Button Checked - UnansweredToVoicemail"
						if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
						{
							$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
							#TURN OFF FORWARDS
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredToVoicemail -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredToVoicemail -UnansweredWaitTime `"$WaitTime`" -SettingsActiveWorkHours"
						}
						else
						{
							$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
							#TURN OFF FORWARDS
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredToVoicemail -UnansweredWaitTime $WaitTime" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredToVoicemail -UnansweredWaitTime `"$WaitTime`""
						}
					}
					else
					{	
						Write-Verbose "Off Radio Button Checked - There is no setting for UnansweredCallForwardDestination. No Voice mail installed."
						Write-Host "INFO: There is no setting for UnansweredCallForwardDestination. No Voice mail installed." -foreground "yellow"
			
						if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
						{
							$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
							$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredWaitTime `"$WaitTime`" -SettingsActiveWorkHours" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredWaitTime `"$WaitTime`" -SettingsActiveWorkHours"
						}
						else
						{
							$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
							$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredWaitTime `"$WaitTime`"" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -DisableForwarding -UnansweredWaitTime `"$WaitTime`""
						}
					}
				}
				elseif($ForwardRadioButton.Checked -eq $true)
				{
					if($ForwardOnComboBox.SelectedItem -eq "My Delegates") #DELEGATE FORWARD
					{
						Write-Verbose "Forward Radio Button Checked"
						if($UpdatedDelegates -gt 0)
						{
							Write-Verbose "Updated Delegates greater than 0"
							Write-Verbose "ForwardOnComboBox SelectedItem = My Delegates"
							if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
							{
								Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
								if($DelegateRingWaitTime -ne "0" -and $DelegateRingWaitTime -ne "5" -and $DelegateRingWaitTime -ne "10")
								{
									Write-Verbose "DelegateRingWaitTime != 0 AND DelegateRingWaitTime != 5 AND DelegateRingWaitTime != 10"
									Write-Host "IMPORTANT INFO: The Wait Time value is $DelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
									$DelegateRingWaitTime = "10"
								}
								Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"Delegates`" -DelegateRingWaitTime $DelegateRingWaitTime -SettingsActiveWorkHours" -foreground "green"
								Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"Delegates`" -DelegateRingWaitTime $DelegateRingWaitTime -SettingsActiveWorkHours"
							}
							else
							{
								Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
								$DelegateRingWaitTime = $Script:UpdatedUserForwardSettings.DelegateRingWaitTime
								if($DelegateRingWaitTime -ne "0" -and $DelegateRingWaitTime -ne "5" -and $DelegateRingWaitTime -ne "10")
								{
									Write-Verbose "DelegateRingWaitTime != 0 AND DelegateRingWaitTime != 5 AND DelegateRingWaitTime != 10"
									Write-Host "IMPORTANT INFO: The Wait Time value is $DelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
									$DelegateRingWaitTime = "10"
								}
								Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"Delegates`" -DelegateRingWaitTime $DelegateRingWaitTime" -foreground "green"
								Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"Delegates`" -DelegateRingWaitTime $DelegateRingWaitTime"

							}
						}
						else
						{
							Write-Host "There are no Delegate members assigned. So User Forwarding Settings were not updated." -foreground "red"
							[System.Windows.Forms.MessageBox]::Show("There are no Delegate members assigned. So User Forwarding Settings were not updated.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
						}
					}
					elseif($ForwardOnComboBox.SelectedItem -eq "Voice Mail") #FORWARD TO VOICEMAIL WORKAROUND - Inital release of PowerShell commands don't support forwarding to Voicemail
					{
						Write-Verbose "Forward Immediate to Voice Mail is set. Using opaque workaround."
						if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
						{
							Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
							$ForwardDestination = $ForwardOnComboBox.SelectedItem
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"sip:${SipAddress};opaque=app:voicemail`" -SettingsActiveWorkHours" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"sip:${SipAddress};opaque=app:voicemail`" -SettingsActiveWorkHours"
						}
						else
						{
							Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
							$ForwardDestination = $ForwardOnComboBox.SelectedItem
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"sip:${SipAddress};opaque=app:voicemail`"" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"sip:${SipAddress};opaque=app:voicemail`""
						}
					}
					else
					{
						if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
						{
							Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
							$ForwardDestination = $ForwardOnComboBox.SelectedItem
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"$ForwardDestination`" -SettingsActiveWorkHours" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"$ForwardDestination`" -SettingsActiveWorkHours"
						}
						else
						{
							Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
							$ForwardDestination = $ForwardOnComboBox.SelectedItem
							Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"$ForwardDestination`"" -foreground "green"
							Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableForwarding `"$ForwardDestination`""
						}
					}
				}
				elseif($SimRingRadioButton.Checked -eq $true) #SIM RING
				{	
					Write-Verbose "SimRing Radio Button Checked"
					if($SimRingComboBox.SelectedItem -eq "My Team-Call Group")  #TEAM CALL SIMRING
					{
						Write-Verbose "SimRingComboBox SelectedItem My Team-Call Group"
						if($UpdatedTeam.count -gt 0) #Check if there are any Team members.
						{
							Write-Verbose "UpdatedTeam count greater than 0"
							Write-Verbose "SimRing UpdatedUserForwardSettings Unanswered To Number"
							if($Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne "" -and $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne $null)
							{
								Write-Verbose "UpdatedUserForwardSettings UnansweredCallForwardDestination -ne null"
								if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours) #Active work hours
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
									$TeamDelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									if($TeamDelegateRingWaitTime -ne "0" -and $TeamDelegateRingWaitTime -ne "5" -and $TeamDelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$TeamDelegateRingWaitTime -ne 0 -and $TeamDelegateRingWaitTime -ne 5 -and $TeamDelegateRingWaitTime -ne 10"
										Write-Host "IMPORTANT INFO: The Wait Time value is $TeamDelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$TeamDelegateRingWaitTime = "10"
									}
									$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours"
								}
								else #All of the time
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
									$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									
									$TeamDelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									if($TeamDelegateRingWaitTime -ne "0" -and $TeamDelegateRingWaitTime -ne "5" -and $TeamDelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$TeamDelegateRingWaitTime -ne 0 -and $TeamDelegateRingWaitTime -ne 5 -and $TeamDelegateRingWaitTime -ne 10"
										Write-Host "IMPORTANT INFO: The Wait Time value is $TeamDelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$TeamDelegateRingWaitTime = "10"
									}
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime"
								}
							}
							elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail)
							{
								Write-Verbose "SimRing UpdatedUserForwardSettings UnansweredToVoicemail = True"
								if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
									$TeamDelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									if($TeamDelegateRingWaitTime -ne "0" -and $TeamDelegateRingWaitTime -ne "5" -and $TeamDelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$TeamDelegateRingWaitTime -ne 0 -and $TeamDelegateRingWaitTime -ne 5 -and $TeamDelegateRingWaitTime -ne 10"
										Write-Host "IMPORTANT INFO: The Wait Time value is $TeamDelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$TeamDelegateRingWaitTime = "10"
									}
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredToVoicemail -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredToVoicemail -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours"
								}
								else
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
									$TeamDelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									if($TeamDelegateRingWaitTime -ne "0" -and $TeamDelegateRingWaitTime -ne "5" -and $TeamDelegateRingWaitTime -ne "10")
									{	
										Write-Verbose "$TeamDelegateRingWaitTime -ne 0 -and $TeamDelegateRingWaitTime -ne 5 -and $TeamDelegateRingWaitTime -ne 10"
										Write-Host "IMPORTANT INFO: The Wait Time value is $TeamDelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$TeamDelegateRingWaitTime = "10"
									}
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredToVoicemail -UnansweredWaitTime $WaitTime" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredToVoicemail -UnansweredWaitTime $WaitTime"
								}

							}
							else #UnansweredCallForwardDestination = None
							{
								Write-Verbose "UnansweredCallForwardDestination = None / Voice mail"
								if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours) #Active work hours
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
									$TeamDelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									if($TeamDelegateRingWaitTime -ne "0" -and $TeamDelegateRingWaitTime -ne "5" -and $TeamDelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$TeamDelegateRingWaitTime -ne 0 -and $TeamDelegateRingWaitTime -ne 5 -and $TeamDelegateRingWaitTime -ne 10"
										Write-Host "IMPORTANT INFO: The Wait Time value is $TeamDelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$TeamDelegateRingWaitTime = "10"
									}
									#$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours"
								}
								else #All of the time
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False / All the time"
									#$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									
									$TeamDelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									if($TeamDelegateRingWaitTime -ne "0" -and $TeamDelegateRingWaitTime -ne "5" -and $TeamDelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$TeamDelegateRingWaitTime -ne 0 -and $TeamDelegateRingWaitTime -ne 5 -and $TeamDelegateRingWaitTime -ne 10"
										Write-Host "IMPORTANT INFO: The Wait Time value is $TeamDelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$TeamDelegateRingWaitTime = "10"
									}
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredWaitTime $WaitTime" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Team`" -TeamDelegateRingWaitTime $TeamDelegateRingWaitTime -UnansweredWaitTime $WaitTime"
								}
							}
							
						}
						else
						{
							Write-Host "There are no Team members assigned. So User Forwarding Settings were not updated." -foreground "red"
							[System.Windows.Forms.MessageBox]::Show("There are no Team members assigned. So User Forwarding Settings were not updated.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
						}
					}
					elseif($SimRingComboBox.SelectedItem -eq "My Delegates") #DELEGATE SIMRING
					{
						Write-Verbose "SimRing to My Delegates"
						if($UpdatedDelegates -gt 0)
						{
							Write-Verbose "UpdatedDelegates greater than 0"
							Write-Host "INFO: SimRing to My Delegates" -foreground "yellow"
							
							Write-Verbose "UpdatedUserForwardSettings UnansweredToVoicemail = False / Unanswered to Number or None"
							Write-Verbose "UnansweredCallForwardDestination = $($Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination)"
							if($Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne "" -and $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne $null) #THERE IS A FORWARD LOCATION SET
							{
								Write-Verbose "UpdatedUserForwardSettings UnansweredCallForwardDestination != null"
								if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
									$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									
									$DelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									if($DelegateRingWaitTime -ne "0" -and $DelegateRingWaitTime -ne "5" -and $DelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$DelegateRingWaitTime -ne 0 -and $DelegateRingWaitTime -ne 5 -and $DelegateRingWaitTime -ne 10)"
										Write-Host "IMPORTANT INFO: The Wait Time value is $DelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$DelegateRingWaitTime = "10"
									}
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime  -SettingsActiveWorkHours" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours"
								}
								else
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
									$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									
									$DelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									if($DelegateRingWaitTime -ne "0" -and $DelegateRingWaitTime -ne "5" -and $DelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$DelegateRingWaitTime -ne 0 -and $DelegateRingWaitTime -ne 5 -and $DelegateRingWaitTime -ne 10)"
										Write-Host "IMPORTANT INFO: The Wait Time value is $DelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$DelegateRingWaitTime = "10"
									}
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime"
								}
							}
							elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail)
							{
								Write-Verbose "UpdatedUserForwardSettings UnansweredToVoicemail = True"
								if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
									$DelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									if($DelegateRingWaitTime -ne "0" -and $DelegateRingWaitTime -ne "5" -and $DelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$DelegateRingWaitTime -ne 0 -and $DelegateRingWaitTime -ne 5 -and $DelegateRingWaitTime -ne 10"
										Write-Host "IMPORTANT INFO: The Wait Time value is $DelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$DelegateRingWaitTime = "10"
									}
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredWaitTime $WaitTime -UnansweredToVoicemail -SettingsActiveWorkHours" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime `"$DelegateRingWaitTime`" -UnansweredWaitTime $WaitTime -UnansweredToVoicemail -SettingsActiveWorkHours"
								}
								else
								{
									Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
									$DelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									if($DelegateRingWaitTime -ne "0" -and $DelegateRingWaitTime -ne "5" -and $DelegateRingWaitTime -ne "10")
									{
										Write-Verbose "$DelegateRingWaitTime -ne 0 -and $DelegateRingWaitTime -ne 5 -and $DelegateRingWaitTime -ne 10" 
										Write-Host "IMPORTANT INFO: The Wait Time value is $DelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
										$DelegateRingWaitTime = "10"
									}
									Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredWaitTime $WaitTime -UnansweredToVoicemail" -foreground "green"
									Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredWaitTime $WaitTime -UnansweredToVoicemail"

								}
							}
							else #UnansweredCallForwardDestination = None 
								{
									Write-Verbose "NO FORWARD LOCATON SET - UnansweredCallForwardDestination = None / Voice mail"
									$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
									
									if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
									{
										Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
										$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
										$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
										
										$DelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
										if($DelegateRingWaitTime -ne "0" -and $DelegateRingWaitTime -ne "5" -and $DelegateRingWaitTime -ne "10")
										{
											Write-Verbose "$DelegateRingWaitTime -ne 0 -and $DelegateRingWaitTime -ne 5 -and $DelegateRingWaitTime -ne 10"
											Write-Host "IMPORTANT INFO: The Wait Time value is $DelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
											$DelegateRingWaitTime = "10"
										}
										Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours" -foreground "green"
										Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours"
									}
									else
									{
										Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
										$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
										$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
										
										$DelegateRingWaitTime = $Script:UpdatedUserForwardSettings.TeamDelegateRingWaitTime
										if($DelegateRingWaitTime -ne "0" -and $DelegateRingWaitTime -ne "5" -and $DelegateRingWaitTime -ne "10")
										{
											Write-Verbose "$DelegateRingWaitTime -ne 0 -and $DelegateRingWaitTime -ne 5 -and $DelegateRingWaitTime -ne 10"
											Write-Host "IMPORTANT INFO: The Wait Time value is $DelegateRingWaitTime which is larger than the maximum of 10 seconds which is supported by the Set-CsUserCallForwardingSettings command. The value has been changed to 10 seconds in order to avoid error." -foreground "Magenta"
											$DelegateRingWaitTime = "10"
										}
										Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredWaitTime $WaitTime" -foreground "green"
										Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"Delegates`" -TeamDelegateRingWaitTime $DelegateRingWaitTime -UnansweredWaitTime $WaitTime"
									}
								}
							
						}
						else
						{
							Write-Host "There are no Delegate members assigned. So User Forwarding Settings were not updated." -foreground "red"
							[System.Windows.Forms.MessageBox]::Show("There are no Delegate members assigned. So User Forwarding Settings were not updated.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
						}
						
					}
					else #SimRing to Number
					{
						Write-Verbose "SimRing to Number"
						if($Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne "" -and $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination -ne $null) #Unanswered to Number
						{
							Write-Verbose "UpdatedUserForwardSettings UnansweredCallForwardDestination != null"
							if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
							{
								Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
								$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
								$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
								
								$SimRingNumber = $SimRingComboBox.SelectedItem
								Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity "sip:${SipAddress}" -EnableSimulRing `"$SimRingNumber`" -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours" -foreground "green"
								Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"$SimRingNumber`" -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours"
							}
							else
							{
								Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
								$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
								$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
								
								$SimRingNumber = $SimRingComboBox.SelectedItem
								Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity "sip:${SipAddress}" -EnableSimulRing `"$SimRingNumber`" -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime" -foreground "green"
								Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"$SimRingNumber`" -UnansweredToOther `"$UnansweredLocation`" -UnansweredWaitTime $WaitTime"
							}
						}
						elseif($Script:UpdatedUserForwardSettings.UnansweredToVoicemail)
						{
							Write-Verbose "UpdatedUserForwardSettings UnansweredToVoicemail = True"
							if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
							{
								Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
								$SimRingNumber = $SimRingComboBox.SelectedItem
								$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
								Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity "sip:${SipAddress}" -EnableSimulRing $SimRingNumber -UnansweredToVoicemail -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours" -foreground "green"
								Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"$SimRingNumber`" -UnansweredToVoicemail -UnansweredWaitTime $WaitTime -SettingsActiveWorkHours"
							}
							else
							{
								Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
								$SimRingNumber = $SimRingComboBox.SelectedItem
								$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
								Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity "sip:${SipAddress}" -EnableSimulRing $SimRingNumber -UnansweredToVoicemail -UnansweredWaitTime $WaitTime" -foreground "green"
								Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"$SimRingNumber`" -UnansweredToVoicemail -UnansweredWaitTime $WaitTime"
							}
						}
						else
						{
							Write-Verbose "UpdatedUserForwardSettings UnansweredToVoicemail = False"
							if($Script:UpdatedUserForwardSettings.SettingsActiveWorkhours)
							{
								Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = True"
								$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
								$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
								
								$SimRingNumber = $SimRingComboBox.SelectedItem
								Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity "sip:${SipAddress}" -EnableSimulRing `"$SimRingNumber`" -UnansweredWaitTime `"$WaitTime`" -SettingsActiveWorkHours" -foreground "green"
								Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"$SimRingNumber`" -UnansweredWaitTime `"$WaitTime`" -SettingsActiveWorkHours"
							}
							else
							{
								Write-Verbose "UpdatedUserForwardSettings SettingsActiveWorkhours = False"
								$UnansweredLocation = RemoveSip $Script:UpdatedUserForwardSettings.UnansweredCallForwardDestination
								$WaitTime = $Script:UpdatedUserForwardSettings.UnansweredWaitTime
								
								$SimRingNumber = $SimRingComboBox.SelectedItem
								Write-Host "RUNNING: Set-CsUserCallForwardingSettings -Identity "sip:${SipAddress}" -EnableSimulRing `"$SimRingNumber`" -UnansweredWaitTime `"$WaitTime`"" -foreground "green"
								Invoke-Expression "Set-CsUserCallForwardingSettings -Identity `"sip:${SipAddress}`" -EnableSimulRing `"$SimRingNumber`" -UnansweredWaitTime `"$WaitTime`""
							}
						}
						
					}
				}
			}
			Write-Host "INFO: Finished updating settings." -foreground "yellow"
		}
	}
}




function CheckUserVoicemail([string] $User)
{
	$UserSettings = Get-CsUser -Identity $User
	$ExUmEnabled = $UserSettings.ExUmEnabled
	$HostedVoicemail = $UserSettings.HostedVoicemail
	if($HostedVoicemail -eq "" -or $HostedVoicemail -eq $null)
	{
		$HostedVoicemail = $false
	}
	Write-Host "INFO: Is On Prem Exchange UM Enabled: " $ExUmEnabled -foreground "yellow"
	Write-Host "INFO: Is Hosted UM Enabled: " $HostedVoicemail -foreground "yellow"
	$UserForwardSettings = @{"ExUmEnabled"="$ExUmEnabled";"HostedVoicemail"="$HostedVoicemail"}
	return $UserForwardSettings
}


$Script:UsersListboxState = $false
$Script:FilterButtonState = $false
$Script:FilterTextBoxState = $false
$Script:GetDefaultForwardButtonState = $false
$Script:OffRadioButtonState = $false
$Script:ForwardRadioButtonState = $false
$Script:ForwardOnComboBoxState = $false
$Script:SimRingRadioButtonState = $false
$Script:SimRingComboBoxState = $false
$Script:UnansweredCallsWillGoToLinkState = $false
$Script:TheseSettingsWillApplyLinkState = $false
$Script:TeamCallGroupLabelState = $false
$Script:DelegateGroupLabelState = $false
$Script:SetUsersListboxState = $false
$Script:SetUserForwardButtonState = $false
	
function ToolIsBusy()
{
	$Script:UsersListboxState = $UsersListbox.Enabled
	$Script:FilterButtonState = $FilterButton.Enabled
	$Script:FilterTextBoxState = $FilterTextBox.Enabled
	$Script:GetDefaultForwardButtonState = $GetDefaultForwardButton.Enabled
	$Script:OffRadioButtonState = $OffRadioButton.Enabled
	$Script:ForwardRadioButtonState = $ForwardRadioButton.Enabled
	$Script:ForwardOnComboBoxState = $ForwardOnComboBox.Enabled
	$Script:SimRingRadioButtonState = $SimRingRadioButton.Enabled
	$Script:SimRingComboBoxState = $SimRingComboBox.Enabled
	$Script:UnansweredCallsWillGoToLinkState = $UnansweredCallsWillGoToLink.Enabled
	$Script:TheseSettingsWillApplyLinkState = $TheseSettingsWillApplyLink.Enabled
	$Script:TeamCallGroupLabelState = $TeamCallGroupLabel.Enabled
	$Script:DelegateGroupLabelState = $DelegateGroupLabel.Enabled
	$Script:SetUsersListboxState = $SetUsersListbox.Enabled
	$Script:SetUserForwardButtonState = $SetUserForwardButton.Enabled
	
	$UsersListbox.Enabled = $false
	$FilterButton.Enabled = $false
	$FilterTextBox.Enabled = $false
	$GetDefaultForwardButton.Enabled = $false
	$OffRadioButton.Enabled = $false
	$ForwardRadioButton.Enabled = $false
	$ForwardOnComboBox.Enabled = $false
	$SimRingRadioButton.Enabled = $false
	$SimRingComboBox.Enabled = $false
	$UnansweredCallsWillGoToLink.Enabled = $false
	$TheseSettingsWillApplyLink.Enabled = $false
	$TeamCallGroupLabel.Enabled = $false
	$DelegateGroupLabel.Enabled = $false
	$SetUsersListbox.Enabled = $false
	$SetUserForwardButton.Enabled = $false
	
	
}

function ToolIsIdle()
{
	$UsersListbox.Enabled = $Script:UsersListboxState
	$FilterButton.Enabled = $Script:FilterButtonState
	$FilterTextBox.Enabled = $Script:FilterTextBoxState
	$GetDefaultForwardButton.Enabled = $Script:GetDefaultForwardButtonState
	$OffRadioButton.Enabled = $Script:OffRadioButtonState
	$ForwardRadioButton.Enabled = $Script:ForwardRadioButtonState
	$ForwardOnComboBox.Enabled = $Script:ForwardOnComboBoxState
	$SimRingRadioButton.Enabled = $Script:SimRingRadioButtonState
	$SimRingComboBox.Enabled = $Script:SimRingComboBoxState
	$UnansweredCallsWillGoToLink.Enabled = $Script:UnansweredCallsWillGoToLinkState 
	$TheseSettingsWillApplyLink.Enabled = $Script:TheseSettingsWillApplyLinkState 
	$TeamCallGroupLabel.Enabled = $Script:TeamCallGroupLabelState
	$DelegateGroupLabel.Enabled = $Script:DelegateGroupLabelState
	$SetUsersListbox.Enabled = $Script:SetUsersListboxState
	$SetUserForwardButton.Enabled = $Script:SetUserForwardButtonState
}


$Script:CurrentUserForwardSettings = Load-DefaultForwardSettings
$Script:UpdatedUserForwardSettings = Load-DefaultForwardSettings


$objForm.Add_Shown({
$objForm.Activate()
#Update the selected index of the User List Box
$numberOfItems = $UsersListbox.Items.count
if($numberOfItems -gt 0)
{
	Write-Host "INFO: Setting users list box selected index to 0" -foreground "yellow"
	$UsersListbox.SelectedIndex = 0
}
})

[void] $objForm.ShowDialog()	



# SIG # Begin signature block
# MIIcZgYJKoZIhvcNAQcCoIIcVzCCHFMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUoBjDgBJ60bGUB0S0rme3ac4O
# EgqggheVMIIFHjCCBAagAwIBAgIQDGWW2SJRLPvqOO0rxctZHTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE5MDIwNjAwMDAwMFoXDTIwMDIw
# NjEyMDAwMFowWzELMAkGA1UEBhMCQVUxDDAKBgNVBAgTA1ZJQzEQMA4GA1UEBxMH
# TWl0Y2hhbTEVMBMGA1UEChMMSmFtZXMgQ3Vzc2VuMRUwEwYDVQQDEwxKYW1lcyBD
# dXNzZW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDHPwqNOkuXxh8T
# 7y2cCWgLtpW30x/3rEUFnrlCv2DFgULLfZHFTd+HWhCiTUMHVESj+X8s+cmgKVWN
# bmEWPri590V6kfUmjtC+4/iKdVpvjgwrwAm6O6lHZ91y4Sn90A7eUV/EvUmGREVx
# uFk2s7jD/cYjTzm0fACQBuPz5sVjTzgFzbZMndPcptB8uEjtIF/k6BGCy7XyAMn6
# 0IncNguxGZBsS/CQQlsXlVhTnBn0QQxa7nRcpJQs/84OXjDypgjW6gVOf3hOzfXY
# rXNR54nqIh/VKFKz+PiEIW11yLW0608cI0xEE03yBOg14NGIapNBwOwSpeLMlQbH
# c9twu9BhAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32
# ZXUOWDAdBgNVHQ4EFgQU2P05tP7466o6clrA//AUqWO4b2swDgYDVR0PAQH/BAQD
# AgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Ax
# hi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNy
# bDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEEeDB2
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYBBQUH
# MAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1
# cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEB
# CwUAA4IBAQCdaeq4xJ8ISuvZmb+ojTtfPN8PDWxDIsWos6e0KJ4sX7jYR/xXiG1k
# LgI5bVrb95YQNDIfB9ZeaVDrtrhEBu8Z3z3ZQFcwAudIvDyRw8HQCe7F3vKelMem
# TccwqWw/UuWWicqYzlK4Gz8abnSYSlCT52F8RpBO+T7j0ZSMycFDvFbfgBQk51uF
# mOFZk3RZE/ixSYEXlC1mS9/h3U9o30KuvVs3IfyITok4fSC7Wl9+24qmYDYYKh8H
# 2/jRG2oneR7yNCwUAMxnZBFjFI8/fNWALqXyMkyWZOIgzewSiELGXrQwauiOUXf4
# W7AIAXkINv7dFj2bS/QR/bROZ0zA5bJVMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1
# U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQD
# ExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcN
# MjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2Vy
# dCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid
# 2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sj
# lOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjf
# DPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzL
# fnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR
# 93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckw
# EgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgw
# OqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIE
# MCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYI
# YIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQY
# MBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1a
# JLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUP
# UbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1
# UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjF
# Emifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM
# 1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhs
# RDKyZqHnGKSaZFHvMIIGajCCBVKgAwIBAgIQAwGaAjr/WLFr1tXq5hfwZjANBgkq
# hkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBB
# c3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAwWhcNMjQxMDIyMDAwMDAwWjBH
# MQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERpZ2lD
# ZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBTqZ8fZFnmfGt/a4ydVfiS457V
# WmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWRn8YUOawk6qhLLJGJzF4o9GS2
# ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRVfRiGBYxVh3lIRvfKDo2n3k5f
# 4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3vJ+P3mvBMMWSN4+v6GYeofs/s
# jAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA8bLOcEaD6dpAoVk62RUJV5lW
# MJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGjggM1MIIDMTAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCCAb8G
# A1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAoBggrBgEFBQcCARYcaHR0
# cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsGAQUFBwICMIIBVh6CAVIA
# QQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMA
# YQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4A
# YwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAA
# UwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAA
# QQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkA
# YQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIA
# YQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUA
# LjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOYspkH7R7for5XDStnAs0w
# HQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9MH0GA1UdHwR2MHQwOKA2oDSG
# Mmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEu
# Y3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcnQwDQYJKoZIhvcN
# AQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI//+x1GosMe06FxlxF82pG7xa
# FjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7easGAm6mlXIV00Lx9xsIOUGQVr
# NZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8OxwYtNiS7Dgc6aSwNOOMdgv420X
# Ewbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQNJsQOfxu19aDxxncGKBXp2JPl
# VRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNtomHpigtt7BIYvfdVVEADkitr
# wlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbNMIIFtaADAgECAhAG/fkDlgOt
# 6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0wNjExMTAwMDAwMDBa
# Fw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/JM/xNRZFcgZ/tLJz4FlnfnrUk
# FcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPsi3o2CAOrDDT+GEmC/sfHMUiA
# fB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ8DIhFonGcIj5BZd9o8dD3QLo
# Oz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNugnM/JksUkK5ZZgrEjb7Szgau
# rYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJrGGWxwXOt1/HYzx4KdFxCuGh+
# t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3owggN2MA4GA1UdDwEB/wQEAwIB
# hjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUHAwIGCCsGAQUFBwMDBggrBgEF
# BQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIBxTCCAbQGCmCGSAGG/WwAAQQw
# ggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wtY3Bz
# LXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUA
# cwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMA
# bwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYA
# IAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQA
# IAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUA
# bQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkA
# dAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAA
# aABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG
# /WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsNC5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMB0GA1UdDgQW
# BBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYun
# pyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+ybcoJKc4HbZbKa9Sz1LpMUer
# Vlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6hnKtOHisdV0XFzRyR4WUVtHr
# uzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5PsQXSDj0aqRRbpoYxYqioM+Sb
# OafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke/MV5vEwSV/5f4R68Al2o/vsH
# OE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qquAHzunEIOz5HXJ7cW7g/DvXwK
# oO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQnHcUwZ1PL1qVCCkQJjGCBDsw
# ggQ3AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNI
# QTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0ECEAxlltkiUSz76jjtK8XLWR0w
# CQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcN
# AQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
# IwYJKoZIhvcNAQkEMRYEFOPs1SMwV+qXQbPcXieULH2Uo1H2MA0GCSqGSIb3DQEB
# AQUABIIBAAFHedoNUhIIWk8RvOrRZ7+Mf7HcTmlh0wUru2bS7rzj0i7zfFRSQa1x
# yTswUCev7oAiMiR556CV+U1Nr2YZe+fwZyyVJzwiV5g2LRHIOA7RNSSejjCIAsvn
# XuKGGOStnhi6bjqsMuIWvZ+RCdZJXw1Ok/EgfUZiJF3A/3pQwVayKrQPSF4PFyeZ
# aClKoLB6J/LPLrS+VHBcuEmoWZ8Bstr+Ztc3s/a/JdU8qQ9G91x1vx8bB7FUYDG+
# VjafMXfQip0W7vFLnw8Az99fDPFQQSU6bagZop6utCkKdtMXmcGbgwveqqEdsIEA
# 8tLD8K7xxZE3yTr0tO/fHgD0bM2zqo+hggIPMIICCwYJKoZIhvcNAQkGMYIB/DCC
# AfgCAQEwdjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkw
# FwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1
# cmVkIElEIENBLTECEAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE5MDgxODA0NTM1
# NlowIwYJKoZIhvcNAQkEMRYEFL6wbufmo6X9T0TH6izCiDRcZ5w+MA0GCSqGSIb3
# DQEBAQUABIIBABO99uhaBQpZvPfw1l6gUPsvbVwKtiecLVUHIajbIDjBD9aefGvX
# cTGF1ldvHgIgLbXihdfkxkMmjJsqsu6a7vJx6znsQ35M/dCL9SaBaQj/RDzeuOCw
# /TZlNcYTi6UDHphflnxMbCABLxUP/avIGZklIVHcurJhhTAU2Wvz77ss8dNyaGc/
# N7M2VRm9gl1JS3n1PLQGUeFjB5si39ONz2uaV7SgqpLsb9aAVsC9n4IHYI6hegPt
# 6AKNTuu3T8YaSXMkQ4jlpduhpW3USyC/HA/OzEGcPNhFXQBKz8fqndR/Vq/+KAJ0
# tQEilkG3kRelhegygfsWuRceTIGPbLM335Q=
# SIG # End signature block
