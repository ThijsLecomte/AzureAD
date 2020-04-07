<#
.SYNOPSIS
    Get Azure AD Licensing errors and send email report

.DESCRIPTION
    This scripts emails a list of all the users that currently have Azure AD Licensing errors due to AAD group licensing.
    This scripts prompts for credentials, these will be used to gather information about Azure AD Groups and will be used to send an email.
    This means this user should have permissions to request group information and permissions to send email through the address specified in the ReportSender Variable

    This script requires the MSOnline PowerShell Module
    https://www.powershellgallery.com/packages/MSOnline/1.1.183.17

    This scripts creates a log file each time the script is executed. 
    It deleted all the logs it created that are older than 30 days. This value is defined in the MaxAgeLogFiles variable.

.PARAMETER LogPath
    This defines the path of the logfile. By default: "C:\Windows\Temp\CustomScript\Get-AADLicenseErrors..txt"
    You can overwrite this path by calling this script with parameter -logPath (see examples)

.PARAMETER ReportSender
    Emailaddress through which the report will be sent.

.PARAMETER ReportRecipient
    Emailaddress of the recipient for the report

.PARAMETER SMTPserver
    SMTP servers that will be used to send the email
    Default: smtp.office365.com

.PARAMETER SMTPPort
    Port used on the SMTPserver to send email
    Default: 587

.PARAMETER SMTPSLL
    Set require SSL for SMTP yes or no
    Default: True

.EXAMPLE
    Use script to authenticate with O365
    ..\Get-AADLicenseErrors..ps1 -ReportSender 'example@contoso.com' -ReportRecipient 'Example2@contoso.com' -SMTPServer "smtp.office365.com" -SMTPPort 587 -SMTPSSL True

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\Get-AADLicenseErrors..ps1 -logPath "C:\Windows\Temp\CustomScripts\Get-AADLicenseErrors.txt"

.NOTES
    File Name  : Get-AADLicenseErrors.ps1  
    Author     : Thijs Lecomte 
    Company    : Orbid NV
#>

#region Parameters
#Define Parameter LogPath
param (
    [Parameter(Mandatory=$True,Position=0)]
    [String]$ReportSender,
    [Parameter(Mandatory=$True,Position=1)]
    [String]$ReportRecipient,
    [String]$SMTPserver = "smtp.office365.com",
    [double]$SMTPPort = 587,
    [string]$SMTPSSL = "True",
    [string]$LogPath = "C:\Windows\Temp\CustomScripts\Get-AADLicenseErrors.txt"
)
#endregion

#region variables
$MaxAgeLogFiles = 30

#region Log file creation
#Create Log file
  Try{
    #Create log file based on logPath parameter followed by current date
    $date = Get-Date -Format yyyyMMddTHHmmss
    $date = $date.replace("/","").replace(":","")
    $logpath = $logpath.insert($logpath.IndexOf(".txt")," $date")
    $logpath = $LogPath.Replace(" ","")
    New-Item -Path $LogPath -ItemType File -Force -ErrorAction Stop

    #Delete all log files older than x days (specified in $MaxAgelogFiles variable)
    $limit = (Get-Date).AddDays(-$MaxAgeLogFiles)
    Get-ChildItem -Path $logPath.substring(0,$logpath.LastIndexOf("\")) -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force
    
  } catch {
    #Throw error if creation of loge file fails
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup($_.Exception.Message,0,"Creation Of LogFile failed",0x1)
    exit
  }
#endregion

#region functions
#Define Log function
Function Write-Log {
    Param ([string]$logstring)

    $DateLog = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $WriteLine = $DateLog + "|" + $logstring
    try {
        Add-Content -Path $LogPath -Value $WriteLine -ErrorAction Stop
    } catch {
        Start-Sleep -Milliseconds 100
        Write-Log $logstring
    }
    Finally{
        Write-Host $logstring
    }
}

Function Connect-MSOL(){
    Write-Log "[INFO] - Starting Function Connect-MSO"
    Try{
        $Cred = (Get-Credential)

        $null = Connect-MsolService -Credential $cred
        Write-Log "[INFO] - Connected to MSOL service"

        return $cred
    }
    Catch{
        Write-Log "[ERROR] - Error connecting to MSOL, exiting"
        Write-Log "$($_.Exception.Message)"
        Exit
    }
    Write-Log "[INFO] - Exiting Function Connect-MSO"
}

Function Populate-ErrorMesages {
    Write-Log "[INFO] - Starting Function Populate-ErrorMesages"

    $errormessages = @()
    $errormessages += @{Name="CountViolation";Description="There aren't enough available licenses for one of the products that's specified in the group. You need to either purchase more licenses for the product or free up unused licenses from other users or groups."} 
    $errormessages += @{Name="MutuallyExclusiveViolation";Description="One of the products that's specified in the group contains a service plan that conflicts with another service plan that's already assigned to the user via a different product. Some service plans are configured in a way that they can't be assigned to the same user as another, related service plan."}
    $errormessages += @{Name="DependencyViolation";Description="One of the products that's specified in the group contains a service plan that must be enabled for another service plan, in another product, to function. This error occurs when Azure AD attempts to remove the underlying service plan. For example, this can happen when you remove the user from the group."}
    $errormessages += @{Name="ProhibitedInUsageLocationViolation";Description="Some Microsoft services aren't available in all locations because of local laws and regulations."}
    Write-Log "[INFO] - Added $($errormessages.count) errormessages"
    Write-Log "[INFO] - Ending Function Populate-ErrorMesages"

    return $errormessages
}

Function Get-ErrorGroups(){
    Write-Log "[INFO] - Starting Get-ErrorGroups"

    #find groups with license errors
    $errorgroups = Get-MSOLgroup -HasLicenseErrorsOnly $true

    Write-Log "[INFO] Found $($errorgroups.count) groups with errors"
    Write-Log "[INFO] - Ending Get-ErrorGroups"
    return $errorgroups
}

Function Get-LicenseErrors($errorgroups){
    Write-Log "[INFO] - Starting Get-LicenseErrors"
    $errors = @()
    foreach($errorgroup in $errorgroups){
        #get all user members of the group
        Write-Log "[INFO] - Checking members of $($errorgroup.Name)"
        Try{
            $groupMembers = Get-MsolGroupMember -All -GroupObjectId $errorgroup.objectid 
            Write-Log "[INFO] - Got groupmembers"
        }
        Catch{
            Write-Log "[ERROR] - Getting groupmembers"
            Write-Log "$(_.Exception.Message)"
        }
        
        Write-Log "[INFO] - Found $($groupMembers.count) members"
        foreach($member in $groupMembers){
            Try{
                #Get user object
                $user = Get-MsolUser -ObjectId $member.ObjectId 
                Write-Log "[INFO] - Got user full details $($user.UserPrincipalName)"
            }
            Catch{
                Write-Log "[ERROR] - Error getting user"
                Write-Log "$(_.Exception.Message)"
            }
            
            #Check if user has errors and if errors are because of gropu we are currently checking
            if($user.IndirectLicenseErrors -and $user.IndirectLicenseErrors.ReferencedObjectId -eq $($errorgroup.ObjectId)){
                Write-Log "[INFO] - Found user with error - $($user.UserPrincipalName)"
                $errors += @{UPN="$($user.UserPrincipalName)";Error="$($user.IndirectLicenseErrors.Error)";GroupName="$($errorgroup.DisplayName)";GroupDescription="$($errorgroup.Description)"}
            }
            
        }
    }
    Write-Log "[INFO] - Ending Get-LicenseErrors"
    return $errors
}

Function Send-Mail($errors, $Credential, $errormessages){
    Write-Log "[INFO] - Starting Send-Mail"

    #Check if there are errors
    if($errors.count){
        #Create body of mail, add all errors
        $body = "<html><body><table><h3>Licensing Errors</h3>
        <tr>
            <th align='left'>UPN Username</th>
            <th align='left'>License Error</th> 
            <th align='left'>Group Name</th>
            <th align='left'>Group Description</th>
        </tr>"
        foreach($error in $errors){
            $body += "
            <tr>
                <td>$($error.UPN)</td>
                <td>$($error.Error)</td> 
                <td>$($error.GroupName)</td>
                <td>$($error.GroupDescription)</td>
            </tr>"
        }

        $body += "</table>"

        #Populate body with error messages, but only show explanation for current errors
        foreach($message in $errormessages){
            if($errors.Error -contains $message.Name){
                $body += "
                    <h4>$($message.Name)</h4>
                    <div>$($message.description)</div>
                "
            }
        }

        $body += "</body></html>"
    }
    else{
        $body += "<html><body><h4>There are currently no licensing errors</4></body></html>"
    }

    Try{
        #Check if SSL should be used
        if($SMTPSSL){
            Send-MailMessage -To $ReportRecipient -From $ReportSender -Subject ("Azure AD Group Licensing Errors " +  (Get-Date -format d)) -Credential $Credential -Body $body -BodyAsHtml -smtpserver $SMTPServer -usessl -Port $SMTPPort 
        }
        else{
            Send-MailMessage -To $ReportRecipient -From $ReportSender -Subject ("Azure AD Group Licensing Errors " +  (Get-Date -format d)) -Credential $Credential -Body $body -BodyAsHtml -smtpserver $SMTPServer -Port $SMTPPort  
        }
        
    }
    Catch{
        Write-Log "[ERROR] - Sending message"
        Write-Log "$(_.Exception.Message)"
    }


    Write-Log "[INFO] - Ending Send-Mail"
}
#endregion

#region Operational Script
Write-Log "[INFO] - Starting script"
Try{
    $Credential = Connect-MSOL

    $groups = Get-ErrorGroups

    $errors = Get-LicenseErrors -errorgroups $groups
}
Catch{
    Write-Log "[ERROR] - Signing into Exchange Online"
    Write-Log "$($_.Exception.Message)"
}
Finally{
    #Remove all current PS Sessions
    Get-PSSession | Remove-PSSession
}

$errormessages  = Populate-ErrorMesages

Send-Mail -errors $errors -Credential $Credential -errormessages $errormessages
Write-Log "[INFO] - Stopping script"
#endregion