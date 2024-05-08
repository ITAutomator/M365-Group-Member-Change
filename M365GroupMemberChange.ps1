#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####
# Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
# Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
if (!(Test-Path $scriptCSV))
{
    ######### Template
    "GroupNameOrEmail,UserNameOrEmail,AddRemove" | Add-Content $scriptCSV
    "mygroup@contoso.com,user1@contoso.com,Add" | Add-Content $scriptCSV
    ######### 
	$ErrOut=201; Write-Host "Err $ErrOut : Couldn't find '$(Split-Path $scriptCSV -leaf)'. Template CSV created. Edit CSV and run again.";Pause; Exit($ErrOut)
}
# ----------Fill $entries with contents of file or something
$entries=@(import-csv $scriptCSV)
$entriescount = $entries.count
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in M365"
Write-Host ""
Write-Host ""
Write-Host "CSV: $(Split-Path $scriptCSV -leaf) ($($entriescount) entries)"
$entries | Format-Table
Write-Host "-----------------------------------------------------------------------------"
PressEnterToContinue
$no_errors = $true
$error_txt = ""
$results = @()
# Load required modules
$module= "Microsoft.Graph.Groups" ; Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module -checkver $false; Write-Host $lm_result
if ($lm_result.startswith("ERR")) {
    Write-Host "ERR: Load-Module $($module) failed. Suggestion: Open PowerShell $($PSVersionTable.PSVersion.Major) as admin and run: Install-Module $($module)";PressEnterToContinue; Return $false
}
# Connect
$myscopes=@()
$myscopes+="User.ReadWrite.All"
$myscopes+="GroupMember.ReadWrite.All"
$myscopes+="Group.ReadWrite.All"
$connected_ok = ConnectMgGraph -myscopes $myscopes
$domain_mg = Get-MgDomain -ErrorAction Ignore| Where-object IsDefault -eq $True | Select-object -ExpandProperty Id
# connect failed
if (-not ($connected_ok))
{
    Write-Host "Connection failed."
}
else
{ # connect OK
    Write-Host "----- Group Summary"
    # groups
    $i=0
    $grps = @()
    $mygroups = @($entries.GroupNameOrEmail | Sort-Object | Select-Object -Unique) # unique group names
    foreach ($g in $mygroups)
    { # each g
        $i++
        write-host "----- $($i) of $($mygroups.count): " -NoNewline
        $groupinfo = GroupInfo $g
        if (-not $groupinfo.id) 
        { # group bad
            Write-Host "'$($g)' Group not found ERR"  -ForegroundColor Red
        } # group bad
        else
        { # group ok
            Write-Host "$($groupinfo.Name) <$($groupinfo.Mail)> " -NoNewline
            Write-Host "[$($groupinfo.Type)]" -NoNewline -ForegroundColor Yellow
            Write-Host " OK"  -ForegroundColor Green
            $grp_obj=[pscustomobject][ordered]@{
                Id                       = $groupinfo.Id
                Type                     = $groupinfo.Type
                Name                     = $groupinfo.Name
                Mail                     = $groupinfo.Mail
                MailEnabled              = $groupinfo.MailEnabled
                MembershipType           = $groupinfo.MembershipType
            }
            #### Add to results
            $grps += $grp_obj
        } # group ok
    } # each g
    # 
    Write-Host "----- Users"
    # entries
    $processed=0
    $choiceLoop=0
    $i=0
    $conn_exch=$false
    foreach ($x in $entries)
    { # each entry
        $i++
        write-host "-----" $i of $entriescount $x
        if ($choiceLoop -ne 1)
        { # Process all not selected yet, Ask
            $choices = @("&Yes","Yes to &All","&No","No and E&xit") 
            $choiceLoop = AskforChoice -Message "Process entry $($i)?" -Choices $choices -DefaultChoice 1
        } # Process all not selected yet, Ask
        if (($choiceLoop -eq 0) -or ($choiceLoop -eq 1))
        { # Process
            $processed++
            #######
            ####### Start code for object $x
            #region Object X
            <#                 
            Mg-graph Cannot Update a mail-enabled security groups and or distribution list.
            Use exch for 2, use mg for the other 2
            https://learn.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0&tabs=http
            ###################################################################################################
            # Group Type            Module Add                         Remove
            # ##################### ###### ########################### ########################################
            # Microsoft 365         mg     New-MgGroupMember           Remove-MgGroupMemberDirectoryObjectByRef
            # Security              mg     New-MgGroupMember           Remove-MgGroupMemberDirectoryObjectByRef
            # Mail-enabled security exch   Add-DistributionGroupMember Remove-DistributionGroupMember
            # Distribution          exch   Add-DistributionGroupMember Remove-DistributionGroupMember
            ###################################################################################################
            #>
            $UserNameOrEmail = $x.UserNameOrEmail
            $user = Get-MgUser -Filter "(mail eq '$($UserNameOrEmail)') or (displayname eq '$($UserNameOrEmail)')"
            if (-not $user)
            { # user bad
                Write-Host "User not found: $($x.UserNameOrEmail) ERR"  -ForegroundColor Red
            } # user bad
            else
            { # user ok
                $group = $grps | Where-Object {($_.Name -eq $x.GroupNameOrEmail) -or ($_.Mail -eq $x.GroupNameOrEmail)}
                if (-not $group) 
                { # group bad
                    Write-Host "Group not found: $($x.GroupNameOrEmail) ERR"  -ForegroundColor Red
                } # group bad
                else
                { # group ok
                    $groupinfo = $group.Type
                    $isMember = Get-MgGroupMember -GroupId $group.Id | Where-Object { $_.Id -eq $user.Id }
                    If ($x.AddRemove -eq "Add")
                    { # Add
                        if ($isMember) {
                            Write-Host "User already in [$($groupinfo)] group. OK" -ForegroundColor Yellow
                        } # is a member
                        else
                        { # not a member
                            if ($group.Type -in ("Distribution","Mail-enabled security")) {
                                if (-not $conn_exch)
                                {
                                    $conn_exch = ConnectExchangeOnline -domain $domain_mg
                                }
                                Add-DistributionGroupMember -Identity $group.Id -Member $user.id
                            }
                            else { # Mg
                                New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.id
                            }
                            Write-Host "User added to [$($groupinfo)] group. OK" -ForegroundColor Green
                        } # not a member
                    } # Add
                    Elseif ($x.AddRemove -eq "Remove")
                    { # Remove
                        if (-not $isMember) {
                            Write-Host "User already removed from [$($groupinfo)] group. OK" -ForegroundColor Yellow
                        } # not a member
                        else
                        { # is a member
                            if ($group.Type -in ("Distribution","Mail-enabled security")) {
                                if (-not $conn_exch)
                                {
                                    $conn_exch = ConnectExchangeOnline -domain $domain_mg
                                }
                                Remove-DistributionGroupMember -Identity $group.Id -Member $user.id -Confirm:$false
                            }
                            else { # Mg
                                Remove-MgGroupMemberDirectoryObjectByRef -GroupId $group.Id -DirectoryObjectId $user.Id 
                            }
                            Write-Host "User removed from [$($groupinfo)] group. OK" -ForegroundColor Green
                        } # is a member
                    } # Remove
                    Else
                    {
                        Write-Host "AddRemove column has invalid data (should be Add or Remove): $($x.AddRemove) ERR"  -ForegroundColor Red
                    }
                } # group ok
            } # user ok
            #endregion Object X
            ####### End code for object $x
            #######
        } # Process
        if ($choiceLoop -eq 2)
        {
            write-host ("Entry "+$i+" skipped.")
        }
        if ($choiceLoop -eq 3)
        {
            write-host "Aborting."
            break
        }
    } # each entry
    WriteText "------------------------------------------------------------------------------------"
    $message ="Done. $($processed) of $($entriescount) entries processed. Press [Enter] to exit."
    WriteText $message
    WriteText "------------------------------------------------------------------------------------"
	# Transcript Save
    Stop-Transcript | Out-Null
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
    $TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
    If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
    Move-Item $Transcript $TranscriptTarget -Force
    # Transcript Save
} # connect OK
PressEnterToContinue