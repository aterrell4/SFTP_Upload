# Start Logging of session
$datetime = get-date -format yyyy-MM-dd-hhmm
$logpath= "C:\auto-sftp\sftp_transfer_"+$datetime+".powershell_log"
Start-Transcript -path $logpath
$DT = (Get-Date).AddDays(-1)
$FileText = "null"
$ErrorState = "0"
$FileUploadCount = 0
Write-Host "Initial Error State: $ErrorState"

# Prepare for Sending Emails
$ErrorMessageSubject = "An Error Occurred During SFTP Operations."
$htmlerrorbody = "<body>
			<h1>An Error Occurred During SFTP Upload.</h1>
			<p><strong>Time Of Failed Run:</strong> $DT</p>"
$SuccessMessageSubject = "SFTP Operations Completed without error"
$htmlSuccessbody = "<body>
			<h1>The scheduled SFTP upload operation has completed</h1>
			<p><strong>Start Time Of Run:</strong> $DT</p>"
function Email-Notification 
{
	param(
			[Parameter()]
			[string] $MessageSubject, [string] $HTMLBody, [string] $LogTime, [string] $LogPreview, [int] $FilesSent
		)
	Import-Module -Name "ExchangeOnlineManagement"
	$Modules = Get-Module
	if ("ExchangeOnlineManagement" -notin $Modules.Name)
	{
		Write-Host -BackgroundColor Black -ForegroundColor Red "ExchangeOnlineManagement Module Is Not Imported!"
	}
	# Microsoft Graph API info from Enterprise App Registration in AzureAD
	# Manage URL https://aad.portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps 
	$AppID = "FILL_IN_APP_ID"
	$SecretKey = "FILL_IN_SECRET"
	$AzTenant = "FILL_IN_AZURE_TENANT"

	# Construct URI and body needed for authentication
	$uri = "https://login.microsoftonline.com/$AzTenant/oauth2/v2.0/token"
	$Azbody = @{
		client_id     = $AppId
		scope         = "https://graph.microsoft.com/.default"
		client_secret = $SecretKey
		grant_type    = "client_credentials"
	}

	# Use WebRequest API to generate login to extract auth token
	$AZopen = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $Azbody -UseBasicParsing
	$AzToken = ($AZopen.Content | ConvertFrom-Json).access_token

	# Create Header For Subsequent Commands
	$AZHeaders = @{
				'Content-Type'  = "application\json"
				'Authorization' = "Bearer $AZToken" 
				}
			
	# Configure E-mail Message
	# Header with styles
	$htmlhead = "<html>
					<style>
						BODY{font-family: Calibri; font-size: 12pt;}
						H1{font-size: 24px;}
						H2{font-size: 20px; padding-top: 10px;}
						H3{font-size: 16px; padding-top: 8px;}
					</style>"

	$Emailbody = $htmlhead + $HTMLBody + "<p><stong>Total Number Of Files Successfully Uploaded: $FileUploadCount</strong></p>
			<h2>Please see attached log for details on files uploaded.</h2>
			<h3><p>Preview:</P></h3><p>$FileText</p></body></html>"
	# Build Message Details
	$MessageFromAddress = "FILL_IN_FROM_ADDRESS"
	$MessageRecipient = "FILL_IN_TO_ADDRESS"
	$MessageRecipient2 = "FILL_IN_TO_ADDRESS_SECONDARY"
	$FileText = Get-Content $logpath

	# Ingest Error Log and convert to Base64 and attach.
	Stop-Transcript
	$AttachmentPath = "$logpath"
	$AttachmentBase64 = [convert]::ToBase64String([system.io.file]::readallbytes($AttachmentPath))
	Start-Transcript -path $logpath -Append

	# Build the actual email data in JSON form to send via WebRequest API. 
	$MessageDetails = @{
			"URI"         = "https://graph.microsoft.com/v1.0/users/$MessageFromAddress/sendMail"
			"Headers"     = $AZHeaders
			"Method"      = "POST"
			"ContentType" = 'application/json'
			"Body" = (@{
					"message" = @{
						"subject" = $MessageSubject
						"body"    = @{
						"contentType" = 'HTML' 
						"content"     = $Emailbody }
						"attachments" = @(
											@{
												"@odata.type" = "#microsoft.graph.fileAttachment"
												"name" = "sftp_transfer_"+$datetime+".powershell_log"
												"contenttype" = "text/plain"
												"contentBytes" = $AttachmentBase64 
											} 
										)  
						"toRecipients" = @(
											@{
												"emailAddress" = @{"address" = $MessageRecipient }
											},
											@{
												"emailAddress" = @{"address" = $MessageRecipient2 }
											} 
										)
								}
					}) | ConvertTo-JSON -Depth 6
					}
	# Call Rest to send JSON data to Office365 Graph API.
	try {
			Invoke-RestMethod @MessageDetails
		}
	Catch
		{
			$streamReader = [System.IO.streamReader]::new($_.Exception.Response.GetResponseStream())
			$ErrResp = $streamReader.ReadToEnd() | ConvertFrom-JSON
			$streamReader.Close()
			$ErrResp.error.message
		}
}

# Define local paths and the time to look back to for comparing. Modify the AddHours() value to increase the search back time. 



try
{
	$localdatapaths = Get-Item "ABSOLUTEPATH","ABSOLUTEPATH_2","ABSOLUTEPATH_3" -erroraction stop
}
Catch
{
	$ErrorState = "1"
}

if ($ErrorState -ne "1")
{
Write-Host -ForegroundColor Cyan "Last run time is $DT."

# Load WinSCP .NET assembly for SFTP Interaction
Add-Type -Path ".\WinSCPnet.dll"

# Setup TLS1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Set up session options and plaintext credentials. Look for way to secure/encrypt these! 
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = "FILL_IN_SFTP_HOSTNAME"
    UserName = "FILL_IN_SFTP_USERNAME"
    Password = "FILL_IN_SFTP_PASSWORD" # Still Don't Like Harcoded Password Here. Need to store as secure object.
    SshHostKeyFingerprint = "ssh-rsa 2048 RSA_FINGERPRINT_HERE"
}

# Session Transfer Options
$transferOptions = New-Object WinSCP.TransferOptions
$transferOptions.OverwriteMode = "Overwrite"
$transferOptions.PreserveTimestamp = $False
$FileExistError = "0. OK"



# Session prototype 
$session = New-Object WinSCP.Session
# Connect with credentials above
try
{
	$session.Open($sessionOptions)

	# Define Remote Directories and arm the counter. We use -1 as the insertion point since arrays contain data starting at element 0.
	# For each of the local paths above at line 8, make sure there is a remote path below. Otherwise after the array indexes beyond the last 
	# element the remaining files from the local paths will dump into the last directory specified in the array.
	$remotePath = @("REMOTEPATH","REMOTEPATH_2", "REMOTEPATH_3")
	$count = -1
	
	# Loop for checking entire directory
	foreach ($localdatapath in $localdatapaths)
	{
		$filelist = Get-ChildItem -Path $localdatapath
		$compare = $filelist.LastWriteTime
		$count = $count + 1
	
		# Loop for populating every file in the directories.  
		foreach($filecount in $filelist)
		{
			$fileinfoname = $filecount.name
			$fileinfodate = $filecount.LastWriteTime
			Write-Host -BackgroundColor Black -ForegroundColor Yellow "Evaluating $localdatapath\$fileinfoname . Last Write Time is: $fileinfodate"
			if ( $filecount.LastWriteTime -gt $DT )
			{
				Write-Host -BackgroundColor Black -ForegroundColor Green "The File $fileinfoname is newer than $DT ."
				$UploadTarget = "$($remotePath[$count])/$fileinfoname"
				Write-Host -BackgroundColor Black -ForegroundColor Green "Attempting to upload file to the following remote path: $UploadTarget"
			
				# Try statement to begin pushing the files up. 
				try
				{
					$session.RemoveFiles("$UploadTarget")
					$session.PutFiles("$localdatapath\$fileinfoname", $UploadTarget, $False, $transferOptions).Check()
				}
			
				# Error Handling
				catch
				{
					$ErrorState = "1"
					Write-Host "An Upload task error occurred." -BackgroundColor Black -ForegroundColor Cyan
					Write-Host $_.ScriptStackTrace	-BackgroundColor Black -ForegroundColor Cyan
					Write-Host $_.Exception.Message -BackgroundColor Black -ForegroundColor Cyan
					$uploadErrorStack = $_.ScriptStackTrace
					$uploadError = $_.Exception.Message
					Write-Host "Upload Error State: $ErrorState"
				}
			
				# If no upload error occurred, check for the existence and test remote file presence.
				if ($session.FileExists($UploadTarget))
				{
					Write-Host -BackgroundColor Black -ForegroundColor Green "File $UploadTarget exists. Upload Successful."
					$FileUploadCount = $FileUploadCount + 1
					Write-Host -BackgroundColor Black -ForegroundColor Green "Uploaded File Iteration: $FileUploadCount"
				}
				else
				{
					$ErrorState = "1"
					Write-Host -BackgroundColor Black -ForegroundColor Red "File check failed! Check if it really exists on the SFTP Server."
					$FileExistError = "Failed to find $UploadTarget on remote server!"
					Write-Host "File Presence Error State: $ErrorState"
				}
			}
		
			# Blanket statement for ignoring files that don't match the modify date requirement.
			else
			{
				Write-Host -BackgroundColor Black -ForegroundColor Magenta "File is not newer than last run. Ignoring."
			}
		}
	}
}

# Error Handling
Catch
{
	$ErrorState = "1"
	Write-Host "A Server error occurred." -BackgroundColor Black -ForegroundColor Red
	Write-Host $_.ScriptStackTrace	-BackgroundColor Black -ForegroundColor Red
	Write-Host $_.Exception.Message -BackgroundColor Black -ForegroundColor Red
	$loginErrorStack = $_.ScriptStackTrace
	$loginError = $_.Exception.Message
	Write-Host "Server Connection Error State: $ErrorState"

}

# Clean Up
finally
{
	Write-Host "========================"
	Write-Host "===  End Of Summary  ==="
	Write-Host "========================"
		
	Write-Host "================="
	Write-Host "===   Debug   ==="
	Write-Host "================="
	Write-Host "Script Final Error State: $ErrorState"
	Write-Host "Script Final File Upload Count: $FileUploadCount"
	if($ErrorState -eq "1")
	{
		$FileText = Get-Content $logpath -tail 8
		Email-Notification -MessageSubject $ErrorMessageSubject -HTMLBody $htmlerrorbody -LogTime $DT -LogPreview -$FileText
	}
	else
	{
		$FileText = Get-Content $logpath -tail 8
		
		Email-Notification -MessageSubject $SuccessMessageSubject -HTMLBody $htmlSuccessbody -LogTime $DT -LogPreview -$FileText -FilesSent $FileUploadCount
	}
}
$session.Dispose()
}
else
{
	Write-Host "A Pre-Upload Error Occurred." -BackgroundColor Black -ForegroundColor Red
	Write-Host $_.ScriptStackTrace	-BackgroundColor Black -ForegroundColor Red
	Write-Host $_.Exception.Message -BackgroundColor Black -ForegroundColor Red
	Write-Host "Path Error State: $ErrorState"
	$FileText = Get-Content $logpath -tail 8
	Email-Notification -MessageSubject $ErrorMessageSubject -HTMLBody $htmlerrorbody -LogTime $DT -LogPreview -$FileText
}
Stop-Transcript
