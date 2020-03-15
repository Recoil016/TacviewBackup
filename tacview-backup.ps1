#Tacview file backup by M16A3NoRecoilHax#7490

#-----------------------START OF CONFIG-----------------------

#Tacview Directory
$workingdir = "${HOME}\Documents\Tacview\"

#Log Directory
$logdir = "${HOME}\Documents\Tacview\"

#Timeframe in hours. Files older than $timeframe get deleted, files newer than get uploaded.
$timeframe = 12

#FTP credentials
$ftp = 'ftp://ftp.example.asdf/backupdirectory/'
$user = 'username'
$pass = 'password'

#------------------------END OF CONFIG------------------------

$version = 'v0.6.3'

$credentials = New-Object System.Net.NetworkCredential($user,$pass)

function Get-TimeStamp {
	$date = Get-Date -Format "yyyy-MM-dd HH:mm:ss K"
	return $date
}

function Write-Log{
	param (
		[string]$text = $null,
		[string]$logfile = $stdlogfile,
		[string]$loglevel = "INFO:"
	)
	$logpath = "$logdir$logfile"
	Write-Output "$(Get-TimeStamp) $loglevel $text" | Out-File -LiteralPath $logpath -Append -Encoding ASCII
}

Write-Host "$(Get-TimeStamp) Starting cleanup and backup of Tacview files. Script version: $version"



#CLEANUP
Write-Host "`n$(Get-TimeStamp) Cleaning up Tacview files older than $timeframe hours..."
Write-Log "Starting deletion of files older than $timeframe hours. Script version: $version" -logfile "tacview-delete.log"

#Deletes files (and logs which files were deleted)
Get-ChildItem -Path "$workingdir" -Recurse -filter "*.acmi" | Where-Object {($_.LastWriteTime-lt (Get-Date).AddHours(-$timeframe))} | ForEach-Object {
	Remove-Item $_.FullName
	Write-Host "$(Get-TimeStamp) Deleted: $($_.Name)"
	Write-Log "Deleted $($_.Name)" -logfile "tacview-delete.log"
}

Write-Host "`n$(Get-TimeStamp) Cleanup finished!"
Write-Log "Cleanup finished." -logfile "tacview-delete.log"



#COMPRESSION
Write-Host "`n$(Get-TimeStamp) Compressing Tacview files newer than $timeframe hours..."
Write-Log "Starting compression of files newer than $timeframe hours. Script version: $version" -logfile "tacview-compression.log"

#loop for all detected UNCOMPRESSED Tacview files
foreach($item in (Get-ChildItem $workingdir "*.txt.acmi"))
{
    #cutting off file extension
    $namewoext = $item.Name.Substring(0,$item.Name.Length -9)

    Write-Host "`n$(Get-TimeStamp) Found uncompressed file: $namewoext.txt.acmi"
    Write-Log "Found uncompressed file: $namewoext.txt.acmi" -logfile "tacview-compression.log"

    #get size of uncompressed file
    Write-Host "$(Get-TimeStamp) Checking if file is currently being written..."
    $size = (Get-Item -Path "$workingdir\$namewoext.txt.acmi").length
    #Write-Host $size
    Start-Sleep -s 2
    $size2 = (Get-Item -Path "$workingdir\$namewoext.txt.acmi").length
    #Write-Host $size2

    #check if file is being written, abort if yes
    if ($size -ne $size2) {
        Write-Host "$(Get-TimeStamp) File is still being written. Skipping."
        Write-Log "Skipping $namewoext.txt.acmi: File is being written." -logfile "tacview-compression.log"
        break
    }

    if (!(Test-Path "$workingdir\$namewoext.zip.acmi")) {   
        Write-Host "$(Get-TimeStamp) Compressing $namewoext.txt.acmi"
        Write-Log "Compressing $namewoext.txt.acmi" -logfile "tacview-compression.log"    
        Compress-Archive -LiteralPath "$workingdir\$namewoext.txt.acmi" -CompressionLevel Optimal -Update -DestinationPath "$workingdir\$namewoext.zip"
        Rename-Item -LiteralPath "$workingdir\$namewoext.zip" -NewName "$namewoext.zip.acmi"
    } else {
        Write-Host "$(Get-TimeStamp) Skipping $namewoext.txt.acmi: Compressed Version already exists!"
        Write-Log "Skipping $namewoext.txt.acmi: Compressed Version already exists." -logfile "tacview-compression.log"
    }
}

Write-Host "`n$(Get-TimeStamp) Compression finished!"
Write-Log "Compression finished." -logfile "tacview-compression.log"



#UPLOAD
Write-Host "`n$(Get-TimeStamp) Uploading compressed Tacview files newer than $timeframe hours..."
Write-Log "Starting upload of compressed Tacview files newer than $timeframe hours. Script version: $version" -logfile "tacview-upload.log"

#variables
$success = 0
$try = 0
$upload = 0

#loop for all detected COMPRESSED Tacview files
foreach($item in (Get-ChildItem $workingdir "*.zip.acmi"))
{
    #request file size of remote file
    $url = "$ftp$item"
    $request = [Net.WebRequest]::Create($url)
    $request.Credentials = $credentials
    $request.Method = [System.Net.WebRequestMethods+Ftp]::GetFileSize
    $request.EnableSsl = $true

    Write-Host "`n$(Get-TimeStamp) Current File: $item"
    Write-Log "Current File: $item" -logfile "tacview-upload.log"

    #repeat request until successful response or 2 tries (5 seconds wait time)
    DO {
        try {
            $response = $request.GetResponse()
            $success = 1
        } catch {
            $response = $_.Exception.InnerException.Response

            #Do not put out error message if file not found
            if($response.StatusCode -ne "ActionNotTakenFileUnavailable")
            {
                if($response.StatusCode -eq "Undefined") {
                    Write-Host "$(Get-TimeStamp) ERROR: Undefined error while checking remote file size!"
                    Write-Log "Undefined error while checking remote file size." -logfile "tacview-upload.log" -loglevel "ERROR:"
                } else {
                    Write-Host "$(Get-TimeStamp) ERROR: $($response.StatusDescription)"
                    Write-Log "$($response.StatusDescription)" -logfile "tacview-upload.log" -loglevel "ERROR:"
                }
            }
            
            $success = 0
            $try++
            Start-Sleep -s 5
        }
    } While (($success -eq 0) -and ($try -le 1))

    #execute if successful response
    if ($success -eq 1) {
        #remove unnecessary information from response
        $remotesize = $response.StatusDescription
        $remotesize = $remotesize.Substring($remotesize.LastIndexOf(” “) + 1) 
        
        #convert string to integer
        $remoteint = [int]$remotesize

    #otherwise set size to zero
    } else {
        $remoteint = 0
    }
    
    #get size of local file
    $filename = $item.Name
    $localsize = (Get-Item -Path "$workingdir\$filename").length

    #Write file sizes to log
    Write-Log "Remote: $remoteint Bytes Local: $localsize Bytes" -logfile "tacview-upload.log"

    #start upload if sizes are different
    if($remoteint -ne $localsize){
        Write-Host "$(Get-TimeStamp) Attempting Upload!"
        #upload files
        try {
            $upload++ 

            $request = [System.Net.WebRequest]::Create($url)
            $request.Credentials = $credentials
            $request.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
            $request.EnableSsl = $true

            $content = [System.IO.File]::ReadAllBytes("$workingdir\$filename")
            $request.ContentLength = $content.Length
            
            $requeststream = $request.GetRequestStream()
            $requeststream.Write($content,0,$content.Length)
            $requeststream.Close()
            $requeststream.Dispose()

            Write-Host "$(Get-TimeStamp) Finished upload!"
            Write-Log "Uploaded file: $item" -logfile "tacview-upload.log"

        #log errors
        } catch {
            $upload--
            
            if($_.Exception.InnerException.Response.StatusCode -eq "Undefined") {
                Write-Host "$(Get-TimeStamp) ERROR: Undefined error when attempting to upload!"
                Write-Log "Undefined error when attempting to upload." -logfile "tacview-upload.log" -loglevel "ERROR:"
            } else {
				Write-Host "$(Get-TimeStamp) ERROR: $($_.Exception.InnerException.Response.StatusDescription)"
				Write-Log "$($_.Exception.InnerException.Response.StatusDescription)" -logfile "tacview-upload.log" -loglevel "ERROR:"
            }
        }
    #log if sizes are equal
    } else {
        Write-Host "$(Get-TimeStamp) File skipped, same version already archived!"
        Write-Log "$item skipped, same version already archived!" -logfile "tacview-upload.log"
    }
}

Write-Host "`n$(Get-TimeStamp) Upload finished. $upload file(s) uploaded!"
Write-Log "Upload finished. $upload file(s) uploaded." -logfile "tacview-upload.log"
Start-Sleep 2
#Pause