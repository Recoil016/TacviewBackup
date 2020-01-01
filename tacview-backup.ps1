#Tacview file backup by M16A3NoRecoilHax#7490 2020-01-01
#ONLY VERIFIED FOR PS VERSION 5.1
$version = '2020-01-01'

#Tacview Directory
$workingdir = "${env:homepath}\Documents\Tacview\"

#Log Directory
$logdir = "${env:homepath}\Documents\Tacview\"

#Timeframe in hours. Files older than $timeframe get deleted, files newer than get uploaded.
$timeframe = 12

#FTP credentials
$ftp = 'ftp://ftp.example.asdf/backupdirectory/'
$user = 'username'
$pass = 'password'

#------------------------END OF CONFIG------------------------

$credentials = New-Object System.Net.NetworkCredential($user,$pass)

function Get-TimeStamp {

    $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss K"  
    return $date
    
}

Write-Host "$(Get-TimeStamp) Starting cleanup and backup of Tacview files. Script version: $version"



#CLEANUP
Write-Host "`n$(Get-TimeStamp) Cleaning up Tacview files older than $timeframe hours..."
Write-Output "$(Get-TimeStamp) Starting deletion of files older than $timeframe hours. Script version: $version" | Out-file "$logdir\tacview-delete.log" -append -encoding ASCII

#logs deleted files
Write-Output "$(Get-TimeStamp) Files deleted:" | Out-file "$logdir\tacview-delete.log" -append -encoding ASCII
Get-ChildItem -Path "$workingdir" -Recurse -filter "*.acmi" | Where-Object {($_.LastWriteTime-lt (Get-Date).AddHours(-$timeframe))} | Add-Content "$logdir\tacview-delete.log"

#deletes files
Get-ChildItem -Path "$workingdir" -Recurse -filter "*.acmi" | Where-Object {($_.LastWriteTime-lt (Get-Date).AddHours(-$timeframe))} | Remove-Item

Write-Host "`n$(Get-TimeStamp) Cleanup finished!"
Write-Output "$(Get-TimeStamp) Cleanup finished." | Out-file "$logdir\tacview-delete.log" -append -encoding ASCII



#COMPRESSION
Write-Host "`n$(Get-TimeStamp) Compressing Tacview files newer than $timeframe hours..."
Write-Output "$(Get-TimeStamp) Starting compression of files newer than $timeframe hours. Script version: $version" | Out-file "$logdir\tacview-compression.log" -append -encoding ASCII

#loop for all detected UNCOMPRESSED Tacview files
foreach($item in (Get-ChildItem $workingdir "*.txt.acmi"))
{
    #cutting off file extension
    $namewoext = $item.Name.Substring(0,$item.Name.Length -9)

    Write-Host "`n$(Get-TimeStamp) Found uncompressed file: $namewoext.txt.acmi"
    Write-Output "$(Get-TimeStamp) Found uncompressed file: $namewoext.txt.acmi" | Out-file "$logdir\tacview-compression.log" -append -encoding ASCII

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
        Write-Output "$(Get-TimeStamp) Skipping $namewoext.txt.acmi: File is being written." | Out-file "$logdir\tacview-compression.log" -append -encoding ASCII
        break
    }

    if (!(Test-Path "$workingdir\$namewoext.zip.acmi")) {   
        Write-Host "$(Get-TimeStamp) Compressing $namewoext.txt.acmi"
        Write-Output "$(Get-TimeStamp) Compressing $namewoext.txt.acmi" | Out-file "$logdir\tacview-compression.log" -append -encoding ASCII     
        Compress-Archive -LiteralPath "$workingdir\$namewoext.txt.acmi" -CompressionLevel Optimal -Update -DestinationPath "$workingdir\$namewoext.zip"
        Rename-Item -LiteralPath "$workingdir\$namewoext.zip" -NewName "$namewoext.zip.acmi"
    } else {
        Write-Host "$(Get-TimeStamp) Skipping $namewoext.txt.acmi: Compressed Version already exists!"
        Write-Output "$(Get-TimeStamp) Skipping $namewoext.txt.acmi: Compressed Version already exists." | Out-file "$logdir\tacview-compression.log" -append -encoding ASCII
    }
}

Write-Host "`n$(Get-TimeStamp) Compression finished!"
Write-Output "$(Get-TimeStamp) Compression finished." | Out-file "$logdir\tacview-compression.log" -append -encoding ASCII



#UPLOAD
Write-Host "`n$(Get-TimeStamp) Uploading compressed Tacview files newer than $timeframe hours..."
Write-Output "$(Get-TimeStamp) Starting upload of compressed Tacview files newer than $timeframe hours. Script version: $version" | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII

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
    Write-Output "$(Get-TimeStamp) Current File: $item" | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII

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
                    Write-Host "$(Get-TimeStamp) Error: Undefined Error while checking remote file size!"
                    Write-Output "$(Get-TimeStamp) Error: Undefined Error while checking remote file size." | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII
                } else {
                    Write-Host "$(Get-TimeStamp) Error: $($response.StatusDescription)"
                    Write-Output "$(Get-TimeStamp) Error: $($response.StatusDescription)" | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII
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
    Write-Output "$(Get-TimeStamp) Remote: $remoteint Bytes Local: $localsize Bytes" | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII

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
            Write-Output "$(Get-TimeStamp) Uploaded file: $item" | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII

        #log errors
        } catch {
            $upload--
            
            if($response.StatusCode -eq "Undefined") {
                Write-Host "$(Get-TimeStamp) Error: Undefined Error when attempting to upload!"
                Write-Output "$(Get-TimeStamp) Error: Undefined Error when attempting to upload." | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII
            } else {
				Write-Host "$(Get-TimeStamp) Error: $($_.Exception.InnerException.Response.StatusDescription)"
				Write-Output "$(Get-TimeStamp) Error: $($_.Exception.InnerException.Response.StatusDescription)" | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII
            }
        }
    #log if sizes are equal
    } else {
        Write-Host "$(Get-TimeStamp) File skipped, same version already archived!"
        Write-Output "$(Get-TimeStamp) $item skipped, same version already archived!" | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII
    }
}

Write-Host "`n$(Get-TimeStamp) Upload finished. $upload file(s) uploaded!"
Write-Output "$(Get-TimeStamp) Upload finished. $upload file(s) uploaded." | Out-file "$logdir\tacview-upload.log" -append -encoding ASCII
Start-Sleep 2
#Pause