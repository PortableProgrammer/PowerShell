$counter = 0;
$output = '<html><head/><body>';
$messageText = '';
$baseDir = '\\server\Media\'
$childDirs = 'TV','Kids TV','Movies'
$daysToSearch = 2

# For "daily" shows, we only want to keep the last X number of episodes
$dailyShows = 'TV\The Daily Show with Jon Stewart','TV\Last Week Tonight with John Oliver','Kids TV\Sesame Street'
$dailyShowEpisodesToKeep = 7

function Get-MediaInfo {
    param ($name)

    $returnValue = New-Object -TypeName PSObject
    $returnValue    | Add-Member -Name FoundAudio -Value $false -MemberType NoteProperty -PassThru `
                    | Add-Member -Name FoundEnglishDefaultAudio -Value $false -MemberType NoteProperty -PassThru `
                    | Add-Member -Name FoundEnglishAudio -Value $false -MemberType NoteProperty -PassThru `
                    | Add-Member -Name BestEnglishAudioTrackDetails -Value '' -MemberType NoteProperty -PassThru `
                    | Add-Member -Name BestEnglishAudioTrackID -Value '' -MemberType NoteProperty -PassThru `
                    | Add-Member -Name FoundForeignDefaultAudio -Value $false -MemberType NoteProperty -PassThru `
                    | Add-Member -Name ForeignDefaultAudioTrackID -Value '' -MemberType NoteProperty -PassThru `
                    | Add-Member -Name UnknownAudioTrackID -Value '' -MemberType NoteProperty -PassThru `
                    | Add-Member -Name FoundVideo -Value $false -MemberType NoteProperty -PassThru `
                    | Add-Member -Name VideoFormat -Value ‘’ -MemberType NoteProperty -PassThru `
                    | Add-Member -Name Log -Value $false -MemberType NotePropertya

    $unkLangArr = $null,'','und','unk','unknown'
    $okLangArr = 'en','ja','zh'
    
    $mediaInfo = (C:\Utils\MediaInfo\MediaInfo.exe --output=JSON $name).Replace('"":', '"Empty":') | ConvertFrom-Json

    $mediaInfo.media.track |? “@type” -in ‘Video’,’Audio’ |% {
        $track = $_
        switch ($_.”@type”) {
            “Video” {
                $returnValue.FoundVideo = $true
                $returnValue.VideoFormat = $track.Format
            } # Video
            “Audio” {
                $returnValue.FoundAudio = $true

                # Is this track English and default?
                if ($track.Language -eq 'en' -and $track.Default -eq 'Yes') {
                    $returnValue.FoundEnglishDefaultAudio = $true
                }

                # Is this track English?
                if ($track.Language -eq 'en') {
                    $returnValue.FoundEnglishAudio = $true
                }

                # Is this track Unknown and default?
                if ($track.Language -in $unkLangArr -and $track.Default -eq 'Yes') {
                    $returnValue.UnknownAudioTrackID = $track.ID
                }

                # Is this track foreign and default?
                if ($track.Language -notin $okLangArr -and $track.Language -notin $unkLangArr -and $track.Default -eq 'Yes') {
                    $returnValue.FoundForeignDefaultAudio = $true
                    $returnValue.ForeignDefaultAudioTrackID = $track.ID
                }
            } # Audio
        } # switch

        # If we found English tracks, but none of them are the default, find the best track to use as the default
        $returnValue.BestEnglishAudioTrackDetails = $mediaInfo.media.track |? { $_."@type" -eq 'Audio' -and $_."Language" -eq 'en' } | Sort -Descending Channels,StreamSize | Select -First 1 ID,Title,Channels,Format,StreamSize
        if ($returnValue.BestEnglishAudioTrackDetails -ne '') {
            $returnValue.BestEnglishAudioTrackID = $returnValue.BestEnglishAudioTrackDetails.ID
        }
    }

    return $returnValue
}

# Look for extraneous episodes of any of the daily shows
Write-Progress 'Removing old daily show episodes' -PercentComplete -1
$dailyShows |% {
    Write-Progress 'Removing old daily show episodes' -PercentComplete -1 -Status $_
    $messageText = "Checking for old episodes in $_"
    Write-Host $messageText
    $fileOutput = "<br /><br />$messageText"
    # Assume MKV here, even though it's possible we could have AVI or MP4 by this point in the process.
    # We might end up with a few extra episodes, but they'll get taken care of tomorrow.
    $fileList = $null
    $fileList = gci $($baseDir + $_) -Recurse -Filter '*.mkv' -ErrorAction SilentlyContinue | Sort -Descending Name | Select -Expand FullName
    if ($fileList -and ($fileList -is [Array]) -and ($fileList.Length -gt $dailyShowEpisodesToKeep)) {
        $messageText = "Found $($fileList.Length) episodes, keeping the last $dailyShowEpisodesToKeep"
        Write-Host $messageText
        $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;$messageText"
        $keepList = $fileList | Select -First $dailyShowEpisodesToKeep
        $deleteList = $fileList |? { $_ -NotIn $keepList }
        
        if ($deleteList -and ($deleteList -is [Array])) {
            $messageText = "Removing $($deleteList.Length) old episodes:"
            Write-Warning $messageText
            $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: orange'>WARNING: $messageText</span>"
            $fileOutput += ($deleteList |% { "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: orange'>==&gt;$($_.Replace($baseDir, ''))</span>" })
            Remove-Item $deleteList #-WhatIf
            $output += $fileOutput
        }
    }
}

Write-Progress 'Scanning Files' PercentComplete -1

# Get all files in the child directories
$fileList = gci $($childDirs |% { $baseDir + $_ }) -Recurse |? Attributes -ne ([System.IO.FileAttributes]::Directory)
$fileList = $fileList |? CreationTime -ge ([System.DateTime]::Today.AddDays($daysToSearch * -1))

$fileList | Sort DirectoryName,Name <#| Select -First 100#> |% {
    $fileName = $_.FullName
    $baseName = $_.Name
    $folder = $_.DirectoryName.Replace($baseDir, '')

    $currentPercent = [math]::Round($counter / $fileList.Count * 100, 0)
    Write-Progress "Scanning Files ($($currentPercent)%)" -PercentComplete $currentPercent -Status "$folder\$baseName"
    $messageText = "Checking $folder\$baseName"
    Write-Host $messageText
    $fileOutput = "<br/><br/>$messageText"

    # Get the file properties
    $info = Get-MediaInfo $fileName     

    # Now that we're done with the file, what did we find?
    if (-not $info.FoundVideo) {
        # This is not a video file. No reason to continue.
        $messageText = "Skipping non-video file"
        Write-Warning $messageText
        $fileOutput = "<br /><span style='color: orange'>WARNING: $messageText</span>"
    }
    else {
        # If this file is not an MKV, we need to convert it first
        if ([System.IO.Path]::GetExtension($fileName) -ne ".mkv") {
            $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;Converting to MKV..." 
            Write-Progress "Scanning Files ($($currentPercent)%)" -PercentComplete $currentPercent -Status "$folder\$baseName" -CurrentOperation "Converting to MKV..."
            $newFilename = "$([System.IO.Path]::GetDirectoryName($fileName))\$([System.IO.Path]::GetFileNameWithoutExtension($fileName)).mkv"
            $throwAway = C:\Utils\mkvtoolnix\mkvmerge.exe -o $newFilename $fileName
            Write-Progress "Scanning Files ($($currentPercent)%)" -PercentComplete $currentPercent -Status "$folder\$baseName" -CurrentOperation "Removing old file..."
            $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: green;'>Removing old file...</span>"
            Remove-Item $fileName
            $fileName = $newFilename
            $info = Get-MediaInfo $fileName
        }

        if (-not $info.FoundAudio) {
            # This is bad, and shouldn't happen
            Write-Warning "No audio found!"
            $fileOutput += "<br /><span style='color: red;'>ERROR: No audio found!</span>"
            $info.Log = $true
        }

        if (-not $info.FoundEnglishDefaultAudio) {
            # We didn't find a default English track. This could be ok if we also didn't find a bad track, or if there's another default track
            if ($info.FoundForeignDefaultAudio) {
                 # There's a default foreign track. Warn about it.
                 $messageText = "Foreign default audio found!"
                 Write-Warning $messageText
                 $fileOutput += "<br /><span style='color: orange;'>WARNING: $messageText</span>" 
                 $info.Log = $true

                 # Furthermore, check to see if there are *any* English tracks. If not, this is a real issue.
                 if (-not $info.FoundEnglishAudio) {
                    $messageText = "No English audio found!"
                    Write-Host "ERROR: $messageText" -ForegroundColor Red
                    $fileOutput += "<br /><span style='color: red;'>ERROR: $messageText</span>"
                    $info.Log = $true
                 }
                 else {
                    if ($info.BestEnglishAudioTrackDetails -ne '') {
                        # If there are English tracks, set the English track with the most channels as default
                        $messageText = "Fixing foreign default audio track..."
                        Write-Progress "Scanning Files ($($currentPercent)%)" -PercentComplete $currentPercent -Status "$folder\$baseName" -CurrentOperation $messageText
                        Write-Host $messageText -ForeGroundColor Green
                        $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: green;'>$messageText</span>"
                        $throwAway = C:\Utils\mkvtoolnix\mkvpropedit.exe $fileName --edit track:@$($info.ForeignDefaultAudioTrackID) --set flag-default=0
                        $throwAway = C:\Utils\mkvtoolnix\mkvpropedit.exe $fileName --edit track:@$($info.BestEnglishAudioTrackID) --set flag-default=1
                        if ($($throwAway -like 'Error:*').Length -gt 0) {
                            $messageText = "Could not set best English default audio track: '$($throwAway -like 'Error:*')'" 
                            Write-Warning "ERROR: $messageText" -ForegroundColor Red
                            $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: red;'>ERROR: $messageText)'</span>"
                            $info.Log = $true
                        }
                        else {
                            $messageText = "Set default audio to: $($info.BestEnglishAudioTrackDetails.Format), $($info.BestEnglishAudioTrackDetails.Channels)"
                            if ($messageText -notlike '*ch*') {
                                $messageText += " channels"
                            }
                            if ($info.BestEnglishAudioTrackDetails.Title -notin $null,'') {
                                $messageText += " ($($info.BestEnglishAudioTrackDetails.Title))"
                            }
                            Write-Host "    $messageText" -ForegroundColor Green
                            $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: green;'>$messageText</span>"
                            $messageText = "Fixed!"
                            Write-Host $messageText -ForegroundColor Green
                            $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: green;'>$messageText</span>"
                            $info.Log = $true
                        }
                    }
                 }
            }

            # We found an unknown default track, with no default English track, and no default foreign track. We should assume this is English. This is most prevalent on TV shows
            if ($info.UnknownAudioTrackID -ne '' -and -not $info.FoundForeignDefaultAudio) {
                Write-Progress "Scanning Files ($($currentPercent)%)" -PercentComplete $currentPercent -Status "$folder\$baseName" -CurrentOperation "Fixing unknown audio track..."
                $messageText = "Fixing unknown audio track..."
                Write-Host $messageText -ForeGroundColor Green
                $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: green;'>$messageText</span>"

                $throwAway = C:\Utils\mkvtoolnix\mkvpropedit.exe $fileName --edit track:@$($info.UnknownAudioTrackID) --set language=en
                if ($($throwAway -like 'Error:*').Length -gt 0) {
                    $messageText = "Could not set language: '$($throwAway -like 'Error:*')'" 
                    Write-Warning "ERROR: $messageText" -ForegroundColor Red
                    $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: red;'>ERROR: $messageText)'</span>"
                    $info.Log = $true
                }
                else {
                    $messageText = "Fixed!"
                    Write-Host $messageText -ForegroundColor Green
                    $fileOutput += "<br />&nbsp;&nbsp;&nbsp;&nbsp;<span style='color: green;'>$messageText</span>"
                    $info.Log = $true
                }
            }
        }
    }

    if ($info.Log -or $VerbosePreference -ne 'SilentlyContinue') {
        $output += $fileOutput
    }

    $counter++
}
Write-Progress "Scanning Files (100%)" -Completed
if ($output -notlike '*<body>') {
    $output += "<br /><br />"
}
$output += "Finished Plex maintenance ($($counter.ToString("N0")) files scanned)</body></html>"

$mailTo = 'recipient@domain.tld'
$mailFrom = 'from@domain.tld'
$mailUser = 'username'
$mailPass = 'password'

Send-MailMessage -To $mailTo -From $mailFrom -SmtpServer 'smtp.gmail.com' -UseSsl -Port 587 -Credential (New-Object PSCredential($mailUser, (ConvertTo-SecureString $mailPass -AsPlainText -Force))) -Subject $('Plex Maintenance ' + [DateTime]::Today.ToString('yyyy-MM-dd')) -BodyAsHtml -Body $($output) -Priority $(if ($foundError) { 'High' } else { 'Normal' })
