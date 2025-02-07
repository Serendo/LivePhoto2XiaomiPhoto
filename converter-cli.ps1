﻿function Convert-LivePhotoFolder {
	[CmdletBinding()]
	param
	(
		[Parameter(Position = 0)]
		[string]$foldername,
		[Parameter(Position = 1)]
		[string]$scriptfolder = $PWD
	)
	#Write-Progress -Activity Convert -Status "准备中" -PercentComplete 0
	$shell = New-Object -ComObject Shell.Application
	function Get-FileDate {
		[CmdletBinding()]
		Param (
			$object
		)
		
		$dir = $shell.NameSpace($object.Directory.FullName)
		$file = $dir.ParseName($object.Name)
		
		# First see if we have Date Taken, which is at index 12
		$date = Get-Date-Property-Value $dir $file 12
		
		if ($null -eq $date) {
			# If we don't have Date Taken, then find the oldest date from all date properties
			0 .. 287 | ForEach-Object {
				$name = $dir.GetDetailsof($dir.items, $_)
				
				if ($name -match '(date)|(created)') {
					
					# Only get value if date field because the GetDetailsOf call is expensive
					$tmp = Get-Date-Property-Value $dir $file $_
					if (($null -ne $tmp) -and (($null -eq $date) -or ($tmp -lt $date))) {
						$date = $tmp
					}
				}
			}
		}
		return $date
	}
	
	function Convert-To-JPEG {
		# https://github.com/DavidAnson/ConvertTo-Jpeg
		Param (
			[Parameter(
					   Mandatory = $true,
					   Position = 1,
					   ValueFromPipeline = $true,
					   ValueFromPipelineByPropertyName = $true,
					   ValueFromRemainingArguments = $true,
					   HelpMessage = "Array of image file names to convert to JPEG")]
			[Alias("FullName")]
			[String[]]$Files,
			[Parameter(
					   HelpMessage = "Fix extension of JPEG files without the .jpg extension")]
			[Switch][Alias("f")]
			$FixExtensionIfJpeg,
			[Parameter(
					   HelpMessage = "Remove existing extension of non-JPEG files before adding .jpg")]
			[Switch][Alias("r")]
			$RemoveOriginalExtension
		)
		Begin {
			# Technique for await-ing WinRT APIs: https://fleexlab.blogspot.com/2018/02/using-winrts-iasyncoperation-in.html
			Add-Type -AssemblyName System.Runtime.WindowsRuntime
			$runtimeMethods = [System.WindowsRuntimeSystemExtensions].GetMethods()
			$asTaskGeneric = ($runtimeMethods | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
			Function AwaitOperation ($WinRtTask, $ResultType) {
				$asTaskSpecific = $asTaskGeneric.MakeGenericMethod($ResultType)
				$netTask = $asTaskSpecific.Invoke($null, @($WinRtTask))
				$netTask.Wait() | Out-Null
				$netTask.Result
			}
			$asTask = ($runtimeMethods | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
			Function AwaitAction ($WinRtTask) {
				$netTask = $asTask.Invoke($null, @($WinRtTask))
				$netTask.Wait() | Out-Null
			}
			
			# Reference WinRT assemblies
			[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
			[Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime] | Out-Null
		}
		Process {
			# Summary of imaging APIs: https://docs.microsoft.com/en-us/windows/uwp/audio-video-camera/imaging
			foreach ($file in $Files) {
				Write-Host $file -NoNewline
				try {
					try {
						# Get SoftwareBitmap from input file
						$file = Resolve-Path -LiteralPath $file
						$inputFile = AwaitOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($file)) ([Windows.Storage.StorageFile])
						$inputFolder = AwaitOperation ($inputFile.GetParentAsync()) ([Windows.Storage.StorageFolder])
						$inputStream = AwaitOperation ($inputFile.OpenReadAsync()) ([Windows.Storage.Streams.IRandomAccessStreamWithContentType])
						$decoder = AwaitOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($inputStream)) ([Windows.Graphics.Imaging.BitmapDecoder])
					}
					catch {
						# Ignore non-image files
						Write-Host " [Unsupported]"
						continue
					}
					if ($decoder.DecoderInformation.CodecId -eq [Windows.Graphics.Imaging.BitmapDecoder]::JpegDecoderId) {
						$extension = $inputFile.FileType
						if ($FixExtensionIfJpeg -and ($extension -ne ".jpg") -and ($extension -ne ".jpeg")) {
							# Rename JPEG-encoded files to have ".jpg" extension
							$newName = $inputFile.Name -replace ($extension + "$"), ".jpg"
							$outputFile = AwaitOperation ($inputFolder.CreateFileAsync($newName, [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
							#AwaitAction ($inputFile.RenameAsync($newName))
							AwaitAction($inputFile.CopyAndReplaceAsync($outputFile))
							Write-Host " => $newName"
						}
						else {
							# Skip JPEG-encoded files
							Write-Host " [Already JPEG]"
						}
						continue
					}
					$bitmap = AwaitOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
					
					# Determine output file name
					# Get name of original file, including extension
					$fileName = $inputFile.Name
					if ($RemoveOriginalExtension) {
						# If removing original extension, get the original file name without the extension
						$fileName = $inputFile.DisplayName
					}
					# Add .jpg to the file name
					$outputFileName = $fileName + ".jpg"
					
					# Write SoftwareBitmap to output file
					$outputFile = AwaitOperation ($inputFolder.CreateFileAsync($outputFileName, [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
					$outputStream = AwaitOperation ($outputFile.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])
					$encoder = AwaitOperation ([Windows.Graphics.Imaging.BitmapEncoder]::CreateAsync([Windows.Graphics.Imaging.BitmapEncoder]::JpegEncoderId, $outputStream)) ([Windows.Graphics.Imaging.BitmapEncoder])
					$encoder.SetSoftwareBitmap($bitmap)
					$encoder.IsThumbnailGenerated = $true
					
					# Do it
					AwaitAction ($encoder.FlushAsync())
					Write-Host " -> $outputFileName"
				}
				catch {
					# Report full details
					throw $_.Exception.ToString()
				}
				finally {
					# Clean-up
					if ($inputStream -ne $null) { [System.IDisposable]$inputStream.Dispose() }
					if ($outputStream -ne $null) { [System.IDisposable]$outputStream.Dispose() }
				}
			}
		}
	}
	
	function Get-Date-Property-Value {
		[CmdletBinding()]
		Param (
			$dir,
			$file,
			$index
		)
		
		$value = $dir.GetDetailsof($file, $index) -replace ([char]0x200e) -replace ([char]0x200f)
		if ($value -and $value -ne '') {
			return $value
		}
		return $null
	}
	
	Set-Location $scriptfolder
	$photos = @{ }
	$videos = @{ }
	$files = Get-ChildItem $foldername
	for ($i = 0; $i -lt $files.Length; $i++) {
		$base = $files[$i].DirectoryName + '\' + $files[$i].BaseName
		$extension = $files[$i].Extension
		if ($extension.ToLower() -in @(".jpg", ".jpeg", ".heic")) {
			$photos[$base] = $files[$i]
		}
		if ($extension.ToLower() -eq ".mov") {
			$videos[$base] = $files[$i]
		}
	}
	$processed = 0
	$datetimemap = @{ }
	foreach ($filename in $videos.Keys) {
		if ($photos.ContainsKey($filename)) {
			Write-Progress -Activity Convert -Status "$filename 处理中" -CurrentOperation $processed
			$photo_file = $photos[$filename].fullname
			$photo_file_o = $photo_file
			$converted = $false
			if ($photos[$filename].Extension -eq ".heic") {
				Convert-To-JPEG "$photo_file" -RemoveOriginalExtension -FixExtensionIfJpeg
				.\exiftool\exiftool.exe -overwrite_original -tagsFromFile "$photo_file" "$filename.jpg"
				$photo_file = "$filename.jpg"
				$converted = $true
			}
			$video_file = $videos[$filename].fullname
			$len = $video_file.Length
			Write-Host Converting Photo: $photo_file  Video: $video_file
			.\ffmpeg\bin\ffmpeg.exe -i "$video_file" -vcodec copy -acodec aac -y "$video_file.mp4" 2>$null
			$len = (Get-ChildItem "$video_file.mp4" | Select-Object -First 1).Length
			$fff = "`"$photo_file`"+`"$video_file" + ".mp4`""
			$dt = Get-FileDate $(Get-ChildItem $photo_file)
			$dt_str = get-date $([DateTime]::ParseExact("$dt", "yyyy/M/d H:m", $null)) -Format "yyyyMMdd_HHmm"
			if ($datetimemap.ContainsKey($dt_str)) {
				$num = $datetimemap[$dt_str] + 1
				$datetimemap[$dt_str] = $num
			}
			else {
				$datetimemap[$dt_str] = 1
				$num = 1
			}
			$dirname = $filename.substring(0, $filename.lastIndexOf("\"))
			$newfilename = "$dirname\MVIMG${dt_str}$("{0:D2}" -f $num).jpg"
			echo "newfilename: $newfilename"
			echo $fff
			cmd /c copy /b $fff "`"$newfilename`""  | Out-Null
   			if ($converted) {
                		.\exiftool\exiftool.exe -config mi.config -Orientation=1 -MVIMG=1 -MicroVideo=1 -MicroVideoVersion=1 -MicroVideoPresentationTimestampUs=15000 "-MicroVideoOffset=${len}" -n -overwrite_original "$newfilename" |Out-Null
			} else {
				.\exiftool\exiftool.exe -config mi.config -MVIMG=1 -MicroVideo=1 -MicroVideoVersion=1 -MicroVideoPresentationTimestampUs=15000 "-MicroVideoOffset=${len}" -n -overwrite_original "$newfilename" |Out-Null
			}
   			# modify time
			$final_file = Get-ChildItem "$newfilename"
			$final_file.LastWriteTime = $dt
			
			Remove-Item "$video_file.mp4" -Force
			if ($converted) {
				Remove-Item $photo_file -Force
			}
			##Remove-Item $photo_file -Force
			##Remove-Item $video_file -Force
			$processed += 1
			Write-Progress -Activity Convert -Status "$filename 处理完毕" -CurrentOperation $processed
		}
	}
}
