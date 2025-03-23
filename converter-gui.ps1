[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][reflection.assembly]::Load('System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
[void][reflection.assembly]::Load('System.Runtime.WindowsRuntime, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')


[System.Windows.Forms.Application]::EnableVisualStyles()
$form1 = New-Object 'System.Windows.Forms.Form'
$statusstrip1 = New-Object 'System.Windows.Forms.StatusStrip'
$splitcontainer1 = New-Object 'System.Windows.Forms.SplitContainer'
$buttonSelectFolder = New-Object 'System.Windows.Forms.Button'
$textboxFolder = New-Object 'System.Windows.Forms.TextBox'
$buttonStart = New-Object 'System.Windows.Forms.Button'
$buttonStop = New-Object 'System.Windows.Forms.Button'
$labelInfo = New-Object 'System.Windows.Forms.Label'
$toolstripstatuslabel1 = New-Object 'System.Windows.Forms.ToolStripStatusLabel'
$toolstripprogressbar1 = New-Object 'System.Windows.Forms.ToolStripProgressBar'
$timerJob = New-Object 'System.Windows.Forms.Timer'
$folderbrowserdialog1 = New-Object 'System.Windows.Forms.FolderBrowserDialog'
$toolstripstatuslabel2 = New-Object 'System.Windows.Forms.ToolStripStatusLabel'
$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

$global:livePhotoFileNames = @()
$global:job


function Convert-LivePhotoFolder {
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
			$newfilename = "$dirname\MVIMG_${dt_str}$("{0:D2}" -f $num).jpg"
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

function Stop-ProcessPhotos {
	#TODO: Place custom script here
	Stop-Job $global:job
	$timerJob.Stop()
	$timerJob.Enabled = $false
	#$global:job | Out-Host
	#Receive-Job $global:job | Out-Host
	$buttonStop.Enabled = $false
}

function Update-Progress {
	if ($global:job.ChildJobs.Length -gt 0 -and $global:job.ChildJobs[0].Progress.Count -gt 0) {
		$jobProgressHistory = $global:job.ChildJobs[0].Progress;
		$latestProgress = $jobProgressHistory[$jobProgressHistory.Count - 1];
		$latestPercentComplete = $latestProgress | Select-Object -ExpandProperty CurrentOperation;
		$latestActivity = $latestProgress | Select-Object -ExpandProperty Activity;
		$latestStatus = $latestProgress | Select-Object -ExpandProperty StatusDescription;
		
		$toolstripprogressbar1.Value = [int]$latestPercentComplete
		$toolstripstatuslabel1.Text = "已转换: $($toolstripprogressbar1.Value)/$($toolstripprogressbar1.Maximum)"
		$toolstripstatuslabel2.Text = $latestStatus
		if ($toolstripprogressbar1.Value -ge $toolstripprogressbar1.Maximum) {
			Stop-ProcessPhotos
			$toolstripstatuslabel2.Text = "全部转换完毕"
		}
	}
}



$form1_Load = {
	$toolstripstatuslabel2.Anchor = 'Right'
	$toolstripstatuslabel2.Spring = $true
	$toolstripstatuslabel2.Dock = 'Right'
}


$buttonSelectFolder_Click = {
	$folderbrowserdialog1.ShowDialog()
	$textboxFolder.Text = $folderbrowserdialog1.SelectedPath
	$buttonStart.Enabled = $true
	$photos = @{ }
	$videos = @{ }
	$global:livePhotoFileNames = @()
	# get live photo pairs
	$files = Get-ChildItem -LiteralPath $textboxFolder.Text
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
	foreach ($filename in $videos.Keys) {
		if ($photos.ContainsKey($filename)) {
			# $global:livePhotoFileNames += $filename
			$global:livePhotoFileNames += New-Object PSObject -Property @{
				filename = $filename
				video    = $videos[$filename]
				photo    = $photos[$filename]
			}
		}
	}
	$labelInfo.Text = "找到的livephoto数量: " + $global:livePhotoFileNames.Length
	$toolstripstatuslabel1.Text = "已转换: 0/$(${global:livePhotoFileNames}.Length)"
	$toolstripprogressbar1.Step = 1
	$toolstripprogressbar1.Maximum = $global:livePhotoFileNames.Length
}

$buttonStart_Click = {
	Write-Host starting job
	$textboxFolder.Enabled = $false
	$buttonStop.Enabled = $true
	$buttonStart.Enabled = $false
	$toolstripprogressbar1.Maximum = $global:livePhotoFileNames.Length
	$toolstripprogressbar1.Value = 0
	$global:job = Start-Job -ScriptBlock ${Function:Convert-LivePhotoFolder} -ArgumentList ($textboxFolder.Text, $PWD)
	Start-Sleep 1
	if ($global:job.State -eq "Failed") {
		Receive-Job $job
	}
	$timerJob.Interval = 200
	$timerJob.add_Tick({ Update-Progress })
	$timerJob.Enabled = $true
	$timerJob.Start()
	$timerJob
}

$buttonStop_Click = {
	Stop-ProcessPhotos
	$toolstripstatuslabel2.Text += "..已停止"
}

$Form_StateCorrection_Load=
{
	$form1.WindowState = $InitialFormWindowState
}

$Form_Cleanup_FormClosed=
{
	try
	{
		$form1.remove_Load($form1_Load)
		$buttonSelectFolder.remove_Click($buttonSelectFolder_Click)
		$buttonStart.remove_Click($buttonStart_Click)
		$buttonStop.remove_Click($buttonStop_Click)
		$form1.remove_Load($Form_StateCorrection_Load)
		$form1.remove_FormClosed($Form_Cleanup_FormClosed)
	}
	catch { Out-Null }
	$form1.Dispose()
	$statusstrip1.Dispose()
	$splitcontainer1.Dispose()
	$buttonSelectFolder.Dispose()
	$textboxFolder.Dispose()
	$buttonStart.Dispose()
	$buttonStop.Dispose()
	$labelInfo.Dispose()
	$toolstripstatuslabel1.Dispose()
	$toolstripprogressbar1.Dispose()
	$timerJob.Dispose()
	$folderbrowserdialog1.Dispose()
	$toolstripstatuslabel2.Dispose()
}

$form1.SuspendLayout()
$statusstrip1.SuspendLayout()
$splitcontainer1.BeginInit()
$splitcontainer1.SuspendLayout()
$form1.Controls.Add($statusstrip1)
$form1.Controls.Add($splitcontainer1)
$form1.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13)
$form1.AutoScaleMode = 'Font'
$form1.ClientSize = New-Object System.Drawing.Size(794, 461)
$form1.Name = 'form1'
$form1.StartPosition = 'CenterScreen'
$form1.Text = 'Form'
$form1.add_Load($form1_Load)
$statusstrip1.AllowMerge = $False
$statusstrip1.AutoSize = $False
[void]$statusstrip1.Items.Add($toolstripstatuslabel1)
[void]$statusstrip1.Items.Add($toolstripprogressbar1)
[void]$statusstrip1.Items.Add($toolstripstatuslabel2)
$statusstrip1.LayoutStyle = 'HorizontalStackWithOverflow'
$statusstrip1.Location = New-Object System.Drawing.Point(0, 433)
$statusstrip1.Name = 'statusstrip1'
$statusstrip1.Size = New-Object System.Drawing.Size(794, 28)
$statusstrip1.TabIndex = 4
$splitcontainer1.Dock = 'Fill'
$splitcontainer1.Location = New-Object System.Drawing.Point(0, 0)
$splitcontainer1.Name = 'splitcontainer1'
[void]$splitcontainer1.Panel2.Controls.Add($labelInfo)
[void]$splitcontainer1.Panel2.Controls.Add($buttonStop)
[void]$splitcontainer1.Panel2.Controls.Add($buttonSelectFolder)
[void]$splitcontainer1.Panel2.Controls.Add($buttonStart)
[void]$splitcontainer1.Panel2.Controls.Add($textboxFolder)
$splitcontainer1.Size = New-Object System.Drawing.Size(794, 461)
$splitcontainer1.SplitterDistance = 253
$splitcontainer1.TabIndex = 3
$buttonSelectFolder.Anchor = 'Top, Right'
$buttonSelectFolder.Location = New-Object System.Drawing.Point(420, 12)
$buttonSelectFolder.Name = 'buttonSelectFolder'
$buttonSelectFolder.Size = New-Object System.Drawing.Size(105, 23)
$buttonSelectFolder.TabIndex = 0
$buttonSelectFolder.Text = '选择待转换目录'
$buttonSelectFolder.UseVisualStyleBackColor = $True
$buttonSelectFolder.add_Click($buttonSelectFolder_Click)
$textboxFolder.Anchor = 'Top, Left, Right'
$textboxFolder.Location = New-Object System.Drawing.Point(13, 14)
$textboxFolder.Name = 'textboxFolder'
$textboxFolder.Size = New-Object System.Drawing.Size(401, 20)
$textboxFolder.TabIndex = 1
$buttonStart.Anchor = 'Bottom, Left'
$buttonStart.AutoSize = $True
$buttonStart.Enabled = $False
$buttonStart.Location = New-Object System.Drawing.Point(13, 403)
$buttonStart.Name = 'buttonStart'
$buttonStart.Size = New-Object System.Drawing.Size(84, 23)
$buttonStart.TabIndex = 2
$buttonStart.Text = '开始转换'
$buttonStart.UseVisualStyleBackColor = $True
$buttonStart.add_Click($buttonStart_Click)
$buttonStop.Anchor = 'Bottom, Right'
$buttonStop.AutoSize = $True
$buttonStop.Enabled = $False
$buttonStop.Location = New-Object System.Drawing.Point(441, 403)
$buttonStop.Name = 'buttonStop'
$buttonStop.Size = New-Object System.Drawing.Size(84, 23)
$buttonStop.TabIndex = 3
$buttonStop.Text = '停止'
$buttonStop.UseVisualStyleBackColor = $True
$buttonStop.add_Click($buttonStop_Click)
$labelInfo.AutoSize = $True
$labelInfo.Location = New-Object System.Drawing.Point(13, 62)
$labelInfo.Name = 'labelInfo'
$labelInfo.Size = New-Object System.Drawing.Size(116, 13)
$labelInfo.TabIndex = 1
$labelInfo.Text = '找到的livephoto数量: '
$toolstripstatuslabel1.AutoSize = $False
$toolstripstatuslabel1.Name = 'toolstripstatuslabel1'
$toolstripstatuslabel1.Size = New-Object System.Drawing.Size(110, 23)
$toolstripstatuslabel1.Text = '已转换: 0/0'
$toolstripprogressbar1.Name = 'toolstripprogressbar1'
$toolstripprogressbar1.Size = New-Object System.Drawing.Size(100, 22)
$toolstripstatuslabel2.Margin = '5, 0, 5, 0'
$toolstripstatuslabel2.Name = 'toolstripstatuslabel2'
$toolstripstatuslabel2.RightToLeft = 'No'
$toolstripstatuslabel2.Size = New-Object System.Drawing.Size(0, 28)
$toolstripstatuslabel2.Spring = $True
$toolstripstatuslabel2.TextAlign = 'MiddleRight'
$splitcontainer1.EndInit()
$splitcontainer1.ResumeLayout()
$statusstrip1.ResumeLayout()
$form1.ResumeLayout()

$InitialFormWindowState = $form1.WindowState
$form1.add_Load($Form_StateCorrection_Load)
$form1.add_FormClosed($Form_Cleanup_FormClosed)
$form1.ShowDialog()
