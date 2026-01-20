# Video Rotation Fixer

This is a Windows PowerShell script designed to address an issue I have with video playback rotation in Plex on my nVidia Shield TV Pro.

The script uses `ffmpeg` to physically rotate video file content by reading the files rotation meta-data. So, instead of players using the rotation meta-data to determine the rotation, the video file is always "correct" and can be played in the right rotation without a dependency on the meta-data.

The script analyses the meta-data of each file in the specified source folder, and only applies rotation to those that need it.

The script can be configured to "report only", giving you a summary of what actions it will take, without changing any files. It can also be configured to either replace the original video file with the "fixed" one, or to place the "fixed" file in an alternate location.

## Requirements

### FFMpeg

You'll need to download `ffmpeg` from here: https://www.ffmpeg.org/download.html.

### Windows PowerShell

The script requires Windows PowerShell. If you get this error when trying to run the script:

> powershell cannot be loaded because running scripts is disabled on this system

Run this command in PowerShell:

`Set-ExecutionPolicy RemoteSigned`

## Configuration

You'll need to update these "Configuration" parameters in the script:

| Parameter         | Description                                                  |
| ----------------- | ------------------------------------------------------------ |
| $sourceFolderRoot | Full path to the root of the video files to process. The script will recurse through all sub-folders of this root. |
| $backupFolder     | Full path to a folder where processed files will first be backed-up. Will be created if it doesn't exist. Only used when `$BackupFiles` is `$true`. |
| $reportOnly       | If set to `$true`, the process will generate a log of activity, without actually touching any of your files. Use this first to see what files will be effected. Set to `$false` when you're ready to update your video files. |
| $backupFiles      | If set to `$true`, files that have been identified for rotation will first be backed up to this location. This requires additional disk space. If `$false`, the original files will be replaced with the "fixed" versions, once conversion has been successful. |
| $abortOnError     | If set to `$true` the entire process will halt if an error is encountered. If set to `$false`, the file that caused the error will be skipped, and the process will continue. |
| $ffmpegFolderPath | Set this to the full path of when you installed ffmpeg. This folder must contain both `ffmpeg.exe` and `ffprobe.exe` |



## Run the script

Run the script be following these steps:

1. Open the script in Notepad and configure the parameters are described above.

2. Open windows PowerShell.

3. Navigate to the folder where you've downloaded this script. e.g.

```
cd e:\downloads\
```

4. Run the script:

```
.\ApplyVideoRotation.ps1
```

5. You can review the log file that is created in the same folder as the script.

## Disclaimer

This script was written with some help from ChatGPT, and has been thoroughly tested on my own extensive library of home movies.

Use this script at your own risk. The author accepts no liability for any damages.