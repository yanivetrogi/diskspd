<# 
    https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/dn894707(v%3Dws.11)

    .SYNOPSIS 
    DiskSpd Batch - A Results Automator - by David Klee
    http://www.heraflux.com
    .VERSIONINFO 1.0.5
       -07/08/2016 - Added CPU info to CSV output for further analysis, Also fixed bug in the way XML is output
    
    .VERSIONINFO 1.0.5
       -07/08/2016 - Added CPU info to CSV output for further analysis, Also fixed bug in the way XML is output

    .DESCRIPTION
    The purpose of this script is to drive increasing load to a storage device so a performance profile under varying
    degrees of load can be created with a single portable PoSH script. 
    
    This script executes numerous DiskSpd tests and saves the output to an XML file. 
    The XML file is then loaded and the appropriate data pulled out into a CSV file for future analysis.
    
    Sample usage of this script:
    ./DiskSpdBatch.ps1 -Time 30 -DataFile "e:\diskspd\diskspdtest.dat" -DataFileSize "1000M" -BlockSize "4K" -OutPath "c:\diskspd" -SplitIO "False" -AllowIdle "False" -DiskSpdExe "c:\diskspd"

    
    .PARAMETER Time
    Set your individual test duration in seconds - req minimum 30 seconds per test

    .PARAMETER DataFile
    Workload data file path and name.

    .PARAMETER DataFileSize
    Workload data file size, in MB.

    .PARAMETER BlockSize
    Change the test block size (in KB "K" or MB "M") according to your application workload profile 

    .PARAMETER OutPath
    Path to store output to.

    .PARAMETER DiskspdExe
    DiskSpd Folder and path to diskspd.exe (but leave out the diskspd.exe or trailing \)

    .PARAMETER SplitIO
    Test permutations of %R/%W in a single test

    .PARAMETER AllowIdle
    So as not to overrun SAN controller cache, do you want to have a 20 second
    pause between tests?

    .PARAMETER EntropySize
    Manual set entropy size

    
    powershell.exe -file C:\temp\diskspd\diskspd.ps1 -time 30 -dataFile E:\temp\diskspd\diskspdtest.dat -dataFileSize 1024M -outPath C:\temp\diskspd -BlockSize 4k     -diskdpdExe C:\temp\diskspd -SplitIO N -AllowIdle Y -EntropySize 1G
    powershell.exe -file C:\temp\diskspd\diskspd.ps1 -time 30 -dataFile E:\temp\diskspd\diskspdtest.dat -dataFileSize 1024M -outPath C:\temp\diskspd -BlockSize 8k     -diskdpdExe C:\temp\diskspd -SplitIO N -AllowIdle Y -EntropySize 1G
    powershell.exe -file C:\temp\diskspd\diskspd.ps1 -time 30 -dataFile E:\temp\diskspd\diskspdtest.dat -dataFileSize 1024M -outPath C:\temp\diskspd -BlockSize 64k    -diskdpdExe C:\temp\diskspd -SplitIO N -AllowIdle Y -EntropySize 1G
    powershell.exe -file C:\temp\diskspd\diskspd.ps1 -time 30 -dataFile E:\temp\diskspd\diskspdtest.dat -dataFileSize 1024M -outPath C:\temp\diskspd -BlockSize 512k   -diskdpdExe C:\temp\diskspd -SplitIO N -AllowIdle Y -EntropySize 1G
    powershell.exe -file C:\temp\diskspd\diskspd.ps1 -time 30 -dataFile E:\temp\diskspd\diskspdtest.dat -dataFileSize 1024M -outPath C:\temp\diskspd -BlockSize 1024k  -diskdpdExe C:\temp\diskspd -SplitIO N -AllowIdle Y -EntropySize 1G
    powershell.exe -file C:\temp\diskspd\diskspd.ps1 -time 30 -dataFile E:\temp\diskspd\diskspdtest.dat -dataFileSize 1024M -outPath C:\temp\diskspd -BlockSize 2048k  -diskdpdExe C:\temp\diskspd -SplitIO N -AllowIdle Y -EntropySize 1G
    powershell.exe -file C:\temp\diskspd\diskspd.ps1 -time 30 -dataFile E:\temp\diskspd\diskspdtest.dat -dataFileSize 1024M -outPath C:\temp\diskspd -BlockSize 4096k  -diskdpdExe C:\temp\diskspd -SplitIO N -AllowIdle Y -EntropySize 1G
    powershell.exe -file C:\temp\diskspd\diskspd.ps1 -time 30 -dataFile E:\temp\diskspd\diskspdtest.dat -dataFileSize 1024M -outPath C:\temp\diskspd -BlockSize 8192k  -diskdpdExe C:\temp\diskspd -SplitIO N -AllowIdle Y -EntropySize 1G

#>

param( 
        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false, Position=0)] 
        [int]$time = 30,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false, Position=1)] 
        [string]$dataFile,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false, Position=2)] 
        [string]$dataFileSize = '1024M',

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false, Position=3)] 
        [string]$outPath = 'C:\temp\Diskspd',

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false, Position=4)] 
        [string]$BlockSize = '64K',

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false, Position=5)] 
        [string]$DiskspdExe = 'C:\temp\Diskspd',

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false, Position=6)] 
        [string]$SplitIO = "N",

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false, Position=7)] 
        [string]$AllowIdle = "Y",

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false, Position=8)] 
        [string]$EntropySize = "1G"
    )

$ExecutionPolicy = Get-ExecutionPolicy
if ($EscutionPolicy -ne 'RemoteSigned') {Set-ExecutionPolicy -ExecutionPolicy RemoteSigned; }


if ($PSVersionTable.PSVersion.Major -lt 3) 
{
    Write-host "Error: This script requires Powershell Version 3 and above to run correctly" -ForegroundColor Red;
    Return
    #[Environment]::Exit(1)
} 

#Clear-Host

$datafiledir = split-path -path $dataFile
$sArgs = @()
$outset = @()
$buckets =@()

#Testing that directories for new data files exist
foreach ($folder in $datafiledir) 
{
    if ((test-path $folder) -eq $false) 
    {
        $noDir = $true
    } 
}


# If the output folder or the test folder does not exist 
if ((test-path $outpath) -eq $false) 
{
    Write-Host "The output path $outPath doesnt exist and needs to be created..." -ForegroundColor Yellow #-BackgroundColor yellow
    New-Item -Path $outPath -ItemType Directory
    Write-Host "Created folder "$outPath -ForegroundColor Yellow
}

foreach ($folder in $datafiledir) 
{
    if ((test-path $folder) -eq $false) 
    {
        Write-Host "The test folder $folder doesnt exist and needs to be created..." -ForegroundColor Yellow; # -BackgroundColor yellow
        New-Item -Path $folder -ItemType Directory;
        Write-Host "Created folder " $folder;
    }
}        



#Building datafile args(this is to handle multiple data files in the future)
foreach ($s in $dataFile) 
{
    $sArgs += """$s""" 
}


# variables
$startDT = Get-Date
$invocation = (Get-Variable MyInvocation).Value
$directorypath = Split-Path $invocation.MyCommand.Path
$seqrandSet = @("r","s")   #random or sequential
$opsSet = @(1,2,4,8,16,32,64,128) 


# get CPU cores to determine number of threads (non-hyper-threaded)
$processors = Get-WmiObject -ComputerName localhost Win32_Processor
$Cores = 0
if ( @($processors)[0].NumberOfCores) 
{
    $Cores = @($processors).Count * @($processors)[0].NumberOfCores
} 
else 
{
    $Cores = @($processors).Count
}
$threads = $Cores


[string]$drive_letter = $dataFile.Substring(0,2).ToUpper();
$wql = "SELECT FileSystem, BlockSize,DriveLetter,Label FROM Win32_Volume WHERE DriveLetter = '$drive_letter'";
$disk_info = Get-WmiObject -Query $wql -ComputerName '.' | Select-Object DriveLetter, Label, FileSystem, BlockSize 

#timestamp output file
$filename = "diskspd__" + $env:COMPUTERNAME + "_" + $drive_letter.Substring(0,1) + "_" + $disk_info.Label + "_" + $disk_info.FileSystem + "_" + $disk_info.BlockSize/1024 + "K_" + $BlockSize + "_" + (Get-Date -format '_yyyyMMdd_HHmm') + ".xml"
$outfile = Join-Path -Path $outPath -childpath  $filename
$csvfile = $outfile -replace ".xml", ".csv"


# if opt to perform R&W testing in a single test, set steppoints here
if ($SplitIO -eq "Y") 
{
    $writeperc = @(0,10,20,30,40,50,60,70,80,90,100)
} 
else 
{
    $writeperc = @(0,100)
}

# Determine the number of tests we will be performing
$testCount = $seqrandSet.Count * $opsSet.Count * $writeperc.Count

Write-Host Number of tests to be executed: $testCount
Write-Host "Approximate time to complete test:" ([System.Math]::Ceiling($testCount * $time / 60)) "minute(s)"
$currentDir = Split-Path $myinvocation.mycommand.path

Write-Host ""
Write-Host "DiskSpd test sweep - Now beginning"


$p = New-Object System.Diagnostics.Process
$diskspdExe = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($diskspdExe)
$p.StartInfo.FileName = "$diskspdExe\diskspd.exe"
$p.StartInfo.RedirectStandardError = $true
$p.StartInfo.RedirectStandardOutput = $true
$p.StartInfo.UseShellExecute = $false
$p.StartInfo.CreateNoWindow = $true

#Write-Output "DiskSpd testing started...";
Write-Host "DiskSpd testing started...";

# DiskSpd does not write a proper XML root node
"<DiskSpdTests>" | Out-File $outfile

#Progress meter       
$counter = 1  

# Execute tests in a loop from the array values above
foreach ( $seqrand in $seqrandSet ) 
{
    foreach ( $ops in $opsSet ) 
    {
        foreach ( $writetest in $writeperc ) 
        {
            #sequential or random - ignore -r flag if sequential
            if ( $seqrand -eq "s" ) 
            {
                $rnd = "-si"
            } 
            else 
            {
                $rnd = "-r"
            }

            Write-Progress -Activity "Executing DiskSpd Tests..." -Status "Executing Test $counter of $testCount" -PercentComplete ( ($counter / ($testCount)) * 100 )                        
            
            $arguments = "-c$dataFileSize -w$writetest -t$threads -d$time -o$ops $rnd -b$BlockSize -C1 -Z$EntropySize -W1 -Rxml -L -h $sArgs";

            Write-Host $counter 'diskspd.exe' $arguments;

            $p.StartInfo.Arguments = $arguments;
            $p.Start() | Out-Null;
            $output = $p.StandardOutput.ReadToEnd();
            $error_output = $p.StandardError.ReadToEnd();

            # If there was an error exit here
            if ($error_output -like '*ERROR*')
            {
                Write-Host $error_output -ForegroundColor Yellow;
                return;
            };

            #Fix for MS bug that doesnt correctly label the Tag for Random from DiskSpd
            if ( $seqrand -eq "r" ) 
            {
                $output = $output.Replace('<RandomAccess>false</RandomAccess>','<RandomAccess>true</RandomAccess>')
            }


            $output | Out-File $outfile -Append
            $p.WaitForExit()

            $counter = $counter + 1

            if ($AllowIdle -eq "Y") 
            {
                # Pause to allow I/O idling   
                Start-Sleep -Seconds 2;
            }
        }
    }
}

# Close the XML root node
"</DiskSpdTests>" >> $outfile
#Write-Output "Done DiskSpd testing. Now creating CSV output file."
Write-Host "Done DiskSpd testing. Now creating CSV output file.";

#Export test results as .csv
[xml]$xDoc = Get-Content $outfile;

$timespans = $xDoc.DiskSpdTests.Results.timespan;

$n = 0
$resultobj = @()
$cols_sum = @('BytesCount','IOCount','ReadBytes','ReadCount','WriteBytes','WriteCount')
$cols_avg = @('AverageReadLatencyMilliseconds','ReadLatencyStdev','AverageWriteLatencyMilliseconds','WriteLatencyStdev','AverageLatencyMilliseconds','LatencyStdev')
$cols_ntile = @('0','25','50','75','90','95','99','99.9','99.99'.'99.999','99.9999','99.99999'.'99.999999','100')

foreach($ts in $timespans)
{
    $threads = $ts.Thread.Target
    $buckets = $ts.Latency.Bucket

    #create custom PSObject for output
    $outset = New-Object -TypeName PSObject
    $outset | Add-Member -MemberType NoteProperty -Name TimeSpan -Value $n
    $outset | Add-Member -MemberType NoteProperty -Name TestTimeSeconds -Value $xDoc.DiskSpdTests.Results.timespan[$n].TestTimeSeconds
    $outset | Add-Member -MemberType NoteProperty -Name RequestCount -Value $xDoc.DiskSpdTests.Results[$n].Profile.TimeSpans.TimeSpan.Targets.Target.RequestCount
    $outset | Add-Member -MemberType NoteProperty -Name WriteRatio -Value $xDoc.DiskSpdTests.Results[$n].Profile.TimeSpans.TimeSpan.Targets.Target.WriteRatio
    $outset | Add-Member -MemberType NoteProperty -Name ThreadsPerFile -Value $xDoc.DiskSpdTests.Results[$n].Profile.TimeSpans.TimeSpan.Targets.Target.ThreadsPerFile
    $outset | Add-Member -MemberType NoteProperty -Name FileSize -Value $xDoc.DiskSpdTests.Results[$n].Profile.TimeSpans.TimeSpan.Targets.Target.FileSize
    $outset | Add-Member -MemberType NoteProperty -Name IsRandom -Value $xDoc.DiskSpdTests.Results[$n].Profile.TimeSpans.TimeSpan.Targets.Target.RandomAccess
    $outset | Add-Member -MemberType NoteProperty -Name BlockSize -Value $xDoc.DiskSpdTests.Results[$n].Profile.TimeSpans.TimeSpan.Targets.Target.BlockSize
    $outset | Add-Member -MemberType NoteProperty -Name TestFilePath -Value $xDoc.DiskSpdTests.Results[$n].Profile.TimeSpans.TimeSpan.Targets.Target.Path
    

    #loop through nodes that will be summed across threads
    foreach($col in $cols_sum)
    {
        $outset | Add-Member -MemberType NoteProperty -Name $col -Value ($threads | Measure-Object $col -Sum).Sum
    }

    #generate MB/s and IOP values
    $outset | Add-Member -MemberType NoteProperty -Name MBps -Value (($outset.BytesCount / 1048576) / $outset.TestTimeSeconds)
    $outset | Add-Member -MemberType NoteProperty -Name IOps -Value ($outset.IOCount / $outset.TestTimeSeconds)
    $outset | Add-Member -MemberType NoteProperty -Name ReadMBps -Value (($outset.ReadBytes / 1048576) / $outset.TestTimeSeconds)
    $outset | Add-Member -MemberType NoteProperty -Name ReadIOps -Value ($outset.ReadCount / $outset.TestTimeSeconds)
    $outset | Add-Member -MemberType NoteProperty -Name WriteMBps -Value (($outset.WriteBytes / 1048576) / $outset.TestTimeSeconds)
    $outset | Add-Member -MemberType NoteProperty -Name WriteIOps -Value ($outset.WriteCount / $outset.TestTimeSeconds)

    #loop through nodes that will be averaged across threads
    foreach($col in $cols_avg)
    {
        if($threads.SelectNodes($col))
        {
            $outset | Add-Member -MemberType NoteProperty -Name $col -Value ($threads |Measure-Object $col -Average).Average
        } 
        else 
        {
            $outset | Add-Member -MemberType NoteProperty -Name $col -Value ""
        }
    }
    #loop through ntile buckets and extract values for the declared ntiles
    foreach($bucket in $buckets)
    {
        if($cols_ntile -contains $bucket.Percentile)
        {
            if($bucket.SelectNodes('ReadMilliseconds'))
            {
                $outset | Add-Member -MemberType NoteProperty -Name ("ReadMS_"+$bucket.Percentile) -Value $bucket.ReadMilliseconds
            }
            else
            {
                $outset | Add-Member -MemberType NoteProperty -Name ("ReadMS_"+$bucket.Percentile) -Value ""
            }

            if($bucket.SelectNodes('WriteMilliseconds'))
            {
                $outset | Add-Member -MemberType NoteProperty -Name ("WriteMS_"+$bucket.Percentile) -Value $bucket.WriteMilliseconds
            }
            else
            {
                $outset | Add-Member -MemberType NoteProperty -Name ("WriteMS_"+$bucket.Percentile) -Value ""
            }

            $outset | Add-Member -MemberType NoteProperty -Name ("TotalMS_"+$bucket.Percentile) -Value $bucket.TotalMilliseconds
        }
    }

    #Add some CPU Avg's to CSV file for analysis
    $outset | Add-Member -MemberType NoteProperty -Name AvgUsagePercent -Value $xDoc.DiskSpdTests.Results[$n].TimeSpan.CpuUtilization.Average.UsagePercent
    $outset | Add-Member -MemberType NoteProperty -Name AvgUserPercent -Value $xDoc.DiskSpdTests.Results[$n].TimeSpan.CpuUtilization.Average.UserPercent
    $outset | Add-Member -MemberType NoteProperty -Name AvgKernelPercent -Value $xDoc.DiskSpdTests.Results[$n].TimeSpan.CpuUtilization.Average.KernelPercent
    $outset | Add-Member -MemberType NoteProperty -Name AvgIdlePercent -Value $xDoc.DiskSpdTests.Results[$n].TimeSpan.CpuUtilization.Average.IdlePercent

    $resultobj += $outset
    $n++

}
$resultobj | Export-Csv -Path $csvfile -NoTypeInformation;
