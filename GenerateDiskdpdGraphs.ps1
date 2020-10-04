<#
    Execute the Python script file disksdp_graph_generator.py that accepts a csv file and outputs an xls file
    Delete the csv files that were processed
#>

[string]$python_file_py  = 'C:\GoogleDrive\Team\Pyton\disksdp_graph_generator.py';
[string]$python_file_exe = 'C:\Program Files\Python\Python36\python.exe';
[string]$path            = 'C:\GoogleDrive\Team\Pyton\';

#[System.Diagnostics.Process]::Start("C:\Program Files\Automated QA\TestExecute 8\Bin\TestExecute.exe", "C:\temp\TestProject1\TestProject1.pjs /run /exit /SilentMode")
#[System.Diagnostics.Process]::Start($python_file_exe, "C:\GoogleDrive\Team\Pyton\disksdp_graph_generator.py /'C:\GoogleDrive\Team\Pyton\diskspd__ECAESQLEMTEST_E_Optane_NTFS_2048K_8K__20200402_1336.csv'")
#var proc = Process.Start(@"cmd.exe ",@"/c C:\Users\user2\Desktop\XXXX.reg")

#[System.Diagnostics.Process]::Start($python_file_exe, "C:\GoogleDrive\Team\Pyton\disksdp_graph_generator.py C:\GoogleDrive\Team\Pyton\diskspd__ECAESQLEMTEST_E_Optane_NTFS_4K_8K__20200402_0259.csv")

#$files = Get-ChildItem -Recurse $path -Include *.csv | Where-Object {$_.PSIsContainer -eq $False};
$files = Get-ChildItem $path -Include *.csv | Where-Object {$_.PSIsContainer -eq $False};
foreach($file in $files)
{       
    try
    {     
        [string]$arguments = $python_file_py + " " + $file.FullName ;
        Write-Host $arguments;    
    
        $Process = New-Object System.Diagnostics.Process;    
        $Process.StartInfo.FileName = $python_file_exe;
        $Process.StartInfo.UseShellExecute = $false;
        $Process.StartInfo.CreateNoWindow = $true;    
        $Process.StartInfo.Arguments = $arguments;
        $Process.StartInfo.RedirectStandardError = $true;
        $Process.StartInfo.RedirectStandardOutput = $true;
	    $Process.Start() | Out-Null;   

        $standard_output = $Process.StandardOutput.ReadToEnd();
        $error_output = $Process.StandardError.ReadToEnd();
    
        Write-Host $standard_output -ForegroundColor Yellow;
        Write-Host $error_output -ForegroundColor Green; 
    
        #Delete the file we processed
        #$file.Delete();   
    }    
    catch [Exception] 
    {
        $exception = $_.Exception;
        Write-Host -ForegroundColor Red $exception;                  
    }
}
