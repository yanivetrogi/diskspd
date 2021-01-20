<#
    Execute the Python script file disksdp_graph_generator.py that accepts a csv file and outputs an xls file
    Delete the csv files that were processed
#>

# Full file name for disksdp_graph_generator.py
[string]$python_file_py  = "G:\My Drive\Team\Pyton\disksdp_graph_generator.py";

# Full file name for python.exe
[string]$python_file_exe = 'C:\Program Files\Python\Python36\python.exe';

# Path to the folder containing the csv file(s)
# Note that I had to use the -Recurse switch with Get-ChildItem due to the csv extension filter applied so the command will also return csv files from sub folders
[string]$path            = 'G:\My Drive\Team\Pyton\DXC\Reshut\ECAESQLWFS01\';

$files = Get-ChildItem $path -Include *.csv -Recurse | Where-Object {$_.PSIsContainer -eq $False};
foreach($file in $files)
{       
    try
    {     
        [string]$arguments = """" + $python_file_py + """" + " """ + $file.FullName + """";
        #Write-Host $arguments;    
    
        $Process = New-Object System.Diagnostics.Process;    
        $Process.StartInfo.FileName = $python_file_exe;
        $Process.StartInfo.UseShellExecute = $false;
        $Process.StartInfo.CreateNoWindow = $true;    
        $Process.StartInfo.Arguments = $arguments;
        
        Write-Host $Process.StartInfo.Arguments

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
        Throw;          
    }
}
