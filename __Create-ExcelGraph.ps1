param
(
     $DataFile =  "C:\GoogleDrive\Team\PowerShell\Diskspd\test\diskspd__ECAESQLEMTEST_H_3PAR_NTFS_64K_8k__20200404_1439.csv"
    ,$GraphFile = "C:\GoogleDrive\Team\PowerShell\Diskspd\test\diskspd__ECAESQLEMTEST_H_3PAR_NTFS_64K_8k__20200404_1439_ps.xlsx"
)

function Get-FileExtention($FilePath)
{
    $FilePath.Split('.')[1]
    return
}

# Function matching the range function in python
function Get-Range($Start, $Stop, $Step)
{ 
    $Nums = @()

    for($i = $Start; $i -lt $Stop; $i+=$Step)
    {
        $Nums += $i
    }

    $Nums
    return
}

function Create-Graph($Read_MBps_Rand, $Read_IOps_Rand, $Write_MBps_Rand, $Write_IOps_Rand,$Read_MBps_Seq, $Read_IOps_Seq, $Write_MBps_Seq, $Write_IOps_Seq,$BlockSize=0,$Drive='a',$Path='')
{
    $Nums = Get-Range -Start 0 -Stop 32 -Step 3

    #create new excel file
    $Excel = New-Object -ComObject Excel.Application 
    
    $Workbook = $Excel.Workbooks.Add()
    $Worksheet = $Workbook.Worksheets.Add()
    $Worksheet.Name = "Sheet3"
    
    # Creating Headings
    $Heading = ('Read MBps Rand','Read IOps Rand','Write MBps Rand','Write IOps Rand','Read MBps Seq','Read IOps Seq','Write MBps Seq','Write IOps Seq')
    for($i = 1; $i -le $Heading.Length; $i++)
    {
        $Worksheet.Cells.Item(1, $i) = $Heading[$i - 1]
        $Worksheet.Cells.Item(1, $i).Font.Bold = $True
    }

    $StartRow = 2

    # Creating column for $Read_MBps_Rand
    for($i = 0; $i -lt $Read_MBps_Rand.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 1) = $Read_MBps_Rand[$i]
    }

    # Creating column for $Read_IOps_Rand
    for($i = 0; $i -lt $Read_IOps_Rand.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 2) = $Read_IOps_Rand[$i]
    }

    # Creating column for $Write_MBps_Rand
    for($i = 0; $i -lt $Write_MBps_Rand.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 3) = $Write_MBps_Rand[$i]
    }

    # Creating column for $Write_IOps_Rand
    for($i = 0; $i -lt $Write_IOps_Rand.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 4) = $Write_IOps_Rand[$i]
    }

    # Creating column for $Read_MBps_Seq
    for($i = 0; $i -lt $Read_MBps_Seq.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 5) = $Read_MBps_Seq[$i]
    }

    # Creating column for $Read_IOps_Seq
    for($i = 0; $i -lt $Read_IOps_Seq.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 6) = $Read_IOps_Seq[$i]
    }

    # Creating column for $Write_MBps_Seq
    for($i = 0; $i -lt $Write_MBps_Seq.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 7) = $Write_MBps_Seq[$i]
    }

    # Creating column for $Write_IOps_Seq
    for($i = 0; $i -lt $Write_IOps_Seq.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 8) = $Write_IOps_Seq[$i]
    }
    
    # Creating column for $Nums
    for($i = 0; $i -lt $Nums.Length; $i++)
    {
        $Worksheet.Cells.Item($StartRow + $i, 9) = $Nums[$i]
    }

    # Create a new chart object. In this case an embedded chart.
    $Chart1 = $Worksheet.Shapes.AddChart().Chart
    $Chart2 = $Worksheet.Shapes.AddChart().Chart
    $Chart1.Type = [Microsoft.Office.Interop.Excel.XLChartType]::xlLine
    $Chart2.Type = [Microsoft.Office.Interop.Excel.XLChartType]::xlLine

    $Chart1.Axes(1).CategoryNames = '=Sheet1!$I$2:$I$9'
    $Chart2.Axes(1).CategoryNames = '=Sheet1!$I$2:$I$9'

    # Configure the Read MBps Rand series.
    $Chart2.SeriesCollection().NewSeries.Invoke()
    $Chart2.SeriesCollection(1).Name   = '=Sheet1!$A$1'
    $Chart2.SeriesCollection(1).Values = '=Sheet1!$A$2:$A$9'

    # Configure Write MBps Rand series.
    $Chart2.SeriesCollection().NewSeries.Invoke()
    $Chart2.SeriesCollection(2).Name   = '=Sheet1!$C$1'
    $Chart2.SeriesCollection(2).Values = '=Sheet1!$C$2:$C$9'

    # Configure Read IOps Rand series.
    $Chart1.SeriesCollection().NewSeries.Invoke()
    $Chart1.SeriesCollection(1).Name   = '=Sheet1!$B$1'
    $Chart1.SeriesCollection(1).Values = '=Sheet1!$B$2:$B$9'

    # Configure Write IOps Rand series.
    $Chart1.SeriesCollection().NewSeries.Invoke()
    $Chart1.SeriesCollection(2).Name   = '=Sheet1!$D$1'
    $Chart1.SeriesCollection(2).Values = '=Sheet1!$D$2:$D$9'

    # Configure the Read MBps Seq series.
    $Chart2.SeriesCollection().NewSeries.Invoke()
    $Chart2.SeriesCollection(3).Name   = '=Sheet1!$E$1'
    $Chart2.SeriesCollection(3).Values = '=Sheet1!$E$2:$E$9'

    # Configure Write MBps Seq series.
    $Chart2.SeriesCollection().NewSeries.Invoke()
    $Chart2.SeriesCollection(4).Name   = '=Sheet1!$G$1'
    $Chart2.SeriesCollection(4).Values = '=Sheet1!$G$2:$G$9'

    # Configure Read IOps Seq series.
    $Chart1.SeriesCollection().NewSeries.Invoke()
    $Chart1.SeriesCollection(3).Name   = '=Sheet1!$F$1'
    $Chart1.SeriesCollection(3).Values = '=Sheet1!$F$2:$F$9'

    # Configure Write IOps Seq series.
    $Chart1.SeriesCollection().NewSeries.Invoke()
    $Chart1.SeriesCollection(4).Name   = '=Sheet1!$H$1'
    $Chart1.SeriesCollection(4).Values = '=Sheet1!$H$2:$H$9'

    # Add a chart title and some axis labels.
    $Chart1.HasTitle = $true
    $Chart1.ChartTitle.Text = "IOps using $BlockSize block size on drive $Drive"
    $Chart1.Axes(1).HasTitle = $true
    $Chart1.Axes(1).AxisTitle.Text = "Seconds"
    $Chart1.Axes(2).HasTitle = $true
    $Chart1.Axes(2).AxisTitle.Text = "IOps"

    $Chart2.HasTitle = $true
    $Chart2.ChartTitle.Text = "MBps using $BlockSize block size on drive $Drive"
    $Chart2.Axes(1).HasTitle = $true
    $Chart2.Axes(1).AxisTitle.Text = "Seconds"
    $Chart2.Axes(2).HasTitle = $true
    $Chart2.Axes(2).AxisTitle.Text = "MBps"

    # Set an Excel chart style. Colors with white outline and shadow.
    $Chart1.ChartStyle = 10
    $Chart2.ChartStyle = 10

    # Insert the chart into the worksheet (with an offset).
    $Worksheet.Shapes.Item(1).Top = 5
    $Worksheet.Shapes.Item(1).Left = 500
    $Worksheet.Shapes.Item(2).Top = 250
    $Worksheet.Shapes.Item(2).Left = 500

    # Save file
    $workbook.SaveAs($GraphFile)
}

$AllData = Import-Csv -Path $DataFile
$ReadMBps_Rand = @()
$WriteMBps_Rand = @()
$ReadIOps_Rand = @()
$WriteIOps_Rand = @()
$ReadMBps_Seq = @()
$WriteMBps_Seq = @()
$ReadIOps_Seq = @()
$WriteIOps_Seq = @()
$BlockSize = $AllData[0].BlockSize
$Drive = $AllData[0].TestFilePath[0]

foreach($Data in $AllData)
{
    if([bool]::Parse($Data.IsRandom))
    {
        if($Data.ReadMBps -ne "0")
        {
            $ReadMBps_Rand += $Data.ReadMBps
        }
        if($Data.WriteMBps -ne "0")
        {
            $WriteMBps_Rand += $Data.WriteMBps
        }
        if($Data.ReadIOps -ne "0")
        {
            $ReadIOps_Rand += $Data.ReadIOps
        }
        if($Data.WriteIOps -ne "0")
        {
            $WriteIOps_Rand += $Data.WriteIOps
        }
    }
    else
    {
        if($Data.ReadMBps -ne "0")
        {
            $ReadMBps_Seq += $Data.ReadMBps
        }
        if($Data.WriteMBps -ne "0")
        {
            $WriteMBps_Seq += $Data.WriteMBps
        }
        if($Data.ReadIOps -ne "0")
        {
            $ReadIOps_Seq += $Data.ReadIOps
        }
        if($Data.WriteIOps -ne "0")
        {
            $WriteIOps_Seq += $Data.WriteIOps
        } 
    }
}

Create-Graph -Read_MBps_Rand $ReadMBps_Rand -Read_IOps_Rand $ReadIOps_Rand -Write_MBps_Rand $WriteMBps_Rand -Write_IOps_Rand $WriteIOps_Rand -Read_MBps_Seq $ReadMBps_Seq -Read_IOps_Seq $ReadMBps_Seq -Write_MBps_Seq $WriteMBps_Seq -Write_IOps_Seq $WriteIOps_Seq -BlockSize $BlockSize -Drive $Drive
#$Excel = null
