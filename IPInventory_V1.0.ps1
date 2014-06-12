<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function System-Inventory
{
    Begin  {
           $datenormal = Get-Date
           $date = $datenormal.tostring("yyyyMMdd")

           #net view | Out-File '.\computers.txt' #Change to "Get-ADComputer -Filter *" in place of net view, once DCs updated; use " (for /f %a in ('net view ^| findstr/b \\\\') do @echo %a) > D:\Documents\Computers.txt " + remove // manually in meantime
           $computers = Get-Content .\computertest.txt #DHCPComputer.txt
           
           $excel = New-Object -ComObject Excel.Application
           $excel.Visible = $true
           $excelbook = $excel.Workbooks.Add()
           $excelitem = $excelbook.Worksheets.Item(1)
           $excelitem.Name = "Inventory_" + $date
           $excelbook.Title = "Inventory_" + $date
           $excelname = "Inventory_" + $date
           #$excel.ActiveWorkbook._SaveAs($excelname)

           $excelitem.Cells.Item(1,1) = "Machine Name"
           $excelitem.Cells.Item(1,2) = "IP Address"
           $excelitem.Cells.Item(1,3) = "Manufacture"
           $excelitem.Cells.Item(1,4) = "Model"
           $excelitem.Cells.Item(1,5) = "Owner"
           $excelitem.Cells.Item(1,6) = "OS Version"
           $excelitem.Cells.Item(1,7) = "OS Architecture"
           $excelitem.Cells.Item(1,8) = "CPU"
           $excelitem.Cells.Item(1,9) = "CPU Speed"
           $excelitem.Cells.Item(1,10) = "HDD"
           $excelitem.Cells.Item(1,11) = "RAM"
           $excelitem.Cells.Item(1,12) = "GPU"
           [int]$excelRow = 2
           $range = $excelItem.UsedRange 
           $range.Font.Bold = $True 
           $range.EntireColumn.AutoFit()

           $IP = ""
           $Disk = ""
           $Mem = ""
           $Memory = ""
           $GPU = ""
           $Capacity = ""
           $CapacityTotal = ""
           $CPUSpeed = ""

           }

    Process {
            foreach($computer in $computers) {
                                             $IP = ""
                                             $Disk = ""
                                             $Mem = ""
                                             $Memory = ""
                                             $GPU = ""
                                             $Capacity = ""
                                             $CapacityTotal = ""
                                             $CPUSpeed = ""
                                             Try {
                                                 #$IP = get-wmiobject win32_networkadapterconfiguration -ComputerName $computer -Filter DHCPEnabled="True" -ErrorAction stop
                                                 $OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer -ErrorAction stop
                                                 $System = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer -ErrorAction stop
                                                 $CPU = Get-WmiObject -Class Win32_Processor -ComputerName $computer -ErrorAction stop
                                                 [decimal]$CPUClock = $CPU.MaxClockSpeed
                                                 $CPUClock = $CPUClock/100
                                                 $CPUClock = [decimal]::round($CPUClock)
                                                 [string]$CPUSpeed = ($CPUClock/10) -as [string]
                                                 $CPUSpeed = $CPUSpeed + "GHz"
                                                 
                                                 $DiskList = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $computer -ErrorAction stop -Filter DriveType=3
                                                 $Device = $DiskList.DeviceID
                                                 Foreach($Device in $DiskList) {
                                                                               $Size = [decimal]::round($Device.Size/1gb)
                                                                               $Free = [decimal]::round($Device.Freespace/1gb)
                                                                               $Disk = $Disk + ($Device.DeviceID  + $Size + "GB Total - " + $Free + "GB Free, ")
                                                                               }
                                                 
                                                 $MemoryList = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $computer -ErrorAction stop
                                                 $MemoryID = $MemoryList.DeviceLocator
                                                 Foreach($MemoryID in $MemoryList) {
                                                                                   $Capacity = [decimal]::round($MemoryID.Capacity/1gb)
                                                                                   $Manufacture = $MemoryID.Manufacturer
                                                                                   [decimal]$CapacityTotal = $CapacityTotal + ($MemoryID.Capacity/1gb)
                                                                                   [string]$Mem = $Mem + (" " + $Capacity + "GB " + $Manufacture + " ")
                                                                                   }
                                                 $CapacityTotal = "{0:F2}" -f $CapacityTotal
                                                 $Memory = "(" + $Mem + ") " + $CapacityTotal + " GB Total"

                                                 $VideoList = Get-WmiObject -Class Win32_VideoController -ComputerName $computer -ErrorAction stop
                                                 $VideoID = $VideoList.DeviceID
                                                 [string]$GPU = $VideoList.Caption

                                                 $excelitem.Cells.Item($excelRow,1) = $System.Name
                                                 #$excelitem.Cells.Item($excelRow,2) = $IP.IPaddress
                                                 $excelitem.Cells.Item($excelRow,3) = $System.Manufacturer
                                                 $excelitem.Cells.Item($excelRow,4) = $System.Model
                                                 $excelitem.Cells.Item($excelRow,5) = $System.PrimaryOwnerName
                                                 $excelitem.Cells.Item($excelRow,6) = $OS.Caption
                                                 $excelitem.Cells.Item($excelRow,7) = $OS.OSArchitecture
                                                 $excelitem.Cells.Item($excelRow,8) = $CPU.Name
                                                 $excelitem.Cells.Item($excelRow,9) = $CPUSpeed
                                                 $excelitem.Cells.Item($excelRow,10) = $Disk
                                                 $excelitem.Cells.Item($excelRow,11) = $Memory
                                                 $excelitem.Cells.Item($excelRow,12) = $GPU
                                                                                                  
                                                 }
                                             Catch {
                                                   Write-Host "The computer " $computer " is currently unavailable."
                                                   $computer | Out-File -FilePath ./Failed_Inventory_Request.txt -Append
                                                   }
                                             Finally {
                                                     $excelitem.Cells.Item($excelRow,2) = $computer
                                                     $IP = ""
                                                     $Disk = ""
                                                     $Mem = ""
                                                     $Memory = ""
                                                     $GPU = ""
                                                     $Capacity = ""
                                                     $CapacityTotal = ""
                                                     $CPUSpeed = ""
                                                     $excelRow = $excelRow +1
                                                     }
                                                 }
            }
    End
    {
    $range.EntireColumn.AutoFit()
    $excel.ActiveWorkbook._SaveAs($excelname)
    }
}
