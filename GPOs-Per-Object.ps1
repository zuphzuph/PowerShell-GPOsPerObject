$computers = Get-ADComputer -Filter * -Properties 'Name' | Select 'Name'

$namespace = "root\rsop\computer"
foreach ($computer in $computers){
    Write-host $computer.Name -ForegroundColor Cyan
    
    Try {
        Get-WmiObject -Namespace $namespace -Class RSOP_GPLink -Filter "AppliedOrder <> 0"  -ComputerName $computer.Name -ErrorAction Stop | Foreach-Object {
            $GPO_FILTER = $_.GPO.ToString().Replace("RSOP_GPO.","")

            #$linkOrder = $_.linkOrder 
            #$appliedOrder = $_.appliedOrder 
            $Enabled = $_.Enabled 
            #$noOverride = $_.noOverride 
            #$SourceOU = $_.SOM 
            #$somOrder = $_.somOrder
            
            if ($enabled){
                Get-WmiObject -Namespace $namespace -Class RSOP_GPO -Filter $GPO_FILTER -ComputerName $computer.Name -ErrorAction Stop | Foreach-Object { 
                    Write-host "`t" $_.Name
                }
            }
        }
    }
    Catch [System.UnauthorizedAccessException]{
        Write-host "Unauthorized Access" -ForegroundColor Red
    }
    Catch [Exception]{
        if ($_.Exception.GetType().Name -eq "COMException") {
            Write-host "`tServer unavailable" -ForegroundColor Red
        }
    }
    
    Write-host ""
}