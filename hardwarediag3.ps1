#requires -Version 5.1
<#
Enterprise Windows Hardware Diagnostic Tool v2
Statistical / historical / weighted-health edition.
Non-destructive diagnostics. Run as Administrator for best coverage.
#>
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$OutputFolder='C:\Temp\HardwareDiagnostics'
$CsvFile=Join-Path $OutputFolder 'Hardware_Diagnostic_Report.csv'
$HtmlFile=Join-Path $OutputFolder 'Hardware_Diagnostic_Report.html'
$ExcelFile=Join-Path $OutputFolder 'Hardware_Diagnostic_Report.xlsx'
$HistoryFile=Join-Path $OutputFolder 'Fleet_History.csv'
$SampleCount=10
$SampleIntervalMs=1000
if(!(Test-Path $OutputFolder)){New-Item -ItemType Directory -Path $OutputFolder -Force|Out-Null}

$script:Results=New-Object 'System.Collections.Generic.List[object]'
$script:Measurements=New-Object 'System.Collections.Generic.List[object]'
$script:Scores=@{}
$script:Run=Get-Date
$script:ComputerName=$env:COMPUTERNAME
$script:Serial='';$script:AssetTag=''
$script:Failures=0;$script:Warnings=0;$script:Attempted=0;$script:Completed=0;$script:Score=0

function Num([object]$v,[double]$d=0){try{if($null -eq $v -or "$v" -eq ''){return $d};return [double]$v}catch{return $d}}
function Rnd([object]$v,[int]$n=2){try{return [math]::Round([double]$v,$n)}catch{return 0}}
function Stats([double[]]$a){
 if(!$a -or $a.Count -eq 0){return $null};$s=@($a|Sort-Object);$avg=($s|Measure-Object -Average).Average
 $p50=$s[[math]::Floor(($s.Count-1)*.5)];$p95=$s[[math]::Floor(($s.Count-1)*.95)]
 $ss=0;foreach($x in $s){$ss+=([double]$x-$avg)*([double]$x-$avg)}
 [pscustomobject]@{Count=$s.Count;Average=Rnd $avg;Minimum=Rnd (($s|Measure-Object -Minimum).Minimum);Maximum=Rnd (($s|Measure-Object -Maximum).Maximum);Median=Rnd $p50;P95=Rnd $p95;StdDev=Rnd ([math]::Sqrt($ss/$s.Count))}
}
function StatusThreshold([double]$v,[double]$warn,[double]$fail,[switch]$HigherIsBad){if($HigherIsBad){if($v -ge $fail){'FAIL'}elseif($v -ge $warn){'WARNING'}else{'PASS'}}else{if($v -le $fail){'FAIL'}elseif($v -le $warn){'WARNING'}else{'PASS'}}}
function AddResult{
 param([string]$Category,[string]$Component,[string]$Test,[object]$Value,[string]$Unit='', [string]$Threshold='',
 [ValidateSet('PASS','WARNING','FAIL','INFO','UNKNOWN','NOT_APPLICABLE')][string]$Status='INFO',
 [ValidateSet('CRITICAL','HIGH','MEDIUM','LOW','INFO')][string]$Severity='INFO',
 [ValidateSet('HIGH','MEDIUM','LOW','UNKNOWN')][string]$Confidence='HIGH',[string]$Details='')
 $script:Attempted++;if($Status -notin @('UNKNOWN','NOT_APPLICABLE')){$script:Completed++}
 $script:Results.Add([pscustomobject]@{Timestamp=$script:Run.ToString('yyyy-MM-dd HH:mm:ss');ComputerName=$script:ComputerName;SerialNumber=$script:Serial;AssetTag=$script:AssetTag;Category=$Category;Component=$Component;Test=$Test;Value="$Value";Unit=$Unit;Threshold=$Threshold;Status=$Status;Severity=$Severity;Confidence=$Confidence;Details=$Details})
 if($Status -eq 'FAIL'){$script:Failures++}elseif($Status -eq 'WARNING'){$script:Warnings++}
}
function AddMeasure([string]$Category,[string]$Component,[string]$Metric,[double]$Value,[string]$Unit,[string]$Details=''){$script:Measurements.Add([pscustomobject]@{Timestamp=$script:Run.ToString('yyyy-MM-dd HH:mm:ss');ComputerName=$script:ComputerName;SerialNumber=$script:Serial;Category=$Category;Component=$Component;Metric=$Metric;Value=Rnd $Value;Unit=$Unit;Details=$Details})}
function SetScore([string]$Name,[double]$Score,[double]$Weight){$script:Scores[$Name]=@{Score=[math]::Max(0,[math]::Min(100,$Score));Weight=$Weight}}
function WeightedScore{$tw=0;$v=0;foreach($k in $script:Scores.Keys){$e=$script:Scores[$k];$tw+=$e.Weight;$v+=$e.Score*$e.Weight};if($tw){return Rnd ($v/$tw) 1};return 0}
function SafeCim([string]$Class,[string]$Namespace='root/cimv2'){try{return @(Get-CimInstance -Namespace $Namespace -ClassName $Class -ErrorAction Stop)}catch{return @()}}
function TempC{try{$z=@(Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop);$t=@($z|%{if($_.CurrentTemperature -gt 0){(Num $_.CurrentTemperature)/10-273.15}});if($t){return Rnd (($t|Measure-Object -Average).Average) 1}}catch{};return $null}
function EventCount([string]$Provider,[int]$Days){try{$e=@(Get-WinEvent -FilterHashtable @{LogName='System';ProviderName=$Provider;StartTime=(Get-Date).AddDays(-$Days)} -ErrorAction Stop);return $e.Count}catch{return 0}}
function BatteryHealth{try{$d=@(Get-CimInstance root/wmi BatteryStaticData -ErrorAction Stop);$f=@(Get-CimInstance root/wmi BatteryFullChargedCapacity -ErrorAction Stop);$dc=($d|%{Num $_.DesignedCapacity}|Measure-Object -Sum).Sum;$fc=($f|%{Num $_.FullChargedCapacity}|Measure-Object -Sum).Sum;if($dc -gt 0 -and $fc -gt 0){return Rnd (($fc/$dc)*100) 1}}catch{};return $null}
function AssetTag{try{$x=(Get-ItemProperty 'HKLM:\SOFTWARE\Nearlinx\AssetManagement' -ErrorAction Stop).AssetTag;if($x){return "$x"}}catch{};return ''}
function Trend([string]$name,[double]$score){if(!(Test-Path $HistoryFile)){return @('NO HISTORY',0)};try{$p=@(Import-Csv $HistoryFile|?{$_.ComputerName -eq $script:ComputerName}|sort Timestamp -Descending|select -First 1);if(!$p){return @('NO HISTORY',0)};$d=Rnd ($score-[double]$p.OverallScore) 1;if($d -le -5){return @('DEGRADING',$d)}elseif($d -ge 5){return @('IMPROVING',$d)}else{return @('STABLE',$d)}}catch{return @('NO HISTORY',0)}}
function UpdateGrid{$Grid.DataSource=$null;$Grid.DataSource=$script:Results.ToArray();foreach($r in $Grid.Rows){switch("$($r.Cells['Status'].Value)"){'PASS'{$r.DefaultCellStyle.BackColor=[Drawing.Color]::LightGreen}'WARNING'{$r.DefaultCellStyle.BackColor=[Drawing.Color]::Khaki}'FAIL'{$r.DefaultCellStyle.BackColor=[Drawing.Color]::LightSalmon}'UNKNOWN'{$r.DefaultCellStyle.BackColor=[Drawing.Color]::LightGray}'NOT_APPLICABLE'{$r.DefaultCellStyle.BackColor=[Drawing.Color]::Gainsboro}}}}
function Progress([int]$v,[string]$m){$Progress.Value=[math]::Max(0,[math]::Min(100,$v));$lblStatus.Text=$m;$form.Refresh()}

# GUI
$form=New-Object Windows.Forms.Form;$form.Text='Enterprise Windows Hardware Diagnostic Tool v2';$form.Size=New-Object Drawing.Size(1450,900);$form.StartPosition='CenterScreen';$form.BackColor=[Drawing.Color]::FromArgb(30,30,30)
$btnRun=New-Object Windows.Forms.Button;$btnRun.Text='Run Diagnostics';$btnRun.Location=New-Object Drawing.Point(15,15);$btnRun.Size=New-Object Drawing.Size(155,40)
$btnExcel=New-Object Windows.Forms.Button;$btnExcel.Text='Open Excel';$btnExcel.Location=New-Object Drawing.Point(180,15);$btnExcel.Size=New-Object Drawing.Size(140,40)
$btnHtml=New-Object Windows.Forms.Button;$btnHtml.Text='Open Dashboard';$btnHtml.Location=New-Object Drawing.Point(330,15);$btnHtml.Size=New-Object Drawing.Size(160,40)
$btnFolder=New-Object Windows.Forms.Button;$btnFolder.Text='Open Report Folder';$btnFolder.Location=New-Object Drawing.Point(500,15);$btnFolder.Size=New-Object Drawing.Size(165,40)
$form.Controls.AddRange([Windows.Forms.Control[]]@($btnRun,$btnExcel,$btnHtml,$btnFolder))
$lblStatus=New-Object Windows.Forms.Label;$lblStatus.Text='Ready';$lblStatus.Location=New-Object Drawing.Point(15,70);$lblStatus.Size=New-Object Drawing.Size(1300,25);$lblStatus.ForeColor=[Drawing.Color]::White;$form.Controls.Add($lblStatus)
$Progress=New-Object Windows.Forms.ProgressBar;$Progress.Location=New-Object Drawing.Point(15,100);$Progress.Size=New-Object Drawing.Size(1400,25);$form.Controls.Add($Progress)
$Grid=New-Object Windows.Forms.DataGridView;$Grid.Location=New-Object Drawing.Point(15,140);$Grid.Size=New-Object Drawing.Size(1400,680);$Grid.AutoSizeColumnsMode='Fill';$Grid.ReadOnly=$true;$Grid.AllowUserToAddRows=$false;$Grid.SelectionMode='FullRowSelect';$Grid.BackgroundColor=[Drawing.Color]::White;$Grid.ForeColor=[Drawing.Color]::Black;$Grid.RowHeadersVisible=$false;$form.Controls.Add($Grid)

$btnRun.Add_Click({
try{
$btnRun.Enabled=$false;$script:Results.Clear();$script:Measurements.Clear();$script:Scores.Clear();$script:Failures=0;$script:Warnings=0;$script:Attempted=0;$script:Completed=0;$script:Run=Get-Date;$Grid.DataSource=$null
Progress 3 'Collecting system identity...'
$c=Get-CimInstance Win32_ComputerSystem -ErrorAction Stop;$bios=Get-CimInstance Win32_BIOS -ErrorAction Stop;$cpu=@(Get-CimInstance Win32_Processor -ErrorAction Stop);$os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
$script:ComputerName=$env:COMPUTERNAME;$script:Serial="$($bios.SerialNumber)";$script:AssetTag=AssetTag
AddResult System Identity 'Computer Name' $script:ComputerName -Status INFO;AddResult System Identity Manufacturer $c.Manufacturer -Status INFO;AddResult System Identity Model $c.Model -Status INFO;AddResult System Identity 'Serial Number' $script:Serial -Status INFO;AddResult System Identity 'BIOS Version' $bios.SMBIOSBIOSVersion -Status INFO;AddResult System Identity 'BIOS Release Date' $bios.ReleaseDate -Status INFO;AddResult System Identity 'Operating System' $os.Caption -Status INFO;AddResult System Identity 'Windows Version' $os.Version -Status INFO

# CPU repeated samples
Progress 10 "Sampling CPU ($SampleCount samples)...";$ca=New-Object Collections.Generic.List[double];$ct=New-Object Collections.Generic.List[double]
for($i=1;$i -le $SampleCount;$i++){ $x=@(Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue);if($x){$u=($x|Measure-Object LoadPercentage -Average).Average;if($null-ne$u){$ca.Add([double]$u);AddMeasure Performance CPU Utilization $u '%' "Sample $i"}};$t=TempC;if($null-ne$t){$ct.Add([double]$t);AddMeasure Hardware CPU Temperature $t '°C' "Sample $i"};if($i -lt $SampleCount){Start-Sleep -Milliseconds $SampleIntervalMs}}
$s=Stats $ca.ToArray();if($s){$st=StatusThreshold $s.P95 80 95 -HigherIsBad;AddResult Performance CPU 'Average CPU Utilization' $s.Average '%' '<80% P95' $st MEDIUM HIGH "P95=$($s.P95)%; Max=$($s.Maximum)%; StdDev=$($s.StdDev)%";AddResult Performance CPU 'CPU P95 Utilization' $s.P95 '%' '<80%' $st MEDIUM HIGH;SetScore 'CPU Performance' $(if($s.P95-ge95){45}elseif($s.P95-ge80){70}elseif($s.P95-ge65){85}else{100}) 1}else{AddResult Performance CPU 'CPU Sampling' 'Unavailable' -Status UNKNOWN -Confidence LOW;SetScore 'CPU Performance' 50 1}
AddResult CPU Processor 'CPU Model' $cpu[0].Name -Status INFO;AddResult CPU Processor 'CPU Cores' $cpu[0].NumberOfCores -Status INFO;AddResult CPU Processor 'Logical Processors' $cpu[0].NumberOfLogicalProcessors -Status INFO
if($ct.Count){$ts=Stats $ct.ToArray();$tst=StatusThreshold $ts.P95 85 95 -HigherIsBad;AddResult CPU Thermal 'Thermal Zone P95' $ts.P95 '°C' '<85°C' $tst HIGH MEDIUM 'ACPI thermal-zone reading; package temperature may not be exposed.'}else{AddResult CPU Thermal 'Temperature' 'Unavailable' -Status UNKNOWN -Confidence LOW}

# RAM repeated samples
Progress 27 'Sampling RAM and memory pressure...';$ram=@(Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue);$ru=New-Object Collections.Generic.List[double];$installed=Rnd ((($ram|Measure-Object Capacity -Sum).Sum)/1GB) 2;AddResult RAM Memory 'Installed RAM' "$installed GB" -Status INFO
for($i=1;$i -le $SampleCount;$i++){
    $z=Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    if($z.TotalVisibleMemorySize){
        $totalMem=Num $z.TotalVisibleMemorySize
        $freeMem=Num $z.FreePhysicalMemory
        $pct=(($totalMem-$freeMem)/$totalMem)*100
        $ru.Add($pct)
        AddMeasure Performance RAM 'Memory Usage' $pct '%' "Sample $i"
    }
    if($i -lt $SampleCount){
        Start-Sleep -Milliseconds $SampleIntervalMs
    }
}
$rs=Stats $ru.ToArray();if($rs){$rst=StatusThreshold $rs.P95 80 92 -HigherIsBad;AddResult Performance RAM 'Average Memory Usage' $rs.Average '%' '<80% P95' $rst MEDIUM HIGH "P95=$($rs.P95)%; Max=$($rs.Maximum)%";AddResult Performance RAM 'Memory P95 Usage' $rs.P95 '%' '<80%' $rst MEDIUM HIGH;SetScore 'RAM Performance' $(if($rs.P95-ge92){40}elseif($rs.P95-ge80){70}elseif($rs.P95-ge65){85}else{100}) 1}else{SetScore 'RAM Performance' 50 1}
$ri=@($ram|?{$_.Status -and $_.Status -notin @('OK','Unknown')});if($ri.Count){AddResult RAM Memory 'Module Health' "$($ri.Count) abnormal module(s)" -Status FAIL -Severity HIGH;SetScore 'RAM Hardware' 30 1}else{AddResult RAM Memory 'Module Health' 'Normal' -Status PASS;SetScore 'RAM Hardware' 100 1};foreach($m in $ram){AddResult RAM 'Memory Module' Module "$(Rnd ($m.Capacity/1GB) 2) GB | $($m.Speed) MHz | $($m.Manufacturer) | $($m.PartNumber) | $($m.DeviceLocator)" -Status INFO}
foreach($d in 1,7,30){$n=EventCount 'Microsoft-Windows-MemoryDiagnostics-Results' $d;AddResult RAM 'Memory Diagnostics' "Events Last $d Days" $n events 0 $(if($n){'WARNING'}else{'PASS'}) MEDIUM HIGH}

# Storage
Progress 42 'Checking storage health and free space...';$disks=@(Get-PhysicalDisk -ErrorAction SilentlyContinue);$ss=New-Object Collections.Generic.List[double];foreach($d in $disks){$ds=if($d.HealthStatus -eq 'Healthy'){'PASS'}elseif($d.HealthStatus -eq 'Unknown'){'UNKNOWN'}else{'FAIL'};$score=if($ds-eq'PASS'){100}elseif($ds-eq'UNKNOWN'){60}else{20};$ss.Add($score);AddResult Storage Disk 'Disk Health' "$($d.FriendlyName) | $($d.MediaType) | $(Rnd ($d.Size/1GB) 2) GB | Health=$($d.HealthStatus) | Operational=$($d.OperationalStatus)" -Status $ds -Severity HIGH HIGH};$lds=@(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction SilentlyContinue);foreach($d in $lds){if($d.Size){$free=Rnd (($d.FreeSpace/$d.Size)*100) 1;$ds=StatusThreshold $free 15 5;$sc=if($free-lt5){30}elseif($free-lt15){70}else{100};$ss.Add($sc);AddResult Storage Disk "$($d.DeviceID) Free Space" $free '%' '>15%' $ds MEDIUM HIGH "$(Rnd ($d.FreeSpace/1GB) 2) GB free of $(Rnd ($d.Size/1GB) 2) GB"}};SetScore Storage $(if($ss.Count){($ss|Measure-Object -Average).Average}else{50}) 2

# GPU
Progress 50 'Checking graphics hardware...';$gpu=@(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue);$gs=New-Object Collections.Generic.List[double];foreach($g in $gpu){$st=if($g.Status-eq'OK'){'PASS'}else{'WARNING'};$gs.Add($(if($st-eq'PASS'){100}else{70}));AddResult GPU Graphics Adapter "$($g.Name) | Driver=$($g.DriverVersion) | Resolution=$($g.CurrentHorizontalResolution)x$($g.CurrentVerticalResolution)" -Status $st HIGH HIGH};SetScore GPU $(if($gs.Count){($gs|Measure-Object -Average).Average}else{50}) .5

# Battery
Progress 56 'Checking battery health...';$bat=@(Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue);if($bat){foreach($b in $bat){AddResult Battery Battery 'Charge Level' $b.EstimatedChargeRemaining '%' -Status INFO};$bh=BatteryHealth;if($null-ne$bh){$bst=if($bh-lt40){'FAIL'}elseif($bh-lt60){'WARNING'}else{'PASS'};$bsc=if($bh-lt40){30}elseif($bh-lt60){60}elseif($bh-lt80){80}else{100};AddResult Battery Battery 'Estimated Battery Health' $bh '%' '>=80%' $bst HIGH HIGH 'Full-charge capacity versus design capacity.';SetScore Battery $bsc 1.5}else{AddResult Battery Battery Health 'Capacity data unavailable' -Status UNKNOWN -Confidence LOW;SetScore Battery 60 1.5}}else{AddResult Battery Battery 'Battery Present' 'No battery detected' -Status NOT_APPLICABLE -Confidence HIGH}

# Security
Progress 63 'Checking TPM, Secure Boot and BitLocker...';$sec=New-Object Collections.Generic.List[double];try{$t=Get-Tpm -ErrorAction Stop;$sec.Add($(if($t.TpmPresent-and$t.TpmReady){100}elseif($t.TpmPresent){70}else{20}));AddResult Security TPM Present $t.TpmPresent -Status $(if($t.TpmPresent){'PASS'}else{'FAIL'}) HIGH HIGH;AddResult Security TPM Ready $t.TpmReady -Status $(if($t.TpmReady){'PASS'}else{'WARNING'}) MEDIUM HIGH}catch{AddResult Security TPM Status 'Unable to query' -Status UNKNOWN};try{$sb=Confirm-SecureBootUEFI -ErrorAction Stop;$sec.Add($(if($sb){100}else{70}));AddResult Security 'Secure Boot' Status $(if($sb){'ENABLED'}else{'DISABLED'}) -Status $(if($sb){'PASS'}else{'WARNING'}) HIGH HIGH}catch{AddResult Security 'Secure Boot' Status 'Not available / Legacy BIOS' -Status UNKNOWN};try{$bl=Get-BitLockerVolume -MountPoint C: -ErrorAction Stop;$on=$bl.ProtectionStatus-eq1;$sec.Add($(if($on){100}else{60}));AddResult Security BitLocker Protection $(if($on){'ON'}else{'OFF'}) -Status $(if($on){'PASS'}else{'WARNING'}) HIGH HIGH}catch{AddResult Security BitLocker Protection 'Unable to query' -Status UNKNOWN};SetScore Security $(if($sec.Count){($sec|Measure-Object -Average).Average}else{60}) 1

# PnP
Progress 70 'Checking Device Manager errors...';$pnp=@(Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue|?{$_.ConfigManagerErrorCode -and $_.ConfigManagerErrorCode-ne0});if(!$pnp.Count){AddResult Windows 'Device Manager' 'PnP Errors' 0 devices 0 PASS HIGH HIGH;SetScore 'Device Manager' 100 .75}else{AddResult Windows 'Device Manager' 'PnP Errors' $pnp.Count devices 0 WARNING HIGH HIGH;foreach($p in $pnp){AddResult Windows 'Device Manager' 'Problem Device' "$($p.Name) | Code=$($p.ConfigManagerErrorCode)" -Status WARNING -Severity HIGH};SetScore 'Device Manager' ([math]::Max(30,100-15*$pnp.Count)) .75}

# Network samples
Progress 76 'Sampling gateway latency and connectivity...';$ad=@(Get-NetAdapter -ErrorAction SilentlyContinue);$up=@($ad|?{$_.Status-eq'Up'});AddResult Network Adapters 'Active Adapters' $up.Count adapters '>0' $(if($up.Count){'PASS'}else{'FAIL'}) HIGH HIGH;foreach($a in $ad){AddResult Network Adapter $a.Name "$($a.InterfaceDescription) | $($a.Status) | $($a.MacAddress)" -Status $(if($a.Status-eq'Up'){'PASS'}else{'INFO'})};$gw=$null;try{$gw=Get-NetRoute -DestinationPrefix '0.0.0.0/0'|sort RouteMetric|select -First 1}catch{};$la=New-Object Collections.Generic.List[double];$ok=0;if($gw){for($i=1;$i-le$SampleCount;$i++){try{$r=Test-Connection $gw.NextHop -Count 1 -ErrorAction Stop;if($r){$la.Add([double]$r.ResponseTime);$ok++}}catch{};if($i-lt$SampleCount){Start-Sleep -Milliseconds 250}}};if($la.Count){$ls=Stats $la.ToArray();$lst=StatusThreshold $ls.P95 50 100 -HigherIsBad;$succ=Rnd (($ok/$SampleCount)*100) 1;AddResult Network Gateway 'Average Latency' $ls.Average ms '<50ms P95' $lst MEDIUM HIGH;AddResult Network Gateway 'P95 Latency' $ls.P95 ms '<50ms' $lst MEDIUM HIGH;AddResult Network Gateway 'Packet Success' $succ '%' '>=95%' $(if($succ-ge95){'PASS'}elseif($succ-gt0){'WARNING'}else{'FAIL'}) HIGH HIGH;SetScore Network $(if($succ-lt50){30}elseif($ls.P95-ge100){60}elseif($ls.P95-ge50){80}else{100}) 1}else{AddResult Network Gateway Reachability 'No successful replies' -Status FAIL -Severity HIGH;SetScore Network 30 1};$internet=Test-NetConnection 8.8.8.8 -Port 53 -InformationLevel Quiet -WarningAction SilentlyContinue;AddResult Network Internet 'Internet Connectivity' $(if($internet){'PASS'}else{'FAIL'}) -Status $(if($internet){'PASS'}else{'FAIL'}) HIGH HIGH

# Peripherals as detection, not false failures
Progress 84 'Checking peripheral detection...';$cam=@(Get-PnpDevice|?{$_.FriendlyName-match'Camera|Webcam'});$bt=@(Get-PnpDevice|?{$_.FriendlyName-match'Bluetooth'});$aud=@(Get-PnpDevice|?{$_.Class-in @('AudioEndpoint','MEDIA')});$usb=@(Get-PnpDevice -Class USB);$mon=@(Get-CimInstance root/wmi WmiMonitorID -ErrorAction SilentlyContinue);AddResult Peripherals Webcam Detection $cam.Count devices N/A $(if($cam){'PASS'}else{'NOT_APPLICABLE'}) LOW MEDIUM 'Detection only; no image test.';AddResult Peripherals Bluetooth Detection $bt.Count devices N/A $(if($bt){'PASS'}else{'NOT_APPLICABLE'}) LOW MEDIUM;AddResult Peripherals Audio Detection $aud.Count devices N/A $(if($aud){'PASS'}else{'NOT_APPLICABLE'}) LOW MEDIUM;AddResult Peripherals Display Detection $mon.Count monitors '>=1' $(if($mon){'PASS'}else{'UNKNOWN'}) MEDIUM MEDIUM;AddResult Peripherals USB 'USB Controllers' $usb.Count controllers '>0' $(if($usb){'PASS'}else{'UNKNOWN'}) MEDIUM HIGH;SetScore Peripherals 100 .5

# Windows / management
Progress 89 'Checking Windows Update and management...';$wu=Get-Service wuauserv -ErrorAction SilentlyContinue;if($wu){AddResult Windows 'Windows Update' Service $wu.Status Running $(if($wu.Status-eq'Running'){'PASS'}else{'WARNING'}) MEDIUM HIGH;SetScore 'Windows Update' $(if($wu.Status-eq'Running'){100}else{70}) .5}else{AddResult Windows 'Windows Update' Service 'Not found' -Status UNKNOWN;SetScore 'Windows Update' 50 .5};$en=@(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Enrollments\*' -ErrorAction SilentlyContinue);AddResult Management Intune Enrollment $(if($en.Count){'ENROLLED'}else{'NOT DETECTED'}) -Status $(if($en.Count){'PASS'}else{'WARNING'}) MEDIUM MEDIUM;SetScore Management $(if($en.Count){100}else{70}) .5

# Event windows
Progress 94 'Calculating 24h / 7d / 30d event statistics...';$evScores=New-Object Collections.Generic.List[double];foreach($e in @(@('WHEA Hardware Errors','Microsoft-Windows-WHEA-Logger'),@('Disk Errors','disk'),@('Kernel-Power 41','Microsoft-Windows-Kernel-Power'))){foreach($d in 1,7,30){$n=EventCount $e[1] $d;$st=if(!$n){'PASS'}elseif($d-eq1-and$n-ge5){'FAIL'}else{'WARNING'};AddResult Events $e[0] "Last $d Days" $n events 0 $st HIGH HIGH 'Time-windowed event count';if($d-eq30){$evScores.Add($(if(!$n){100}elseif($n-lt5){85}elseif($n-lt20){65}else{35}))}}};SetScore 'Event Log Health' $(if($evScores.Count){($evScores|Measure-Object -Average).Average}else{60}) .75

# Final scoring / coverage / trend
Progress 97 'Calculating weighted scores, coverage and trend...';$script:Score=WeightedScore;$coverage=if($script:Attempted){Rnd (($script:Completed/$script:Attempted)*100) 1}else{0};$overall=if($script:Score-lt50-or$script:Failures-gt0){'FAIL'}elseif($script:Score-lt75-or$script:Warnings-gt0){'WARNING'}else{'PASS'};AddResult Summary Overall 'Health Score' $script:Score '/100' '>=75' $overall HIGH HIGH 'Weighted component model';AddResult Summary Overall 'Diagnostic Coverage' $coverage '%' '>=90%' $(if($coverage-ge90){'PASS'}elseif($coverage-ge75){'WARNING'}else{'FAIL'}) MEDIUM HIGH;AddResult Summary Overall Failures $script:Failures tests 0 INFO;AddResult Summary Overall Warnings $script:Warnings tests 0 INFO;foreach($k in $script:Scores.Keys){$e=$script:Scores[$k];AddResult Scores Component $k (Rnd $e.Score 1) '/100' '' INFO MEDIUM HIGH}
$tr=Trend 'Overall' $script:Score;$trend=$tr[0];$delta=$tr[1];AddResult Trend Historical 'Device Trend' $trend '' '' INFO MEDIUM HIGH;if($trend-ne'NO HISTORY'){AddResult Trend Historical 'Score Change' $delta points '±5 stable' $(if($trend-eq'DEGRADING'){'WARNING'}elseif($trend-eq'IMPROVING'){'PASS'}else{'INFO'}) MEDIUM HIGH}

# Append fleet history
$rec=[ordered]@{Timestamp=$script:Run.ToString('yyyy-MM-dd HH:mm:ss');ComputerName=$script:ComputerName;SerialNumber=$script:Serial;AssetTag=$script:AssetTag;OverallScore=Rnd $script:Score 1;Coverage=$coverage;Failures=$script:Failures;Warnings=$script:Warnings};foreach($k in $script:Scores.Keys){$rec[$k]=Rnd $script:Scores[$k].Score 1};[pscustomobject]$rec|Export-Csv $HistoryFile -NoTypeInformation -Append -Encoding UTF8

UpdateGrid;$script:Results.ToArray()|Export-Csv $CsvFile -NoTypeInformation -Encoding UTF8
if(Get-Module -ListAvailable ImportExcel){try{Remove-Item $ExcelFile -Force -ErrorAction SilentlyContinue;$script:Results.ToArray()|Export-Excel $ExcelFile -WorksheetName Diagnostics -AutoSize -BoldTopRow -FreezeTopRow;if($script:Measurements.Count){$script:Measurements.ToArray()|Export-Excel $ExcelFile -WorksheetName Measurements -AutoSize -BoldTopRow -FreezeTopRow};Import-Csv $HistoryFile|Export-Excel $ExcelFile -WorksheetName FleetHistory -AutoSize -BoldTopRow -FreezeTopRow}catch{}}

# HTML dashboard
$rows=foreach($x in $script:Results){$cl=switch($x.Status){'PASS'{'pass'}'FAIL'{'fail'}'WARNING'{'warning'}'UNKNOWN'{'unknown'}'NOT_APPLICABLE'{'na'}default{'info'}};"<tr class='$cl'><td>$($x.Category)</td><td>$($x.Component)</td><td>$($x.Test)</td><td>$($x.Value)</td><td>$($x.Unit)</td><td>$($x.Status)</td><td>$($x.Confidence)</td><td>$($x.Details)</td></tr>"}
$scoreRows=foreach($k in $script:Scores.Keys){"<tr><td>$k</td><td>$(Rnd $script:Scores[$k].Score 1)/100</td></tr>"}
@"
<!doctype html><html><head><meta charset='utf-8'><title>Hardware Diagnostic v2</title><style>body{font-family:Segoe UI,Arial;background:#1e1e1e;color:white;margin:30px}.summary{display:flex;gap:15px;flex-wrap:wrap}.card{background:#2d2d2d;padding:18px;border-radius:8px;min-width:160px}.score{font-size:28px;font-weight:bold}table{border-collapse:collapse;width:100%;background:white;color:#222;margin:20px 0}th{background:#0078d7;color:white;padding:8px;text-align:left}td{padding:7px;border:1px solid #ccc}.pass{background:#d9f2d9}.fail{background:#ffd6d6}.warning{background:#fff2b3}.unknown,.na{background:#e5e5e5}</style></head><body><h1>Enterprise Windows Hardware Diagnostic Tool v2</h1><p>$script:ComputerName | $script:Serial | $($script:Run.ToString('yyyy-MM-dd HH:mm:ss'))</p><div class='summary'><div class='card'>Overall<div class='score'>$($script:Score)/100</div></div><div class='card'>Status<div class='score'>$overall</div></div><div class='card'>Coverage<div class='score'>$coverage%</div></div><div class='card'>Trend<div class='score'>$trend</div></div><div class='card'>Failures<div class='score'>$($script:Failures)</div></div><div class='card'>Warnings<div class='score'>$($script:Warnings)</div></div></div><h2>Component Scores</h2><table><tr><th>Component</th><th>Score</th></tr>$($scoreRows -join "`n")</table><h2>Diagnostics</h2><table><tr><th>Category</th><th>Component</th><th>Test</th><th>Value</th><th>Unit</th><th>Status</th><th>Confidence</th><th>Details</th></tr>$($rows -join "`n")</table><p>Non-destructive diagnostic. Physical/stress testing is not included.</p></body></html>
"@|Set-Content $HtmlFile -Encoding UTF8
Progress 100 'Diagnostics completed.';[Windows.Forms.MessageBox]::Show("Diagnostics completed.`r`n`r`nStatus: $overall`r`nHealth Score: $($script:Score)/100`r`nCoverage: $coverage%`r`nTrend: $trend`r`nFailures: $($script:Failures)`r`nWarnings: $($script:Warnings)`r`n`r`nCSV: $CsvFile`r`nExcel: $ExcelFile`r`nHTML: $HtmlFile`r`nHistory: $HistoryFile",'Diagnostic Complete','OK','Information')|Out-Null
}catch{[Windows.Forms.MessageBox]::Show("Diagnostic error:`r`n$($_.Exception.Message)`r`nLine: $($_.InvocationInfo.ScriptLineNumber)",'Error','OK','Error')|Out-Null;$lblStatus.Text='Diagnostic failed.'}finally{$btnRun.Enabled=$true}
})
$btnExcel.Add_Click({if(Test-Path $ExcelFile){Invoke-Item $ExcelFile}elseif(Test-Path $CsvFile){Invoke-Item $CsvFile}else{[Windows.Forms.MessageBox]::Show('Run diagnostics first.')}})
$btnHtml.Add_Click({if(Test-Path $HtmlFile){Invoke-Item $HtmlFile}else{[Windows.Forms.MessageBox]::Show('Run diagnostics first.')}})
$btnFolder.Add_Click({if(Test-Path $OutputFolder){Invoke-Item $OutputFolder}})
[void]$form.ShowDialog()
