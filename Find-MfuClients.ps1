# ----------------------------
# Config
# ----------------------------

$printerIP   = "192.168.200.46"
$output      = "result.csv"
$subnets     = @("192.168.100.", "192.168.101.")
$maxParallel = 30

"Computer,Printer,Port" | Out-File $output -Encoding ascii

# ----------------------------
# Генерация IP
# ----------------------------
$ips = foreach ($sub in $subnets) {
    1..254 | ForEach-Object { "$sub$_" }
}

$total     = $ips.Count
$completed = 0
$jobs      = @()

# ----------------------------
# Функция обработки jobs
# ----------------------------
function Process-Jobs {

    $jobs | Wait-Job | Out-Null

    foreach ($j in $jobs) {

        $result = Receive-Job $j

        if ($result -like "FOUND:*") {
            Write-Host "НАЙДЕН → $($result.Split(':')[1])" -ForegroundColor Green
        }
        elseif ($result -like "ERROR:*") {
            Write-Host "ОТКАЗ → $($result.Split(':')[1])" -ForegroundColor Red
        }

        $script:completed++
        $percent = [int](($script:completed / $script:total) * 100)

        Write-Progress -Id 1 `
            -Activity "Scanning network" `
            -Status "$script:completed / $script:total" `
            -PercentComplete $percent
    }

    $jobs | Remove-Job
    $script:jobs = @()
}

# ----------------------------
# Основной цикл
# ----------------------------
foreach ($ip in $ips) {

    $jobs += Start-Job -ScriptBlock {

        param($pc, $printerIP, $output)

        function Write-Result {
            param($pc, $printerName, $port)
            "$pc,$printerName,$port" | Out-File $output -Append -Encoding ascii
        }

        try {
            $printers = Get-WmiObject Win32_Printer -ComputerName $pc -ErrorAction Stop
            $ports    = Get-WmiObject Win32_TCPIPPrinterPort -ComputerName $pc -ErrorAction SilentlyContinue

            $reg         = [wmiclass]"\\$pc\root\default:StdRegProv"
            $HKLM        = 2147483650
            $monitorsKey = "SYSTEM\CurrentControlSet\Control\Print\Monitors"
            $monitors    = $reg.EnumKey($HKLM, $monitorsKey).sNames

            foreach ($printer in $printers) {

                # 1️⃣ Проверка PortName
                if ($printer.PortName -like "*$printerIP*") {
                    Write-Result $pc $printer.Name $printer.PortName
                    return "FOUND:$pc"
                }

                # 2️⃣ Проверка TCPIPPrinterPort
                if ($ports) {
                    $portMatch = $ports | Where-Object {
                        $_.Name -eq $printer.PortName -and
                        $_.HostAddress -eq $printerIP
                    }

                    if ($portMatch) {
                        Write-Result $pc $printer.Name $portMatch.HostAddress
                        return "FOUND:$pc"
                    }
                }

                # 3️⃣ Проверка через реестр
                if ($monitors) {

                    foreach ($monitor in $monitors) {

                        $portsKey  = "$monitorsKey\$monitor\Ports"
                        $portsList = $reg.EnumKey($HKLM, $portsKey)

                        if ($portsList.ReturnValue -ne 0 -or -not $portsList.sNames) {
                            continue
                        }

                        foreach ($portName in $portsList.sNames) {

                            if ($portName -ne $printer.PortName) { continue }

                            $portKey = "$portsKey\$portName"

                            $ipValue   = $reg.GetStringValue($HKLM, $portKey, "IPAddress").sValue
                            $hostValue = $reg.GetStringValue($HKLM, $portKey, "HostName").sValue

                            if ($ipValue -eq $printerIP -or $hostValue -eq $printerIP) {
                                Write-Result $pc $printer.Name $printerIP
                                return "FOUND:$pc"
                            }
                        }
                    }
                }
            }

            return "OK:$pc"
        }
        catch {
            return "ERROR:$pc"
        }

    } -ArgumentList $ip, $printerIP, $output


    # Ограничение параллельности
    if ($jobs.Count -ge $maxParallel) {
        Process-Jobs
    }
}

# Обработка остатка
Process-Jobs

Write-Progress -Id 1 -Activity "Scanning network" -Completed
Write-Host "Сканирование завершено."
