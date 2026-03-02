# Export-RDP-ByDate.ps1
param(
    [string]$ExportPath = "",
    [string]$Date = ""
)

try {
    # Устанавливаем правильную кодировку для консоли
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8

    # Получаем путь где лежит скрипт
    $ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

    # Обрабатываем параметр даты
    if ([string]::IsNullOrEmpty($Date)) {
        # Если дата не указана - используем сегодня
        $TargetDate = (Get-Date).Date
        $DateDescription = "Today"
    } else {
        # Парсим указанную дату
        try {
            $TargetDate = [DateTime]::Parse($Date).Date
            $DateDescription = "Specified date: $Date"
        }
        catch {
            Write-Host "Error: Invalid date format '$Date'. Please use format YYYY-MM-DD" -ForegroundColor Red
            return
        }
    }

    # Если путь экспорта не указан, создаем автоматическое имя файла
    if ([string]::IsNullOrEmpty($ExportPath)) {
        $ExportPath = "RDP_Log_$($TargetDate.ToString('yyyy-MM-dd')).json"
    }

    # Создаем полный путь для файла логов
    if (Split-Path -Path $ExportPath -IsAbsolute) {
        $FullExportPath = $ExportPath
    } else {
        $FullExportPath = Join-Path -Path $ScriptDirectory -ChildPath $ExportPath
    }

    # Создаем папку для логов если нужно
    $LogFolder = Split-Path -Path $FullExportPath -Parent
    if (!(Test-Path $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }

    # Устанавливаем временной диапазон для указанной даты
    $StartTime = $TargetDate
    $EndTime = $TargetDate.AddDays(1)

    Write-Host "Collecting RDP data for: $($StartTime.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan

    # Получаем события (1024 - успешные подключения, 1026 и 1105 - отключения)
    $Events = Get-WinEvent -LogName "Microsoft-Windows-TerminalServices-RDPClient/Operational" -ErrorAction SilentlyContinue | 
        Where-Object { 
            ($_.Id -eq 1024 -or $_.Id -eq 1026 -or $_.Id -eq 1105) -and 
            $_.TimeCreated -ge $StartTime -and 
            $_.TimeCreated -lt $EndTime
        }

    Write-Host "Found events: $($Events.Count) (1024: $(($Events | Where-Object { $_.Id -eq 1024 }).Count), 1026: $(($Events | Where-Object { $_.Id -eq 1026 }).Count), 1105: $(($Events | Where-Object { $_.Id -eq 1105 }).Count))" -ForegroundColor Gray

    # Группируем события по Correlation ActivityID (исключаем пустые ActivityID)
    $GroupedEvents = $Events | Where-Object { $_.ActivityId -ne $null } | Group-Object -Property @{Expression={$_.ActivityId}}

    $SessionData = foreach ($Group in $GroupedEvents) {
        $ActivityID = $Group.Name
        $EventsInGroup = $Group.Group | Sort-Object TimeCreated
        
        # Ищем подключение (1024) и отключения (1026 или 1105) в этой группе
        $ConnectionEvent = $EventsInGroup | Where-Object { $_.Id -eq 1024 } | Select-Object -First 1
        $DisconnectionEvents = $EventsInGroup | Where-Object { $_.Id -eq 1026 -or $_.Id -eq 1105 } | Sort-Object TimeCreated
        
        # Берем последнее событие отключения (приоритет для 1026, если есть)
        $DisconnectionEvent = $null
        if ($DisconnectionEvents) {
            # Сначала ищем событие 1026
            $DisconnectionEvent1026 = $DisconnectionEvents | Where-Object { $_.Id -eq 1026 } | Select-Object -Last 1
            if ($DisconnectionEvent1026) {
                $DisconnectionEvent = $DisconnectionEvent1026
            } else {
                # Если нет 1026, берем 1105
                $DisconnectionEvent = $DisconnectionEvents | Select-Object -Last 1
            }
        }
        
        if ($ConnectionEvent) {
            try {
                $ConnectionXml = [xml]$ConnectionEvent.ToXml()
                $ConnectionDataValues = @{}
                if ($ConnectionXml.Event.EventData -ne $null -and $ConnectionXml.Event.EventData.Data -ne $null) {
                    $ConnectionXml.Event.EventData.Data | ForEach-Object {
                        if ($_.Name -ne $null) {
                            $ConnectionDataValues[$_.Name] = $_.'#text'
                        }
                    }
                }
                
                $DisconnectionDataValues = @{}
                $DisconnectionTime = $null
                $Duration = $null
                $DisconnectionType = $null
                
                if ($DisconnectionEvent) {
                    try {
                        $DisconnectionXml = [xml]$DisconnectionEvent.ToXml()
                        if ($DisconnectionXml.Event.EventData -ne $null -and $DisconnectionXml.Event.EventData.Data -ne $null) {
                            $DisconnectionXml.Event.EventData.Data | ForEach-Object {
                                if ($_.Name -ne $null) {
                                    $DisconnectionDataValues[$_.Name] = $_.'#text'
                                }
                            }
                        }
                        $DisconnectionTime = $DisconnectionEvent.TimeCreated.ToString("yyyy-MM-ddTHH:mm:ss.fffffff")
                        $DisconnectionType = if ($DisconnectionEvent.Id -eq 1026) { "1026" } else { "1105" }
                        
                        # Вычисляем длительность сессии
                        $ConnectionTime = $ConnectionEvent.TimeCreated
                        $DisconnectionTimeObj = $DisconnectionEvent.TimeCreated
                        $Duration = ($DisconnectionTimeObj - $ConnectionTime).ToString("hh\:mm\:ss")
                    }
                    catch {
                        Write-Warning "Error processing disconnection event $($DisconnectionEvent.RecordId): $($_.Exception.Message)"
                    }
                }
                
                # Формируем объединенный объект сессии
                @{
                    ActivityID = $ActivityID
                    Connection = @{
                        Time = $ConnectionEvent.TimeCreated.ToString("yyyy-MM-ddTHH:mm:ss.fffffff")
                        Data = $ConnectionDataValues
                        EventRecordID = $ConnectionEvent.RecordId
                        Level = $ConnectionEvent.LevelDisplayName
                    }
                    Disconnection = if ($DisconnectionEvent) {
                        @{
                            Time = $DisconnectionTime
                            Data = $DisconnectionDataValues
                            EventRecordID = $DisconnectionEvent.RecordId
                            Level = $DisconnectionEvent.LevelDisplayName
                            EventID = $DisconnectionType
                        }
                    } else {
                        $null
                    }
                    SessionInfo = @{
                        Server = if ($ConnectionDataValues['Server']) { $ConnectionDataValues['Server'] } else { "Unknown" }
                        User = if ($ConnectionDataValues['User']) { $ConnectionDataValues['User'] } else { "Unknown" }
                        Duration = $Duration
                        Status = if ($DisconnectionEvent) { "Completed" } else { "Active" }
                    }
                    MachineName = $ConnectionEvent.MachineName
                }
            }
            catch {
                Write-Warning "Error processing connection event $($ConnectionEvent.RecordId): $($_.Exception.Message)"
                continue
            }
        }
    }

    # Статистика по сессиям
    $TotalSessions = $SessionData.Count
    $CompletedSessions = ($SessionData | Where-Object { $_.Disconnection -ne $null }).Count
    $ActiveSessions = ($SessionData | Where-Object { $_.Disconnection -eq $null }).Count

    # Статистика по типам отключений
    $Disconnection1026Count = ($SessionData | Where-Object { $_.Disconnection -ne $null -and $_.Disconnection.EventID -eq "1026" }).Count
    $Disconnection1105Count = ($SessionData | Where-Object { $_.Disconnection -ne $null -and $_.Disconnection.EventID -eq "1105" }).Count

    # Структура для экспорта
    $ExportObject = @{
        ExportTime = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
        QueryPeriod = @{
            Start = $StartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
            End = $EndTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
            Description = $DateDescription
        }
        Statistics = @{
            TotalSessions = $TotalSessions
            CompletedSessions = $CompletedSessions
            ActiveSessions = $ActiveSessions
            CompletionRate = if ($TotalSessions -gt 0) { 
                [math]::Round(($CompletedSessions / $TotalSessions) * 100, 2) 
            } else { 0 }
            DisconnectionTypes = @{
                Event1026 = $Disconnection1026Count
                Event1105 = $Disconnection1105Count
            }
        }
        Sessions = $SessionData
    }

    # Экспорт в JSON с правильной кодировкой
    $ExportObject | ConvertTo-Json -Depth 8 | Set-Content -Path $FullExportPath -Encoding UTF8

    # Информация о времени выполнения
    $ExecutionTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "`nScript executed: $ExecutionTime" -ForegroundColor Gray
}
catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}