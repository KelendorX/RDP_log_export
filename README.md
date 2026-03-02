# RDP Connections Log Exporter

Набор скриптов для экспорта логов подключений по протоколу RDP (Remote Desktop Protocol) из журнала событий Windows.

## 📋 Описание

Данный инструмент позволяет извлекать и анализировать события RDP-подключений из журнала "Microsoft-Windows-TerminalServices-RDPClient/Operational". Скрипты собирают информацию о сессиях, группируют связанные события по ActivityID и экспортируют данные в структурированном JSON-формате.

## 📁 Состав пакета

- `rdp_log_export_BYDATE.bat` - пакетный файл для запуска PowerShell скрипта с интерактивным вводом даты
- `RDP_Log_ByDate.ps1` - основной PowerShell скрипт для сбора и анализа RDP-событий

## 🔧 Требования

- Windows (с поддержкой PowerShell)
- Права администратора для чтения журналов событий
- Журнал "Microsoft-Windows-TerminalServices-RDPClient/Operational" должен содержать события

## 📊 Анализируемые события

Скрипт обрабатывает следующие Event ID:
- **1024** - Успешное подключение к RDP-серверу
- **1026** - Отключение от RDP-сервера
- **1105** - Завершение RDP-сессии

## 🚀 Использование

### Запуск через BAT-файл (рекомендуется)

1. Запустите `rdp_log_export_BYDATE.bat`
2. Введите дату в формате `YYYY-MM-DD` или нажмите Enter для использования текущей даты
3. Результат будет сохранен в JSON-файл с именем `RDP_Log_YYYY-MM-DD.json`

### Прямой запуск PowerShell скрипта

```powershell
# Для текущей даты
.\RDP_Log_ByDate.ps1

# Для конкретной даты
.\RDP_Log_ByDate.ps1 -Date "2026-03-03"

# С указанием пути для экспорта
.\RDP_Log_ByDate.ps1 -Date "2026-03-03" -ExportPath "C:\Logs\rdp_log.json"
