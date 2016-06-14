<#
.Synopsis
    Скрипт выполняет резервное копирование сайт-коллекции вместе с её настройками Doctrix
.DESCRIPTION
    Требуемые права доступа:
    Для резервирования сайт-коллекции нужно иметь права локального администратора на сервере (???)
    Для резервирования только настроек Doctrix нужно быть db_owner базы Doctrix
.EXAMPLE
    Backup-DataAndConfig -SiteURL 'https://doctrix-test.mycompany.ru:5566' -DoctrixServerInstance 'dbserver' -DoctrixDataBase 'Test_Doctrix_Config' -BackUpRootFolder 'E:\Doctrix Backups' -RetentionDays 8
.EXAMPLE
    Backup-DataAndConfig -SiteURL 'https://doctrix-test.mycompany.ru:5566' -DoctrixServerInstance 'dbserver' -DoctrixDataBase 'Test_Doctrix_Config' -BackUpRootFolder 'E:\Doctrix Backups' -RetentionDays -1 -SkipContentBackup
#>
param(
# Адрес коллекции сайтов для снятия резервной копии.
[Parameter(Mandatory=$true)]
[uri]$SiteURL,
# Имя инстанса MSSQL, можно использовать алиас
[Parameter(Mandatory=$true)]
[string]$DoctrixServerInstance,
[Parameter(Mandatory=$true)]
[string]$DoctrixDataBase = "ECMTest_Doctrix_Config",
# Корневая папка, где будет создана подпапка для сохранения бэкапа.
[Parameter(Mandatory=$true)]
[string]$BackUpRootFolder = "D:\Temp\DoctrixBackup",
# Количество дней, бэкапы старше которого будут удалены. При отрицательном значении ничего не удаляется
[Parameter(Mandatory=$false)]
[int]$RetentionDays = -1,
# Имя файла для записи лога
# Todo: Реализовать проверку и использование этого параметра
[Parameter(Mandatory=$false)]
[string]$LogFile,
[switch]$SkipContentBackup,
[switch]$DoNotCompress
)

# Флаг, включаемый при критических ошибках
$errorFlag = $false

# Сохранение текущего значения ErrorAction после всех операций, его изменяющих
$currentErrorAction = $ErrorActionPreference

cd $PSScriptRoot

# Запись даты бэкапа для формирования файлов и определения старых бэкапов
$startDate = Get-Date
$dateBackup = Get-Date $startDate -Format "yy_MM_dd_HH-mm-ss"

# Настройка и инициализация лога
. .\Function-Write-Log.ps1
if ([string]::IsNullOrEmpty($LogFile)) {
    $LogFile = Join-Path $PSScriptRoot "\log\backup-$dateBackup.log"
}

Write-Log -Message "Бэкап запущен в $startDate" -Path $LogFile

# Добавляем модуль SharePoint Powershell
Add-PSSnapin microsoft.sharepoint.powershell -ErrorAction SilentlyContinue

######################################################################################
Write-Log "------------Параметры бэкапа------------" -Path $LogFile
######################################################################################
Write-Log "Сайт-коллекция: $SiteURL" -Path $LogFile
Write-Log "БД Doctrix: $DoctrixServerInstance\$DoctrixDataBase" -Path $LogFile
Write-Log "Путь к корневой папке для бэкапов: $BackUpRootFolder" -Path $LogFile
if ($RetentionDays -gt 0) {
    Write-Log "Удаляются бэкапы старше: $RetentionDays дней" -Path $LogFile -Level Warn
} elseif ($RetentionDays -eq 0) {
    Write-Log "Удаляются все старые бэкапы (RetentionDays=0)!" -Path $LogFile -Level Warn
} else {
    Write-Log "Cтарые бэкапы не удаляются" -Path $LogFile
}
if ($SkipContentBackup) {
    Write-Log "Пропустить бэкап сайт-коллекции" -Path $LogFile
}
if ($DoNotCompress) {
    Write-Log "Пропустить архивацию бэкапа" -Path $LogFile
}

######################################################################################
Write-Log "------------Создание папки для хранения бэкапа------------" -Path $LogFile
######################################################################################

# Формирование пути для бэкапа
$siteBackupPathSegment = $SiteURL.host
if ($SiteURL.Port) {$siteBackupPathSegment += ("_"+$SiteURL.Port)}
if ($SiteURL.LocalPath -ne "/") {$siteBackupPathSegment += ($SiteURL.LocalPath.Replace('/','_'))}
$BackUpFolder = Join-Path (Join-Path $BackUpRootFolder $siteBackupPathSegment) $dateBackup

Write-Log "Путь к папке для сохранения бэкапа: $BackUpFolder" -Path $LogFile

if ((Test-Path $BackUpFolder) -eq $false)
{
    try {
        New-Item -Path $BackUpFolder -ItemType Directory -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Log "При создании папки для бэкапа возникла ошибка:`n $($_.Exception.Message)" -Path $LogFile -Level Error
        return
    }
    Write-Log "Папки не было, создали новую" -Path $LogFile
}
Else
{
    Write-Log "Папка уже существует" -Path $LogFile
}

######################################################################################
Write-Log "------------Создание бэкапа сайт-коллекции------------" -Path $LogFile
######################################################################################
if (!$SkipContentBackup)
{
    $BackupFile = Join-Path $BackUpFolder ("SiteContent.spbak")
    Write-Log "Бэкап сайт-коллекции в файл $BackupFile" -Path $LogFile
    try {
        Backup-SPSite -Identity $SiteURL -Path $BackupFile -ErrorAction Stop
    }
    catch
    {
        $errorFlag = $true
        Write-Log "Ошибка бэкапа сайт-коллекции!`n$($_.Exception.Message)" -Path $LogFile -Level Error
    }
    
    if (!$DoNotCompress -and (Test-Path $BackupFile -PathType Leaf))
    {
        $BackupFileZip = $BackupFile.Replace(".spbak", ".zip")
        #Add-Type -assembly "system.io.compression.filesystem"
        Write-Log "Архивирование файла бэкапа..." -Path $LogFile
        try {
        # Todo: избавиться от использования 7z для универсальности (?)
        # Todo: Сделать архивирование всех файлов, а не только бэкапа сайта
            &'C:\Program Files\7-Zip\7z.exe' a $BackupFileZip $BackupFile 2>&1
        }
        catch {
            $errorFlag = $true
            Write-Log "Ошибка архивирования файла бэкапа:`n $($_.Exception.Message)" -Path $LogFile -Level Error
        }
        if (Test-Path $BackupFileZip)
        {
            Write-Log "Архивирование завершено, удаление несжатого бэкапа..." -Path $LogFile
            try {
                Remove-Item $BackupFile -ErrorAction Stop
            }
            catch {
                $errorFlag = $true
                Write-Log "При удалении несжатого бэкапа возникла ошибка:`n $($_.Exception.Message)" -Path $LogFile -Level Error
            }
        }
    }
    Write-Log "Бэкап сайт-коллекции завершён." -Path $LogFile
}
else {
    Write-Log "Бэкап сайт-коллекции пропущен согласно параметрам скрипта." -Path $LogFile
}

######################################################################################
Write-Log "------------Сохранение настроек Doctrix для сайта------------" -Path $LogFile
######################################################################################
$site = Get-SPsite -Identity $SiteURL
Write-Log "Настройки Doctrix для сайта с ID $($site.ID)" -Path $LogFile

$filePath = Join-Path $BackUpFolder "ListSettings_$($site.ID).xml"
Write-Log "Сохраняем настройки Doctrix в $filePath" -Path $LogFile
Write-Log "Начало сохранения настроек Doctrix..." -Path $LogFile

$ErrorActionPreference = 'Stop'

try {
    $conn = New-Object System.Data.SqlClient.SqlConnection("Server=$DoctrixServerInstance;DataBase=$DoctrixDataBase;Integrated Security=True")
    $conn.Open()
    $cmd = New-Object System.Data.SqlClient.SqlCommand("dbo.ExportSiteSettings", $conn)
    $cmd.CommandType = [System.Data.CommandType]::StoredProcedure
    $cmd.Parameters.AddWithValue("@SiteID", $site.ID) | Out-Null
    $param = $cmd.Parameters.Add("@Settings", [System.Data.SqlDbType]::Xml, 2000000000)
    $param.Direction = [System.Data.ParameterDirection]::Output
    $cmd.ExecuteNonQuery() | Out-Null
    $param.Value | Out-File -FilePath:$filePath
}
catch {
    $ErrorActionPreference = $currentErrorAction
    $errorFlag = $true
    Write-Log "При экспорте настроек Doctrix возникла ошибка!:`n $($_.Exception.Message)" -Path $LogFile -Level Error
}
finally {
    $conn.Close()
}

Write-Log "Настройки Doctrix сохранены." -Path $LogFile

######################################################################################
Write-Log "------------Сохранение правил Doctrix для сайта------------" -Path $LogFile
######################################################################################
<#
.Synopsis
    Функция выгрузки правил Doctrix в файл
#>
function Export-DTSiteRules(
    [string] $ServerInstance,
    [string] $DataBase,
    [guid]   $SiteID,
    [string] $BackupPath)
{

        $Connection = New-Object System.Data.SqlClient.SqlConnection("Server=$ServerInstance;DataBase=$DataBase;Integrated Security=True")
        $Connection.Open()
        $Command = New-Object System.Data.SqlClient.SqlCommand("dbo.ExportSiteRules", $Connection)
        $Command.CommandType = [System.Data.CommandType]::StoredProcedure
    
        # SiteID
        $Command.Parameters.AddWithValue("@SiteID", $SiteID) | Out-Null

        # Settings
        $Parameter = $Command.Parameters.Add("@Settings", [System.Data.SqlDbType]::Xml, 2000000000)
        $Parameter.Direction = [System.Data.ParameterDirection]::Output
    
        $Command.ExecuteNonQuery() | Out-Null
        $Parameter.Value | Out-File -FilePath $BackupPath
        $Connection.Close()
}

try
{
    $ErrorActionPreference = 'Stop'  
    $RulesBackupFile = Join-Path $BackUpFolder "SiteRules_$($site.ID).xml"
    Write-Log "Сохраняем правила Doctrix в $RulesBackupFile." -Path $LogFile
    Write-Log "Начало экспорта правил Doctrix..." -Path $LogFile

    Export-DTSiteRules                              `
        -ServerInstance       $DoctrixServerInstance       `
        -DataBase             $DoctrixDataBase             `
        -SiteID               ($site.ID)            `
        -BackupPath           $RulesBackupFile
}
catch
{
        $ErrorActionPreference = $currentErrorAction
        $errorFlag = $true
        Write-Log "Произошла ошибка при экспорте правил Doctrix:`n$($_.Exception.Message)" -Path $LogFile -Level Error
}  
$ErrorActionPreference = $currentErrorAction
Write-Log "Экспорт правил Doctrix завершён." -Path $LogFile

# Если бэкап прошел без ошибок, при необходимости, выполняем удаление устаревших бэкапов
if (!$errorFlag -and $RetentionDays -ge 0)
{
    $delBeforeDate = $startDate.AddDays(-$RetentionDays)
    Write-Log "Удаляем бэкапы старше $RetentionDays дней, созданные до $delBeforeDate" -Path $LogFile
    Get-ChildItem (Join-Path $BackUpRootFolder $siteBackupPathSegment) | where {$_.CreationTime -lt $delBeforeDate} | Remove-Item -confirm:$false -recurse
}

# Todo: Проверка на наличие файлов в папке бэкапа. Если ни одного нет - выводить что бэкап не выполнен.
if ($errorFlag) { Write-Log "Бэкап выполнен с ошибками!" -Path $LogFile -Level Warn}
else { Write-Log "Бэкап выполнен успешно." -Path $LogFile }