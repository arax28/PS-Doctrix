Param(
# Файл выгрузки оргструктуры
[string]$inputFile# = "D:\Scripts\Работа с оргструктурой\Orgstructure.2016-02-17-15-59-47.xml"
)

# Получение отдельных блоков данных, в разные переменные для удобства
$employees = Select-Xml -Path $inputFile -XPath "//EmployeeDto" | Select-Object -ExpandProperty node
$departments = Select-Xml -Path $inputFile -XPath "//DepartmentDto" | Select-Object -ExpandProperty node
$roles = Select-Xml -Path $inputFile -XPath "//RoleDto" | Select-Object -ExpandProperty node

$xmlSrc = new-object XML
$xmlSrc.Load($inputFile)

<#
.Synopsis
   Возвращает список подразделений указанного пользователя
.EXAMPLE
   Get-EmployeeDepartments -employeeTitle "Komarova" -matchTitle:$true
#>
function Get-EmployeeDepartments
(
    $employeeLogin,
    $employeeTitle="",
    $matchTitle=$false
)
{
    if ($employeeTitle -ne "") {
        if ($matchTitle) {
            $employee = $employees | where {$_.Title -match $employeeTitle}   
        }
        else {
            $employee = $employees | where {$_.Title -eq $employeeTitle}    
        }
    }
    elseif ($employeeLogin -ne "") {
        $employee = $employees | where {$_.LoginName -eq $employeeLogin} 
    }
    
    return $employee.DepartmentUniqueNames
}

function Get-DepartmentPath {
<#
.Synopsis
   Функция возвращает структуру вышестоящих подразделений для указанного подразделения
.DESCRIPTION
   Функция возвращает структуру вышестоящих подразделений для указанного подразделения в виде массива строк
.EXAMPLE
   Get-DepartmentPath -DepartmentUniqueName "1a98410d-7e29-11e0-b528-0022648a31f2"
#>
    [CmdletBinding()]
    [OutputType([string[]])]
    Param (
     # Уникальный идентификатор подразделения
     [Parameter(Mandatory=$True)]
      [string]$DepartmentUniqueName,

     # Выводить ли полные имена ролей. Включено по умолчанию. !Обработка значения $False пока не реализовано!
     [Parameter(Mandatory=$false)]
      [switch]$ShowTitles=$True
      )
    BEGIN {
        if ($ShowTitles = $false) {
            throw [System.NotImplementedException]
        }
    }
    PROCESS {
        #$department = Select-XML -Path $input -XPath "DepartmentDto[@UniqueName = $DepartmentUniqueName]" | Select-Object -ExpandProperty node
        $department = $departments | where {$_.UniqueName -eq $DepartmentUniqueName}
        $departmentPath = @($department.Title)
        while ($department.ParentDepartmentUniqueName -ne $null) {
            $department = $departments | where {$_.UniqueName -eq $department.ParentDepartmentUniqueName}
            $departmentPath += $department.Title
        }

        [array]::Reverse($departmentPath)
        return $departmentPath
    }
}

<#
.Synopsis
   Функция возвращает UniqueName подразделения по частичному Title с возможностью интерактивного выбора из нескольких результатов
.DESCRIPTION
   Длинное описание
.EXAMPLE
   Пример использования этого командлета
.EXAMPLE
   Еще один пример использования этого командлета
#>
function Find-DepartmentByName
{
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # Справочное описание параметра 1
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Param1,

        # Справочное описание параметра 2
        [int]
        $Param2
    )

    Begin
    {
    }
    Process
    {
    }
    End
    {
    }
}


function Out-OrgXML
{
<#
.Synopsis
   Функция для формирования файла для последующего импорта в оргструктуру СЭД
.DESCRIPTION
   Длинное описание
.EXAMPLE
   Out-OrgXML -LiteralPath "c:\temp\temp.xml" -DepartmentsXML $XmlElementCollection
#>
    [CmdletBinding()]
    Param
    (
        # Путь к файлу для сохранения оргструктуры
        [Parameter(Mandatory=$true)]
        [string]$LiteralPath,

        # Массив элементов подразделений
        [Parameter(Mandatory=$false)]
        $DepartmentsXML,

        # Массив элементов ролей
        [Parameter(Mandatory=$false)]
        $RolesXML,

        # Массив элементов сотрудников
        [Parameter(Mandatory=$false)]
        $EmployeesXML
    )

    Begin
    {

    }
    Process
    {
        
        try {
            if ($DepartmentsXML -eq $null -and $RolesXML -eq $null -and $EmployeesXML -eq $null) {
                Write-Error "Нужно передать хотя бы один массив элементов оргструктуры для формирования файла!"
                return 0
            }
            Write-Verbose "Подготовка настроек для вывода XML"
            [System.XMl.XmlWriterSettings]$settings = New-Object System.XMl.XmlWriterSettings
            $settings.Indent = $true
            $settings.IndentChars = "`t"
            $XmlWriter = [System.XMl.XmlWriter]::Create([string]$LiteralPath, $settings)
            Write-Verbose "Подготовка базовой структуры XML для файла оргструктуры"
            $XmlWriter.WriteStartDocument()
            $XmlWriter.WriteStartElement("OrgstructureDto")
            $XmlWriter.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")
            $XmlWriter.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")

            $XmlWriter.WriteStartElement("DepartmentDtoList")
            if ($DepartmentsXML) {
                Write-Verbose "Записываем массив элементов подразделений"
                foreach ($DepartmentXML in $DepartmentsXML) {
                    $XmlWriter.WriteStartElement("DepartmentDto")
                    $XmlWriter.WriteRaw($DepartmentXML.innerXml)
                    $XmlWriter.WriteEndElement() # DepartmentDto
                }
            
            }
            $XmlWriter.WriteEndElement() # DepartmentDtoList

            $XmlWriter.WriteStartElement("RoleDtoList")
            if ($RolesXML) {
                Write-Verbose "Записываем массив элементов ролей"
                 foreach ($RoleXML in $RolesXML) {
                    $XmlWriter.WriteStartElement("RoleDto")
                    $XmlWriter.WriteRaw($RoleXML.innerXml)
                    $XmlWriter.WriteEndElement() # RoleDto
                }
            }
            $XmlWriter.WriteEndElement() # RoleDtoList

            $XmlWriter.WriteStartElement("EmployeeDtoList")
            if ($EmployeesXML) {
                Write-Verbose "Записываем массив элементов сотрудников"
                foreach ($EmployeetXML in $EmployeesXML) {
                    $XmlWriter.WriteStartElement("EmployeeDto")
                    $XmlWriter.WriteRaw($EmployeetXML.innerXml)
                    $XmlWriter.WriteEndElement() # EmployeeDto
                }
                
            }
            $XmlWriter.WriteEndElement() # EmployeeDtoList

            Write-Verbose "Закрытие документа"
            $XmlWriter.WriteEndElement() # OrgstructureDto
            $XmlWriter.WriteEndDocument()
        }
        finally {
            $XmlWriter.Close()
        }
    }
    End
    {
    }
}

function Add-EmployeeToRole {
<#
.Synopsis
   Добавляет указанного сотрудника в заданную роль
.DESCRIPTION
   Добавляет указанного сотрудника в заданную роль. Обратно возвращает XML element с описанием сотрудника.
.EXAMPLE
   Add-EmployeeToRole -RoleUniqueName "РУКОВОДИТЕЛЬПОДРАЗДЕЛЕНИЯОТДЕЛРАЗВИТИЯПРОДАЖРЕГИОНОВ" -UserLoginName "PonyEx\username" -Verbose
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    Param
    (
        # Уникальное имя роли (параметр UniqueName из XML)
        [Parameter(Mandatory=$true)]
        [string]$RoleUniqueName,

        # Имя пользователя в формате domainname\username (параметр LoginName из XML)
        [Parameter(Mandatory=$true)]
        [string]$UserLoginName
    )
    Begin
    {
        $role = $xmlSrc.OrgstructureDto.RoleDtoList.RoleDto | where {$_.UniqueName -eq $RoleUniqueName}
        Write-Verbose "Добавляем пользователя '$UserLoginName' в роль '$($role.Title)'"
    }
    Process
    {
        #$employee = (Select-Xml -Path $inputFile -XPath "//EmployeeDto[LoginName = `"$UserLoginName`" ]")
        Write-Debug "Ищем в исходном файле пользователя с логином '$UserLoginName' (точное совпадение)"
        $employee = $xmlSrc.OrgstructureDto.EmployeeDtoList.EmployeeDto | where {$_.LoginName -like "*$UserLoginName"}
        if (!$employee) {
            Write-Error "Пользователь с логином '$UserLoginName' не найден!"
            continue
        }
        elseif ($employee.count -gt 1) {
            Write-Error "Пользователь с логином '$UserLoginName' найден в нескольких экземлярах, невозможно однозначно определить добавляемого пользователя"
            continue
        }

        
        if ($($employee.RoleUniqueNames | where {$_.string -eq $RoleUniqueName}) -eq $null) {
            $newRole = $xmlSrc.CreateElement("string")
            $newRole.InnerText = $RoleUniqueName
            $employee.RoleUniqueNames.AppendChild($newRole) | Out-Null
        } else {
            Write-Verbose "Пользователь '$UserLoginName' уже включен в роль '$RoleUniqueName'"
        }
        return $employee.Clone()
    }
    End
    {
    }
}


function Get-EmptyManagerRoles
{
<#
.Synopsis
   Выводит список ролей с префиксом "Руководитель подразделения", в которых нет ни одного сотрудника.
.DESCRIPTION
   Выводит список ролей с префиксом "Руководитель подразделения", в которых нет ни одного сотрудника.
   Для каждой роли выводится путь по подразделениям.
   Никаких дополнительных параметров не требует.
.EXAMPLE
   Get-EmptyManagerRoles
#>
    [CmdletBinding()]
        $managerRoles = $xmlSrc.OrgstructureDto.RoleDtoList.RoleDto | where {$_.Title -like "Руководитель подразделения*"}
        foreach ($managerRole in $managerRoles) { 
            $managers = $xmlSrc.OrgstructureDto.EmployeeDtoList.EmployeeDto | 
                where {$_.RoleUniqueNames.InnerText -match $managerRole.UniqueName}
            if (!$managers) {
                Write-Output "$($managerRole.Title) не назначен"
                Write-Output "Роль можно найти по следующей иерархии подразделений:"
                Write-Output $(Get-DepartmentPath $managerRole.DepartmentUniqueName)
                Write-Output "-------------------"
            }
        }
}

function Get-MultiManagerRoles
{
<#
.Synopsis
   Выводит список ролей с префиксом "Руководитель подразделения", в которые включено несколько сотрудников.
.DESCRIPTION
   Выводит список ролей с префиксом "Руководитель подразделения", в которые включено несколько сотрудников.
   Для каждой найденной роли выводится количество и список сотрудников, а также путь по подразделениям.
   Никаких дополнительных параметров не требует
.EXAMPLE
   Get-MultiManagerRoles
#>
    [CmdletBinding()]
    Param()

    Begin
    {
    }
    Process
    {
        $managerRoles = $xmlSrc.OrgstructureDto.RoleDtoList.RoleDto | where {$_.Title -like "Руководитель подразделения*"}
        Write-Verbose "Нашли всего $($managerRoles.Count) ролей с префиксом 'Руководитель подразделения', начинаем проверять количество сотрудников в этих ролях"
        foreach ($managerRole in $managerRoles) { 
            $managers = $xmlSrc.OrgstructureDto.EmployeeDtoList.EmployeeDto | 
                where {$_.RoleUniqueNames.InnerText -contains $managerRole.UniqueName}
            if ($managers.Count -gt 1) {
                Write-Output "Роль '$($managerRole.Title)' назначена на $($managers.Count) сотрудников:"
                $managers | %{Write-Output " $($_.Title)"}
                Write-Output "Роль можно найти по следующей иерархии подразделений:"
                Write-Output $(Get-DepartmentPath $managerRole.DepartmentUniqueName)
                Write-Output "-------------------"
            }
            else {
                Write-Verbose "Роль '$($managerRole.Title)' назначена на $($managers.Count) сотрудников, пропускаем её."
            }
        }
    }
    End
    {
    }
}


function Get-DeptsWithNoManagerRole
{
<#
.Synopsis
   Выводит список подразделений, в которых не существует роли с префиксом "Руководитель подразделения".
.DESCRIPTION
   Выводит список подразделений, в которых не существует роли с префиксом "Руководитель подразделения".
   Для каждой роли выводится путь по подразделениям.
.EXAMPLE
   Get-DeptsWithNoManagerRole
.EXAMPLE
   Get-DeptsWithNoManagerRole -SearchRootDeptID "1a98410d-7e29-11e0-b528-0022648a31f2"
#>
    [CmdletBinding()]
    Param(
    #Название подразделения, с которого начинать просмотр. Используется для исключения подразделений, в которых не используется автоматическое создание роли руководителя. Требуется указывать идентификатор подразделения, а не его название.
    [Parameter(Mandatory=$false)]
        [string]$SearchRootDeptID
    )

    Begin
    {
    }
    Process
    {
        if ($SearchRootDeptID) {
            Write-Debug "ID стартового подразделения: $SearchRootDeptID"
            Write-Verbose "Проверяем ограниченную иерархию оргструктуры, начиная с заданного подразделения"
            $DepartmentScope = Get-XMLDeptDescendants -StartDeptUniqueName $SearchRootDeptID
        } else {
            Write-Verbose "Проверяем все подразделения"
            $DepartmentScope = $xmlSrc.OrgstructureDto.DepartmentDtoList.DepartmentDto
        }

        foreach ($dept in $DepartmentScope) {
                $rolesInDept = $xmlSrc.OrgstructureDto.RoleDtoList.RoleDto | where {$_.DepartmentUniqueName -eq $dept.UniqueName}
                if ($rolesInDept -and $rolesInDept.count -eq 1) {
                    Write-Verbose "В подразделении '$($dept.Title)' найдена единственная роль"
                    if ($rolesInDept.Title -like "Руководитель подразделения*") {
                        Write-Verbose "И это роль руководителя подразделения"
                    }
                    else {
                        Write-Verbose "И это не роль руководителя подразделения!"
                        Write-Output "Подразделение '$($dept.Title)' не содержит роли руководителя!"
                        Write-Output "Роль можно найти по следующей иерархии подразделений:"
                        Write-Output $(Get-DepartmentPath $dept.UniqueName)
                        Write-Output "-------------------"
                    }
                }
                elseif (!$rolesInDept) {
                    Write-Verbose "В подразделении '$($dept.Title)' не найдено ролей"
                    Write-Output "Подразделение '$($dept.Title)' не содержит роли руководителя!"
                    Write-Output "Роль можно найти по следующей иерархии подразделений:"
                    Write-Output $(Get-DepartmentPath $dept.UniqueName)
                    Write-Output "-------------------"
                }
                $rolesInDept = $null
         } #foreach ($dept in $DepartmentScope)
    }
    End
    {
    }
}

<#
.Synopsis
   Возвращает подразделения, входящие в иерархию, начиная с указанного
.DESCRIPTION
   Возвращает линейный набор узлов XML, соответствующий подразделениям, входящим в иерархию, начиная с указанного.
   Поддерживается Verbose вывод.
.EXAMPLE
   Get-XMLDeptDescendants -StartDeptUniqueName "1a98410d-7e29-11e0-b528-0022648a31f2"
#>
function Get-XMLDeptDescendants
{
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    Param
    (
        # UniqueName подразделения, для которого нужно получить все нижестоящие подразделения 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [string]$StartDeptUniqueName
    )

    Begin
    {
    }
    Process
    {
        $startDepartment = $xmlSrc.OrgstructureDto.DepartmentDtoList.DepartmentDto | where {$_.UniqueName -eq $StartDeptUniqueName}
        if (!$startDepartment) {
            throw [System.ArgumentException]
        }
        Write-Verbose "Подразделение $($startDepartment.Title)"
        
        $ChildDepartments = $xmlSrc.OrgstructureDto.DepartmentDtoList.DepartmentDto | where {$_.ParentDepartmentUniqueName -eq $StartDeptUniqueName}
        $childXML = @()
        foreach ($ChildDepartment in $ChildDepartments) {
            Get-XMLDeptDescendants $ChildDepartment.UniqueName
            
        }
        
        return [System.Xml.XmlElement]$startDepartment
    }
    End
    {
    }
}



function Clear-XMLSection {
# Функция для удаления всех записей из секции XML
}

function Clear-XMLDepartmentMembers {
# Функция для удаления членов из подразделений
}

function Clear-XMLRoleMembers {
# Функция для удаления членов из ролей
}

<#
.Synopsis
   Возвращает список функциональных ролей в виде набора объектов XML Element
.EXAMPLE
   Get-FunctionalRolesXML
#>
function Get-FunctionalRolesXML
{
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement[]])]
    Param()

    Begin
    {
    }
    Process
    {
        $FunctionalRoles = $xmlSrc.OrgstructureDto.RoleDtoList.RoleDto | where {$_.DepartmentUniqueName -eq $null}
        return $FunctionalRoles
    }
    End
    {
    }
}

<#
.Synopsis
   Возвращает элемент (или набор элементов) оргструктуры, соответсвующий пользователю с указанным именем
.DESCRIPTION
   Возвращает элемент оргструктуры, соответсвующий пользователю с указанным именем. При поиске по частичному совпадению может возвращаться набор элементов
.EXAMPLE
   Get-EmployeeByName -EmployeeName "Иванов Иван" -PartialMatch
#>
function Get-EmployeeByName 
{
    [CmdletBinding()]
    #[OutputType([System.Xml.XmlElement])]
    Param
    (
        # Имя сотрудника
        [Parameter(Mandatory=$true)]
        [string]$EmployeeName,

        # Включает возможность поиска по частичному совпадению
        [switch]$PartialMatch
    )

    Begin
    {
    }
    Process
    {
        if ($PartialMatch) {
            Write-Verbose "Ищем запись сотрудника в оргструктуре по имени `"$EmployeeName`""
            $results = $xmlSrc.OrgstructureDto.EmployeeDtoList.EmployeeDto | where {$_.Title -match $EmployeeName}
        }
        else {
            Write-Verbose "Ищем запись сотрудника в оргструктуре по частичному совпадению имени `"$EmployeeName`""
            $results = $xmlSrc.OrgstructureDto.EmployeeDtoList.EmployeeDto | where {$_.Title -eq $EmployeeName}
        }
        Write-Debug "Найдено $($results.count) записей"
        return $results
    }
    End
    {
    }
}

<#
.Synopsis
   Выводит XML элементы сотрудников, имеющих одинаковые логины
.DESCRIPTION
   Выводит XML элементы сотрудников, имеющих одинаковые логины. Параметров не требует.
.EXAMPLE
   Get-DuplicateLoginEmployeesXML
#>
function Get-DuplicateLoginEmployeesXML
{
    [CmdletBinding()]
    Param()
    Begin
    {
    }
    Process
    {
        $employeeLogins = $xmlSrc.OrgstructureDto.EmployeeDtoList.EmployeeDto.LoginName
        Write-Verbose "Выбираем уникальные логины сотрудников для сравнения"
        $uniqueLogins = $employeeLogins | select -Unique
        Write-Verbose "Получаем список дублированных логинов с помощью сравнения полного списка логинов со списком уникальных, отсеивая повторные дубликаты"
        $duplicateLogins = Compare-Object -ReferenceObject $uniqueLogins -DifferenceObject $employeeLogins | select -Property InputObject -Unique
        Write-Verbose "Получаем XML элементы по всем сотрудникам, имеющим задублированные логины, и выводим итоговый список"
        $xmlSrc.OrgstructureDto.EmployeeDtoList.EmployeeDto | where {$_.LoginName -in $duplicateLogins.InputObject} | Sort-Object -Property LoginName
    }
    End
    {
    }
}