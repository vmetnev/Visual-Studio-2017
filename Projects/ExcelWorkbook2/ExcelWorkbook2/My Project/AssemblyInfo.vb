﻿Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' Управление общими сведениями о сборке осуществляется с помощью 
' набора атрибутов. Отредактируйте эти значения атрибутов, чтобы изменить сведения,
' сопоставленные со сборкой.

' Проверьте значения атрибутов сборки

<Assembly: AssemblyTitle("ExcelWorkbook2")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("ExcelWorkbook2")> 
<Assembly: AssemblyCopyright("Copyright ©  2017")> 
<Assembly: AssemblyTrademark("")> 

' Установка значения False в параметре ComVisible делает типы в этой сборке невидимыми 
' для компонентов COM.  Если необходимо обратиться к типу в этой сборке через 
' COM, следует установить атрибут ComVisible в TRUE для этого типа.
<Assembly: ComVisible(False)>

'Следующий GUID служит для идентификации библиотеки типов, если этот проект видим для COM
<Assembly: Guid("d0e010ab-3728-4344-a19a-d91370b6ed5b")> 

' Сведения о версии сборки состоят из следующих четырех значений:
'
'      Основной номер версии
'      Дополнительный номер версии 
'   Номер сборки
'      Версия
'
' Можно задать все значения или принять номера сборки и редакции по умолчанию 
' с помощью знака '*', как показано ниже:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
