Attribute VB_Name = "modCHF"
Option Explicit
Option Compare Text

'last update 10.01.2017 - Added modified function - Replace_OneStrDirectorPodr
'last update 09.01.2017 - Function change - ReturnChiefsDepartmentFIO
'                       - Function change - LCaseString
'last update 29.12.2016 - Added modified function - DetailedGetInfoIdToValue
'                         Added operator enum - SortingMethod
'last update 28.12.2016

'----------Оператор перечисления в качестве переменных указаны параметры для функции Replace_DirectorPodr,
'которая возвращает сведения о руководителе подразделения
Public Enum ParamTypeDepartmentName
    DepartmentShortName = 1
    DepartmentFullName = 2
    DepartmentChiefFIO = 3
    DepartmentChiefPostShortName = 5
End Enum

'---------Оператор перечисления в качестве переменных зададим Падежи
'хотел написать по-английски падежи, но ....
Public Enum DifferentPadeg
    Im = 1
    Rod = 2
    Dat = 3
    Vin = 4
    Tv = 5
    Pr = 6
End Enum

'---------Оператор способа сортировки
Public Enum SortingMethod
    desc
    Asc
End Enum

'-------------------получение строки с дbректором подразделения по штатной единице сотрудника--------------
Public Function Replace_OneStrDirectorPodr(NumbOption As Integer, NameOption As String) As String
    'Слитый реквизит - shtat_podr_info
    '_0='Object_Code_charval=''01'';
    'ShortName_charval='Название подразделения';
    'fullname_charval='Полное название подразделения';
    'Director_FIO_Personal='Фамилия Имя отчество';
    'Director_intval=id_сотрудника;
    'Podr_Director_ExecPost_Shtat_Shortname='Должность сотрудника';';
    
    If Not Nvl(CStr(NameOption), "") = "" Then
        Dim WrdArray1() As String
        Dim WrdArray2() As String
        WrdArray1() = Split(CStr(NameOption), ";")
        If UBound(WrdArray1) > 0 Then
            WrdArray2() = Split(WrdArray1(NumbOption), "=")
            If Not CStr(WrdArray2(1)) = "''" Or Not CStr(WrdArray2(1)) = "Null" Then
                Replace_OneStrDirectorPodr = CStr(Replace(WrdArray2(1), "'", ""))
            Else
                Replace_OneStrDirectorPodr = "Null"
            End If
        End If
    Else
        Replace_OneStrDirectorPodr = "Null"
    End If
End Function


'----------------Функция возвращает полное/краткое наименование подразделения ---------------------------
Public Function ReturnChiefsDepartmentName(ByVal rs As ADODB.Recordset, SelectionParams As ParamTypeDepartmentName, Optional DepartmentPadeg As DifferentPadeg = Im) As String
    If (SelectionParams >= 1 Or SelectionParams <= 2) Or (SelectionParams = 5) Then
        If Not CastToString(CastToString(Replace_DirectorPodr(CastToLong(SelectionParams), rs("shtat_podr_info").Value))) = "Null" Then
            ReturnChiefsDepartmentName = GetPodrPadeg(CastToString(Replace_DirectorPodr(CastToLong(SelectionParams), rs("shtat_podr_info").Value)), DepartmentPadeg)
        End If
    End If
End Function

'----------------Функция возвращает ФИО руководителя подразделения---------------------------
Public Function ReturnChiefsDepartmentFIO(ByVal rs As ADODB.Recordset, Optional SelectionParams As FIOFormatEnum = ffSurnameNamePatronomic, Optional FIOPadeg As DifferentPadeg = Im) As String
    If Not CastToString(Replace_DirectorPodr("3", rs("shtat_podr_info").Value)) = "Null" Then
    'Проверка, возможна ситуация когда у нас несколько вложенных подразделений
    'будем пробегать все по цепочке снизу-вверх
    Dim PodrArray() As String, AllStringArray() As String, i As Integer, CountStringPodr As Integer, TempString As String
    Dim WrdArray3() As String
    AllStringArray() = Split(CastToString(rs("shtat_podr_info").Value), ";")
    TempString = ""
    CountStringPodr = -1
    ReDim Preserve PodrArray(0)
    For i = 0 To UBound(AllStringArray)
        If (i Mod 7) = 0 Then
            CountStringPodr = CountStringPodr + 1
            If CStr(AllStringArray(i)) <> "" Then
                ReDim Preserve PodrArray(CountStringPodr)
                PodrArray(CountStringPodr) = AllStringArray(i)
            End If
        Else
            If CStr(AllStringArray(i)) <> "" Then
                PodrArray(CountStringPodr) = PodrArray(CountStringPodr) & ";" & AllStringArray(i)
            End If
        End If
    Next i
    'На выходе получим массив строк содержащий сведения о вышестоящем подразделении
    
    i = 0
    Do While Not i > UBound(PodrArray)
        Erase WrdArray3
'        'WrdArray3() = Split(CastToString(Replace_DirectorPodr("3", rs("shtat_podr_info").Value)), " ")
        WrdArray3() = Split(CastToString(Replace_OneStrDirectorPodr("3", PodrArray(i))), " ")
        If UBound(WrdArray3) = 2 Then
            ReturnChiefsDepartmentFIO = MakeFIOShortCorrectly(CastToString(WrdArray3(0)), CastToString(WrdArray3(1)), CastToString(WrdArray3(2), ""), FIOPadeg, SelectionParams)
            Exit Do
        End If
        i = i + 1
    Loop
    End If
End Function

'---------------Функция возвращает либо ФИО руководителя, либо Должность руководителя организации,
'которого нет в штате. Данные извлекаются из опций.
'Функция реагирует на параметры:
'    ffNPSurname, ffSurnameNamePatronomic, ffSurnameNP = вернет ФИО в разных форматах
'    Post = вернет должность
Public Function ReturnGeneralChief(FieldsValue As String, bs As IBusinessServer) As String

    If CastToString(FieldsValue) = "ffNPSurname" Then
        If Not CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 0)) = "" Then
            ReturnGeneralChief = MakeFIOShortOneString(CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 0)), 1, ffNPSurname)
        Else
            ReturnGeneralChief = CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 4))
        End If
    ElseIf CastToString(FieldsValue) = "ffSurnameNamePatronomic" Then
        If Not CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 0)) = "" Then
            ReturnGeneralChief = MakeFIOShortOneString(CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 0)), 1, ffSurnameNamePatronomic)
        Else
            ReturnGeneralChief = CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 4))
        End If
    ElseIf CastToString(FieldsValue) = "ffSurnameNP" Then
        If Not CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 0)) = "" Then
            ReturnGeneralChief = MakeFIOShortOneString(CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 0)), 1, ffSurnameNP)
        Else
            ReturnGeneralChief = CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_FIO", 4))
        End If
    ElseIf CastToString(FieldsValue) = "Post" Then
        If Not CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_POST", 0)) = "" Then
        ReturnGeneralChief = GetPostPadeg(CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_POST", 0)), 1)
        Else
            ReturnGeneralChief = CastToString(bs.GetOption("ORG_STRUCTURE_BOSS_POST", 4))
        End If
    End If
End Function

'Функция ничего не возвращает. Она пробегает по Закладкам, которые определены для Руководителей которые должны
'подписывать документ
'вызывается вот так: Call FieldRecordingChiefs(параметры)
Public Function FieldRecordingChiefs(ByVal FRC_RS As ADODB.Recordset, _
                                    FRC_QueryDate As Date, _
                                    FRC_bs As IBusinessServer)
    Dim NameFieldChief As String, ValueFieldChief As String, WriteString As String
    Dim TempString As String
    Dim Set1 As bookmark, LastSymbol As Long
    'берем все закладки страницы
    ActiveDocument.Range.InsertAfter (vbCrLf)
    For Each Set1 In ActiveDocument.Bookmarks
        'если закладка существует
        If ActiveDocument.Bookmarks.Exists(Set1.Name) = True Then
            'извлекаем имя кода
            NameFieldChief = DeleteCharacters(Set1.Name, True)
            'извлекаем имя поля
            ValueFieldChief = DeleteCharacters(Set1.Name, False)
            
            'Проверяем, чтобы оба значения не были пустыми
            If (Not Nvl(NameFieldChief, "") = "") And (Not Nvl(ValueFieldChief, "") = "") Then
                
                'Проверяем, если вконце строки имеется число,
                'то это дублирование строки параметров
                If IsNumeric(CStr(Right(ValueFieldChief, 1))) Then
                    LastSymbol = CInt(Right(ValueFieldChief, 1))
                    If LastSymbol >= 0 Then
                        TempString = Trim(ValueFieldChief)
                        ValueFieldChief = ""
                        ValueFieldChief = Left(TempString, Len(TempString) - 1)
                    End If
                End If
                
                'получаем результат
                If CastToString(NameFieldChief) = "Управляющий" Then
                    PutToBkm CastToString(Set1.Name), CastToString(ReturnGeneralChief(ValueFieldChief, FRC_bs))
                Else
                    WriteString = CastToString(ReturnChiefCustomValue(FRC_RS, FRC_QueryDate, FRC_bs, NameFieldChief, ValueFieldChief))
                    If Not CastToString(WriteString, "") = "" Then
                        PutToBkm CastToString(Set1.Name), CastToString(WriteString)
                    'Не понятно как использовать исключение
                    'Else
                    '    PutToBkm CastToString(Set1.Name), CastToString("Некорректный параметр")
                    End If
                End If
            End If
        End If
    Next
End Function

'Функция возвращает английскую подстроку, либо русскую подстроку
'используется для автоматической замены закладок рекордсета Chief
Public Function DeleteCharacters(Stxt As String, Optional LanguageEnglishCharacters As Boolean = True) As String
    Dim i As Integer, a As String
    For i = Len(Stxt) To 1 Step -1
        a = Mid(Stxt, i, 1)
        If LanguageEnglishCharacters = True Then
            'English characters
            If a Like "[a-zA-Z0-9_]" Then Stxt = Replace(Stxt, a, "")
        ElseIf LanguageEnglishCharacters = False Then
            'Russian characters
            If a Like "[А-яЁё]" Then Stxt = Replace(Stxt, a, "")
        End If
    Next
    
    'return rus/eng string
    DeleteCharacters = Trim(Stxt)
End Function

'Функция возвращает паспорт сотрудника свернутый в строку формата серия/номер/кем выдан/когда выдан
'в случае чего можно будет добавить формат
Public Function ReturnPersonalPasport(IdPersonal As Long, _
                                    ByVal rs As ADODB.Recordset, _
                                    QueryDate As Date, _
                                    bs As IBusinessServer, _
                                    Optional StyleFormat As Integer = 1) As String
    Dim PasportString As String
    Dim TempString As String, TempArray, TexpTempArray, CountTemp As Integer
    TempArray = Array("serdoc", "numdoc", "whogive", "date_begin")
    If StyleFormat = 1 Then
        TexpTempArray = Array("серия", "номер", "кем выдан", "дата выдачи")
    End If
    CountTemp = 0
    Do While UBound(TempArray) >= CountTemp
        TempString = CastToString(GetInfoIdToValue(IdPersonal, "REC_Personal", "passportinternal", CastToString(TempArray(CountTemp)), QueryDate, bs), "")
        If Not TempString = "" Then
            PasportString = PasportString & " " & CastToString(TexpTempArray(CountTemp)) & " " & TempString
        End If
        CountTemp = CountTemp + 1
    Loop
    
    If Not CastToString(PasportString, "") = "" Then
        ReturnPersonalPasport = PasportString
    End If
End Function

'------------------Функция возращает единичное значение по коду Chief, например, вернет Фамилия, Имя
Public Function ReturnChiefCustomValue(ByVal rs As ADODB.Recordset, _
                                    QueryDate As Date, _
                                    bs As IBusinessServer, _
                                    NameChief As String, _
                                    ValueChief) As String
    rs.MoveFirst
    Do While Not rs.EOF
        If Nvl(CastToString(rs("chief_code").Value), "") = CastToString(NameChief) Then
            If CastToString(ValueChief) = "ffNPSurname" Then
                ReturnChiefCustomValue = MakeFIOShortCorrectly(CastToString(rs("surname").Value), _
                                                            CastToString(rs("name").Value), _
                                                            CastToString(rs("patronymic").Value), _
                                                            1, _
                                                            ffNPSurname)
            ElseIf CastToString(ValueChief) = "ffSurnameNamePatronomic" Then
                ReturnChiefCustomValue = MakeFIOShortCorrectly(CastToString(rs("surname").Value), _
                                                            CastToString(rs("name").Value), _
                                                            CastToString(rs("patronymic").Value), _
                                                            1, _
                                                            ffSurnameNamePatronomic)
            ElseIf CastToString(ValueChief) = "ffSurnameNP" Then
                ReturnChiefCustomValue = MakeFIOShortCorrectly(CastToString(rs("surname").Value), _
                                                            CastToString(rs("name").Value), _
                                                            CastToString(rs("patronymic").Value), _
                                                            1, _
                                                            ffSurnameNP)
            ElseIf Not Nvl(CastToString(rs(ValueChief).Value), "") = "" Then
                ReturnChiefCustomValue = CastToString(rs(ValueChief).Value)
            Else
                ReturnChiefCustomValue = "Некорректный параметр"
            End If
            rs.MoveLast
            rs.MoveNext
        Else
            rs.MoveNext
        End If
    Loop
End Function

Public Function InformationEmployee(ByVal rs As ADODB.Recordset, QueryDate As Date, bs As IBusinessServer, TransmissionString As String, _
                                    Optional idPers As Boolean = False, _
                                    Optional Tabn As Boolean = False, _
                                    Optional SurnameName As Boolean = False, _
                                    Optional Name As Boolean = False, _
                                    Optional Patronomyc As Boolean = False, _
                                    Optional idShtat As Boolean = False, _
                                    Optional ShtatCode As Boolean = False, _
                                    Optional ShtatShorname As Boolean = False, _
                                    Optional ShtatFullname As Boolean = False, _
                                    Optional PodrFullname As Boolean = False, _
                                    Optional ChiefFullname As Boolean = False) As String()
    'перебираем рекордсет руководителей
    Dim Masv() As String
    Dim idPodr As Long, CountMasv As Integer
    rs.MoveFirst
    CountMasv = 0
    Do While Not rs.EOF
        If Nvl(CastToString(rs("chief_code").Value), "") = CastToString(TransmissionString) Then
            If idPers = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("id_personal").Value, "")
                CountMasv = CountMasv + 1
            End If
            If Tabn = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("tabn").Value, "")
                CountMasv = CountMasv + 1
            End If
            If SurnameName = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("surname").Value, "")
                CountMasv = CountMasv + 1
            End If
            If Name = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("name").Value, "")
                CountMasv = CountMasv + 1
            End If
            If Patronomyc = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("patronymic").Value, "")
                CountMasv = CountMasv + 1
            End If
            If idShtat = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("id_shtat").Value, "")
                CountMasv = CountMasv + 1
            End If
            If ShtatCode = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("shtat_code").Value, "")
                CountMasv = CountMasv + 1
            End If
            If ShtatShorname = True Then
                ReDim Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("shtat_shortname").Value, "")
                CountMasv = CountMasv + 1
            End If
            If ShtatFullname = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(GetInfoIdToValue(CastToLong(rs("id_shtat").Value), "REC_SHTAT", "fullname", "charval", QueryDate, bs))
                CountMasv = CountMasv + 1
            End If
            If PodrFullname = True Then
                ReDim Masv(CountMasv)
                idPodr = CastToLong(GetInfoIdToValue(CastToLong(rs("id_shtat").Value), "REC_SHTAT", "parent_object", "intval", QueryDate, bs))
                Masv(CountMasv) = CastToString(GetInfoIdToValue(CastToLong(idPodr), "REC_PODR", "fullname", "charval", QueryDate, bs))
                CountMasv = CountMasv + 1
            End If
            If ChiefFullname = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(rs("chief_fullname").Value, "")
                CountMasv = CountMasv + 1
            End If
            Exit Do
         End If
        rs.MoveNext
    Loop
    InformationEmployee = Masv()
End Function

'Функция возвращает значение реквизита на момент времени
Public Function RekvFirm(NameBookMark As String, _
                        idFirm As Long, _
                        ObjectName As String, _
                        RekvName As String, _
                        FieldName As String, _
                        DateQuery As Date, _
                        BsName As IBusinessServer, _
                        Optional SettingName As String = "", _
                        Optional ValueName As String = "") As Integer
    Dim TextFirm As String
    TextFirm = GetInfoIdToValue(idFirm, ObjectName, RekvName, FieldName, DateQuery, BsName, SettingName, ValueName) 'RekvFirmidFirm, "REC_FIRM", "OGRN", "charval", qDate, bs)
    If Not Nvl(TextFirm, "") = "" Then
        PutToBkm NameBookMark, TextFirm
        RekvFirm = 1
    Else
        ActiveDocument.Bookmarks(NameBookMark).Select
        Selection.Font.ColorIndex = wdRed
        PutToBkm NameBookMark, CastToString("В карточке организации не заполнене реквизит" & NameBookMark)
    End If
End Function

'----------Функция изменяет формат даты-----------------------------------
Public Function MyFormatDate(ByVal d, Optional PrintMonth As Boolean = True, Optional PrintYear As Boolean = True) As String
Dim Months
Months = Array(0, "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря")
If IsNull(d) Then
    MyFormatDate = ""
Else
    MyFormatDate = IIf(Len(CastToString(Day(d))) = 1, CastToString(Day(d)), CastToString(Day(d))) & _
                   IIf(PrintMonth, " " & Months(Month(d)), "") & _
                   IIf(PrintYear, " " & CastToString(Year(d)) & " года", "")
End If
End Function

'-----------Функция, приводящая две даты к текстовому формату------------
'на вход подаются значения "дата начала" и "дата конца", в итоге выдаётся "с [дд] по [дд] [месяца] [гггг] года"
'если командировка заканчивается в следующем месяце/году, формат выходного значения меняется на:
' "с [дд] [месяца] по [дд] [месяца] [гггг] года" или "с [дд] [месяца] [гггг] года по [дд] [месяца] [гггг] года"
Public Function GetDatesPeriod(date_begin As Date, date_end As Date)
  'объявляем массив названий месяцев и строковые переменные для хранения дат
  Dim Months() As String, fromString, toString As String
  Months = Split("0,января,февраля,марта,апреля,мая,июня,июля,августа,сентября,октября,ноября,декабря", ",")
  If Day(date_begin) = 2 Then
    fromString = "со "
  Else
    fromString = "с "
  End If
  'Заполняем "дату с" - день
  fromString = fromString & CastToString(Day(date_begin))
    'Проверяем, что командировка/отпуск заканчивается в том же месяце
        If Month(date_begin) <> Month(date_end) Then
            'Если месяц не тот же, добавляем его название в строку "дата с"
            fromString = fromString & " " & Months(Month(date_begin))
        End If
    'Проверяем, что командировка/отпуск заканчивается в том же месяце
        If Year(date_begin) <> Year(date_end) Then
            'Если месяц тот же, добавляем его название в строку "дата с"
            If Month(date_begin) = Month(date_end) Then
                fromString = fromString & " " & Months(Month(date_begin))
            End If
            'Если год тот же, добавляем его в строку "дата с"
            fromString = fromString & " " & Year(date_begin) & " года"
        End If
        'Формируем строку "дата по" - всегда одинаково
        toString = " по " & CastToString(Day(date_end)) & " " & Months(Month(date_end)) & " " & Year(date_end) & " года"
        'выводим результат
        GetDatesPeriod = fromString & toString
'-----------------------------------------------------------------------------------
' проверка работы функции (запускать из кода процедуры)
'        MsgBox (GetDatesPeriod(CastToDate("30.01.2016"), CastToDate("31.01.2016")))
'        MsgBox (GetDatesPeriod(CastToDate("30.01.2016"), CastToDate("10.02.2016")))
'        MsgBox (GetDatesPeriod(CastToDate("30.12.2016"), CastToDate("11.01.2017")))
End Function

'-----------Функция, приводящая две даты к текстовому формату------------
Public Function GetDateString(dateValue As Date)
  'объявляем массив названий месяцев и строковые переменные для хранения дат
  Dim Months() As String
  Months = Split("0,января,февраля,марта,апреля,мая,июня,июля,августа,сентября,октября,ноября,декабря", ",")
  GetDateString = CastToString(DatePart("d", dateValue)) & " " & Months(Month(dateValue)) & " " & Year(dateValue)
End Function

'-----------Функция снижает регистр первого символа строки
Public Function LCaseString(AllString As String)
    If CStr(AllString) <> "" Then
        LCaseString = LCase(Left(AllString, 1)) & Right(AllString, Len(AllString) - 1)
    End If
End Function

'Возвращаемое значение - обрезанное название штатно единицы - "Начальник"
Public Function RemovePodrDouble(shtat_name As String, podr_name As String, Optional BoolLCase As Boolean = True)
    Dim shtat() As String, podr() As String, i As Integer, shtat_output As String, shtat_output_length As Integer
    shtat_name = Trim(shtat_name)
    podr_name = Trim(podr_name)
    shtat = Split(shtat_name, " ", -1)
    shtat_output_length = UBound(shtat)
    shtat_output = shtat(0)
    
    'Обрабатываем случай, если ШЕ входит в организацию без подразделения, то просто передаем наименование
    'ШЕ снижая регистр первой буквы.
    If Len(podr_name) > 0 Then
        podr = Split(podr_name, " ", -1)
        If LCase(shtat(shtat_output_length)) = LCase(GetPodrPadeg(podr(0), 2)) Then
            shtat_output_length = shtat_output_length - 1
        End If
    End If
    
    For i = 1 To shtat_output_length
    shtat_output = shtat_output & " " & shtat(i)
    Next i
    If BoolLCase = True Then
        RemovePodrDouble = LCaseString(shtat_output)
    Else
        RemovePodrDouble = shtat_output
    End If
End Function

'Возвращаемое значение - обрезанное название подразделений - "Отдел закупок Управления Гсслужбы"
Public Function SplitPodrString(podr_name As String, Optional Padeg As Long = 2, Optional SplitStr As String = "")
    Dim podr() As String, podr_output As String, podr_output_length As Integer, i As Integer
    podr_output_length = 0
    If SplitStr <> "" Then
        podr = Split(podr_name, SplitStr, -1)
        podr_output_length = UBound(podr)
    Else
        If Not UBound(Split(podr_name, "->", -1)) = 0 Then
            podr = Split(podr_name, "->", -1)
            podr_output_length = UBound(podr)
        End If
        If Not UBound(Split(podr_name, ", ", -1)) = 0 Then
            podr = Split(podr_name, ", ", -1)
            podr_output_length = UBound(podr)
        Else
            If Not UBound(Split(podr_name, ",", -1)) = 0 Then
                podr = Split(podr_name, ",", -1)
                podr_output_length = UBound(podr)
            End If
        End If
    End If
    'пробежим еще раз
    podr_output = ""
    If podr_output_length >= 1 Then
        'podr_output = podr_output & GetPostPadeg(LCase(podr(0)), 1) & " "
        podr_output = podr_output & GetPostPadeg(LCaseString(podr(0)), 1) & " "
        For i = 1 To podr_output_length - 1
            'podr_output = podr_output & GetPostPadeg(LCase(podr(i)), Padeg) & " "
            podr_output = podr_output & GetPostPadeg(LCaseString(podr(i)), Padeg) & " "
        Next i
        podr_output = podr_output & GetPostPadeg(podr(podr_output_length), Padeg)
    Else
        podr_output = podr_name
    End If
   
    'результат работы
    SplitPodrString = podr_output
End Function

'---------------------------Получение произвольного реквизита по ID----------------------------
Public Function GetInfoIdToValue(ItemId As Long, ItemBsObject As String, ItemPartObject As String, ItemValue As String, sQueryDate As Date, bs As IBusinessServer, Optional SettingCondition As String = "", Optional VariableCondition As String = "")
    'Функция возвращает строку по полю, части, объекту переданных в качестве параметра
    'на дату по указанному id
    Dim TempString As String
    Dim bo_podr As IBSDataObject, rs_podr As SKBS.SKRecordset
    Dim PodrParams As New Params

    'добавляем параметры для получения объекта приложения
    PodrParams.AddParam "id", ItemId
    PodrParams.AddParam "QueryDate", sQueryDate

    'получаем бизнес-объект
    Set bo_podr = bs.GetBsObject(ItemBsObject, PodrParams)

    'получаем нужную часть
    Set rs_podr = bo_podr(ItemPartObject)

    'проверяем, что в рекордсете есть записи, собираем нужные поля
     If Not SettingCondition = "" And Not VariableCondition = "" Then
        'Проверяем наличие строк
        If rs_podr.RecordCount > 0 Then
            'применяем фильтр
            rs_podr.SetFilter SettingCondition & "=" & QuotedStr(VariableCondition)
            'Вполне возможно, что после условия вообще не останется строк, делаем проверку
            If rs_podr.RecordCount > 0 Then
                TempString = CastToString(rs_podr(ItemValue).Value, "")
            End If
        End If
    Else
        If rs_podr.RecordCount > 0 Then
            TempString = CastToString(rs_podr(ItemValue).Value, "")
        End If
    End If

     GetInfoIdToValue = TempString
End Function

'---------------------------Усложненная функция для получения произвольного реквизита по ID----------------------------
'Возникла потребность измвлекать данные из реквизитов извлечением последнего значения с предварительной сортировкой
'Добавил два поля
' SortBy - если поле не пустое, то по нему производится сортировка, обычно это поле date_begin или date_end
' SortingMethod - принимает два значения: desc - по убыванию, asc - по возрастанию (используется по умолчанию)
Public Function DetailedGetInfoIdToValue(bs As IBusinessServer, _
                                        ItemId As Long, _
                                        ItemBsObject As String, _
                                        ItemPartObject As String, _
                                        ItemValue As String, _
                                        sQueryDate As Date, _
                                        SortBy As String, _
                                        Optional prefixSort As SortingMethod = desc, _
                                        Optional SettingCondition As String = "", _
                                        Optional VariableCondition As String = "")
    'Функция возвращает строку по полю, части, объекту переданных в качестве параметра
    'на дату по указанному id
    Dim TempString As String
    Dim bo_podr As IBSDataObject, rs_podr As SKBS.SKRecordset
    Dim PodrParams As New Params

    'добавляем параметры для получения объекта приложения
    PodrParams.AddParam "id", ItemId
    PodrParams.AddParam "QueryDate", sQueryDate

    'получаем бизнес-объект
    Set bo_podr = bs.GetBsObject(ItemBsObject, PodrParams)

    'получаем нужную часть
    Set rs_podr = bo_podr(ItemPartObject)
    
    'Добавим проверку на необходимость сортировки
    If Not CStr(SortBy) = "" Then
        If prefixSort = desc Then
            rs_podr.Sort = CStr(SortBy & " desc")
        ElseIf prefixSort = Asc Then
            rs_podr.Sort = CStr(SortBy & " asc")
        End If
    End If

    'проверяем, что в рекордсете есть записи (у сотрудника заданы
    'паспортные данные на дату QDate), собираем нужные поля
     If Not SettingCondition = "" And Not VariableCondition = "" Then
        'Проверяем наличие строк
        If rs_podr.RecordCount > 0 Then
            'применяем фильтр
            rs_podr.SetFilter SettingCondition & "=" & QuotedStr(VariableCondition)
            'Вполне возможно, что после условия вообще не останется строк, делаем проверку
            If rs_podr.RecordCount > 0 Then
                TempString = CastToString(rs_podr(ItemValue).Value, "")
            End If
        End If
    Else
        If rs_podr.RecordCount > 0 Then
            TempString = CastToString(rs_podr(ItemValue).Value, "")
        End If
    End If

     DetailedGetInfoIdToValue = TempString
End Function
'-----------------Получение ФИО Испонителя из  текущей учётной записи---------------
Public Function GetExecutorFIO(qDate As Date, bs As IBusinessServer)
    Dim surname As String, Name As String, patronymic As String
    Dim user_id As Long
    user_id = bs.CurrentUserID
    Name = CastToString(GetInfoIdToValue(user_id, "SYS_Account", "main", "name", qDate, bs))
    surname = CastToString(GetInfoIdToValue(user_id, "SYS_Account", "main", "surname", qDate, bs))
    patronymic = CastToString(GetInfoIdToValue(user_id, "SYS_Account", "main", "patronymic", qDate, bs))
    If Not Name = "" And Not surname = "" And Not patronymic = "" Then
        GetExecutorFIO = MakeFIOShortCorrectly(surname, Name, patronymic, 1, ffSurnameNP)
    Else
        GetExecutorFIO = "В карточке учетной записи не указан исполнитель"
    End If
End Function
'-----------------Получение телефона Испонителя из  текущей учётной записи---------------
Public Function GetExecutorTelephoneNumber(qDate As Date, bs As IBusinessServer)
        'Вернет телефон исполнителя
    Dim TelNumberStr As String, idIspoln As Long
    Dim TelNumberPersonal As String
    Dim idShtatIsp As Long
    Dim TelNumberShtat As String
    Dim idFirmIsp As Long
    Dim TelNumberFirm As String
    Dim user_id As Long
    user_id = bs.CurrentUserID
    TelNumberStr = ""
    'Вернет id исполнителя

    If Not CastToString(GetInfoIdToValue(user_id, "SYS_Account", "main", "id", qDate, bs, "", "")) = "" Then
        idIspoln = CastToLong(GetInfoIdToValue(user_id, "SYS_Account", "main", "id_personal", qDate, bs, "", ""))
        'если нашли id исполнителя
        If Not idIspoln = -1 Then
            'вернет телефон сотр. из личной карточки из вкладки - Прочие документы - Контакты
            TelNumberPersonal = GetInfoIdToValue(idIspoln, "REC_PERSONAL", "contacts", "charval", qDate, bs, "code", "01")
            'Если телефон не записан ищем в ШЕ
            If TelNumberPersonal = "" Then
                'вернет id штатной единицы исполнителя
                idShtatIsp = CastToLong(GetInfoIdToValue(idIspoln, "REC_PERSONAL", "EXECPOST", "id_shtat", qDate, bs, "Work_workerstatus_code", "1"))
                'вернет телефон из карточки штатной единицы
                TelNumberShtat = GetInfoIdToValue(idShtatIsp, "REC_SHTAT", "telephone", "charval", qDate, bs, "", "")
                'Если телефон не записан в карточки ШЕ
                If TelNumberShtat = "" Then
                    'вернет id организации
                    idFirmIsp = CastToLong(GetInfoIdToValue(idIspoln, "REC_PERSONAL", "EXECPOST", "Work_Firm_id", qDate, bs, "Work_workerstatus_code", "1"))
                    'вернет телефон из карточки организации
                    TelNumberFirm = GetInfoIdToValue(idFirmIsp, "REC_FIRM", "contacts", "charval", qDate, bs, "", "")
                    'Если телефон не записан в карточке организации
                    If TelNumberFirm = "" Then
                        TelNumberStr = "В карточке организации, штатной единицы, личной карточки сотр. отсутствует номер телефона"
                    Else
                        'Выводим номер телефона из карточки Организации
                        TelNumberPersonal = TelNumberFirm
                    End If
                Else
                    'Выводим номер телефона из карточки ШЕ
                    TelNumberPersonal = TelNumberShtat
                End If
            End If
        Else
            TelNumberStr = "В карточке учетной записи не указан исполнитель"
        End If
    Else
        TelNumberStr = "Отсутствует запись об исполнителе"
    End If

    If Not TelNumberStr = "" Then
        GetExecutorTelephoneNumber = CastToString(TelNumberStr)
    Else
        GetExecutorTelephoneNumber = CastToString(TelNumberPersonal)
    End If
End Function

Public Function GetPersonalSex(sotr_id As Long, qDate As Date, bs As BusinessServer)
If GetInfoIdToValue(sotr_id, "REC_PERSONAL", "sex", "Text", qDate, bs) = "М" Then
    GetPersonalSex = True
  Else
    GetPersonalSex = False
  End If
End Function

Public Function GetPersonalSexMG(sotr_id As Long, qDate As Date, bs As BusinessServer)
If GetInfoIdToValue(sotr_id, "REC_PERSONAL", "sex", "Text", qDate, bs) = "М" Then
    GetPersonalSexMG = "М"
  Else
    GetPersonalSexMG = "Ж"
  End If
End Function

Public Function WritePersonalSex(sex As Boolean, Optional MaleString As String = "его", Optional FemaleString As String = "её")
    If sex Then
        WritePersonalSex = MaleString
    Else
        WritePersonalSex = FemaleString
    End If
End Function

'-------------------Представление ФИО в правильном формате (а не как в сервисных функциях)-------------------
Public Function MakeFIOShortCorrectly(surname As String, Name As String, patronymic As String, Optional Padeg As Long = 1, Optional FIOFormat As FIOFormatEnum = ffSurnameNamePatronomic, Optional sotrSexIfNoPatronymic As String = "")
Dim FIO() As String
Dim Result As String
'если есть отчество
If patronymic <> "" Then
    Result = GetFIO_Padeg(surname, Name, patronymic, ffSurnameNamePatronomic, Padeg)
    Select Case FIOFormat
    Case ffNPSurname
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = Left(Name, 1) & "." & Left(patronymic, 1) & ". " & FIO(0)
    Case ffSurnameNP
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = FIO(0) & " " & Left(Name, 1) & "." & Left(patronymic, 1) & "."
    Case Else
        MakeFIOShortCorrectly = Result
    End Select
Else
'если нет отчества
Result = GetFIO_Padeg(surname, Name, "", ffSurnameNamePatronomic, Padeg)
    If sotrSexIfNoPatronymic = "М" Then
        Result = GetFIO_Padeg(surname, Name, "Иванович", ffSurnameNamePatronomic, Padeg)
    End If
    If sotrSexIfNoPatronymic = "Ж" Then
        Result = GetFIO_Padeg(surname, Name, "Ивановна", ffSurnameNamePatronomic, Padeg)
    End If
    Select Case FIOFormat
    Case ffNPSurname
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = Left(Name, 1) & "." & FIO(0)
    Case ffSurnameNP
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = FIO(0) & " " & Left(Name, 1) & "."
    Case Else
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = FIO(0) & " " & FIO(1)
    End Select
End If
End Function

'-------------------Передаём строку ФИО, получаем назад в нужном падеже и сокращении-------------------
Public Function MakeFIOShortOneString(InitialFIO As String, Optional Padeg As Long = 1, Optional FIOFormat As FIOFormatEnum = ffSurnameNamePatronomic, Optional SplitStr As String = " ")
Dim FIO() As String

FIO = Split(InitialFIO, SplitStr, -1)

If UBound(FIO) = 0 Then
    MakeFIOShortOneString = "Введено некорректное ФИО"
Else
    MakeFIOShortOneString = MakeFIOShortCorrectly(FIO(0), FIO(1), Nvl(FIO(2), ""), Padeg, FIOFormat)
End If

End Function

'-------------------Удаляем пробелы из строки-------------------
Public Function RemoveSpaces(InitialString As String, Optional IsShortFIO As Boolean = False)
Dim TempString() As String, i As Integer, OutputStr As String
TempString = Split(InitialString, " ", -1)
OutputStr = ""
For i = 0 To UBound(TempString)
    If i > 1 And IsShortFIO Then
        OutputStr = OutputStr & " " & TempString(i)
    Else
        OutputStr = OutputStr & TempString(i)
    End If
Next i
    RemoveSpaces = OutputStr
End Function

'-------------------получение строки с дbректором подразделения по штатной единице сотрудника--------------
Public Function Replace_DirectorPodr(NumbOption As Integer, NameOption As String) As String
    'Слитый реквизит - shtat_podr_info
    '_0='Object_Code_charval=''01'';
    'ShortName_charval='Название подразделения';
    'fullname_charval='Полное название подразделения';
    'Director_FIO_Personal='Фамилия Имя отчество';
    'Director_intval=id_сотрудника;
    'Podr_Director_ExecPost_Shtat_Shortname='Должность сотрудника';';
    
    If Not Nvl(NameOption, "") = "" Then
        Dim prm_p As New SKGENERALLib.Params
        prm_p.LoadFromString CastToString(NameOption) '
        Dim WrdArray1() As String
        Dim WrdArray2() As String
        Dim i As Integer
        For i = 0 To prm_p.Count - 1
            If Not prm_p.GetValue("_" & CastToString(i), "") = -1 Then
                WrdArray1() = Split(prm_p.GetValue("_" & CastToString(i), ""), ";")
                WrdArray2() = Split(WrdArray1(NumbOption), "=")
            End If
        Next i
        If Not CastToString(WrdArray2(1)) = "''" Or Not CastToString(WrdArray2(1)) = "Null" Then
            Replace_DirectorPodr = CastToString(Replace(WrdArray2(1), "'", ""))
        Else
            Replace_DirectorPodr = "Null"
        End If
    Else
        Replace_DirectorPodr = "Null"
    End If
End Function

'-------------------форматирование строки с дректором подразделения по штатной единице сотрудника--------------
Public Function getDirectorPodrAndFIO(shtat_podr_info As String)
    Dim FIOArray() As String
    'Извлекаем ФИО руководителя
    FIOArray = Split(CastToString(Replace_DirectorPodr(3, shtat_podr_info)))
    'Расширяем массив до трех элементов с сохранением содержимого
    ReDim Preserve FIOArray(4)
    'Извлекаем наименование должности
    FIOArray(4) = Replace_DirectorPodr(5, shtat_podr_info)
    getDirectorPodrAndFIO = FIOArray
End Function

'-------------------получение ФИО директора структурного подразделения-------------------
Public Function getDirectiorPodrFIO(shtat_podr_info As String, FIOFormat As FIOFormatEnum, Optional soglFIO As String = "")

Dim directorPodrData() As String
directorPodrData = getDirectorPodrAndFIO(shtat_podr_info)
If Not CastToString(directorPodrData) = "" Then
Select Case FIOFormat
      Case ffNPSurname
             getDirectiorPodrFIO = CastToString(Mid(directorPodrData(1), 1, 1) & "." & Mid(directorPodrData(2), 1, 1) & ". " & directorPodrData(0))
      Case ffSurnameNamePatronomic
             getDirectiorPodrFIO = CastToString(directorPodrData(0) & " " & directorPodrData(1) & " " & directorPodrData(2))
      Case ffSurnameNP
             getDirectiorPodrFIO = CastToString(directorPodrData(0) & " " & Mid(directorPodrData(1), 1, 1) & "." & Mid(directorPodrData(2), 1, 1) & ".")
      Case Else
             getDirectiorPodrFIO = CastToString(Mid(directorPodrData(1), 1, 1) & "." & Mid(directorPodrData(2), 1, 1) & ". " & directorPodrData(0))
End Select
Else
    getDirectiorPodrFIO = soglFIO
End If
End Function
'-------------------получение названия структурного сотрудника-------------------
Public Function getDirectiorPodrPodr(shtat_podr_info As String, Optional soglDolgn As String = "")
    Dim directorPodrData() As String
    directorPodrData = getDirectorPodrAndFIO(shtat_podr_info)
    If Not CastToString(directorPodrData(4)) = "" Then
        getDirectiorPodrPodr = directorPodrData(4)
    Else
        getDirectiorPodrPodr = soglDolgn
    End If
End Function

' Функция возвращает значение является ли сотрудник госслужащим или нет
Public Function GetPersonalGSStatusByShtatId(shtatId As Long, qDate As Date, bs As IBusinessServer)
    Dim postCodeAnswer As String
    postCodeAnswer = GetInfoIdToValue(shtatId, "REC_SHTAT", "post", "code", qDate, bs)
    If GetPostCategory_GS(postCodeAnswer) <> "" Then
        GetPersonalGSStatusByShtatId = True 'является госслужащим
    Else
        GetPersonalGSStatusByShtatId = False 'не является госслужащим
    End If
End Function

' Функция возвращает значение, является ли сотрудник госслужащим, служащим, работником
Public Function GetPersonalStatusByShtatId(shtatId As Long, qDate As Date, bs As IBusinessServer)
    Dim postCodeAnswer As String, profCodeAnswer As String
    Dim MaskCodeGS As String, MaskCodeSotr As String, MaskCodeJob As String
    MaskCodeGS = "##-#-#-###" 'маска для ГС
    MaskCodeSotr = "2####" 'маска для служащих
    postCodeAnswer = GetInfoIdToValue(shtatId, "REC_SHTAT", "post", "code", qDate, bs)
    If Not postCodeAnswer = "" Then
        If postCodeAnswer Like MaskCodeGS Then
            GetPersonalStatusByShtatId = 1 'является госслужащим
        ElseIf postCodeAnswer Like MaskCodeSotr Then
            GetPersonalStatusByShtatId = 2 'является служащим
        Else
            GetPersonalStatusByShtatId = 0 'иначе ошибка
        End If
    Else
        MaskCodeJob = "1####" 'маска для работников
        profCodeAnswer = GetInfoIdToValue(shtatId, "REC_SHTAT", "prof", "code", qDate, bs)
        If profCodeAnswer <> "" And profCodeAnswer Like MaskCodeJob Then
            GetPersonalStatusByShtatId = 3 'является работником
        Else
            GetPersonalStatusByShtatId = 0 'иначе ошибка 100%
        End If
    End If
End Function
