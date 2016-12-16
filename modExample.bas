Attribute VB_Name = "modExample"
'    'дата приказа
'    Dim qDate As Date
'    qDate = CastToDate(MainRS("date_begin").Value)

'    'Табельный номер
'    If Not IsNull(DataRS("Personal_tabn").Value) Then
'      PutToBkm "tabn", CastToString(DataRS("Personal_tabn").Value)
'    End If
'
'    'Дата приказа
'    If Not IsNull(MainRS("date_begin").Value) Then
'      PutToBkm "date_order", MyFormatDate(CastToString(MainRS("date_begin").Value))
'    End If
'
'    'Номер приказа
'    If Not IsNull(MainRS("number").Value) Then
'      PutToBkm "order_number", CastToString(MainRS("number").Value)
'    End If
'
'    'подразделение
'    Dim NamePodr As String, NamePost As String
'    NamePodr = SplitPodrString(CastToString(DataRS("shtat_shtat_path").Value))
'    PutToBkm "podr", GetPodrPadeg(NamePodr, 2)

'    'Выводим ФИО сотрудника
'    PutToBkm "FIO", MakeFIOShortOneString(CastToString(DataRS("personal_fio").Value), 1, ffSurnameNamePatronomic)
'    PutToBkm "FIO2", MakeFIOShortOneString(CastToString(DataRS("personal_fio").Value), 5, ffSurnameNamePatronomic)
'
'    'Должность
'    NamePost = RemovePodrDouble(CastToString(DataRS("execpost_shortname").Value), NamePodr)
'    PutToBkm "execpost", GetPostPadeg(NamePost, 1)

'     'паспорт
'    Dim tmpS As String
'    tmpS = ReturnPersonalPasport(CastToLong(DataRS("id").Value), DataRS, qDate, bs)
'    If Not CastToString(tmpS) = "" Then
'        PutToBkm "pasport", CastToString(tmpS)
'    Else
'        ActiveDocument.Bookmarks("pasport").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "pasport", CastToString("В карточке сотрудника не указан паспорт")
'    End If

'    ' Компенсация
'    If (Not IsNull(DataRS("pay").Value) And (CastToLong(DataRS("pay").Value)) > 0) Then
'        PutToBkm "pays", CastToDouble(DataRS("pay").Value) & " " & gennumbercase(Fix(CastToDouble(DataRS("pay").Value)), "рубль", "рубля", "рублей")
'    Else
'        PutToBkm "pays", "[Не указана сумма компенсации]"
'    End If
'
'    'ним\ней
'    Dim TextStr1
'    TextStr1 = GetRekvDataFromDataCacheQE(bs, "REC_PERSONAL", DataRS("id").Value, "sex,boolval", DataRS("qdate").Value)
'    If UBound(TextStr1, 2) >= 0 Then
'      PutToBkm "it_her", IIf(CastToBool(TextStr1(1, 0), False), "л", "ла")
'    Else
'      PutToBkm "it_her", "л"
'    End If

'    'Руководитель структурного подразделения
'    If Not IsNull(DataRS("SHTAT_PODR_INFO").Value) Then
'        PutToBkm "FioDirPodr", CastToString(ReturnChiefsDepartmentFIO(DataRS, ffSurnameNP, Rod))
'        PutToBkm "ShtatDirPodr", LCaseString(CastToString(ReturnChiefsDepartmentName(DataRS, DepartmentChiefPostShortName, Rod)))
'        PutToBkm "PodrDirPodr", CastToString(ReturnChiefsDepartmentName(DataRS, DepartmentFullName, Rod))
'    End If

'    'ФИО исполнителя
'    Dim FioIsp As String
'    FioIsp = GetExecutorFIO(qDate, bs)
'    If Not IsNull(FioIsp) Then
'        PutToBkm "isp", CastToString(FioIsp)
'    Else
'        ActiveDocument.bookmarks("isp").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "isp", CastToString("В карточке учетной записи не указан сотрудник")
'    End If
'
'    'Телефон исполнителя
'    Dim TelIsp As String
'    TelIsp = GetExecutorTelephoneNumber(qDate, bs)
'    If Not IsNull(TelIsp) Then
'        PutToBkm "telisp", CastToString(TelIsp)
'    Else
'        ActiveDocument.bookmarks("telisp").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "telisp", CastToString("В карточке сотрудника не указан телефон")
'    End If

'    'Краткое наименование организации
'    Dim qDate As Date
'    Dim idFirm As Long
'    Dim NameFirm As String
'    qDate = CastToDate(MainRS("date_begin").Value)
'    idFirm = MainRS.Fields("id_firm").Value
'    NameFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "shortname", "charval", qDate, bs, "", "")
'    If Not Nvl(NameFirm, "") = "" Then
'        PutToBkm "firm_shortname", NameFirm
'        PutToBkm "firm_shortname2", NameFirm
'    Else
'        ActiveDocument.Bookmarks("firm_shortname").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_shortname", CastToString("В карточке организации не указано краткое название организации")
'        ActiveDocument.Bookmarks("firm_shortname2").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_shortname2", CastToString("В карточке организации не указано краткое название организации")
'    End If
'
'    'Полное наименование организации
'    NameFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "fullname", "charval", qDate, bs, "", "")
'    If Not Nvl(NameFirm, "") = "" Then
'        PutToBkm "firm_fullname", NameFirm
'    Else
'        ActiveDocument.Bookmarks("firm_fullname").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_fullname", CastToString("В карточке организации не указано краткое название организации")
'    End If
'
'    'Полный почтовый адрес
'    Dim AddressFirm As String
'    AddressFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "firm_address", "address_text_brief", qDate, bs, "type_code", "00001")
'    If Not Nvl(AddressFirm, "") = "" Then
'        PutToBkm "firm_address", AddressFirm
'    Else
'        ActiveDocument.Bookmarks("firm_address").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_address", CastToString("В карточке организации не указан полный почтовый адрес")
'    End If
'
'    'основание действий руководителя компании
'    Dim EmailFirm As String
'    EmailFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "otherinfo", "fullname_typeotherinfo", qDate, bs, "code_typeotherinfo", "02")
'    If Not Nvl(EmailFirm, "") = "" Then
'        PutToBkm "firm_osn", EmailFirm
'    Else
'        ActiveDocument.Bookmarks("firm_osn").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_osn", CastToString("В карточке организации не указано - Основание действий руководителя")
'    End If
'
'    'ИНН компании
'    Dim INNFirm As String
'    INNFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "INN", "charval", qDate, bs)
'    If Not Nvl(INNFirm, "") = "" Then
'        PutToBkm "firm_INN", INNFirm
'    Else
'        ActiveDocument.Bookmarks("firm_INN").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_INN", CastToString("В карточке организации не указан ИНН")
'    End If
'  End If

'    'Пример использования функции,
'    'необходимо задать имена закладок в следующем виде
'    '    ГлЮристchief_fullname
'    '    ГлЮристffNPSurname
'    '    ДирУпрПерсchief_fullname
'    '    ДирУпрПерсffNPSurname
'    '    УправляющийffNPSurname
'    '    УправляющийPost
'    'Обязательно! выполнение данного кода, необходимо выполнять в последнюю очередь
'    If HeaderRS.RecordCount > 0 Then
'        HeaderRS.MoveFirst
'        'Заполнение руководителей которые подписывают документ
'        Call FieldRecordingChiefs(HeaderRS, qDate, bs)
'    End If
