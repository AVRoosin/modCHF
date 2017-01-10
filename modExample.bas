Attribute VB_Name = "modExample"

'Last update Fidel on 10.01.2017

'    'Руководитель структурного подразделения
'    Dim FioChiefs As String, PostChiefs As String, PodrChiefs As String
'    If Not IsNull(DataRS("SHTAT_PODR_INFO").Value) Then
'        FioChiefs = CStr(ReturnChiefsDepartmentFIO(DataRS, ffSurnameNP, Rod))
'        If FioChiefs <> "" Then
'            PutToBkm "FioDirPodr", CStr(FioChiefs)
'        Else
'            PutToBkm "FioDirPodr", "[В карточке подразделения отсутствует руководитель]"
'        End If
'
'        PostChiefs = CStr(ReturnChiefsDepartmentName(DataRS, DepartmentChiefPostShortName, Rod))
'        If PostChiefs <> "" Then
'            PutToBkm "ShtatDirPodr", LCaseString(CStr(PostChiefs))
'        Else
'            PutToBkm "ShtatDirPodr", "[В карточке подразделения отсутствует руководитель]"
'        End If
'
'        PodrChiefs = CStr(ReturnChiefsDepartmentName(DataRS, DepartmentFullName, Rod))
'        If PodrChiefs <> "" Then
'            PutToBkm "PodrDirPodr", CStr(PodrChiefs)
'        Else
'            PutToBkm "PodrDirPodr", "[В карточке подразделения отсутствует руководитель]"
'        End If
'    End If

'   'Должность руководителя в основании ПФ
'    Dim ChiefFIO As String, ChiefNameExecPost As String
'    ChiefFIO = MakeFIOShortOneString(CastToString(ReturnChiefCustomValue(HeaderRS, qDate, bs, "ДирУпрПерс", "ffSurnameNamePatronomic")), 2, ffSurnameNamePatronomic)
'    ChiefNameExecPost = GetPostPadeg(CastToString(ReturnChiefCustomValue(HeaderRS, qDate, bs, "ДирУпрПерс", "chief_fullname")), 2)
'    'должность ДирУпрПерс в основании
'    If Not IsNull(ChiefNameExecPost) Then
'      PutToBkm "ChiefNameExecPost", CastToString(ChiefNameExecPost)
'    End If
'
'    'ФИО ДирУпрПерс в основании
'    If Not IsNull(ChiefFIO) Then
'      PutToBkm "ChiefFIO", CastToString(ChiefFIO)
'    End If

'    'паспорт
'    Dim tmpS As String
'    tmpS = ReturnPersonalPasport(CastToLong(DataRS("id").Value), DataRS, qDate, bs)
'    If Not CastToString(tmpS) = "" Then
'        PutToBkm "pasport", CastToString(tmpS)
'    Else
'        ActiveDocument.Bookmarks("pasport").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "pasport", CastToString("В карточке сотрудника не указан паспорт")
'    End If

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

'    'Изменение слова в зависимости от числа, например, рубль/рубля/рублей
'    ' Fix - отбрасывает дробную часть числа
'    'gennumbercase(Fix(CastToDouble(DataRS("pay").Value)), "рубль", "рубля", "рублей")

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

' 'КПП компании
'    Dim KppFirm As String
'    KppFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "INN", "charval2", qDate, bs)
'    If Not Nvl(KppFirm, "") = "" Then
'        PutToBkm "firm_KPP", KppFirm
'    Else
'        ActiveDocument.Bookmarks("firm_KPP").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_KPP", CastToString("В карточке организации не указан КПП")
'    End If

'        'БанковскСчет организации
'        Dim NameBank As String
'        NameBank = GetInfoIdToValue(idFirm, "REC_FIRM", "bank", "fullname", qDate, bs)
'        If Not Nvl(NameBank, "") = "" Then
'            PutToBkm "NameBank", NameBank
'        Else
'            ActiveDocument.Bookmarks("NameBank").Select
'            Selection.Font.ColorIndex = wdRed
'            PutToBkm "NameBank", CastToString("В карточке организации не банк")
'        End If
'
'        'БанковскСчет организации
'        Dim SchetFirm As String
'        SchetFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "bank", "charval", qDate, bs, "fullname", "Сберегательный банк РФ")
'        If Not Nvl(SchetFirm, "") = "" Then
'            PutToBkm "firm_schet", SchetFirm
'        Else
'            ActiveDocument.Bookmarks("firm_schet").Select
'            Selection.Font.ColorIndex = wdRed
'            PutToBkm "firm_schet", CastToString("В карточке организации не указан счет")
'        End If
'
'         'Корреспондирующий счет организации
'        Dim CorrespondingScore As String
'        CorrespondingScore = GetInfoIdToValue(idFirm, "REC_FIRM", "bank", "charval", qDate, bs, "fullname", "Сберегательный банк РФ")
'        If Not Nvl(CorrespondingScore, "") = "" Then
'            PutToBkm "firm_CorrespondingScore", CorrespondingScore
'        Else
'            ActiveDocument.Bookmarks("firm_CorrespondingScore").Select
'            Selection.Font.ColorIndex = wdRed
'            PutToBkm "firm_CorrespondingScore", CastToString("В карточке организации не указан корреспондирующий счет")
'        End If
'
'        'БИК организации
'        Dim BikFirm As String
'        BikFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "bank", "code", qDate, bs)
'        If Not Nvl(BikFirm, "") = "" Then
'            PutToBkm "firm_BIK", BikFirm
'        Else
'            ActiveDocument.Bookmarks("firm_BIK").Select
'            Selection.Font.ColorIndex = wdRed
'            PutToBkm "firm_BIK", CastToString("В карточке организации не указан БИК")
'        End If

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
