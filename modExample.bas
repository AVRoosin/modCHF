Attribute VB_Name = "modExample"
'    '���� �������
'    Dim qDate As Date
'    qDate = CastToDate(MainRS("date_begin").Value)

'    '��������� �����
'    If Not IsNull(DataRS("Personal_tabn").Value) Then
'      PutToBkm "tabn", CastToString(DataRS("Personal_tabn").Value)
'    End If
'
'    '���� �������
'    If Not IsNull(MainRS("date_begin").Value) Then
'      PutToBkm "date_order", MyFormatDate(CastToString(MainRS("date_begin").Value))
'    End If
'
'    '����� �������
'    If Not IsNull(MainRS("number").Value) Then
'      PutToBkm "order_number", CastToString(MainRS("number").Value)
'    End If
'
'    '�������������
'    Dim NamePodr As String, NamePost As String
'    NamePodr = SplitPodrString(CastToString(DataRS("shtat_shtat_path").Value))
'    PutToBkm "podr", GetPodrPadeg(NamePodr, 2)

'    '������� ��� ����������
'    PutToBkm "FIO", MakeFIOShortOneString(CastToString(DataRS("personal_fio").Value), 1, ffSurnameNamePatronomic)
'    PutToBkm "FIO2", MakeFIOShortOneString(CastToString(DataRS("personal_fio").Value), 5, ffSurnameNamePatronomic)
'
'    '���������
'    NamePost = RemovePodrDouble(CastToString(DataRS("execpost_shortname").Value), NamePodr)
'    PutToBkm "execpost", GetPostPadeg(NamePost, 1)

'     '�������
'    Dim tmpS As String
'    tmpS = ReturnPersonalPasport(CastToLong(DataRS("id").Value), DataRS, qDate, bs)
'    If Not CastToString(tmpS) = "" Then
'        PutToBkm "pasport", CastToString(tmpS)
'    Else
'        ActiveDocument.Bookmarks("pasport").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "pasport", CastToString("� �������� ���������� �� ������ �������")
'    End If

'    ' �����������
'    If (Not IsNull(DataRS("pay").Value) And (CastToLong(DataRS("pay").Value)) > 0) Then
'        PutToBkm "pays", CastToDouble(DataRS("pay").Value) & " " & gennumbercase(Fix(CastToDouble(DataRS("pay").Value)), "�����", "�����", "������")
'    Else
'        PutToBkm "pays", "[�� ������� ����� �����������]"
'    End If
'
'    '���\���
'    Dim TextStr1
'    TextStr1 = GetRekvDataFromDataCacheQE(bs, "REC_PERSONAL", DataRS("id").Value, "sex,boolval", DataRS("qdate").Value)
'    If UBound(TextStr1, 2) >= 0 Then
'      PutToBkm "it_her", IIf(CastToBool(TextStr1(1, 0), False), "�", "��")
'    Else
'      PutToBkm "it_her", "�"
'    End If

'    '������������ ������������ �������������
'    If Not IsNull(DataRS("SHTAT_PODR_INFO").Value) Then
'        PutToBkm "FioDirPodr", CastToString(ReturnChiefsDepartmentFIO(DataRS, ffSurnameNP, Rod))
'        PutToBkm "ShtatDirPodr", LCaseString(CastToString(ReturnChiefsDepartmentName(DataRS, DepartmentChiefPostShortName, Rod)))
'        PutToBkm "PodrDirPodr", CastToString(ReturnChiefsDepartmentName(DataRS, DepartmentFullName, Rod))
'    End If

'    '��� �����������
'    Dim FioIsp As String
'    FioIsp = GetExecutorFIO(qDate, bs)
'    If Not IsNull(FioIsp) Then
'        PutToBkm "isp", CastToString(FioIsp)
'    Else
'        ActiveDocument.bookmarks("isp").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "isp", CastToString("� �������� ������� ������ �� ������ ���������")
'    End If
'
'    '������� �����������
'    Dim TelIsp As String
'    TelIsp = GetExecutorTelephoneNumber(qDate, bs)
'    If Not IsNull(TelIsp) Then
'        PutToBkm "telisp", CastToString(TelIsp)
'    Else
'        ActiveDocument.bookmarks("telisp").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "telisp", CastToString("� �������� ���������� �� ������ �������")
'    End If

'    '������� ������������ �����������
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
'        PutToBkm "firm_shortname", CastToString("� �������� ����������� �� ������� ������� �������� �����������")
'        ActiveDocument.Bookmarks("firm_shortname2").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_shortname2", CastToString("� �������� ����������� �� ������� ������� �������� �����������")
'    End If
'
'    '������ ������������ �����������
'    NameFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "fullname", "charval", qDate, bs, "", "")
'    If Not Nvl(NameFirm, "") = "" Then
'        PutToBkm "firm_fullname", NameFirm
'    Else
'        ActiveDocument.Bookmarks("firm_fullname").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_fullname", CastToString("� �������� ����������� �� ������� ������� �������� �����������")
'    End If
'
'    '������ �������� �����
'    Dim AddressFirm As String
'    AddressFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "firm_address", "address_text_brief", qDate, bs, "type_code", "00001")
'    If Not Nvl(AddressFirm, "") = "" Then
'        PutToBkm "firm_address", AddressFirm
'    Else
'        ActiveDocument.Bookmarks("firm_address").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_address", CastToString("� �������� ����������� �� ������ ������ �������� �����")
'    End If
'
'    '��������� �������� ������������ ��������
'    Dim EmailFirm As String
'    EmailFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "otherinfo", "fullname_typeotherinfo", qDate, bs, "code_typeotherinfo", "02")
'    If Not Nvl(EmailFirm, "") = "" Then
'        PutToBkm "firm_osn", EmailFirm
'    Else
'        ActiveDocument.Bookmarks("firm_osn").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_osn", CastToString("� �������� ����������� �� ������� - ��������� �������� ������������")
'    End If
'
'    '��� ��������
'    Dim INNFirm As String
'    INNFirm = GetInfoIdToValue(idFirm, "REC_FIRM", "INN", "charval", qDate, bs)
'    If Not Nvl(INNFirm, "") = "" Then
'        PutToBkm "firm_INN", INNFirm
'    Else
'        ActiveDocument.Bookmarks("firm_INN").Select
'        Selection.Font.ColorIndex = wdRed
'        PutToBkm "firm_INN", CastToString("� �������� ����������� �� ������ ���")
'    End If
'  End If

'    '������ ������������� �������,
'    '���������� ������ ����� �������� � ��������� ����
'    '    �������chief_fullname
'    '    �������ffNPSurname
'    '    ����������chief_fullname
'    '    ����������ffNPSurname
'    '    �����������ffNPSurname
'    '    �����������Post
'    '�����������! ���������� ������� ����, ���������� ��������� � ��������� �������
'    If HeaderRS.RecordCount > 0 Then
'        HeaderRS.MoveFirst
'        '���������� ������������� ������� ����������� ��������
'        Call FieldRecordingChiefs(HeaderRS, qDate, bs)
'    End If
