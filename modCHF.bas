Attribute VB_Name = "modCHF"
Option Explicit
Option Compare Text

'last update 14.12.2016
'last update 14.12.2016 test

'������� ���������� ������� ���������� ��������� � ������ ������� �����/�����/��� �����/����� �����
'� ������ ���� ����� ����� �������� ������
Public Function ReturnPersonalPasport(IdPersonal As Long, _
                                    ByVal RS As ADODB.Recordset, _
                                    QueryDate As Date, _
                                    bs As IBusinessServer, _
                                    Optional StyleFormat As Integer = 1) As String
    Dim PasportString As String
    Dim TempString As String, TempArray, TexpTempArray, CountTemp As Integer
    TempArray = Array("serdoc", "numdoc", "whogive", "date_begin")
    If StyleFormat = 1 Then
        TexpTempArray = Array("�����", "�����", "��� �����", "���� ������")
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

'------------------������� ��������� ��������� �������� �� ���� Chief, ��������, ������ �������, ���
Public Function ReturnChiefCustomValue(ByVal RS As ADODB.Recordset, _
                                    QueryDate As Date, _
                                    bs As IBusinessServer, _
                                    NameChief As String, _
                                    ValueChief As String) As String
    RS.MoveFirst
    Do While Not RS.EOF
        If Nvl(CastToString(RS("chief_code").Value), "") = CastToString(NameChief) Then
            ReturnChiefCustomValue = CastToString(RS(ValueChief).Value)
            RS.MoveLast
            RS.MoveNext
        Else
            RS.MoveNext
        End If
    Loop
End Function

Public Function InformationEmployee(ByVal RS As ADODB.Recordset, QueryDate As Date, bs As IBusinessServer, TransmissionString As String, _
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
    '���������� ��������� �������������
    Dim Masv() As String
    Dim idPodr As Long, CountMasv As Integer
    RS.MoveFirst
    CountMasv = 0
    Do While Not RS.EOF
        If Nvl(CastToString(RS("chief_code").Value), "") = CastToString(TransmissionString) Then
            If idPers = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("id_personal").Value, "")
                CountMasv = CountMasv + 1
            End If
            If Tabn = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("tabn").Value, "")
                CountMasv = CountMasv + 1
            End If
            If SurnameName = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("surname").Value, "")
                CountMasv = CountMasv + 1
            End If
            If Name = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("name").Value, "")
                CountMasv = CountMasv + 1
            End If
            If Patronomyc = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("patronymic").Value, "")
                CountMasv = CountMasv + 1
            End If
            If idShtat = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("id_shtat").Value, "")
                CountMasv = CountMasv + 1
            End If
            If ShtatCode = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("shtat_code").Value, "")
                CountMasv = CountMasv + 1
            End If
            If ShtatShorname = True Then
                ReDim Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("shtat_shortname").Value, "")
                CountMasv = CountMasv + 1
            End If
            If ShtatFullname = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(GetInfoIdToValue(CastToLong(RS("id_shtat").Value), "REC_SHTAT", "fullname", "charval", QueryDate, bs))
                CountMasv = CountMasv + 1
            End If
            If PodrFullname = True Then
                ReDim Masv(CountMasv)
                idPodr = CastToLong(GetInfoIdToValue(CastToLong(RS("id_shtat").Value), "REC_SHTAT", "parent_object", "intval", QueryDate, bs))
                Masv(CountMasv) = CastToString(GetInfoIdToValue(CastToLong(idPodr), "REC_PODR", "fullname", "charval", QueryDate, bs))
                CountMasv = CountMasv + 1
            End If
            If ChiefFullname = True Then
                ReDim Preserve Masv(CountMasv)
                Masv(CountMasv) = CastToString(RS("chief_fullname").Value, "")
                CountMasv = CountMasv + 1
            End If
            Exit Do
         End If
        RS.MoveNext
    Loop
    InformationEmployee = Masv()
End Function

'������� ���������� �������� ��������� �� ������ �������
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
        PutToBkm NameBookMark, CastToString("� �������� ����������� �� ��������� ��������" & NameBookMark)
    End If
End Function

'----------������� �������� ������ ����-----------------------------------
Public Function MyFormatDate(ByVal d, Optional PrintMonth As Boolean = True, Optional PrintYear As Boolean = True) As String
Dim Months
Months = Array(0, "������", "�������", "�����", "������", "���", "����", "����", "�������", "��������", "�������", "������", "�������")
If IsNull(d) Then
    MyFormatDate = ""
Else
    MyFormatDate = IIf(Len(CastToString(Day(d))) = 1, CastToString(Day(d)), CastToString(Day(d))) & _
                   IIf(PrintMonth, " " & Months(Month(d)), "") & _
                   IIf(PrintYear, " " & CastToString(Year(d)) & " ����", "")
End If
End Function

'-----------�������, ���������� ��� ���� � ���������� �������------------
'�� ���� �������� �������� "���� ������" � "���� �����", � ����� ������� "� [��] �� [��] [������] [����] ����"
'���� ������������ ������������� � ��������� ������/����, ������ ��������� �������� �������� ��:
' "� [��] [������] �� [��] [������] [����] ����" ��� "� [��] [������] [����] ���� �� [��] [������] [����] ����"
Public Function GetDatesPeriod(date_begin As Date, date_end As Date)
  '��������� ������ �������� ������� � ��������� ���������� ��� �������� ���
  Dim Months() As String, fromString, toString As String
  Months = Split("0,������,�������,�����,������,���,����,����,�������,��������,�������,������,�������", ",")
  If Day(date_begin) = 2 Then
    fromString = "�� "
  Else
    fromString = "� "
  End If
  '��������� "���� �" - ����
  fromString = fromString & CastToString(Day(date_begin))
    '���������, ��� ������������/������ ������������� � ��� �� ������
        If Month(date_begin) <> Month(date_end) Then
            '���� ����� �� ��� ��, ��������� ��� �������� � ������ "���� �"
            fromString = fromString & " " & Months(Month(date_begin))
        End If
    '���������, ��� ������������/������ ������������� � ��� �� ������
        If Year(date_begin) <> Year(date_end) Then
            '���� ����� ��� ��, ��������� ��� �������� � ������ "���� �"
            If Month(date_begin) = Month(date_end) Then
                fromString = fromString & " " & Months(Month(date_begin))
            End If
            '���� ��� ��� ��, ��������� ��� � ������ "���� �"
            fromString = fromString & " " & Year(date_begin) & " ����"
        End If
        '��������� ������ "���� ��" - ������ ���������
        toString = " �� " & CastToString(Day(date_end)) & " " & Months(Month(date_end)) & " " & Year(date_end) & " ����"
        '������� ���������
        GetDatesPeriod = fromString & toString
'-----------------------------------------------------------------------------------
' �������� ������ ������� (��������� �� ���� ���������)
'        MsgBox (GetDatesPeriod(CastToDate("30.01.2016"), CastToDate("31.01.2016")))
'        MsgBox (GetDatesPeriod(CastToDate("30.01.2016"), CastToDate("10.02.2016")))
'        MsgBox (GetDatesPeriod(CastToDate("30.12.2016"), CastToDate("11.01.2017")))
End Function

'-----------�������, ���������� ��� ���� � ���������� �������------------
Public Function GetDateString(dateValue As Date)
  '��������� ������ �������� ������� � ��������� ���������� ��� �������� ���
  Dim Months() As String
  Months = Split("0,������,�������,�����,������,���,����,����,�������,��������,�������,������,�������", ",")
  GetDateString = CastToString(DatePart("d", dateValue)) & " " & Months(Month(dateValue)) & " " & Year(dateValue)
End Function

'������� ������� ������� ������� ������� ������
Public Function LCaseString(AllString As String)
    LCaseString = LCase(Left(AllString, 1)) & Right(AllString, Len(AllString) - 1)
End Function

'������������ �������� - ���������� �������� ������ ������� - "���������"
Public Function RemovePodrDouble(shtat_name As String, podr_name As String, Optional BoolLCase As Boolean = True)
    Dim shtat() As String, podr() As String, i As Integer, shtat_output As String, shtat_output_length As Integer
    shtat_name = Trim(shtat_name)
    podr_name = Trim(podr_name)
    shtat = Split(shtat_name, " ", -1)
    shtat_output_length = UBound(shtat)
    shtat_output = shtat(0)
    
    '������������ ������, ���� �� ������ � ����������� ��� �������������, �� ������ �������� ������������
    '�� ������ ������� ������ �����.
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

'������������ �������� - ���������� �������� ������������� - "����� ������� ���������� ��������"
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
    '�������� ��� ���
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
   
    '��������� ������
    SplitPodrString = podr_output
End Function

'---------------------------��������� ������������� ��������� �� ID----------------------------
Public Function GetInfoIdToValue(ItemId As Long, ItemBsObject As String, ItemPartObject As String, ItemValue As String, sQueryDate As Date, bs As IBusinessServer, Optional SettingCondition As String = "", Optional VariableCondition As String = "")
    '������� ���������� ������ �� ����, �����, ������� ���������� � �������� ���������
    '�� ���� �� ���������� id
    Dim TempString As String
    Dim bo_podr As IBSDataObject, rs_podr As SKBS.SKRecordset
    Dim PodrParams As New Params

    '��������� ��������� ��� ��������� ������� ����������
    PodrParams.AddParam "id", ItemId
    PodrParams.AddParam "QueryDate", sQueryDate

    '�������� ������-������
    Set bo_podr = bs.GetBsObject(ItemBsObject, PodrParams)

    '�������� ������ �����
    Set rs_podr = bo_podr(ItemPartObject)

    '���������, ��� � ���������� ���� ������ (� ���������� ������
    '���������� ������ �� ���� QDate), �������� ������ ����
     If Not SettingCondition = "" And Not VariableCondition = "" Then
        '��������� ������� �����
        If rs_podr.RecordCount > 0 Then
            '��������� ������
            rs_podr.SetFilter SettingCondition & "=" & QuotedStr(VariableCondition)
            '������ ��������, ��� ����� ������� ������ �� ��������� �����, ������ ��������
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
'-----------------��������� ��� ���������� ��  ������� ������� ������---------------
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
        GetExecutorFIO = "� �������� ������� ������ �� ������ �����������"
    End If
End Function
'-----------------��������� �������� ���������� ��  ������� ������� ������---------------
Public Function GetExecutorTelephoneNumber(qDate As Date, bs As IBusinessServer)
        '������ ������� �����������
    Dim TelNumberStr As String, idIspoln As Long
    Dim TelNumberPersonal As String
    Dim idShtatIsp As Long
    Dim TelNumberShtat As String
    Dim idFirmIsp As Long
    Dim TelNumberFirm As String
    Dim user_id As Long
    user_id = bs.CurrentUserID
    TelNumberStr = ""
    '������ id �����������

    If Not CastToString(GetInfoIdToValue(user_id, "SYS_Account", "main", "id", qDate, bs, "", "")) = "" Then
        idIspoln = CastToLong(GetInfoIdToValue(user_id, "SYS_Account", "main", "id_personal", qDate, bs, "", ""))
        '���� ����� id �����������
        If Not idIspoln = -1 Then
            '������ ������� ����. �� ������ �������� �� ������� - ������ ��������� - ��������
            TelNumberPersonal = GetInfoIdToValue(idIspoln, "REC_PERSONAL", "contacts", "charval", qDate, bs, "code", "01")
            '���� ������� �� ������� ���� � ��
            If TelNumberPersonal = "" Then
                '������ id ������� ������� �����������
                idShtatIsp = CastToLong(GetInfoIdToValue(idIspoln, "REC_PERSONAL", "EXECPOST", "id_shtat", qDate, bs, "Work_workerstatus_code", "1"))
                '������ ������� �� �������� ������� �������
                TelNumberShtat = GetInfoIdToValue(idShtatIsp, "REC_SHTAT", "telephone", "charval", qDate, bs, "", "")
                '���� ������� �� ������� � �������� ��
                If TelNumberShtat = "" Then
                    '������ id �����������
                    idFirmIsp = CastToLong(GetInfoIdToValue(idIspoln, "REC_PERSONAL", "EXECPOST", "Work_Firm_id", qDate, bs, "Work_workerstatus_code", "1"))
                    '������ ������� �� �������� �����������
                    TelNumberFirm = GetInfoIdToValue(idFirmIsp, "REC_FIRM", "contacts", "charval", qDate, bs, "", "")
                    '���� ������� �� ������� � �������� �����������
                    If TelNumberFirm = "" Then
                        TelNumberStr = "� �������� �����������, ������� �������, ������ �������� ����. ����������� ����� ��������"
                    Else
                        '������� ����� �������� �� �������� �����������
                        TelNumberPersonal = TelNumberFirm
                    End If
                Else
                    '������� ����� �������� �� �������� ��
                    TelNumberPersonal = TelNumberShtat
                End If
            End If
        Else
            TelNumberStr = "� �������� ������� ������ �� ������ �����������"
        End If
    Else
        TelNumberStr = "����������� ������ �� �����������"
    End If

    If Not TelNumberStr = "" Then
        GetExecutorTelephoneNumber = CastToString(TelNumberStr)
    Else
        GetExecutorTelephoneNumber = CastToString(TelNumberPersonal)
    End If
End Function

Public Function GetPersonalSex(sotr_id As Long, qDate As Date, bs As BusinessServer)
If GetInfoIdToValue(sotr_id, "REC_PERSONAL", "sex", "Text", qDate, bs) = "�" Then
    GetPersonalSex = True
  Else
    GetPersonalSex = False
  End If
End Function

Public Function GetPersonalSexMG(sotr_id As Long, qDate As Date, bs As BusinessServer)
If GetInfoIdToValue(sotr_id, "REC_PERSONAL", "sex", "Text", qDate, bs) = "�" Then
    GetPersonalSexMG = "�"
  Else
    GetPersonalSexMG = "�"
  End If
End Function

Public Function WritePersonalSex(sex As Boolean, Optional MaleString As String = "���", Optional FemaleString As String = "�")
    If sex Then
        WritePersonalSex = MaleString
    Else
        WritePersonalSex = FemaleString
    End If
End Function

'-------------------������������� ��� � ���������� ������� (� �� ��� � ��������� ��������)-------------------
Public Function MakeFIOShortCorrectly(surname As String, Name As String, patronymic As String, Optional Padeg As Long = 1, Optional FIOFormat As FIOFormatEnum = ffSurnameNamePatronomic, Optional sotrSexIfNoPatronymic As String = "")
Dim FIO() As String
Dim Result As String
'���� ���� ��������
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
'���� ��� ��������
Result = GetFIO_Padeg(surname, Name, "", ffSurnameNamePatronomic, Padeg)
    If sotrSexIfNoPatronymic = "�" Then
        Result = GetFIO_Padeg(surname, Name, "��������", ffSurnameNamePatronomic, Padeg)
    End If
    If sotrSexIfNoPatronymic = "�" Then
        Result = GetFIO_Padeg(surname, Name, "��������", ffSurnameNamePatronomic, Padeg)
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

'-------------------������� ������ ���, �������� ����� � ������ ������ � ����������-------------------
Public Function MakeFIOShortOneString(InitialFIO As String, Optional Padeg As Long = 1, Optional FIOFormat As FIOFormatEnum = ffSurnameNamePatronomic, Optional SplitStr As String = " ")
Dim FIO() As String

FIO = Split(InitialFIO, SplitStr, -1)

If UBound(FIO) = 0 Then
    MakeFIOShortOneString = "������� ������������ ���"
Else
    MakeFIOShortOneString = MakeFIOShortCorrectly(FIO(0), FIO(1), Nvl(FIO(2), ""), Padeg, FIOFormat)
End If

End Function

'-------------------������� ������� �� ������-------------------
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

'-------------------��������� ������ � �b�������� ������������� �� ������� ������� ����������--------------
Public Function Replace_DirectorPodr(NumbOption As Integer, NameOption As String) As String
    '������ �������� - shtat_podr_info
    '_0='Object_Code_charval=''01'';
    'ShortName_charval='�������� �������������';
    'fullname_charval='������ �������� �������������';
    'Director_FIO_Personal='������� ��� ��������';
    'Director_intval=id_����������;
    'Podr_Director_ExecPost_Shtat_Shortname='��������� ����������';';
    
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

'-------------------�������������� ������ � ��������� ������������� �� ������� ������� ����������--------------
Public Function getDirectorPodrAndFIO(shtat_podr_info As String)
    Dim FIOArray() As String
FIOArray = Split(CastToString(Replace_DirectorPodr(3, shtat_podr_info)))
ReDim Preserve FIOArray(4)
FIOArray(4) = Replace_DirectorPodr(5, shtat_podr_info)
getDirectorPodrAndFIO = FIOArray
End Function

'-------------------��������� ��� ��������� ������������ �������������-------------------
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
'-------------------��������� �������� ������������ ����������-------------------
Public Function getDirectiorPodrPodr(shtat_podr_info As String, Optional soglDolgn As String = "")
    Dim directorPodrData() As String
    directorPodrData = getDirectorPodrAndFIO(shtat_podr_info)
    If Not CastToString(directorPodrData) = "" Then
    getDirectiorPodrPodr = directorPodrData(4)
    Else
        getDirectiorPodrPodr = soglDolgn
    End If
End Function

' ������� ���������� �������� �������� �� ��������� ����������� ��� ���
Public Function GetPersonalGSStatusByShtatId(shtatId As Long, qDate As Date, bs As IBusinessServer)
    Dim postCodeAnswer As String
    postCodeAnswer = GetInfoIdToValue(shtatId, "REC_SHTAT", "post", "code", qDate, bs)
    If GetPostCategory_GS(postCodeAnswer) <> "" Then
        GetPersonalGSStatusByShtatId = True '�������� �����������
    Else
        GetPersonalGSStatusByShtatId = False '�� �������� �����������
    End If
End Function

' ������� ���������� ��������, �������� �� ��������� �����������, ��������, ����������
Public Function GetPersonalStatusByShtatId(shtatId As Long, qDate As Date, bs As IBusinessServer)
    Dim postCodeAnswer As String, profCodeAnswer As String
    Dim MaskCodeGS As String, MaskCodeSotr As String, MaskCodeJob As String
    MaskCodeGS = "##-#-#-###" '����� ��� ��
    MaskCodeSotr = "2####" '����� ��� ��������
    postCodeAnswer = GetInfoIdToValue(shtatId, "REC_SHTAT", "post", "code", qDate, bs)
    If Not postCodeAnswer = "" Then
        If postCodeAnswer Like MaskCodeGS Then
            GetPersonalStatusByShtatId = 1 '�������� �����������
        ElseIf postCodeAnswer Like MaskCodeSotr Then
            GetPersonalStatusByShtatId = 2 '�������� ��������
        Else
            GetPersonalStatusByShtatId = 0 '����� ������
        End If
    Else
        MaskCodeJob = "1####" '����� ��� ����������
        profCodeAnswer = GetInfoIdToValue(shtatId, "REC_SHTAT", "prof", "code", qDate, bs)
        If profCodeAnswer <> "" And profCodeAnswer Like MaskCodeJob Then
            GetPersonalStatusByShtatId = 3 '�������� ����������
        Else
            GetPersonalStatusByShtatId = 0 '����� ������ 100%
        End If
    End If
End Function
