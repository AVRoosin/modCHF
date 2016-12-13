Attribute VB_Name = "modCHF"
Option Explicit
Option Compare Text

'last update 08.12.2016

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
        If month(date_begin) <> month(date_end) Then
            '���� ����� �� ��� ��, ��������� ��� �������� � ������ "���� �"
            fromString = fromString & " " & Months(month(date_begin))
        End If
    '���������, ��� ������������/������ ������������� � ��� �� ������
        If Year(date_begin) <> Year(date_end) Then
            '���� ����� ��� ��, ��������� ��� �������� � ������ "���� �"
            If month(date_begin) = month(date_end) Then
                fromString = fromString & " " & Months(month(date_begin))
            End If
            '���� ��� ��� ��, ��������� ��� � ������ "���� �"
            fromString = fromString & " " & Year(date_begin) & " ����"
        End If
        '��������� ������ "���� ��" - ������ ���������
        toString = " �� " & CastToString(Day(date_end)) & " " & Months(month(date_end)) & " " & Year(date_end) & " ����"
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
  GetDateString = CastToString(DatePart("d", dateValue)) & " " & Months(month(dateValue)) & " " & Year(dateValue)
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
    Dim surname As String, name As String, patronymic As String
    Dim user_id As Long
    user_id = bs.CurrentUserID
    name = CastToString(GetInfoIdToValue(user_id, "SYS_Account", "main", "name", qDate, bs))
    surname = CastToString(GetInfoIdToValue(user_id, "SYS_Account", "main", "surname", qDate, bs))
    patronymic = CastToString(GetInfoIdToValue(user_id, "SYS_Account", "main", "patronymic", qDate, bs))
    If Not name = "" And Not surname = "" And Not patronymic = "" Then
        GetExecutorFIO = MakeFIOShortCorrectly(surname, name, patronymic, 1, ffSurnameNP)
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
Public Function MakeFIOShortCorrectly(surname As String, name As String, patronymic As String, Optional Padeg As Long = 1, Optional FIOFormat As FIOFormatEnum = ffSurnameNamePatronomic, Optional sotrSexIfNoPatronymic As String = "")
Dim FIO() As String
Dim Result As String
'���� ���� ��������
If patronymic <> "" Then
    Result = GetFIO_Padeg(surname, name, patronymic, ffSurnameNamePatronomic, Padeg)
    Select Case FIOFormat
    Case ffNPSurname
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = Left(name, 1) & "." & Left(patronymic, 1) & ". " & FIO(0)
    Case ffSurnameNP
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = FIO(0) & " " & Left(name, 1) & "." & Left(patronymic, 1) & "."
    Case Else
        MakeFIOShortCorrectly = Result
    End Select
Else
'���� ��� ��������
Result = GetFIO_Padeg(surname, name, "", ffSurnameNamePatronomic, Padeg)
    If sotrSexIfNoPatronymic = "�" Then
        Result = GetFIO_Padeg(surname, name, "��������", ffSurnameNamePatronomic, Padeg)
    End If
    If sotrSexIfNoPatronymic = "�" Then
        Result = GetFIO_Padeg(surname, name, "��������", ffSurnameNamePatronomic, Padeg)
    End If
    Select Case FIOFormat
    Case ffNPSurname
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = Left(name, 1) & "." & FIO(0)
    Case ffSurnameNP
        FIO = Split(Result, " ", -1)
        MakeFIOShortCorrectly = FIO(0) & " " & Left(name, 1) & "."
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
    MakeFIOShortOneString = MakeFIOShortCorrectly(FIO(0), FIO(1), nvl(FIO(2), ""), Padeg, FIOFormat)
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
    
    If Not nvl(NameOption, "") = "" Then
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
