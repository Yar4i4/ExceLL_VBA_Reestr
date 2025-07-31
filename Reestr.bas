Attribute VB_Name = "Module1"
Option Explicit

Sub Reestr()
    Dim Vb As Workbook: Set Vb = ThisWorkbook
    Dim Vsh As Worksheet: Set Vsh = Vb.Sheets(1)
    Dim fd As FileDialog, files As Variant, fp As Variant
    Dim Ob As Workbook, Ob1 As Worksheet, Ob2 As Worksheet
    Dim foundScenario1 As Boolean, foundScenario2 As Boolean
    Dim c As String
    Dim rowInsert As Long

    ' 1) ������� ������ ������ (A2:AF �� ��������� ������)
    With Vsh
        Dim lr As Long
        lr = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lr >= 2 Then .Range("A2:AF" & lr).ClearContents
    End With

    ' 2) ������ ������ ������
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .InitialFileName = Vb.Path & Application.PathSeparator
        .title = "�������� ����� ��� ������ � ������"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel files", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If Not .Show Then Exit Sub ' ����� "������"

        Dim i As Long
        ReDim files(1 To .SelectedItems.Count)
        For i = 1 To .SelectedItems.Count
            files(i) = .SelectedItems(i)
        Next i
    End With

    Application.ScreenUpdating = False

    ' 3) ������� ��������� ������
    For Each fp In files
        ' ���������� ����-������, ���� �� ��� ������
        If LCase(fp) = LCase(Vb.FullName) Then
            GoTo NextFile
        End If

        On Error Resume Next ' ��������� ���������, ���� ���� �� �����������
        Set Ob = Workbooks.Open(fp, ReadOnly:=True)
        If Err.Number <> 0 Then
            MsgBox "�� ������� ������� ����: " & fp, vbCritical, "������ ��������"
            Err.Clear
            GoTo NextFile
        End If
        On Error GoTo 0 ' ���������� ����������� ��������� ������

        If Ob.Sheets.Count < 2 Then
            MsgBox "���� '" & Ob.Name & "' �������� ����� 2 ������ � ����� ��������.", vbExclamation
            GoTo OhSkip
        End If

        Set Ob1 = Ob.Sheets(1)
        Set Ob2 = Ob.Sheets(2)

        ' --- �������� ������� ---
        c = Clean(Ob1.Range("A16").Text)
        If InStr(c, "��������workcode") = 0 Then
            If MsgBox("� ����� '" & Ob.Name & "' �� ����� '" & Ob1.Name & "' � ������ A16 ����������� ������ ���� ����� / Work Code�." & vbCrLf & vbCrLf & "���������� ���� ����?", vbYesNo + vbExclamation, "������ �������") = vbNo Then GoTo Finished
            GoTo OhSkip
        End If

        foundScenario1 = False: foundScenario2 = False
        If Ob1.Range("A20").MergeCells And Ob1.Range("A20").MergeArea.Address = "$A$20:$F$20" Then
            c = Clean(Ob1.Range("A20").Text)
            If InStr(c, "���������������������������") > 0 Then foundScenario1 = True
        End If
        If Not foundScenario1 And Ob1.Range("A21").MergeCells And Ob1.Range("A21").MergeArea.Address = "$A$21:$F$21" Then
            c = Clean(Ob1.Range("A21").Text)
            If InStr(c, "���������������������������") > 0 Then foundScenario2 = True
        End If
        If Not foundScenario1 And Not foundScenario2 Then
            If MsgBox("� ����� '" & Ob.Name & "' ����������� ������ �������������� ����������, ���� � ������� A20:F20 ��� A21:F21." & vbCrLf & vbCrLf & "���������� ���� ����?", vbYesNo + vbExclamation, "������ �������") = vbNo Then GoTo Finished
            GoTo OhSkip
        End If

        Dim foundExp As Boolean: foundExp = False
        If InStr(Clean(Ob2.Range("A7").Text), "���������") > 0 Then foundExp = True
        If Not foundExp And InStr(Clean(Ob2.Range("A8").Text), "���������") > 0 Then foundExp = True
        If Not foundExp Then
            If MsgBox("� ����� '" & Ob.Name & "' ����������� ������ ������ ���� � ������� A7 ��� A8." & vbCrLf & vbCrLf & "���������� ���� ����?", vbYesNo + vbExclamation, "������ �������") = vbNo Then GoTo Finished
            GoTo OhSkip
        End If

        ' --- ������ ���������� ---
        rowInsert = Vsh.Cells(Vsh.Rows.Count, "A").End(xlUp).Row + 1
        If rowInsert < 2 Then rowInsert = 2

        If foundScenario1 Then
            Call FillLine(Vsh, Ob1, Ob2, rowInsert, isSecondLine:=False)
        ElseIf foundScenario2 Then
            Call FillLine(Vsh, Ob1, Ob2, rowInsert, isSecondLine:=False)
            Call FillLine(Vsh, Ob1, Ob2, rowInsert + 1, isSecondLine:=True)
        End If

OhSkip:
        Ob.Close False
NextFile:
    Next fp

Finished:
    Application.ScreenUpdating = True
'    MsgBox "��������� ���������.", vbInformation
End Sub

Private Sub FillLine(wsDst As Worksheet, ws1 As Worksheet, ws2 As Worksheet, rw As Long, isSecondLine As Boolean)
    Dim srcC11 As String, p() As String, midNum As String, last3 As String
    Dim arrB As Variant, vals() As String
    Dim i As Long
    Dim lst As String
    
    Dim dataRow1 As Long: dataRow1 = 18
    Dim dataRow2 As Long: dataRow2 = 14

    srcC11 = ws1.Range("C11").Text
    srcC11 = Replace(Replace(srcC11, "_", "-"), "�", "-")
    p = Split(srcC11, "-")
    If UBound(p) >= 3 Then midNum = p(3) Else midNum = ""
    If UBound(p) >= 4 Then last3 = Left(p(4), 3) Else last3 = ""

    With wsDst
        .Cells(rw, "A").Formula = "=ROW()-8"
        .Cells(rw, "B").FormulaR1C1 = "=CONCATENATE(RC[15],RC[16],RC[12])"
        .Cells(rw, "C").FormulaR1C1 = "=CONCATENATE(""COR-P3"",""-"",RC[7],""-0"",RC[4],""-"",RC[5])"
        .Cells(rw, "D").FormulaR1C1 = "=RC[13]"
        .Cells(rw, "E").FormulaR1C1 = "=RC[13]"
        .Cells(rw, "F").Value = ""
        
        ' --- ������ � ������� G � ��������������� ---
        .Cells(rw, "G").NumberFormat = "@" ' ������������� ��������� ������ ��� ������
        If IsNumeric(midNum) And midNum <> "" Then
            .Cells(rw, "G").Value = Format(midNum, "0000") ' ����������� ����� �� 4 ������ � ������
        Else
            .Cells(rw, "G").Value = midNum ' ���� �� ����� (��� �����), ��������� ��� ����
        End If
        ' --------------------------------------------
        
        .Cells(rw, "H").Value = last3
        .Cells(rw, "I").Value = ws1.Range("E11").Text
        .Cells(rw, "J").Value = "RSR"
        .Cells(rw, "L").Value = ws1.Range("G11").Text
        
        Dim vv As String: vv = UCase(Trim(.Cells(rw, "L").Value))
        If vv = "A1" Or vv = "�1" Then .Cells(rw, "K").Value = "TYPE 1" Else .Cells(rw, "K").Value = "TYPE 2"

        Dim lookup As String: lookup = UCase(Trim(.Cells(rw, "H").Value))
        Select Case lookup
            Case "CIV": .Cells(rw, "M").Value = "CIVIL"
            Case "U/G": .Cells(rw, "M").Value = "UNDERGROUND PIPING"
            Case "PIP": .Cells(rw, "M").Value = "PIPING"
            Case "STR": .Cells(rw, "M").Value = "STRUCTURES"
            Case "PKG": .Cells(rw, "M").Value = "PACKAGES"
            Case "EQP": .Cells(rw, "M").Value = "EQUIPMENT (STATIC AND ROTARY)"
            Case "ELE": .Cells(rw, "M").Value = "ELECTRICAL"
            Case "I&C": .Cells(rw, "M").Value = "INSTRUMENTATION AND CONTROL"
            Case "PAI": .Cells(rw, "M").Value = "PAINTING"
            Case "INS": .Cells(rw, "M").Value = "INSULATION"
            Case "HSE": .Cells(rw, "M").Value = "SAFETY"
            Case "WHS": .Cells(rw, "M").Value = "WAREHOUSE"
            Case "ADM": .Cells(rw, "M").Value = "ADMINISTRATION/LOGISTICS"
            Case "COM": .Cells(rw, "M").Value = "COMMISSIONING"
            Case "HVA": .Cells(rw, "M").Value = "HVAC"
            Case "PIL": .Cells(rw, "M").Value = "PILING WORK"
            Case "TCF": .Cells(rw, "M").Value = "TEMPORARY FACILITIES"
        End Select

        .Cells(rw, "N").Value = ws1.Cells(dataRow1, "A").Value
        .Cells(rw, "O").Value = ws1.Cells(dataRow1, "C").Value
        .Cells(rw, "P").Value = ws1.Range("N13").Value
        .Cells(rw, "Q").Value = ws1.Range("Q13").Value
        .Cells(rw, "R").Value = ws1.Cells(dataRow1, "B").Value
        .Cells(rw, "S").Value = ws1.Cells(dataRow2, "C").MergeArea.Cells(1, 1).Text

        Dim lastRowB As Long
        lastRowB = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
        If lastRowB >= 8 Then
            arrB = ws2.Range("B8:B" & lastRowB).Value
            If IsArray(arrB) Then
                ReDim vals(1 To UBound(arrB, 1))
                For i = 1 To UBound(arrB, 1)
                    vals(i) = CStr(arrB(i, 1))
                Next
                lst = UniqueSortJoin(vals, ", ")
                .Cells(rw, "T").Value = lst
            End If
        End If

        .Cells(rw, "U").Value = ws1.Cells(dataRow1, "D").Value
        .Cells(rw, "V").Value = ws1.Cells(dataRow1, "E").Value
        .Cells(rw, "W").Value = ws1.Cells(dataRow1, "F").Value
        .Cells(rw, "X").Value = ws1.Cells(dataRow1, "O").Value

        .Cells(rw, "Y").FormulaR1C1 = "=IFERROR(RC[-2]-RC[-3],0)"
        .Cells(rw, "Z").FormulaR1C1 = "=IFERROR(RC[-1]*RC[-2],0)"
        .Cells(rw, "AA").FormulaR1C1 = "=IFERROR(ROUND(RC[-4]*RC[-3],2),0)"
        .Cells(rw, "AF").FormulaR1C1 = "=CONCATENATE(RC[-16],RC[-15],RC[-14],RC[-18])"
        
        Call FormatLine(wsDst, rw)
        
    End With
End Sub


Private Function Clean(s As String) As String
    Dim chars As Variant
    chars = Array(" ", ".", ",", "/", "\", vbTab, vbCr, vbLf)
    s = LCase(s)
    Dim i As Long
    For i = LBound(chars) To UBound(chars)
        s = Replace(s, chars(i), "")
    Next i
    Clean = s
End Function

Private Function UniqueSortJoin(arr As Variant, sep As String) As String
    Dim dict As Object, i As Long, key
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    For i = LBound(arr) To UBound(arr)
        key = Trim(arr(i))
        If Len(key) > 0 Then dict(key) = 1
    Next
    
    If dict.Count = 0 Then Exit Function
    
    Dim keys() As String
    ReDim keys(0 To dict.Count - 1)
    i = 0
    For Each key In dict.keys
        keys(i) = CStr(key)
        i = i + 1
    Next
    
    If dict.Count > 1 Then Call QuickSort(keys, LBound(keys), UBound(keys))
    UniqueSortJoin = Join(keys, sep)
End Function

Private Sub QuickSort(arr As Variant, ByVal L As Long, ByVal R As Long)
    Dim i As Long, j As Long, mid As Variant, tmp As Variant
    i = L: j = R: mid = arr((L + R) \ 2)
    Do While i <= j
        Do While arr(i) < mid: i = i + 1: Loop
        Do While arr(j) > mid: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If L < j Then QuickSort arr, L, j
    If i < R Then QuickSort arr, i, R
End Sub


Private Sub FormatLine(ws As Worksheet, rw As Long)
    Dim lineRange As Range
    Set lineRange = ws.Range("A" & rw & ":AF" & rw)

    ' 1. ��������� ����� ����� (�����, ������� � �.�.)
    With lineRange
        .NumberFormat = "General" ' ������� ������������� ����� ��� ����
        .VerticalAlignment = xlCenter
        .IndentLevel = 0
        .WrapText = False
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
            .Color = 0 ' ������
        End With
        With .Interior
            .Color = 16777215 ' �����
            .Pattern = xlNone
        End With
    End With
    
    ' 2. ������������� ���������� ������� ��� ���� �����
    Dim border As Variant
    For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        With lineRange.Borders(border)
            .LineStyle = xlDot
            .Weight = xlHairline
            .Color = RGB(32, 55, 100) ' -�����-����� (#203764)
        End With
    Next border

    ' 3. ������������� �������������� ������������
    lineRange.HorizontalAlignment = xlCenter
    ws.Range("B" & rw & ",C" & rw & ",F" & rw & ",M" & rw & ",N" & rw & ",O" & rw & ",S" & rw & ",T" & rw).HorizontalAlignment = xlLeft
    ws.Range("AB" & rw & ":AF" & rw).HorizontalAlignment = xlLeft
    
    ' 4. ������������� �������� ������ � ������������ �� ������� ����
    With ws.Range("V" & rw & ":AA" & rw)
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0.00"
    End With
    
    ' 5. ������������� ��������� ������ ��� ������� G
    ' ��� ������ �����������, ��� � ������� G ����� ��������� ������, ���� ����� ����� "��������"
    ws.Range("G" & rw).NumberFormat = "@"

End Sub
