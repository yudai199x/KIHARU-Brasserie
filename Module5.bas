Attribute VB_Name = "Module1"
Option Explicit
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long
Dim Filepath As String
Dim str As String
Dim flag As Boolean
Dim flag2 As Boolean
Dim maxrow As Long

Sub Pull_Data_from_Excel_with_ADODB()

  Dim rqDay As Date
  Dim cnStr As String
  Dim query As String
  Dim rs As New ADODB.Recordset
  
  Application.ScreenUpdating = False
  rqDay = FileToDate(Filepath)
  cnStr = _
  "Provider=Microsoft.ACE.OLEDB.12.0;" & _
  "Data Source=" & Filepath & ";" & _
  "Extended Properties=Excel 12.0"
  query = _
  "SELECT * FROM [SKU MASTER$]" & _
  " WHERE Request_Sent_Date = #" & rqDay & "#" & _
  " AND GDC = 'Asurion JP';"
  
  Call rs.Open(query, cnStr, adOpenStatic, adLockReadOnly)
  With ThisWorkbook.Worksheets("Sheet1")
    .Range(.Range("A2"), .Range("A" & Cells.Rows.Count)).EntireRow.Delete
    .Range("A2").CopyFromRecordset rs
  End With
  Application.ScreenUpdating = True
  
End Sub

Function FileToDate(File As String) As Date

  Dim s As String
  Dim y As String
  Dim m As String
  Dim d As String
  
  s = Replace(Mid(File, InStrRev(File, "_") + 1), ".xlsx", "")
  y = Left(s, 4)
  m = Mid(s, 5, 2)
  d = Right(s, 2)

  FileToDate = CDate(y & "/" & m & "/" & d)

End Function

Sub P_R_RA()

  Dim val As String
  Dim val2 As String
  Dim CustomList() As String

  Call Pull_Data_from_Excel_with_ADODB
  With ThisWorkbook
    With .Worksheets("Sheet1")
      .Activate
      .Range("A:BT").CurrentRegion.RemoveDuplicates Columns:=9, Header:=xlYes
      maxrow = .Cells(Rows.Count, 9).End(xlUp).Row * 3 - 3
      For i = 2 To maxrow Step 3
        If Right(.Cells(i, 9), 2) = "-R" Then
          .Rows(i).Insert
          .Rows(i + 2).Insert
          k = 0
          l = 2
          m = 2
          n = 1
          val = "-R"
          flag = True
          flag2 = True
        ElseIf Right(.Cells(i, 9), 3) = "-RA" Then
          .Rows(i & ":" & i + 1).Insert
          k = 0
          l = 1
          m = 1
          n = 2
          val = "-RA"
          flag = True
          flag2 = False
        Else
          .Rows(i + 1 & ":" & i + 2).Insert
          k = 1
          l = 2
          m = 1
          n = 0
          flag = False
          flag2 = True
        End If
        For j = k To l Step m
          If m = 2 Then
            If Not flag2 Then
              flag = False
              flag2 = True
            End If
          End If
          If flag Then
            If j = k Then val2 = ""
            If j = l Then val2 = "-R"
            If m = 2 Then
              flag2 = False
            End If
            .Cells(i + j, 9) = Replace(.Cells(i + n, 9), val, val2)
          End If
          If flag2 Then
            If j = k Then val2 = "-R"
            If j = l Then
              If m = 2 Then
                val2 = "A"
              Else
                val2 = "-RA"
              End If
            End If
            .Cells(i + j, 9) = .Cells(i + n, 9).Value & val2
          End If
        Next
      Next
      i = 2
      Do While .Cells(i, 9).Value <> ""
        If WorksheetFunction.CountIf(.Range("I:I"), .Cells(i, 9)) > 1 _
        And .Cells(i, 1) = "" Then
          .Rows(i).EntireRow.Delete
          i = i - 1
        Else
          i = i + 1
        End If
      Loop
      maxrow = .Cells(Rows.Count, 9).End(xlUp).Row
      str = ""
      For i = 2 To maxrow
        If Right(.Cells(i, 9), 3) = "-RA" Then
          str = str _
              & Replace(.Cells(i, 9), "-RA", "") & "," _
              & Replace(.Cells(i, 9), "-RA", "-R") & "," _
              & .Cells(i, 9).Value & ","
        End If
      Next
      If str <> "" Then
        CustomList = Split(Left(str, Len(str) - 1), ",")
        Application.AddCustomList Listarray:=CustomList
        n = Application.CustomListCount
        .Range("A:BT").Sort Key1:=Range("I1"), OrderCustom:=n + 1, Header:=xlYes
        Application.DeleteCustomList n
        .Sort.SortFields.Clear
      End If
    End With
    With .Worksheets("Sheet2")
      .Range(.Range("A3"), .Range("A" & Cells.Rows.Count)).EntireRow.Delete
    End With
  End With
  
End Sub

Sub ATK_BOM()

  Dim sttpos As Long
  Dim endpos As Long
  Dim colNo() As String
  Dim initd() As String
  Dim lastd() As String
  Dim parts() As String
  Dim parts2() As String
  
  Filepath = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xlsx")
  If Filepath = "False" Then Exit Sub
  
  colNo = Split("5,6,10,11", ",")
  initd = Split("RMA/Replacement,Refurbished,Refurb,RMA/REFURB", ",")
  lastd = Split("Refurb,REFURBISHED,Ref,Refu,Refur,REFURISHED", ",")
  
  Call P_R_RA
  With ThisWorkbook
    maxrow = .Worksheets("Sheet1").Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To maxrow
      For j = 9 To 11
        .Worksheets("Sheet2").Cells(i + 1, j - 5) = _
        .Worksheets("Sheet1").Cells(i, j).Value
      Next
      .Worksheets("Sheet2").Cells(i + 1, 11) = _
      .Worksheets("Sheet1").Cells(i, 59).Value
    Next
    With .Worksheets("Sheet2")
      .Activate
      .Cells(1, 6) = Replace(Mid(Filepath, InStrRev(Filepath, "_") + 1), ".xlsx", "") & "クール"
      For i = 3 To maxrow + 1
        If Right(.Cells(i, 4), 2) = "-R" Then
          j = 1
          str = "_SALV"
        ElseIf Right(.Cells(i, 4), 3) = "-RA" Then
          j = 2
          str = "_Refurb"
        Else
          j = 0
        End If
        .Cells(i, 1) = i - 2
        .Cells(i, 3) = .Cells(i - j, 4).Value & "-BASE"
        If j > 0 Then
          For k = 0 To 3
            If k = 2 Then
              str = ""
            End If
            l = CLng(colNo(k))
            .Cells(i, l) = _
            .Cells(i - j, l).Value & str
          Next
        Else
          If .Cells(i, 5) = "" Then
            If .Cells(i + 1, 5) <> "" Then
              k = 1
              For l = 5 To 6
                .Cells(i, l) = Left(.Cells(i + 1, l), InStrRev(.Cells(i + 1, l), ",") - 1)
              Next
            Else
              k = 2
              For l = 5 To 6
                For m = LBound(initd) To UBound(initd)
                  If Left(UCase(.Cells(i + 2, l)), Len(initd(m))) = UCase(initd(m)) Then
                    .Cells(i, l) = Mid(.Cells(i + 2, l), InStr(.Cells(i + 2, l), ",") + 1)
                  End If
                Next
                For m = LBound(lastd) To UBound(lastd)
                  If Right(UCase(.Cells(i + 2, l)), Len(lastd(m))) = UCase(lastd(m)) Then
                    .Cells(i, l) = Left(.Cells(i + 2, l), InStrRev(.Cells(i + 2, l), ",") - 1)
                  End If
                Next
              Next
            End If
          End If
          If .Cells(i, 11) = "" Then .Cells(i, 11) = .Cells(i + k, 11).Value
          sttpos = InStr(.Cells(i, 5), .Cells(i, 11)) + Len(.Cells(i, 11)) + 2
          endpos = InStr(Mid(.Cells(i, 5), sttpos), ",") - 1
          If endpos > 0 Then
            flag = True
            .Cells(i, 10) = Left(Mid(.Cells(i, 5), sttpos), endpos)
          End If
          If Right(.Cells(i, 10), 1) <> "G" Then
            flag2 = True
            parts = Split(.Cells(i, 6), ",")
            For j = LBound(parts) To UBound(parts)
              If InStr(parts(j), "G") > 0 Then
                flag2 = False
                If Right(parts(j), 1) = "G" _
                Or Right(parts(j), 2) = "GB" Then
                  flag = True
                  .Cells(i, 10) = parts(j)
                Else
                  parts2 = Split(parts(j), " ")
                  For k = LBound(parts2) To UBound(parts2)
                    If Right(parts2(k), 2) = "GB" Then
                      .Cells(i, 10) = parts2(k)
                    End If
                  Next
                End If
              Else
                If flag2 Then
                  flag = False
                  .Cells(i, 10) = ""
                End If
              End If
            Next
          End If
          If flag Then
            If Right(.Cells(i, 10), 1) <> "B" Then
              .Cells(i, 10) = .Cells(i, 10).Value & "B"
            End If
          End If
          For j = 0 To 3
            k = CLng(colNo(j))
            .Cells(i, k) = Replace(.Cells(i, k), " ", "")
            If j < 2 Then
              .Cells(i, k) = Replace(.Cells(i, k), ",", "_")
            End If
          Next
        End If
      Next
    End With
  End With
  
End Sub








