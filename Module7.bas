Attribute VB_Name = "Module1"
Option Explicit

Public flag As Boolean
Public flag2 As Boolean
Public maxrow2 As Long
Public wsName2 As String

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long
Dim o As Long
Dim p As Long
Dim q As Long
Dim r As Long

Dim Count As Long
Dim maxrow As Long
Dim maxcol As Long
Dim maxcol2 As Long
Dim maxrow3 As Long
Dim maxrow4 As Long
Dim r_UB1 As Integer
Dim r_UB2 As Integer
Dim UB1 As Integer
Dim UB2 As Integer
Dim wsName As String
Dim Filepath As String
Dim itemNo() As String

Dim rsm() As String
Dim tCD() As String
Dim rCD() As String
Dim rty() As String
Dim stg() As String

Dim wb As Workbook
Dim ws As Worksheet

Sub open_file()
  
  Filepath = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xlsx")
  If Filepath <> "False" Then
    Set wb = Workbooks.Open(Filepath)
  End If
End Sub

Sub close_workbook()
  
  If Filepath <> "False" Then
    wb.Close
  End If
End Sub

Sub seigou(wsName2 As String)
  
  Call open_file
  
  If Filepath = "False" Then Exit Sub
  With wb.Worksheets(wsName2)

    maxrow = .Cells(Rows.Count, 1).End(xlUp).Row
    maxcol = .Cells(2, Columns.Count).End(xlToLeft).Column
    
    ReDim rsm(maxrow - 3, maxcol - 2)
    ReDim tCD(maxrow - 3)
    ReDim stg(maxcol - 2)
    
    k = 0
    For i = 2 To maxrow
      If i = 2 Then
        flag = True
      Else
        flag = False
      End If
      l = 0
      For j = 1 To maxcol
        If j = 1 Then
          If Not flag Then
            flag2 = True
            tCD(k) = _
            .Cells(i, j).Value
          Else
            flag2 = False
          End If
        Else
          If flag Then
            stg(l) = _
            .Cells(i, j).Value
          Else
            rsm(k, l) = _
            .Cells(i, j).Value
          End If
          l = l + 1
        End If
      Next
      If flag2 Then
        k = k + 1
      End If
    Next
  End With
  
  Call close_workbook
  Call open_file
  
  If Filepath = "False" Then Exit Sub
  With wb.Worksheets(wsName2)
    maxrow2 = .Cells(Rows.Count, 2).End(xlUp).Row
    maxcol2 = .Cells(2, Columns.Count).End(xlToLeft).Column
     
    UB1 = UBound(tCD)
    ReDim rty(UB1)
    ReDim rCD(UB1, maxcol2 - 4)
    UB2 = UBound(rCD, 2)
    
    For i = 3 To maxrow2
      For j = 0 To UB1
        If tCD(j) = _
           .Cells(i, 3) Then
          rty(j) = _
          .Cells(i, 2).Value
          l = 4
          For k = 0 To UB2
            rCD(j, k) = _
            .Cells(i, l).Value
            l = l + 1
          Next
        End If
      Next
    Next
  End With
  
  Call close_workbook
  With ThisWorkbook.Worksheets("Sheet1")
    
    r_UB1 = UBound(rsm, 1)
    r_UB2 = UBound(rsm, 2)
    
    maxrow3 = .Cells(Rows.Count, 2).End(xlUp).Row
    maxrow4 = .Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To 26 Step 12
      Select Case i
        Case 14
          j = 4
        Case Else
          j = 0
      End Select
      If .Cells(3, i + j) <> "" Then
        Range(.Cells(3, i + j), .Cells(maxrow3, i + j + 5)).ClearContents
      End If
    Next
    
    n = 3
    For i = 0 To r_UB1
      For j = 0 To r_UB2
        If rsm(i, j) = "Yes" Then
          For k = 0 To UB2
            If rCD(i, k) <> "" Then
              m = 0
              For l = 3 To maxrow4
                If _
                  .Cells(l, 10) = tCD(i) And _
                  .Cells(l, 12) = rCD(i, k) And _
                  .Cells(l, 13) = rty(i) And _
                  .Cells(l, 15) = stg(j) Then
                  .Cells(l, 2) = tCD(i)
                  .Cells(l, 4) = rCD(i, k)
                  .Cells(l, 5) = rty(i)
                  .Cells(l, 7) = stg(j)
                  Exit For
                Else
                  m = m + 1
                End If
              Next
              If m = maxrow4 - 2 Then
                .Cells(n, 26) = tCD(i)
                .Cells(n, 28) = rCD(i, k)
                .Cells(n, 29) = rty(i)
                .Cells(n, 31) = stg(j)
                n = n + 1
              End If
            End If
          Next
        End If
      Next
    Next
    j = 3
    For i = 3 To maxrow4
      .Cells(i, 18) = "=B" & j & "=J" & j
      .Cells(i, 20) = "=D" & j & "=L" & j
      .Cells(i, 21) = "=E" & j & "=M" & j
      .Cells(i, 23) = "=G" & j & "=O" & j
      j = j + 1
    Next
  End With
  
End Sub

Sub popUpMenu()
  UserForm1.Show
End Sub

Sub fill_clear()
  With ThisWorkbook.Worksheets("Sheet1")
    maxrow = .Cells(Rows.Count, 9).End(xlUp).Row
    If .Cells(3, 9) <> "" Then
      Range(.Cells(3, 9), .Cells(maxrow, 15)).ClearContents
    End If
  End With
End Sub

Sub GetDataFromSQLServer(SrvrName As String, userName As String)
  Dim objCn As New ADODB.Connection
  Dim objRs As New ADODB.Recordset
  Dim strSql As String
  With objCn
    .CommandTimeout = 0
    .ConnectionString = "Provider=MSOLEDBSQL" _
                      & ";Data Source=" & SrvrName _
                      & ";Initial Catalog=DynamicsLake" _
                      & ";Authentication=ActiveDirectoryInteractive" _
                      & ";User ID=" & userName _
                      & ";Use Encryption for Data=true;"
    .ConnectionTimeout = 0
    .Open
  End With
  
  strSql = "SELECT a.productbaseid,asurepairdiagnosiscode,asurepairdiagnosistypeid,repaircode,RepairType,SortOrder,RepairStepId" _
          & " FROM asuRepairProductDiagnosisMap a" _
          & " LEFT JOIN asuRepairProductDiagnosisRelation c ON a.productbaseid = c.productbaseid" _
          & " LEFT JOIN asuRepairDiagnosisCodeMapping d ON c.recid = d.proddiagrelrefrecid" _
          & " LEFT JOIN asuRepairDiagnosisCodeTable b ON c.diagnosiscoderefrecid = b.recid" _
          & " LEFT JOIN asuRepairDiagnosisStepRelation e ON c.RecId = e.proddiagrelrefrecid" _
          & " WHERE c.productbaseid IN('" & Join(itemNo, "','") & "')" _
          & " AND asurepairdiagnosiscode LIKE 'T%'" _
          & " ORDER BY c.productbaseid, b.asurepairdiagnosiscode,RepairType,RepairStepId,SortOrder,RepairCode;"
  
  Call objRs.Open(strSql, objCn, , adLockReadOnly)
  Worksheets("Sheet1").Range("A2").CopyFromRecordset objRs
  
End Sub

Sub verify365(SrvrName As String, _
              userName As String)

  Dim UB As Integer
  Dim val As String
  Dim val2 As String
  Dim sttrow As Long
  Dim lstrow As Long
  Dim lstcol As Long
  Dim ModelName() As String
  Dim FinalNo() As String
  Dim brNo() As String
  
  With ThisWorkbook.Worksheets("Sheet2")
    maxrow = .Cells(Rows.Count, 2).End(xlUp).Row
    Count = maxrow - 3
    
    j = 0
    val2 = ""
    ReDim itemNo(Count)
    For i = 3 To maxrow
      If i > 3 Then
        If .Cells(i - 1, 1) _
        <> .Cells(i, 1) Then
          val = val _
              & .Cells(i, 1).Value _
              & ","
          val2 = val2 _
              & .Cells(i, 2).Value _
              & ","
        End If
      Else
        val = _
        .Cells(i, 1).Value & ","
      End If
      itemNo(j) _
      = .Cells(i, 2).Value
      j = j + 1
    Next
    ModelName = Split(Left(val, Len(val) - 1), ",")
    FinalNo = Split(Left(val2, Len(val2) - 1), ",")
    
  End With
  
  val = ""
  UB = UBound(itemNo)
  For i = 0 To UB
    val = val _
        & itemNo(i) _
        & ","
  Next
  Set wb = Workbooks.Add
  Set ws = ActiveSheet
  
  Call GetDataFromSQLServer(SrvrName, userName)
  With ws
    .Cells(1, 1) = "productbaseid"
    .Cells(1, 2) = "asurepairdiagnosiscode"
    .Cells(1, 3) = "asurepairdiagnosistypeid"
    .Cells(1, 4) = "repaircode"
    .Cells(1, 5) = "RepairType"
    .Cells(1, 6) = "SortOrder"
    .Cells(1, 7) = "RepairStepId"
    
    With .Sort
      With .SortFields
        .Clear
        .Add Key:=Range("A1") _
        , CustomOrder:=Left(val, Len(val) - 1)
      End With
      .SetRange Range("A:G")
      .Header = xlYes
      .Apply
    End With
    
    k = 0
    sttrow = 2
    maxrow = .Cells(Rows.Count, 1).End(xlUp).Row
    UB = UBound(FinalNo)
    For i = 0 To UB + 1
      If i > 0 Then
        sttrow = lstrow
      End If
      If i < UB + 1 Then
        For j = sttrow To maxrow
          If FinalNo(i) _
          = .Cells(j, 1) Then
            lstrow = j
            Exit For
          End If
        Next
      Else
        lstrow = maxrow + 1
      End If
      With wb
        .Worksheets.Add
        .Sheets(1).Name = ModelName(k)
      End With
      k = k + 1
      l = 2
      m = 0
      For n = sttrow To lstrow
        brNo = Split(l & ",2", ",")
        flag = True
        If n > sttrow Then
          If .Cells(n - 1, 1) _
          <> .Cells(n, 1) Then
            flag = False
          End If
        End If
        
        For o = 0 To 1
        
          p = CLng(brNo(o))
          If p = 2 Then
            q = 1
          Else
            If n = 2 Then
              q = n
            Else
              q = n - 1
            End If
          End If
          
          For r = 1 To 7
          
            Worksheets(1) _
            .Cells(p, m + r) = .Cells(q, r).Value
            
          Next
          If flag _
          Or n = lstrow Then
            Exit For
          Else
            If o = 0 Then
              If m = 0 Then
                m = m + 8
              Else
                m = m + 15
              End If
            Else
              l = 2
            End If
          End If
        Next
        l = l + 1
      Next
      With Worksheets(1)
        maxcol = .Cells(2, Columns.Count).End(xlToLeft).Column
        maxrow2 = .Cells(Rows.Count, 1).End(xlUp).Row
        lstcol = maxcol + 1
        For l = 16 To lstcol Step 15
          For m = 2 To maxrow2
            For n = 1 To 6
              If m = 2 Then
                val = _
                .Cells(m, l + n - 7).Value
              Else
                val = _
                "=RC[-" & l - 1 & "]=RC[-7]"
              End If
              .Cells(m, l + n) = val
            Next
          Next
        Next
      End With
    Next
  End With
  
  UB = UBound(ModelName)
  Count = 0
  For i = 0 To UB
    With Worksheets(i + 1)
      maxcol = .Cells(2, Columns.Count).End(xlToLeft).Column
      maxrow = .Cells(Rows.Count, 1).End(xlUp).Row
      lstcol = maxcol - 5
      For j = 17 To lstcol Step 15
        For k = 3 To maxrow
          For l = 0 To 5
            If .Cells(k, 2 + l) _
            <> .Cells(k, j + l - 7) Then
              Count = Count + 1
              .Cells(k, j + l).Interior.Color = vbRed
            End If
          Next
        Next
      Next
    End With
  Next
  If Count > 0 Then
    MsgBox Count & "個照合一致しません。"
  Else
    MsgBox "照合一致しました。"
  End If
  
End Sub

Sub TableTest(wsName2 As String)

  Dim wb2 As Workbook
  
  Call fill_clear
  Call open_file

  If Filepath = "False" Then Exit Sub
  
  With wb
    Count = .Sheets.Count
    For i = 1 To Count - 1
      With .Worksheets(i)
        wsName = .Name
        maxrow = .Cells(Rows.Count, 1).End(xlUp).Row
        For j = 3 To maxrow
          For k = 1 To 7
            ThisWorkbook.Worksheets("Sheet1").Cells(j, k + 8) _
            = .Cells(j, k).Value
          Next
        Next
      End With
      
      If i = 1 Then
        Call open_file
  
        If Filepath = "False" Then Exit Sub
        With wb.Worksheets(wsName2)
      
          maxrow = .Cells(Rows.Count, 1).End(xlUp).Row
          maxcol = .Cells(2, Columns.Count).End(xlToLeft).Column
          
          ReDim rsm(maxrow - 3, maxcol - 2)
          ReDim tCD(maxrow - 3)
          ReDim stg(maxcol - 2)
          
          j = 0
          For k = 2 To maxrow
            If k = 2 Then
              flag = True
            Else
              flag = False
            End If
            l = 0
            For m = 1 To maxcol
              If m = 1 Then
                If Not flag Then
                  flag2 = True
                  tCD(j) = _
                  .Cells(k, m).Value
                Else
                  flag2 = False
                End If
              Else
                If flag Then
                  stg(l) = _
                  .Cells(k, m).Value
                Else
                  rsm(j, l) = _
                  .Cells(k, m).Value
                End If
                l = l + 1
              End If
            Next
            If flag2 Then
              j = j + 1
            End If
          Next
        End With
        
        Call close_workbook
        Call open_file
        
        If Filepath = "False" Then Exit Sub
        With wb.Worksheets(wsName2)
          maxrow2 = .Cells(Rows.Count, 2).End(xlUp).Row
          maxcol2 = .Cells(2, Columns.Count).End(xlToLeft).Column
           
          UB1 = UBound(tCD)
          ReDim rty(UB1)
          ReDim rCD(UB1, maxcol2 - 4)
          UB2 = UBound(rCD, 2)
          
          For j = 3 To maxrow2
            For k = 0 To UB1
              If tCD(k) = _
                 .Cells(j, 3) Then
                rty(k) = _
                .Cells(j, 2).Value
                l = 4
                For m = 0 To UB2
                  rCD(k, m) = _
                  .Cells(j, l).Value
                  l = l + 1
                Next
              End If
            Next
          Next
        End With
        
        Call close_workbook
        Set wb2 = Workbooks.Add
      End If
      
      With ThisWorkbook.Worksheets("Sheet1")
    
        r_UB1 = UBound(rsm, 1)
        r_UB2 = UBound(rsm, 2)
        
        maxrow3 = .Cells(Rows.Count, 2).End(xlUp).Row
        maxrow4 = .Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To 26 Step 12
          Select Case j
            Case 14
              k = 4
            Case Else
              k = 0
          End Select
          If .Cells(3, j + k) <> "" Then
            Range(.Cells(3, j + k), .Cells(maxrow3, j + k + 5)).ClearContents
          End If
        Next
        
        j = 3
        For k = 0 To r_UB1
          For l = 0 To r_UB2
            If rsm(k, l) = "Yes" Then
              For m = 0 To UB2
                If rCD(k, m) <> "" Then
                  n = 0
                  For o = 3 To maxrow4
                    If _
                      .Cells(o, 10) = tCD(k) And _
                      .Cells(o, 12) = rCD(k, m) And _
                      .Cells(o, 13) = rty(k) And _
                      .Cells(o, 15) = stg(l) Then
                      .Cells(o, 2) = tCD(k)
                      .Cells(o, 4) = rCD(k, m)
                      .Cells(o, 5) = rty(k)
                      .Cells(o, 7) = stg(l)
                      Exit For
                    Else
                      n = n + 1
                    End If
                  Next
                  If n = maxrow4 - 2 Then
                    .Cells(j, 26) = tCD(k)
                    .Cells(j, 28) = rCD(k, m)
                    .Cells(j, 29) = rty(k)
                    .Cells(j, 31) = stg(l)
                    j = j + 1
                  End If
                End If
              Next
            End If
          Next
        Next
        j = 3
        For k = 3 To maxrow4
          .Cells(k, 18) = "=B" & j & "=J" & j
          .Cells(k, 20) = "=D" & j & "=L" & j
          .Cells(k, 21) = "=E" & j & "=M" & j
          .Cells(k, 23) = "=G" & j & "=O" & j
          j = j + 1
        Next
      End With
      
      With wb2
        ThisWorkbook.Worksheets("Sheet1").Copy _
        After:=.Worksheets(i)
        ActiveSheet.Name = wsName
        If i = Count - 1 Then
          Application.DisplayAlerts = False
          .Worksheets(1).Delete
        End If
      End With
      
      Call fill_clear
    Next
  End With
End Sub

Sub popUpMenu2()
  UserForm2.Show
End Sub

Sub IsFileOpen()
  flag2 = True
  Call popUpMenu
End Sub

Public Function ShowDataLinkPropertyDialog() As String
    Dim msd As MSDASC.DataLinks
    Dim con As Connection
 
    Set msd = New MSDASC.DataLinks
    Set con = msd.PromptNew
    If con Is Nothing Then
        ShowDataLinkPropertyDialog = ""
    Else
        ShowDataLinkPropertyDialog = con.ConnectionString
    End If
    Set con = Nothing
End Function

Sub DataLinkProperty()
  Call ShowDataLinkPropertyDialog
End Sub


