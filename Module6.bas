Attribute VB_Name = "Module1"
Option Explicit

Dim Filepath As String
Dim wb As Workbook

Sub open_file()
  
  Filepath = Application.GetOpenFilename(FileFilter:="Excelƒtƒ@ƒCƒ‹,*.xlsx")
  If Filepath <> "False" Then
    Set wb = Workbooks.Open(Filepath)
  End If
End Sub

Sub close_workbook()
  
  If Filepath <> "False" Then
    wb.Close
  End If
End Sub

Sub ADOExcelSQLServer(SrvrName As String, userName As String)
  Dim objCn As New ADODB.Connection
  Dim objRs As New ADODB.Recordset
  Dim mapp() As String
  Dim str(4) As String
  Dim strSql As String
  Dim MAX_UB As Integer
  Dim maxrow(4) As Long
  Dim maxrow2 As Long
  Dim maxrow3 As Long
  
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  
  Dim productbaseid() As String
  Dim repaircode() As String
  Dim RepairType() As String
  Dim asurepairdiagnosiscode() As String
  Dim RepairStepId() As String
  
  With ThisWorkbook
    With .Worksheets("Sheet1")
      For i = 0 To 4
        maxrow(i) = .Cells(Rows.Count, i + 1).End(xlUp).Row
        MAX_UB = maxrow(i) - 2
        
        If maxrow(i) > 1 Then
          str(i) = ""
          j = 0
          If i = 0 Then
            ReDim productbaseid(MAX_UB)
            For k = 2 To maxrow(i)
              productbaseid(j) = .Cells(k, 1).Value
              j = j + 1
            Next
          ElseIf i = 1 Then
            ReDim repaircode(MAX_UB)
            If str(0) <> "--" Then str(i) = "AND "
            For k = 2 To maxrow(i)
              repaircode(j) = .Cells(k, 2).Value
              j = j + 1
            Next
          ElseIf i = 2 Then
            ReDim RepairType(MAX_UB)
            If str(0) <> "--" _
            Or str(1) <> "--" Then str(i) = "AND "
            For k = 2 To maxrow(i)
              RepairType(j) = .Cells(k, 3).Value
              j = j + 1
            Next
          ElseIf i = 3 Then
            ReDim asurepairdiagnosiscode(MAX_UB)
            If str(0) <> "--" _
            Or str(1) <> "--" _
            Or str(2) <> "--" Then str(i) = "AND "
            For k = 2 To maxrow(i)
              asurepairdiagnosiscode(j) = .Cells(k, 4).Value
              j = j + 1
            Next
          ElseIf i = 4 Then
            ReDim RepairStepId(MAX_UB)
            If str(0) <> "--" _
            Or str(1) <> "--" _
            Or str(2) <> "--" _
            Or str(3) <> "--" Then str(i) = "AND "
            For k = 2 To maxrow(i)
              RepairStepId(j) = .Cells(k, 5).Value
              j = j + 1
            Next
          End If
        Else
          str(i) = "--"
        End If
      Next
    End With
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
    strSql = "SELECT a.productbaseid, asurepairdiagnosiscode,asurepairdiagnosistypeid,repaircode,RepairType,SortOrder,RepairStepId" _
            & " FROM asuRepairProductDiagnosisMap a" _
            & " LEFT JOIN asuRepairProductDiagnosisRelation c ON a.productbaseid = c.productbaseid" _
            & " LEFT JOIN asuRepairDiagnosisCodeMapping d ON c.recid = d.proddiagrelrefrecid" _
            & " LEFT JOIN asuRepairDiagnosisCodeTable b ON c.diagnosiscoderefrecid = b.recid" _
            & " LEFT JOIN asuRepairDiagnosisStepRelation e ON c.RecId = e.proddiagrelrefrecid" _
            & " WHERE" _
            & vbCrLf & " " & str(0) & "c.productbaseid IN('" & Join(productbaseid, "','") & "')" _
            & vbCrLf & " " & str(1) & "repaircode IN('" & Join(repaircode, "','") & "')" _
            & vbCrLf & " " & str(2) & "repairtype IN('" & Join(RepairType, "','") & "')" _
            & vbCrLf & " " & str(3) & "asurepairdiagnosiscode IN('" & Join(asurepairdiagnosiscode, "','") & "')" _
            & vbCrLf & " " & str(4) & "RepairStepId IN('" & Join(RepairStepId, "','") & "')" _
            & vbCrLf & " ORDER BY c.productbaseid, b.asurepairdiagnosiscode,RepairType,RepairStepId,SortOrder,RepairCode;"
  
    Call objRs.Open(strSql, objCn, , adLockReadOnly)
    With .Worksheets("Sheet2")
      If .Cells(2, 1) <> "" Then
        maxrow2 = .Cells(Rows.Count, 1).End(xlUp).Row
        Range(.Cells(2, 1), .Cells(maxrow2, 7)).ClearContents
      End If
      .Range("A2").CopyFromRecordset objRs
      maxrow2 = .Cells(Rows.Count, 1).End(xlUp).Row
      k = 0
      MAX_UB = maxrow2 - 2
      ReDim mapp(MAX_UB, 6)
      For i = 2 To maxrow2
        l = 0
        For j = 1 To 7
          mapp(k, l) = .Cells(i, j).Value
          l = l + 1
        Next
        k = k + 1
      Next
    End With
  End With
  
  Call open_file
  If Filepath = "False" Then Exit Sub
  With wb.Worksheets(1)
    For i = 8 To 16 Step 8
      For j = 1 To 7
        .Cells(1, i + j) = .Cells(1, j).Value
      Next
    Next
    maxrow3 = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To maxrow3
      If .Cells(i, 1).Font.Color = vbRed Then Exit For
    Next
    For j = 0 To MAX_UB
      For k = i To maxrow3
        If .Cells(k, 1) = mapp(j, 0) And _
           .Cells(k, 2) = mapp(j, 1) And _
           .Cells(k, 3) = mapp(j, 2) And _
           .Cells(k, 4) = mapp(j, 3) And _
           .Cells(k, 5) = mapp(j, 4) And _
           .Cells(k, 6) = mapp(j, 5) And _
           .Cells(k, 7) = mapp(j, 6) Then
          For l = 0 To 6
            .Cells(k, l + 9) = mapp(j, l)
            .Cells(k, l + 17) = "=RC[-16]=RC[-8]"
          Next
          Exit For
        End If
      Next
    Next
  End With
  
End Sub

Sub popUpMenu()
  UserForm1.Show
End Sub



