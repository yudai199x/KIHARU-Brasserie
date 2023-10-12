Attribute VB_Name = "Module1"
Option Explicit
Public flag As Boolean

Dim i As Long
Dim j As Long
Dim maxrow As Long

Sub 受付(y As Integer, _
         m As Integer, _
         d As Integer, _
         c As Integer, _
         md As String)
                  
  Dim dt As String
  Dim no As String
  Dim str As String
  Dim sir As String
  Dim dgt As String
  Dim cnt As Long
  
  With Worksheets("ファイル作成")
    maxrow = .Cells(Rows.Count, 3).End(xlUp).Row
    cnt = Len(CStr(maxrow))
    str = "0"
    For i = 1 To cnt
      str = _
      str & "0"
    Next
    For i = 3 To maxrow
      dt = y _
         & Format(m, "00") _
         & Format(d, "00") _
         & Format(c, str)
      no = Format(c, str)
      .Cells(i, 1) = "RMA" + dt
      .Cells(i, 2) = md + dt
      If flag Then
        sir = .Cells(1, 6).Value
        dgt = .Cells(1, 7).Value
        While Len(sir) < CInt(dgt) - Len(no)
          sir = _
          sir & "x"
        Wend
        sir = _
        sir & no
      Else
        sir = dt
      End If
      .Cells(i, 4) = sir
      c = c + 1
    Next
  End With
End Sub

Sub rcpt(filePath As String)
  
  Dim model As String
  Dim unkfile As String
  Dim appNo() As String
  Dim refNo() As String
  Dim skuNo() As String
  Dim sirNo() As String
  
  With Worksheets("ファイル作成")
    maxrow = .Cells(Rows.Count, 1).End(xlUp).Row
    model = .Cells(1, 5).Value
    
    ReDim appNo(maxrow - 3)
    ReDim refNo(maxrow - 3)
    ReDim skuNo(maxrow - 3)
    ReDim sirNo(maxrow - 3)
    
    j = 0
    For i = 3 To maxrow
      appNo(j) = .Cells(i, 1).Value
      refNo(j) = .Cells(i, 2).Value
      skuNo(j) = .Cells(i, 3).Value
      sirNo(j) = .Cells(i, 4).Value
      j = j + 1
    Next
    
    For i = 3 To maxrow
      For j = 1 To 4
        .Cells(i, j) = ""
      Next
    Next
  End With
  
  Worksheets("EDI受付フォーマット").Activate
  For i = LBound(appNo) To UBound(appNo)
    Cells(4, 1) = "BGN*13*" & appNo(i) & "*20210923*1030***FT*7~"
    Cells(5, 1) = "N9*DO*" & refNo(i) & "~"
    Cells(10, 1) = "BLI*ZZ*1*1*EA****BP*" & skuNo(i) & "~"
    Cells(12, 1) = "N9*SE*" & sirNo(i) & "~"
    Cells(13, 1) = "PID*F****" & model & "~"
    
    With CreateObject("ADODB.Stream")
      .Charset = "UTF-8"
      .Open
      For j = 1 To 20
          .WriteText Cells(j, 1), 1
      Next
      .SaveToFile filePath & "\" & refNo(i) & ".txt", 2
      .Close
    End With
  Next
  
  Worksheets("ファイル作成").Activate
  unkfile = filePath & "\" & ".txt"
  If Dir(unkfile) <> "" Then
    Kill unkfile
  End If
  
End Sub

Sub start()
  UserForm1.Show
End Sub

Sub start2()
  UserForm2.Show
End Sub

