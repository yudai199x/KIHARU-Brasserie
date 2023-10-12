Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub ’IŽD()
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  Dim wb As Workbook
  Dim wbPath As String
  Dim wbName As String
  Dim Count As Long
  Dim Model As String
  Dim gPN() As String
  Dim compalPN() As String
  Dim DAXname() As String
  Dim maxrow As Long
  Dim Maxsht As Long
  
  wbPath = "C:\Users\yudai.fujii\Documents\ATK_BOM\"
  wbName = Dir(wbPath & "Pixel Fold  ATK BOM v1.0.xlsx")
  
  Do While wbName <> ""
    Workbooks.Open wbPath & wbName
    With Worksheets(1)
      maxrow = .Cells(Rows.Count, 7).End(xlUp).Row
      Model = .Cells(1, 4).Value
      Count = maxrow - 3
      ReDim gPN(Count)
      ReDim compalPN(Count)
      ReDim DAXname(Count)
      j = 1
      For i = 4 To maxrow
        gPN(j) = .Cells(i, 7).Value
        compalPN(j) = .Cells(i, 8).Value
        DAXname(j) = .Cells(i, 10).Value
        j = j + 1
      Next
      If Count Mod 2 = 0 Then
        k = 0
      Else
        k = 1
      End If
      Maxsht = (k + Count) / 2
    End With
    
    Set wb = Workbooks.Add
    l = 1
    For i = 1 To Maxsht
      ThisWorkbook.Worksheets(1).Copy After:=wb.Worksheets(i)
      ActiveSheet.Name = i
      For j = 0 To 40 Step 40
        If l > Count Then Exit For
        Cells(2, j + 7) = Model
        Cells(4, j + 2) = gPN(l)
        Cells(4, j + 21) = compalPN(l)
        Cells(6, j + 2) = DAXname(l)
        l = l + 1
      Next
    Next
    Application.DisplayAlerts = False
    Worksheets(1).Delete
    wb.SaveAs "C:\Users\yudai.fujii\Documents\’IŽD\" & Model & ".xlsx"
    wb.Close
    Workbooks(wbName).Close
    wbName = Dir()
  Loop
  
End Sub

