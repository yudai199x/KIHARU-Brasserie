Attribute VB_Name = "Module1"
Option Explicit
Sub SKUàÍóóçÏê¨()
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  Dim m As Long
  Dim n As Long
  Dim maxrow As Long
  Dim rcnt As Long
  Dim rrow As Long
  Dim flag As Long
  
  maxrow = Worksheets("ëŒè€SKU").Cells(Rows.Count, 1).End(xlUp).Row
  
  Dim iNo() As String
  Dim desc() As String
  Dim rcode() As String
  Dim rtype(2) As String
  ReDim iNo(maxrow - 2) As String
  ReDim desc(maxrow - 2) As String
  ReDim rcode(1, 2, maxrow - 2) As String
  
  j = 0
  For i = 2 To maxrow
    iNo(j) = Worksheets("ëŒè€SKU").Cells(i, 1).Value
    desc(j) = Worksheets("ëŒè€SKU").Cells(i, 2).Value
    j = j + 1
  Next
  
  n = 0
  For i = 0 To maxrow - 2
    If Left(desc(i), 6) = "Pixel3" Then j = 2
    If Left(desc(i), 6) = "Pixel4" Then j = 8
    If Left(desc(i), 6) = "Pixel5" Then j = 14
    If Left(desc(i), 6) = "Pixel6" Then j = 20
    If Left(desc(i), 6) = "Pixel7" Then j = 26
    If Left(desc(i), 7) = "Pixel3a" Then j = j + 4
    If Left(desc(i), 7) = "Pixel4a" Then j = j + 4
    If Left(desc(i), 7) = "Pixel5a" Then j = j + 4
    If Left(desc(i), 7) = "Pixel6a" Then j = j + 4
    If Left(desc(i), 8) = "Pixel3XL" Then j = j + 2
    If Left(desc(i), 8) = "Pixel4XL" Then j = j + 2
    If Left(desc(i), 9) = "Pixel4aXL" Then j = j + 4
    If Left(desc(i), 9) = "Pixel6Pro" Then j = j + 2
    If Left(desc(i), 9) = "Pixel7Pro" Then j = j + 2
    
    For k = 0 To 1
      l = 10
      rcnt = 0
      
      Do While Worksheets("T-Codes").Cells(j + k, l).Value <> ""
        rcnt = rcnt + 1
        l = l + 1
      Loop
      
      For m = 0 To rcnt - 1
        rcode(k, m, n) = Worksheets("T-Codes").Cells(j + k, m + 10).Value
      Next
    Next
    n = n + 1
  Next
  
  rtype(0) = "1to1"
  rtype(1) = "KH"
  rtype(2) = "DOA"
  
  k = 0
  Do While k < maxrow - 1
    For i = 0 To 1
      flag = 0
      rcnt = 0
      For j = 0 To 2
        If rcode(i, j, k) <> "" Then rcnt = rcnt + 1
      Next
      l = Worksheets("Mapping").Cells(Rows.Count, 1).End(xlUp).Row + 1
      For m = 0 To 8
        rrow = rcnt * 3 - 1
        If Left(iNo(k), 1) = "G" Then
          rrow = rcnt - 1
          flag = 1
        End If
        If m > rrow Then Exit For
        Worksheets("Mapping").Cells(l + m, 1) = iNo(k)
        Worksheets("Mapping").Cells(l + m, 3) = "IW"
        Worksheets("Mapping").Cells(l + m, 7) = "3001"
        If i = 0 Then Worksheets("Mapping").Cells(l + m, 2) = "T005"
        If i = 1 Then Worksheets("Mapping").Cells(l + m, 2) = "T085"
        For j = 0 To 2
          If m Mod rcnt = j Then
            Worksheets("Mapping").Cells(l + m, 4) = rcode(i, j, k)
            Worksheets("Mapping").Cells(l + m, 6) = j + 1
          End If
        Next
        If flag = 1 Then
          j = 1
        Else
          If m < rcnt Then
            j = 0
          Else
            If m < rcnt * 2 Then
              j = 1
            Else
              j = 2
            End If
          End If
        End If
        Worksheets("Mapping").Cells(l + m, 5) = rtype(j)
      Next
    Next
    k = k + 1
  Loop
End Sub
