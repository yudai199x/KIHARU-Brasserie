Attribute VB_Name = "Module1"
Option Explicit
Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
Dim i As Long
Dim cb As Variant
Dim flag As Boolean
Dim ofst As Long
Dim ofst2 As Long
Sub ClearClipboard()
  OpenClipboard (0&)
  EmptyClipboard
  CloseClipboard
End Sub
Sub ExeCapture()
  Do While True
    cb = Application.ClipboardFormats
    If cb(1) <> -1 Then
      For i = 1 To UBound(cb)
        If cb(i) = xlClipboardFormatBitmap Then
          If ofst > 45 Then
            ofst2 = 45
          Else
            ofst2 = ofst
          End If
          ActiveWindow.SmallScroll Down:=ofst2
          ActiveSheet.Paste Destination:=Cells(1, 1).Offset(ofst, 0)
          ofst = ofst + 45
          Call ClearClipboard
        End If
      Next
    End If
    DoEvents
    If flag Then
      Exit Do
      DoEvents
    End If
  Loop
End Sub
Sub StartCapture(SheetName As String)
  Sheets.Add(After:=Sheets(Sheets.Count)).Name = SheetName
  ofst = 0
  flag = False
  ActiveWindow.Zoom = 50
  Call ClearClipboard
  Call ExeCapture
End Sub
Sub StopCapture()
  flag = True
End Sub
Sub start()
  UserForm1.Show
End Sub

