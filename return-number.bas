'The code below reffers to retrieving only numbers from a given pop-up text in SAP.
'=====================================================================================================
Dim control As Boolean
Dim controlval As String
Dim resultval As String
Dim z As Integer

resultval = ""
controlval = Session.findById("wnd[3]/usr/lbl[10,2]").DisplayedText

For z = 1 To Len(controlval)
control = IsNumeric(Mid(controlval, z, 1))
If control = True Then resultval = resultval & Mid(controlval, z, 1)
If Len(resultval) = 16 Then Exit For
Next z

Cells(P, 3) = resultval
'=====================================================================================================
