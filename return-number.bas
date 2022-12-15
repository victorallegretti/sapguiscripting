'The code below reffers to retrieving only numbers from a given pop-up text in SAP.
'=====================================================================================================
Dim control As Boolean
Dim controlval As String
Dim resultval As String
Dim z As Integer

resultval = ""
controlval = Session.findById("wnd[3]/usr/lbl[10,2]").DisplayedText 'Here you place the text you want to retrieve only numbers from.

For z = 1 To Len(controlval)
control = IsNumeric(Mid(controlval, z, 1))
If control = True Then resultval = resultval & Mid(controlval, z, 1)
  If Len(resultval) = 16 Then Exit For 'Whenever the 16 chars are reached, stop the search. You can easily discover how many chars a given text has by placing it in Notepad++ The number of chars will appear at the botton.
Next z

  Range("A5") = resultval ' Place the result in A5 Excel Sheet
'=====================================================================================================
