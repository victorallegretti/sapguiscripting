Option Explicit 'This first line will force VBA to read the declared variables

Sub MacroSample()
'We always start the VBA with a Sub statement that must be written without spaces or numbers following with ().
Dim StartTime As Double 'StartTime is a variable we use for counting the minutes the macro took to run once the code is executed.
'As Double will return the smallest and largest intervals from a number which we want to use to convert into minutes later on.
Dim MinutesElapsed As String 'This is string return from the time the macro took to run. Strings are used for write down texts, popup dialog titles and so on.
Dim Fim As Double, P As Double, C As Double, W As Integer, K As Integer 'Declaring the main Letters used for Columns,
'reading the amount of lines to be processed, counting how many session from SAP we have.
Dim cApp As Object 'cAPP is basically the .exe from SAP GUI Launchpad
Dim cConn As Object 'Connections open (QAS3, QAS2...)
Dim Session As Object 'Each window opened in SAP creates a session, we want to count and read which transaction/window we want the macro to run in SAP
Dim Ask(32, 3) As String
Dim SUBRC As Integer
Dim Width As Long
Dim PassportPreSystemId As String

StartTime = Timer 'Counter starts now.

Const i As Double = 5 'A constant i will be the line we want the macro to start reading our values, such as customer number.
''Remember that for programming, the count begins with 0, so we want the 6th line to be read, we then declare 5.


'************Reads if SAP is opened**************
Set cApp = GetObject("SAPGUI").GetScriptingEngine
'************************************************
'Counts how many connections are opened, you must be logged in
Connection:
P = 1
For W = 1 To cApp.Children.Count
  Set cConn = cApp.Children(W - 1)
  For K = 1 To cConn.Children.Count
    Set Session = cConn.Children(K - 1).Info
    Ask(P, 1) = "[" & P & "] " & Session.SystemName & Session.Client & " | " & Session.user & " | " & Session.Transaction
    Ask(P, 2) = W
    Ask(P, 3) = K
    Ask(0, 1) = Ask(0, 1) & Ask(P, 1) & vbCrLf
    P = P + 1
      Next
Next

'********Now we create a message box (pop-up) returning the values we found on the selection shown above.
Msgbox:
Ask(0, 1) = Ask(0, 1) & "[0] - Cancel"
P = InputBox(Ask(0, 1), "Select a Session", W - 1)
If P = 0 Then Exit Sub 'If the user types 0 then the macro will stop.
W = Ask(P, 2)
K = Ask(P, 3)
Set Session = cApp.Children(W - 1).Children(K - 1)

'*********Count of how many lines are going to be processed, up to 65k.

For Fim = i To 65000
 If Cells(Fim, 2) = "" Then Exit For 'Cells statement will always refer to Line, Column So we count Column B
Next
Fim = Fim - 1

'Message box for confirming the amount found in previous selection.

If Msgbox("Do you want to procedure with " & (Fim - i + 1) & " registers?", vbYesNo) = vbNo Then
    Exit Sub
End If

'=====================
For P = i To Fim ' Setting P equals i to Fim means the line now can be declared as Cells(P(Line), 1(COLUMN))

GoSub Mainloop 'This will be the main loop we want the macro to run, it will go line by line until it reaches the last line found under Fim statement above
    Next
Msgbox "Done.", vbOKOnly 'This message box will be called once there no more line to be processed
    GoTo Fim

Mainloop:
 DoEvents
If Cells(P, 1) = 0 Then 'If status column is blank, then we select the customer column
    Cells(P, 2).Select

'Now we can paste our source code from the script1.vbs recorded in SAP from the session.findbyid statement.
Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NXD02"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").Text = "98765"
Session.findById("wnd[1]").sendVKey 0
Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").Text = Cells(P, 3)
Session.findById("wnd[0]/tbar[0]/btn[11]").press

    Cells(P, 1) = "OK" 'Writes OK ok Status column in Excel
    Cells(P, 4) = Session.findById("wnd[0]/sbar").Text 'Write the message bar text from SAP to Excel

End If 'Ends the IF statement from the Cells P,2 selection

Return 'Return to main loop

GoTo Mainloop 'This will force the mainloop to be read in case the selection fails preventing dump on the macro.

Fim:
'Determine how many seconds code took to run
 MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")


'Notify user in seconds
Msgbox "This Macro ran in " & MinutesElapsed & " minutes", vbInformation

ActiveWorkbook.Save
End Sub
