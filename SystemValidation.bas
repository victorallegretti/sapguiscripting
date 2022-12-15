'***************************************************************************************
'***************************************************************************************
'The Below will displayed an error message and stop Macro execution if the system chosen is, for example, Production or any
'other that the scripting should not run.
Sub SystemValidation()
'We always start the VBA with a Sub statement that must be written without spaces or numbers following with ().
Dim Fim As Double, P As Double, C As Double, W As Integer, K As Integer 'Declaring the main Letters used for Columns,
'reading the amount of lines to be processed, counting how many session from SAP we have.
Dim cApp As Object 'cAPP is basically the .exe from SAP GUI Launchpad
Dim cConn As Object 'Connections open (QAS3, QAS2...)
Dim Session As Object 'Each window opened in SAP creates a session, we want to count and read which transaction/window we want the macro to run in SAP
Dim Ask(32, 3) As String
Dim SUBRC As Integer
Dim Width As Long
Dim PassportPreSystemId As String
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

'*********The System Validation Starts here******************
If Session.PassportPreSystemId = "PRD_PassaportID-Example" Then
  Msgbox ("You've chosen to run the Macro in '& PassportPreSystemId &', this is an UAT version, please select a validad session instead."), vbOKOnly
  Ask(0, 1) = ""
  GoTo Connection
End If
          
End Sub
