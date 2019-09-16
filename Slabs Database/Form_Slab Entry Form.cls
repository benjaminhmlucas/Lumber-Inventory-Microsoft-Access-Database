VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Slab Entry Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
''-----------------Buttons for Navigation---------------------------'''
'closes open windows, appends purchased records to sold slabs and deletes purchased records from slab entry

'------------------------------------------------------------
'Sell Slab Button
'Checks for repeat IDs on Sold Slabs table and sends error message if there is. Also Checks for sold date and customer name,
'displays error message if one is not present. If all of these conditions are met then all records with purchased field checked
'are appended to sold table and deleted from slab entry table
'------------------------------------------------------------
Private Sub Command289_Click()
    On Error GoTo Command289_Click_Err
        'if date isn't empty and name isn't empty move slab, otherwise give user error message and exit sub
        If (Text537 <> "") Then
            If (Text543 <> "") Then
                Dim rsetID As Recordset
                Set rsetID = CurrentDb.OpenRecordset("Sold Slab IDs")
                rsetID.MoveLast
                If rsetID.RecordCount > 0 Then
                    rsetID.MoveFirst
                    Do Until rsetID.EOF = True
                        If Me![ID#] = rsetID![ID#] Then
                            MsgBox "You are trying to sell a duplicate ID! This ID already Exists in the Sold Slabs table. Please give slab a new ID before selling."
                            Exit Sub
                        End If
                        rsetID.MoveNext
                    Loop
                End If
                DoCmd.SetWarnings False
                CloseAllObjects
                DoCmd.OpenQuery "SellSelected"
                DoCmd.OpenQuery "DeleteSelected"
                DoCmd.OpenForm "Slab Entry Form"
                DoCmd.SetWarnings True
            Else
                MsgBox "You must enter Customer Name!"
                Exit Sub
            End If
        Else
            MsgBox "You must enter Date Sold!"
            Exit Sub
        End If
    
Command289_Click_Exit:
        Exit Sub
    
Command289_Click_Err:
        MsgBox Error$
        Resume Command289_Click_Exit

End Sub
'----------------------------------------------------------------------------------------------------------------------------------
'Save Record Button
'Save Changes Button, checks sold slabs ID# column and ensures that the ID for the Slab Entry table doesn't match a sold slab ID.
'Saves new record then closes all windows and reopens slab entry form
'----------------------------------------------------------------------------------------------------------------------------------
Private Sub Command396_Click()
    Dim rsetID As Recordset
    Set rsetID = CurrentDb.OpenRecordset("Sold Slab IDs")
    rsetID.MoveLast
    If rsetID.RecordCount > 0 Then
        rsetID.MoveFirst
        Do Until rsetID.EOF = True
            If Me![ID#] = rsetID![ID#] Then
                MsgBox "You are trying to save an ID that is being used in the Sold Slabs table! Please give slab a new ID before saving"
                Exit Sub
            End If
            rsetID.MoveNext
        Loop
    End If
    On Error GoTo Err_ErrorHandler
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Slab Entry Form"
    DoCmd.SetWarnings True
Exit_ErrorHandler:
    Exit Sub
Err_ErrorHandler:
    Select Case Err.Number
        Case 3316
            MsgBox "You are trying to save a record with an ID that is already taken.  Please choose a different ID."
            Resume Exit_ErrorHandler
        Case 3201
            MsgBox "You cannot change the wood code from here.  Please alter Wood Code from Species Form"
            Resume Exit_ErrorHandler
        Case Else
            MsgBox "You are trying to save a record with an ID that is already taken.  Please choose a different ID."
            Resume Exit_ErrorHandler
    End Select
End Sub
'----------------------------------------------------------------------
'Delete Slab Button
'Click event with password for delete slab button on Slab Entry Form
'----------------------------------------------------------------------
Private Sub Command403_Click()
    Dim strPasswd

    strPasswd = InputBox("Enter Password", "Are You Sure You Want To Delete This Slab???")

    'Check to see if there is any entry made to input box, or if
    'cancel button is pressed. If no entry made then exit sub.

    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "You must enter password to proceed", vbInformation, "Required Data"
        Exit Sub
    End If

    'If correct password is entered delete record
    'If incorrect password entered give message and exit sub

    If strPasswd = "Hardwood22" Then
        
        On Error Resume Next
        DoCmd.SetWarnings False
        RunCommand acCmdDeleteRecord
        CloseAllObjects
        DoCmd.OpenForm "Slab Entry Form"
        DoCmd.SetWarnings True
    Else
        MsgBox "Sorry, password is incorrect", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
End Sub


'************************************
'Page Selectors
'************************************
'opens form and closes open windows
Private Sub Command378_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Sold Slabs Form"
    DoCmd.SetWarnings True
End Sub
'opens form and closes open windows
Private Sub Command381_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Species Price Form"
    DoCmd.SetWarnings True
End Sub
'************************************
'Open Reports
'************************************
'open report but doesn't close other windows
Private Sub Command390_Click()
    DoCmd.OpenReport "Show In Stock Slabs", acViewReport
End Sub
'open report but doesn't close other windows
Private Sub Command393_Click()
    DoCmd.OpenReport "Show In Stock Dry Slabs", acViewReport
End Sub

