VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_In Stock Bundle Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'********************************************************************
'Drop Down List selector function
'Requeries Width(in) combobox after update and then set width value to first record and first column then thickness value to first column
'********************************************************************
Private Sub Combo386_AfterUpdate()
    Combo625.Requery
    Me![Width(in)] = Me!Combo625.ItemData(0)
    Me![Width(in)] = Me!Combo625.Column(1)
    Me![Thick(in)] = Combo386.Column(1)
    MsgBox "Thickness resets Width(in).  Please choose a Width."
    SendKeys "{TAB}"
End Sub

'********************************************************************
'Delete Button
'Click event with password for delete slab button on Slab Entry Form
'********************************************************************
Private Sub Delete_Bundle_Entry_Form_Click()
    Dim strPasswd

    strPasswd = InputBox("Are You Sure You Want To Delete This Bundle???", "Enter Password")

    'Check to see if there is any entry made to input box, or if
    'cancel button is pressed. If no entry made then exit sub.

    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "You must enter password to proceed", vbInformation, "Required Data"
        Exit Sub
    End If

    'If correct password is entered delete record
    'If incorrect password entered give message and exit sub

    If strPasswd = "Hardwood22" Then
        
        DoCmd.SetWarnings False
        On Error GoTo Err_ErrorHandler
        RunCommand acCmdDeleteRecord
        CloseAllObjects
        DoCmd.OpenForm "In Stock Bundle Entry"
        DoCmd.SetWarnings True
    Else
        MsgBox "Sorry, password is incorrect.", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
'********************************************************************************************************************************
'Save Button
'Save Changes Button, checks sold slabs ID# column and ensures that the ID for the Slab Entry table doesn't match a sold slab ID.
'Saves new record then closes all windows and reopens slab entry form
'********************************************************************************************************************************
Private Sub Save_Bundle_Entry_Form_Click()
    Dim rsetID As Recordset
    Set rsetID = CurrentDb.OpenRecordset("Sold IDs")
    rsetID.MoveLast
    If rsetID.RecordCount > 0 Then
        rsetID.MoveFirst
        Do Until rsetID.EOF = True
            If Me![ID#] = rsetID![ID#] Then
                MsgBox "You are trying to save an ID that is being used in the Sold table! Please give bundle a new ID before saving."
                GoTo Exit_ErrorHandler
            End If
            rsetID.MoveNext
        Loop
    End If
    If Combo386 = "How Thick?" Then
        MsgBox "Please Pick a Thickness."
        GoTo Exit_ErrorHandler
    End If
    If Combo625 = "How Wide?" Then
        MsgBox "Please Pick a Width."
        GoTo Exit_ErrorHandler
    End If
        If Combo398 = "How Long?" Then
        MsgBox "Please Pick a Length."
        GoTo Exit_ErrorHandler
    End If
    On Error GoTo Err_ErrorHandler
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "In Stock Bundle Entry"
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
'------------------------------------------------------------
'Command404_Click() - Sell Selected Button
'Checks for repeat IDs on Sold Slabs table and sends error message if there is. Also Checks for sold date and customer name,
'displays error message if one is not present. If all of these conditions are met then all records with purchased field checked
'are appended to sold table and deleted from slab entry table
'------------------------------------------------------------
Private Sub Command404_Click()
    On Error GoTo Command404_Click_Err
        'if date isn't empty and name isn't empty move slab, otherwise give user error message and exit sub
        If ([Date Sold] <> "") Then
            If ([Customer Name] <> "") Then
                Dim rsetID As Recordset
                Set rsetID = CurrentDb.OpenRecordset("Sold IDs")
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
                DoCmd.OpenForm "In Stock Bundle Entry"
                DoCmd.SetWarnings True
            Else
                MsgBox "You must enter Customer Name!"
                Exit Sub
            End If
        Else
            MsgBox "You must enter Date Sold!"
            Exit Sub
        End If
    
Command404_Click_Exit:
        Exit Sub
    
Command404_Click_Err:
        MsgBox Error$
        Resume Command404_Click_Exit
'************************************
'Page Selectors
'************************************
End Sub
'opens form and closes open windows
Private Sub Command1279_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Sold Bundles"
    DoCmd.SetWarnings True
End Sub
'opens form and closes open windows
Private Sub Command1307_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Species Pricing Tracker"
    DoCmd.SetWarnings True
End Sub
'************************************
'Open Reports
'************************************
'open report but doesn't close other windows
Private Sub Command1310_Click()
    DoCmd.OpenReport "All In Stock Report", acViewReport
End Sub
'open report but doesn't close other windows
Private Sub Command1315_Click()
    DoCmd.OpenReport "Dry In Stock Report", acViewReport
End Sub
