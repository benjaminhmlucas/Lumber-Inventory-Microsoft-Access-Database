VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sold Bundles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'*****************************
'Return Bundle Button

'move all records with purchased column unchecked back to Slab Entry Table
Private Sub Command471_Click()
        Dim strPasswd

    strPasswd = InputBox("Enter Password", "Are You Sure You Want To Delete This Slab???")

    'Check to see if there is any entry made to input box, or if
    'cancel button is pressed. If no entry made then exit sub.

    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "You must enter password to proceed", vbInformation, "Required Data"
        Exit Sub
    End If

    'If correct password is entered move record back to Slab Entry Table
    'If incorrect password entered give message and exit sub
    If strPasswd = "Hardwood22" Then
        Set rsetID = CurrentDb.OpenRecordset("In Stock IDs")
        rsetID.MoveLast
        If rsetID.RecordCount > 0 Then
            rsetID.MoveFirst
            Do Until rsetID.EOF = True
                If Me![ID#] = rsetID![ID#] Then
                    MsgBox "That ID is already in the Slab Entry Table! Please change duplicate ID in Slab Entry Form."
                    Exit Sub
                End If
                rsetID.MoveNext
            Loop
        End If
        On Error Resume Next
        DoCmd.SetWarnings False
        CloseAllObjects
        DoCmd.OpenQuery "UnsellUnselected"
        DoCmd.OpenQuery "DeleteUnselected"
        DoCmd.OpenForm "Sold Bundles"
        DoCmd.SetWarnings True
    Else
        MsgBox "Sorry, password is incorrect", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
End Sub
'********************************************************************
'Delete Button
'Click event with password for delete slab button on Slab Entry Form
'********************************************************************
Private Sub Command591_Click()
    Dim strPasswd

    strPasswd = InputBox("Enter Password", "Are You Sure You Want To Delete This Bundle???")

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
        DoCmd.OpenForm "Sold Bundles"
        DoCmd.SetWarnings True
    Else
        MsgBox "Sorry, password is incorrect.", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
End Sub
'************************************
'Page Selectors
'************************************
'opens form and closes open windows
Private Sub Command492_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "In Stock Bundle Entry"
    DoCmd.SetWarnings True
End Sub
'opens form and closes open windows
Private Sub Command495_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Species Pricing Tracker"
    DoCmd.SetWarnings True
End Sub
