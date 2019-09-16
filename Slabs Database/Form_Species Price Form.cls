VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Species Price Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'**********************************
'Page Selectors
'**********************************
'opens form and closes open windows
Private Sub Command153_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Slab Entry Form"
    DoCmd.SetWarnings True
End Sub
'opens form and closes open windows
Private Sub Command156_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Sold Slabs Form"
    DoCmd.SetWarnings True
End Sub
'----------------------------------------
'Save Changes button
'saves changes to Species Form in table
'----------------------------------------
Private Sub Command159_Click()
    On Error GoTo exitSUB
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Species Price Form"
    DoCmd.SetWarnings True
    Exit Sub
exitSUB:
    MsgBox "You are trying to save a record with an Wood Code that is already taken.  Please choose a different Wood Code."
    Exit Sub
End Sub

'----------------------------------------------------------------------
'Delete Slab Button
'Click event with password for delete slab button on Slab Entry Form
'----------------------------------------------------------------------
Private Sub Command274_Click()
    Dim strPasswd

    strPasswd = InputBox("Are You Sure You Want To Delete This Species???  If this species in use then delete action will fail.", "Enter Password")

    'Check to see if there is any entry made to input box, or if
    'cancel button is pressed. If no entry made then exit sub.

    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "You must enter password to proceed", vbInformation, "Required Data"
        Exit Sub
    End If

    'If correct password is entered delete record
    'If incorrect password entered give message and exit sub

    If strPasswd = "Hardwood22" Then
        
        On Error GoTo Err_ErrorHandler
        DoCmd.SetWarnings False
        RunCommand acCmdDeleteRecord
        CloseAllObjects
        DoCmd.OpenForm "Species Price Form"
        DoCmd.SetWarnings True
    Else
        MsgBox "Sorry, password is incorrect", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
Exit_ErrorHandler:
    Exit Sub
Err_ErrorHandler:
    MsgBox "You cannot delete a species that is in use. Database will be corrupted!"
    GoTo Exit_ErrorHandler
End Sub

