VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Species Pricing Tracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'When page opens:
Private Sub Form_Load()
    List887.RowSource = "Features"
End Sub
'*********************************
'List Control options
'*********************************
Private Sub Features_Click()
    List887.RowSource = "Features"
End Sub

Private Sub OneBys_Click()
    List887.RowSource = "1bys"
End Sub

Private Sub TwoBys_Click()
    List887.RowSource = "2bys"
End Sub

Private Sub ThreeBys_Click()
    List887.RowSource = "3bys/Misc"
End Sub
'**********************************
'Page Selectors
'**********************************
'opens form and closes open windows
Private Sub Command907_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "In Stock Bundle Entry"
    DoCmd.SetWarnings True
End Sub
'opens form and closes open windows
Private Sub Command919_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Sold Bundles"
    DoCmd.SetWarnings True
End Sub
'**********************************
'Save button
'**********************************
Private Sub Command67_Click()
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Species Pricing Tracker"
    DoCmd.SetWarnings True
End Sub

'********************************************************************
'Delete Button
'Click event with password for delete slab button on Slab Entry Form
'********************************************************************
Private Sub Command1160_Click()
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
        DoCmd.OpenForm "Species Pricing Tracker"
        DoCmd.SetWarnings True
    Else
        MsgBox "Sorry, password is incorrect.", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
Exit_ErrorHandler:
    Exit Sub
Err_ErrorHandler:
    MsgBox "You cannot delete a species that is in use. Database will be corrupted!"
    GoTo Exit_ErrorHandler
End Sub


