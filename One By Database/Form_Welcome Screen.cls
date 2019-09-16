VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Welcome Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'opens form and closes open windows
Private Sub Command8_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "In Stock Bundle Entry"
    DoCmd.SetWarnings True
End Sub
'opens form and closes open windows
Private Sub Command11_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Sold Bundles"
    DoCmd.SetWarnings True
End Sub
'opens form and closes open windows
Private Sub Command12_Click()
    DoCmd.SetWarnings False
    CloseAllObjects
    DoCmd.OpenForm "Species Pricing Tracker"
    DoCmd.SetWarnings True
End Sub
'open report but doesn't close other windows
Private Sub Command13_Click()
    'DoCmd.SetWarnings False
    'CloseAllObjects
    DoCmd.OpenReport "All In Stock Report", acViewReport
    'DoCmd.SetWarnings True
End Sub
'open report but doesn't close other windows
Private Sub Command14_Click()
    'DoCmd.SetWarnings False
    'CloseAllObjects
    DoCmd.OpenReport "Dry In Stock Report", acViewReport
    'DoCmd.SetWarnings True
End Sub

