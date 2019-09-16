Attribute VB_Name = "Module1"
'*********************************************************************************************
'Data Base Created By Ben Lucas 1/11/2019
'This database is licensed for use only to Ashley River Lumber upon payment for services
'Ashley River Lumber cannot resell or alter this database without permission of author,
'Ben Lucas
'*********************************************************************************************
Option Compare Database

'Iterates through all access objects and closes and save them(Forms, Reports, Tables Queries)
Public Function CloseAllObjects()
On Error GoTo errHandler
Dim aob As AccessObject
 
With CurrentData
    ' "Tables"
        For Each aob In .AllTables
            If aob.IsLoaded Then
                DoCmd.Close acTable, aob.Name, acSaveYes
            End If
        Next aob
    ' "Queries"
        For Each aob In .AllQueries
            If aob.IsLoaded Then
                DoCmd.Close acQuery, aob.Name, acSaveYes
            End If
        Next aob
End With
 
With CurrentProject
    ' "Forms"
        For Each aob In .AllForms
            If aob.IsLoaded Then
                DoCmd.Close acForm, aob.Name, acSaveYes
            End If
        Next aob
    ' "Reports"
        For Each aob In .AllReports
            If aob.IsLoaded Then
                DoCmd.Close acReport, aob.Name, acSaveYes
            End If
        Next aob
    ' "Pages"
        For Each aob In .AllDataAccessPages
            If aob.IsLoaded Then
                DoCmd.Close acDataAccessPage, aob.Name, acSaveYes
            End If
        Next aob
    ' "Macros"
        For Each aob In .AllMacros
            If aob.IsLoaded Then
                DoCmd.Close acMacro, aob.Name, acSaveYes
            End If
        Next aob
    ' "Modules"
        For Each aob In .AllModules
            If aob.IsLoaded Then
                DoCmd.Close acModule, aob.Name, acSaveYes
            End If
        Next aob
End With
 
errExit:
    Exit Function
errHandler:
    MsgBox "Error " & Err.Number & " " & Err.Description
    Resume errExit
End Function
 
'Determines if slab is dry based on a year passing since milling
Function readyToSell(dateMilled As Date, kilnedField)
'    if kilnedField is selected "Yes" is returned
    If kilnedField = -1 Then
        readyToSell = "Yes"
    Else
'       if wood was milled a year ago or more than "Yes" is returned otherwise "No" is returned
        If Now - dateMilled > 365.5 Then
            readyToSell = "Yes"
        Else
            readyToSell = "No"
        End If
    End If
End Function

'Returns feature price for wood options
Function getFeaturePrice(tfColumn, priceColumn As Double)
    If tfColumn = -1 Then
        getFeaturePrice = Format(priceColumn, "Currency")
    Else
        getFeaturePrice = Format(0, "Currency")
    End If
End Function

'Determines Price of Wood by type and linear foot categorie from Pricing Table
Function getPerFootPrice(thickness As Integer, width As Integer, c1X2 As Double, c1X4 As Double, c1X6 As Double, c1X8 As Double, c1X10 As Double, c1X12 As Double, c2X2 As Double, c2X4 As Double, c2X6 As Double, c2X8 As Double, c2X10 As Double, c2X12 As Double, c3X2 As Double, c3X4 As Double, c3X6 As Double, c3X8 As Double, c3X10 As Double, c3X12 As Double, c4X4 As Double, c6X6 As Double)
    Select Case thickness
        Case 1
            Select Case width
                Case 2
                    getPerFootPrice = Format(c1X2, "Currency")
                Case 4
                    getPerFootPrice = Format(c1X4, "Currency")
                Case 6
                    getPerFootPrice = Format(c1X6, "Currency")
                Case 8
                    getPerFootPrice = Format(c1X8, "Currency")
                Case 10
                    getPerFootPrice = Format(c1X10, "Currency")
                Case 12
                    getPerFootPrice = Format(c1X12, "Currency")
                Case Else
                    getPerFootPrice = "Wrong Width!"
            End Select
        Case 2
            Select Case width
                Case 2
                    getPerFootPrice = Format(c2X2, "Currency")
                Case 4
                    getPerFootPrice = Format(c2X4, "Currency")
                Case 6
                    getPerFootPrice = Format(c2X6, "Currency")
                Case 8
                    getPerFootPrice = Format(c2X8, "Currency")
                Case 10
                    getPerFootPrice = Format(c2X10, "Currency")
                Case 12
                    getPerFootPrice = Format(c2X12, "Currency")
                Case Else
                    getPerFootPrice = "Wrong Width!"
            End Select
        Case 3
            Select Case width
                Case 2
                    getPerFootPrice = Format(c3X2, "Currency")
                Case 4
                    getPerFootPrice = Format(c3X4, "Currency")
                Case 6
                    getPerFootPrice = Format(c3X6, "Currency")
                Case 8
                    getPerFootPrice = Format(c3X8, "Currency")
                Case 10
                    getPerFootPrice = Format(c3X10, "Currency")
                Case 12
                    getPerFootPrice = Format(c3X12, "Currency")
                Case Else
                    getPerFootPrice = "Wrong Width!"
            End Select
        Case 4
            Select Case width
                Case 4
                    getPerFootPrice = Format(c4X4, "Currency")
                Case Else
                    getPerFootPrice = "Wrong Width!"
            End Select
        Case 6
            Select Case width
                Case 6
                    getPerFootPrice = Format(c6X6, "Currency")
                Case Else
                    getPerFootPrice = "Wrong Width!"
            End Select
        Case Else
            getPerFootPrice = "Wrong Thickness!"
    End Select
End Function

'Sums price per foot for features
Function getFeaturesPerFootTotal(sinkerPrice As Double, ambrosiaPrice As Double, axeSinkerPrice As Double, birdsEyePrice As Double, burledPrice As Double, curlyPrice As Double, peckyPrice As Double, spaltedPrice As Double, kilnedPrice As Double, dressedPrice As Double)
    getFeaturesPerFootTotal = Format((sinkerPrice + ambrosiaPrice + axeSinkerPrice + birdsEyePrice + burledPrice + curlyPrice + peckyPrice + spaltedPrice + kilnedPrice + dressedPrice), "Currency")
End Function

'adds price per foot column and features price per foot columns
Function getTotalPerFoot(pricePerFoot As Double, featuresPricePerFoot As Double)
    getTotalPerFoot = Format((pricePerFoot + featuresPricePerFoot), "Currency")
End Function

'Highlights item in Listbox that is equal to field in textbox. Used in form button macros to manipulate item selected in list.
Function higlightListItem(list1 As ListBox, textbox1 As TextBox)
    list1 = textbox1
End Function


