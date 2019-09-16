Attribute VB_Name = "Module1"
'*********************************************************************************************
'Data Base Created By Ben Lucas 1/1/2019
'This database is licensed for use only to Ashley River Lumber
'Ashley River Lumber cannot resell or alter this database without permission of author,
'Ben Lucas
'*********************************************************************************************

Option Compare Database
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

'Determines Price of Wood by type and width from SlabPriceGuide Table
Function widthCategoriePrice(width As Double, SizeCategorie8to11 As Double, SizeCategorie12to20 As Double, SizeCategorie21to23 As Double, SizeCategorie24to30 As Double)
    Select Case width
'       Case Is < 8
'          widthCategoriePrice = "Small Piece"
       Case 8 To 11.99
          widthCategoriePrice = Format(SizeCategorie8to11, "Currency")
       Case 12 To 20.99
          widthCategoriePrice = Format(SizeCategorie12to20, "Currency")
       Case 21 To 23.99
          widthCategoriePrice = Format(SizeCategorie21to23, "Currency")
       Case 24 To 30
          widthCategoriePrice = Format(SizeCategorie24to30, "Currency")
'       Case Is > 30
'          widthCategoriePrice = "Priced By Piece"
    End Select
End Function

'Totals add on price for wood options
Function getAddOnPrice(tfColumn, priceColumn As Double)
    If tfColumn = -1 Then
        getAddOnPrice = Format(priceColumn, "Currency")
    Else
        getAddOnPrice = Format(0, "Currency")
    End If
End Function

'Adds all add on prices together
Function getAllPrices(Curly As Double, Burled As Double, Spalted As Double, birdsEye As Double, Pecky As Double, Ambrosia As Double, Sinker As Double, axeSinker As Double, Kilned As Double, Dressed As Double)
    getAllPrices = Format(Curly + Burled + Spalted + Crotch + birdsEye + Pecky + Ambrosia + Sinker + axeSinker + Kilned + Dressed, "Currency")
End Function

'Combine Add On Price and Width Price
Function getTotalPricePerBoardFoot(addOnPrice As Double, widthPrice As Double)
    getTotalPricePerBoardFoot = Format(addOnPrice + widthPrice, "Currency")
End Function

Sub appendSoldSlabs()
    Dim dbs As DAO.Database
    Dim lngRowsAffected As Long
    Dim lngRowsDeleted As Long
    
    Set dbs = CurrentDb
    
    ' Execute runs both saved queries and SQL strings
    dbs.Execute SellSelected, dbFailOnError
    
    ' Get the number of rows affected by the Action query.
    ' You can display this to the user, store it in a table, or trigger an action
    ' if an unexpected number (e.g. 0 rows when you expect > 0).
    lngRowsAffected = dbs.RecordsAffected
      
    dbs.Execute "DELETE FROM 'Sold Slabs' WHERE Bad", dbFailOnError
    lngRowsDeleted = dbs.RecordsAffected
End Sub

'used in form buttons
Option Compare Database
Option Explicit

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

'Highlights item in Listbox that is equal to field in textbox. Used in form button macros to manipulate item selected in list.
Function higlightListItem(list1 As ListBox, textbox1 As TextBox)
    list1 = textbox1
End Function
