Attribute VB_Name = "HIAACore"
Sub GrabData(control As IRibbonControl)
    
' Ask user to enter an API to fetch data from
    api_url = Application.InputBox(Prompt:="Data URL", _
                                   Title:="HIAA Data Access", _
                                   Default:="https://api.weather.gc.ca/collections/climate-hourly/items?f=json&lang=en-CA&limit=10&skipGeometry=false&offset=0&CLIMATE_IDENTIFIER=8202251&LOCAL_MONTH=8&LOCAL_YEAR=2024")
    
    'alternative examples
    'api_url = "https://www.bankofcanada.ca/valet/observations/FXUSDCAD%2CFXEURCAD/json?start_date=2023-01-23&end_date=2023-07-19"
    
    If api_url = "" Then
        MsgBox "Failed to enter a URL."
        Exit Sub
    End If
    
' Create a new HTTP request
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", api_url, False
        .setRequestHeader "Content-type", "application/json"
        '.setRequestHeader "Accept", "application/json"
        '.setRequestHeader "Authorization", "Bearer " & authKey
        '.setRequestHeader "Authorization", "Key " & authKey
        .Send
        
        If .Status <> 200 Then
            Debug.Print .Status
            MsgBox "ERROR: Unable to access data, check data URL."
            Exit Sub
        End If
        
        'store JSON response from API request in a string
        sJSONString = .responsetext
        
        'if length > 1,000,000 then it is too much data to process timely and needs to be shut down
        Debug.Print "JSON string length: " & Len(sJSONString)
        If Len(sJSONString) > 1000000 Then
            MsgBox "ERROR: JSON response is > 1M characters."
            Exit Sub
        End If
        
    End With

' Parse JSON from a string into a dictionary / collection of items
    Dim vJSON
    Dim sState As String
    JSON.Parse sJSONString, vJSON, sState
    
    If sState = "Error" Then
        MsgBox "Invalid JSON"
        End
    End If
    
' JSONs are typically returned with multiple components, such as the number of records returned, time the api was requested, and the actual desired data
' This asks the user what part of the JSON we actually want to process and dump into the spreadsheet
    ' populate a userform with the top level items in the JSON dictionary and open the userform for the user to make a selection
    With api_item
    
        .ListBox1.Clear
        
        ' populate the userform
        For Each thing In vJSON.Keys
            .ListBox1.AddItem thing
        Next
    
        ' set the userform position to the middle of the screen
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
    
    ' place the users selection into the string "api_selection"
    api_selection = api_item.ListBox1.Value
    
    ' if the selection is a 2D array, then it will be of type Dictionary which is an object and requires the "Set" part of "Set vJSON"
    Debug.Print TypeName(vJSON(api_selection))
    If Not (IsNull(api_selection)) Then
        If TypeName(vJSON(api_selection)) = "Dictionary" Then
            Set vJSON = vJSON(api_selection)
        Else
            vJSON = vJSON(api_selection)
        End If
    End If
    
' Convert raw JSON to 2d array and output to active cell
    ' if the dataset is a 2D array, then we loop through both the header and the data
    If TypeName(vJSON) = "Variant()" Or TypeName(vJSON) = "Dictionary" Then
        Dim aData()
        Dim aHeader()
        JSON.ToArray vJSON, aData, aHeader
        
        If UBound(aData, 1) + 1 > 10000 Then
            MsgBox "ERROR: Data is > 10,000 rows."
            Exit Sub
        End If
        
        With ActiveCell
            Output1DArray .Cells(1, 1), aHeader
            Output2DArray .Cells(2, 1), aData
        End With
    ' if the dataset is NOT a 2D array and is instead just a single element, we dump the name and value in adjacent cells
    Else
        With ActiveCell
            .Cells(1, 1) = api_selection
            .Cells(1, 2) = vJSON
        End With
    End If
        
End Sub

Sub Output1DArray(oDstRng As Range, aCells As Variant)

    ' adjust the size of the active cell for the size of the dataset, just number of columns
    ' set the value of the active cells to the values of the dataset
    With oDstRng
        .Parent.Select
        With .Resize(1, UBound(aCells) - LBound(aCells) + 1)
            .NumberFormat = "@"
            .Value = aCells
        End With
    End With

End Sub

Sub Output2DArray(oDstRng As Range, aCells As Variant)

    ' adjust the size of the active cell for the size of the dataset, both number of rows and columns
    ' set the value of the active cells to the values of the dataset
    With oDstRng
        .Parent.Select
        With .Resize( _
                UBound(aCells, 1) - LBound(aCells, 1) + 1, _
                UBound(aCells, 2) - LBound(aCells, 2) + 1)
            .NumberFormat = "@"
            .Value = aCells
        End With
    End With

End Sub
