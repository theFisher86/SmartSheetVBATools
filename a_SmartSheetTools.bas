' ======================================================================================================================================================================
' **Smart Sheet Tools**
' Version 1.2
Public Function SmartSheetToolsVersion() As String
    Debug.Print ("1.2")
    SmartSheetToolsVersion = "1.2"
End Function
' Tools and generalized functions to help with using VBA to talk directly to SmartSheet & and it's API
'
' Code Written by Aaron Fisher of FishcoDesign
'
' Utilizing :
'   VBA-Web 4.0.20 library from : https://github.com/VBA-tools
'   VBA-Jason 2.0.1 library from : https://github.com/VBA-tools/VBA-JSON
'   SmartSheet API 2.0 : http://smartsheet-platform.github.io/api-docs
' ======================================================================================================================================================================
'


Public Function SmartSheetGet(Request As String, Optional CustomToken As String = "1") As WebResponse
'   =======================================================================================
'   This will perform a HTTP Get request and return an Unparsed WebResponse from the
'   SmartSheet API.  The Request should be in the format of what comes AFTER the SmartSheet
'   Base URL (https://api.smartsheet.com/2.0/).  See the SmartSheet API for more information
'   =======================================================================================
'

' Check for sync disable and setup API_Token
    If CustomToken = "1" Then
        API_Token = Range("API_Token")
    Else:
        API_Token = CustomToken
    End If
    If Range("Sync_Disabled") = "TRUE" Then
        boxResponse = MsgBox("SmartSheet Syncing is disabled", vbCritical, "SmartSheetSync Disabled")
        End
    End If

' Setup SmartSheet as a WebClient, setup Authentication & Proxy Stuff, and assign Base URL
    Dim SmartSheet As New WebClient
    Dim Options As New Dictionary
        
    SmartSheet.BaseUrl = "https://api.smartsheet.com/2.0/"                                  ' Set Base URl
    
    SmartSheet.EnableAutoProxy = True                                                      ' Get Proxy Stuff


    Dim Response As WebResponse
    Dim Headers As New Collection

    Set Response = SmartSheet.GetJson(Request)
    Headers.Add WebHelpers.CreateKeyValue("Authorization", "Bearer " & API_Token)          ' Add Authentication Header
    Options.Add "Headers", Headers
    Set Response = SmartSheet.GetJson(Request, Options)

' Return SmartSheet JSON File as Object
    If Response.StatusCode = WebStatusCode.Ok Then
        Set SmartSheetGet = Response
    Else
        boxResponse = MsgBox("Error!" & Chr(10) & _
                            "Status Code : " & Response.StatusCode & Chr(10) & _
                            "Code Description : " & Response.StatusDescription & Chr(10) & _
                            "Content : " & Response.Content, vbCritical, "Error!")

        Debug.Print "Error!"
        Debug.Print Response.StatusCode
        Debug.Print Response.StatusDescription
        Debug.Print Response.Content
        End
    End If

End Function

Public Function BuildDictionary(WebResponse As WebResponse) As Scripting.Dictionary
'   =======================================================================================
'   Pass a WebResponse from SmartSheet into this function from the form /sheets/{sheetID}
'   This function will return a Scripting.Dictionary.  You can utilize it by calling
'   the row number or row Id and then the cell Column Name or Column ID
'   Ex :
'
'    Dim APIReturn As WebResponse                               this is the Sheet ID
'    Set APIReturn = a_SmartSheetTools.SmartSheetGet("sheets/   8944216010188676")
'    Set theDictionary = a_SmartSheetTools.BuildDictionary(APIReturn)
'                                *this is the row ID*       *This is the column name*
'    Debug.Print (theDictionary ("8455186625652612")        ("Version"))
'   -or-                        *this is the row number*    *this is the column ID number*
'    Debug.Print (theDictionary (1)                         ("4873887012939652"))
'
'   You can use any combination of the methods above
'   =======================================================================================

    Dim APIData As Object
    Set APIData = WebHelpers.ParseJson(WebResponse.Content)
    
    Dim Row As Dictionary
    Dim Cell As Dictionary
    Set columnNames = CreateObject("Scripting.Dictionary")
    
    ' Save column names with columnId
    For Each Column In APIData("columns")
        columnNames.Add Column("id"), Column("title")
    Next
    ' Create Dictionary to hold all the rows and their contents
    Set allrows = CreateObject("Scripting.Dictionary")
    For Each Row In APIData("rows")
'        allrows.Add "zrowId", Row("rowID")
'        allrows.Add "znumber", Row("rowNumber")
        ' Create Dictionary to hold row contents
        Set rowcontents = CreateObject("Scripting.Dictionary")
            rowcontents.Add "topLevel", Row
        For Each Cell In Row("cells")
            ' Add entry with columnId associated with the value
            rowcontents.Add Cell("columnId"), Cell("value")
            ' Get the column Name
            colName = columnNames(Cell("columnId"))
            ' Add entry with column Name associated with the value
            rowcontents.Add colName, Cell("value")
        Next Cell
        ' Add row to allRows Dictionary by rowID
        allrows.Add Row("id"), rowcontents
        ' Add row to allRows Dictionary by rowNumber
        allrows.Add Row("rowNumber"), rowcontents
    Next Row
    Set BuildDictionary = allrows
End Function

Public Function CountData(ColumnName As String, RawData As WebResponse) As Scripting.Dictionary
'   =======================================================================================
'   This function will parse the RawData WebResponse (obtained from SmartSheetGet) and
'   will return as a Dictionary with all of the unique values in the ColumnName as keys
'   and the counts of those values as the value.  Used for charts and the like.
'   =======================================================================================
'

' Double check to make sure the data is loaded
    If RawData Is Nothing Then
        Debug.Print ("Data Not Loaded")
'       Uncomment the 3 lines below for Production
'       boxResponse = MsgBox("Data Not Loaded", vbCritical, "Data Not Loaded!")
'       formUpdateData.Show
'       End
    End If
    
    ' Parse Data
    Dim Parsed As Object
    Set Parsed = WebHelpers.ParseJson(RawData.Content)
    
    ' Build Dictionary for Column Names
    Set colNames = CreateObject("Scripting.Dictionary")
    Set ParsedCols = Parsed("columns")
    For Each col In ParsedCols
        colNames.Add col("title"), col("id")
    Next
    
    ' Setup Dictionaries to hold counted values
    Set colCounts = CreateObject("Scripting.Dictionary")
    
    ' Get into each Row
    Set ParsedRows = Parsed("rows")
    For Each r In ParsedRows
        Debug.Print ("Row ID : " & r("id"))
        Debug.Print ("Row Number : " & r("rowNumber"))
        For Each C In r("cells")
            ' Count values in column and assign to colCounts Dictionary
            ' If a new value is found a new key will be created
            If C("columnId") = colNames(ColumnName) Then
                If colCounts.Exists(C("value")) Then
                    colCounts.Item(C("value")) = colCounts.Item(C("value")) + 1
                Else
                    colCounts.Add C("value"), 1
                End If
            End If
        Next
    Next
    
' You may disable the For Next segment below for Production
    For Each Line In colCounts
        Debug.Print (Line & " : " & colCounts(Line))
    Next
    
    Set CountData = colCounts
End Function

Public Function CheckForNewVersion(FileName As String)
' ===============================================================================================
' This function will access the File Repository SmartSheet and check to see if there is a newer
' version of FileName than the current version on the PC.  It will also automatically check to
' make sure this Workbook is using the latest version of SmartSheetTools.  Will return TRUE if
' a new version of either file is available, FALSE if this is the current version.
' ===============================================================================================
    Dim currentVersion As Single
    Dim APIReturn As WebResponse
    Set APIReturn = a_SmartSheetTools.SmartSheetGet("sheets/8944216010188676")
    Set repoRows = a_SmartSheetTools.BuildDictionary(APIReturn)
    ' Check to see if filename matches actual File Name on PC
    If Not FileName = ThisWorkbook.Name Then
        boxResponse = MsgBox("You have renamed this file from it's default name.  This could cause problems.  Please close and rename this file as " & FileName & Chr(10) & ". This Workbook will close when you click OK.", vbCritical, "File Name Error")
        ThisWorkbook.Close
        End
    End If
        
    For Each Row In repoRows
        ' Check current file version
        If repoRows(Row)("Current Version") = True And repoRows(Row)("File Name") = FileName Then
            currentVersion = repoRows(Row)("Version")
            Debug.Print ("Current Version of " & FileName & " : " & currentVersion)
            CheckForNewVersion = currentVersion
        End If
        ' Check SmartSheetTools version
        If repoRows(Row)("Current Version") = True And repoRows(Row)("File Name") = "SmartSheetTools" Then
            currentSmartSheetTools = repoRows(Row)("Version")
            Debug.Print ("Current SmartSheetToolsVersion : " & currentSmartSheetTools)
            If currentSmartSheetTools > a_SmartSheetTools.SmartSheetToolsVersion Then
                boxResponse = MsgBox("Your SmartSheetTools is out of date.  Download and import the latest version from the SmartSheet File Repository and import into this workbook.  Ask Aaron Fisher if you need help.", vbCritical, "SmartSheetTools Out of Date!")
            End If
        End If
    Next Row
    
End Function
Public Function RowPost(ByVal Cells As Scripting.Dictionary, Optional Location As String = "toBottom", Optional rowID As String = "", Optional CustomToken As String = "1")
'   This will post a Row to SmartSheet.  Please note that the Cells Dictionary that is passed must be in the format
'   listed below:     Column ID #         Value
'        myCells.Add "3106015123138436", "OEM"
'        myCells.Add "5357814936823684", "O137"

' Check for sync disable and setup API_Token
    If CustomToken = "1" Then
        API_Token = Range("API_Token")
    Else:
        API_Token = CustomToken
    End If
    If Range("Sync_Disabled") = "TRUE" Then
        boxResponse = MsgBox("SmartSheet Syncing is disabled", vbCritical, "SmartSheetSync Disabled")
        End
    End If

' Setup Client
    Dim SmartSheet As New WebClient
    SmartSheet.BaseUrl = "https://api.smartsheet.com/2.0/"
    SmartSheet.EnableAutoProxy = True

' Setup Request
    'Grab SSID
    SheetID = Cells("ssid")
    Dim PostRequest As New WebRequest
    PostRequest.Resource = "sheets/{sID}/rows"
    If rowID = "" Or rowID = "False" Then
        PostRequest.Method = WebMethod.HTTPpost
        Debug.Print ("using POST")
    Else
        PostRequest.Method = WebMethod.HttpPut
        Debug.Print ("using PUT")
    End If
    PostRequest.AddUrlSegment "sID", SheetID
    PostRequest.Format = WebFormat.JSON                         ' This handles the content-type Header and all the other necessary JSON stuff
    PostRequest.AddHeader "Authorization", "Bearer " & API_Token
    
' Build String from Dictionary
    Dim bodyString As String
    
    ' Setup the beginning of the bodyString
    If rowID = "" Then
        bodyString = "{""" & Location & """:true, ""cells"": ["
    Else:
        bodyString = "{""id"": """ & rowID & """, """ & Location & """:true, ""cells"": ["
    End If
    firstOne = True
    
    '{"toTop":true, "cells": [{"columnId": 3106015123138436, "value" : "OEM"}, {"columnId": 5357814936823684, "value" : "O137"}, {"columnId": 5200324928530308, "value" : "Bitcoin mining chip"}, {"columnId": 2865909271422852, "value" : ""}, {"columnId": 2402327681361796, "value" : ""}]}]
    ' Build the cells portion
    For Each C In Cells
        If C <> "ssid" Then
            If firstOne = True Then
                bodyString = bodyString & "{""columnId"": """ & C & """, ""value"" : """ & Cells(C) & """}"
                firstOne = False
            Else:
                bodyString = bodyString & ", {""columnId"": """ & C & """, ""value"" : """ & Cells(C) & """}"
            End If
        End If
    Next
    
    ' Finish off the body string
    bodyString = bodyString & "]}]"
    Debug.Print (bodyString)
    ' Set bodyString as the Body
    PostRequest.Body = bodyString

    
    Dim Response As WebResponse
    Set Response = SmartSheet.Execute(PostRequest)
    Set RowPost = Response

End Function

Public Function ColIDByName(APIResponse As Object, ColumnName As String)
    Set APIData = JsonConverter.ParseJson(APIResponse.Content)
    For Each d In APIData("data")
        If d("title") = ColumnName Then
            colID = d("id")
            ColIDByName = colID
        End If
    Next
    
End Function

Function CollectionToArray(C As Collection) As Variant()
    Dim A() As Variant
    ReDim A(0 To C.Count - 1)
    Dim i As Integer
    For i = 1 To C.Count
        A(i - 1) = C.Item(i)
    Next
    CollectionToArray = A

End Function

Public Function SmartSheetDateToExcelDate(SmartSheetDate As String)
    ExcelDate = Left(SmartSheetDate, 4) & "/" & Mid(SmartSheetDate, 6, 2) & "/" & Mid(SmartSheetDate, 9, 2)
    SmartSheetDateToExcelDate = ExcelDate
End Function

Public Function Prefill(PRID As String, SheetID As String, ByVal Map As Scripting.Dictionary)
        Dim SearchResponse As WebResponse
        Dim RowResponse As WebResponse
        
        Set SearchResponse = SearchSmartSheet(PRID, SheetID)
        Set SearchData = WebHelpers.ParseJson(SearchResponse.Content)
        Debug.Print "Json response : "
        Debug.Print WebHelpers.ConvertToJson(SearchData)
        
        Set RowObject = SearchData("results")
    ' Check number of rows returned, throw error if multiple
        If RowObject.Count > 1 Then
            errMessage = MsgBox("ERROR! : Multiple rows retuned, please check PRID and/or enter manually", vbCritical, "Error!")
        Else:
    ' Assign Row ID
            For Each i In RowObject
                Debug.Print (i("objectId"))
                Row_id = i("objectId")
            Next
        End If
        Debug.Print ("Row ID :" & Row_id)
    ' Get Row
        Set RowResponse = SmartSheetGet("sheets/" & SheetID & "/rows/" & Row_id)
        Set RowData = JsonConverter.ParseJson(RowResponse.Content)
        Set RowData = RowData("cells")
        
        ' Map the data
        Debug.Print ("Fields mapped: ")
        For Each C In RowData
            SScolID = C("columnId")
            If Map.Exists(SScolID) = True Then
                Range(Map(SScolID)) = C("value")
                Debug.Print (Map(SScolID))
            End If
        Next
        
        Final = MsgBox("Data mapped from SmartSheet!", vbDefaultButton1, "Data Map Success!")
End Function

Public Function ProcessAPI(APIObject As WebResponse, returnFields As String)
    Set ParsedResponse = JsonConverter.ParseJson(APIObject.Content)
    Debug.Print (ParsedResponse(returnFields))
    ProcessAPI = ParsedResponse(returnFields)
End Function

Public Function SearchSmartSheet(query As String, SheetID As String)
' Perform Search Query on SmartSheet Sheet
    Debug.Print (SheetID & "-" & query)
    Set APIReturn = SmartSheetGet("/search/sheets/" & SheetID & "?query=" & query)
    Debug.Print APIReturn.Content
    Set SearchSmartSheet = APIReturn
End Function

Public Function GetColumnOptions(ColumnID As String, APIReturn As WebResponse)
    Dim ColOptions As New Collection
    Set APIData = JsonConverter.ParseJson(APIReturn.Content)
    For Each i In APIData("data")
        If i("id") = ColumnID Then
            For Each o In i("options")
                ColOptions.Add o
                Debug.Print o
            Next
        End If
    Next
    Set GetColumnOptions = ColOptions
    

End Function


Public Function SearchForPRID(PRID As String, SheetID As String)
    Dim APIReturn As WebResponse
    Set APIReturn = SearchSmartSheet(PRID, SheetID)
    Dim APIData As Object
    Set APIData = JsonConverter.ParseJson(APIReturn.Content)
' Check to see if PRID is already added and generate error if it is.
    If APIData("totalCount") = 1 Then
        Debug.Print ("PRID Exists")
        For Each i In APIData("results")
            For Each C In i("contextData")
                Debug.Print (C)
                thing = MsgBox("This PRID has already been added under " & C & ".", vbCritical, "Existing PRID Found")
                
            Next
            Debug.Print "objecttype : " & i("objectType")
            Debug.Print "objectid : " & i("objectId")
            If i("objectType") = "row" Then
                rowID = i("objectId")
                Debug.Print ("rowid instance 1 : " & rowID)
            End If
            
        Next
        SearchForPRID = rowID
        Debug.Print ("rowid instance 2 : " & rowID)
    ElseIf APIData("totalCount") > 1 Then
        boxResponse = MsgBox("multiple prids returned", vbCritical, "error")
        End
    Else
        Debug.Print ("SearchForPRID - Success")
        SearchForPRID = "False"
    End If
End Function

Public Function CycleFields(record, fields)
    ' Cycle Through fields in LookupFields Array
    For Each field In fields
        Debug.Print (record(field))
        ' Write value to current field
        ActiveCell.Value = (record(field))
        ' Move right one column
        ActiveCell.Offset(0, 1).Select
    Next
    ' Move down a row
    ActiveCell.Offset(1, 0).Select
    ' Count size of array and move left that many columns
    FieldCount = UBound(fields) - LBound(fields) + 1
    ActiveCell.Offset(0, FieldCount * -1).Select
End Function








