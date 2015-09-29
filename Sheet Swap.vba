Option Explicit
Dim swApp           As SldWorks.SldWorks
'------------------------------------------------------------------------------'
Sub modifyAndCheckin()

'local object declarations'
Dim swDrawing       As SldWorks.DrawingDoc
Dim swPart          As SldWorks.ModelDoc2
Dim swExtension     As SldWorks.ModelDocExtension
Dim swModel         As SldWorks.ModelDoc2
Dim PDMConnection   As IPDMWConnection
Dim PDMPart         As PDMWDocument
Dim PDMDrawing      As PDMWDocument
Dim checkInDocument As PDMWDocument
Dim swCustPropMgr   As SldWorks.CustomPropertyManager

'local data declarations'
Dim errors          As Long
Dim warnings        As Long
Dim CP_Finish       As String
Dim CP_Change       As String
Dim CP_ChangeDate   As String
Dim CP_DrawnBy      As String
Dim CP_DrawnDate    As String
Dim CP_Material     As String
Dim drawingName     As String
Dim modelName       As String
Dim modelnumber()   As String
Dim j               As Integer
Dim bool            As Boolean

'local constant declarations'
Const inputFile     As String = "C:\Users\jpettit\Desktop\SCRIPTS\filesToChange.txt"
Const outputFile    As String = "C:\Users\jpettit\Desktop\SCRIPTS\fileChangeOutput.txt"
Const tempDir       As String = "X:\Engineering\TEMP\"
Const pdmName       As String = "jpettit"
Const pdmLogin      As String = "CDGshoxs!"
Const pdmServer     As String = "SHOXS1"

'set this constant to TRUE to enable test mode, where nothing will be checked in
'and vendor files won't be saved. Also won't close items after modification
Const testMode      As Boolean = True

'Custom property values to be written to each file'
CP_Finish = "002"
CP_Change = "CHANGED FINISH SPECIFICATION"
CP_ChangeDate = UCase(Format(Now, "d-MMM-yy"))
CP_DrawnBy = "JP"
CP_DrawnDate = Format(Now, "mm/d/yy")
CP_Material = "6061-T6 ALLOY"

'initialize objects and start the PDM connection'
Set swApp = Application.SldWorks
Open outputFile For Output As #2
Set PDMConnection = CreateObject("PDMWorks.PDMWConnection")
PDMConnection.Login pdmName, pdmLogin, pdmServer

'function call which returns array of part numbers to change'
'Part numbers are read from an external file'
modelnumber() = readData(inputFile)
Debug.Print UBound(modelnumber()) + 1 & " PARTS TO CHANGE"

'main part number inspection loop. loops through each part number read from'
'the external file, opens the part, modifies it, and checks it in'
For j = LBound(modelnumber) To UBound(modelnumber)

    'set the drawing and model names, and find the PDM objects they represent
    'if the drawing or part cant be set, they aren't in the vault, and we need
    'to skip to the next loop
    drawingName = modelnumber(j) + ".SLDDRW"
    modelName = modelnumber(j) + ".SLDPRT"
    Set PDMPart = PDMConnection.GetSpecificDocument(modelName)
    If PDMPart Is Nothing Then
        Debug.Print modelnumber(j) & " PART NOT IN VAULT"
        Print #2, modelnumber(j) & ", PART NOT IN VAULT"
        GoTo nextLoop
    End If
    Set PDMDrawing = PDMConnection.GetSpecificDocument(drawingName)
    If PDMDrawing Is Nothing Then
        Debug.Print modelnumber(j) & " DRAWING NOT IN VAULT"
        Print #2, modelnumber(j) & ", DRAWING NOT IN VAULT"
        GoTo nextLoop
    End If

    'if the initatior of the PDM connection is not already the owner of the'
    'part and drawing document, take ownership here.
    'tests for test mode, and ownership availabiltiy. if unavailable, skips to
    'next loop
    If testMode = True Then
        If PDMPart.Owner = "" Then
            PDMPart.TakeOwnership
        Else
            If PDMPart.Owner <> pdmName Then
                Debug.Print modelnumber(j) & " PART OWNERSHIP NOT AVAILABLE"
                Print #2, modelnumber(j) & ", PART OWNERSHIP NOT AVAILABLE"
                GoTo nextLoop
            End If
        End If
        If PDMDrawing.Owner = "" Then
            PDMDrawing.TakeOwnership
        Else
            If PDMDrawing.Owner <> pdmName Then
                Debug.Print modelnumber(j) & " DRAWING OWNERSHIP NOT AVAILABLE"
                Print #2, modelnumber(j) & ", PART OWNERSHIP NOT AVAILABLE"
                GoTo nextLoop
            End If
        End If
    End If

    'save the model and drawing retrived from PDM into the temp directory'
    swApp.QuitDoc drawingName
    swApp.QuitDoc modelName
    PDMDrawing.Save (tempDir)
    PDMPart.Save (tempDir)

    'open the part document that was saved in the temp directory'
    Set swPart = swApp.OpenDoc6(tempDir + modelName, _
        swDocPART, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)

    'model is now open. anything that wants to be modified with the model'
    'should happen here. Following lines set up some objects'
    Set swExtension = swPart.Extension
    Set swCustPropMgr = swExtension.CustomPropertyManager("")

    'in case any of the custom properties we want to modify for the model don't'
    'already exist, we add them to the custom properties here'
    bool = swCustPropMgr.Add2("Finish", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("Description of Change", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("Date of Change", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("DrawnBy", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("DrawnDate", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("Material", swCustomInfoType_e.swCustomInfoText, " ")

    'modification of the custom proerties. if any of these don't want to be'
    'changed, we can comment them out.'
    bool = swCustPropMgr.Set("Finish", CP_Finish)
    bool = swCustPropMgr.Set("Description of Change", CP_Change)
    bool = swCustPropMgr.Set("Date of Change", CP_ChangeDate)
    bool = swCustPropMgr.Set("DrawnBy", CP_DrawnBy)
    bool = swCustPropMgr.Set("DrawnDate", CP_DrawnDate)
    bool = swCustPropMgr.Set("Material", CP_Material)

    'save the part now that the modifications are complete'
    bool = swPart.Save3(1, errors, warnings)

    'open the drawing. any modifications to the drawing should happen now'
    Set swDrawing = swApp.OpenDoc6(tempDir + drawingName, _
        swDocDRAWING, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)

    'call the changedrawingsheet function, which performs the actual work of
    'changing the drawing sheet for the drawing object it gets passed
    changeDrawingSheet swDrawing

    'work on the drawing is finished. save it in the temp directory'
    bool = swDrawing.Save3(17, errors, warnings)

    'only quit, checkin and issue vendor files if not in test mode'
    If testMode = False Then
        'close both the drawing and the part. they must be closed for the check in
        'function to work correctly
        swApp.QuitDoc swDrawing.GetTitle
        swApp.QuitDoc swPart.GetTitle

        'check in both the drawing and the part document to pdm. these calls are
        'currently set to up the revision level, but they don't have to
        Set checkInDocument = PDMConnection.CheckIn( _
            tempDir + drawingName, _
            PDMDrawing.project, _
            PDMDrawing.Number, _
            PDMDrawing.Description, _
            "", _
            Default, _
            "", _
            PDMDrawing.GetStatus, _
            False, _
            "")
        Set checkInDocument = PDMConnection.CheckIn( _
            tempDir + modelName, _
            PDMPart.project, _
            PDMPart.Number, _
            PDMPart.Description, _
            "", _
            Default, _
            "", _
            PDMPart.GetStatus, _
            False, _
            "")

        'save the vendor files for the model and drawing by calling the vendor
        'files function. This gets passed only the number and pdmconnection, and
        'checks the drawing and model out fresh to ensure accuracy
        errors = saveVendorFiles(modelnumber(j), PDMConnection)
    End If

    'the work has succeded at this point. should write to a file here or delete
    'the line in the existing input file
    Debug.Print modelnumber(j) + " FINISHED"
    Print #2, modelnumber(j) + ", FINISHED"

'loop back to the next model number that was read from the input file the
'GOTO to eject from the loop points here.
nextLoop: Next j

'cleanup by logging out of pdm. the vendor files script saves over the
'files left in temp and then deletes them, but this is kind of a shoddy way
'to clean up the files in each loop...'
PDMConnection.Logout
Close #2

End Sub
'------------------------------------------------------------------------------'
Sub changeDrawingSheet(swDrawing As SldWorks.DrawingDoc)

'local object declarations'
Dim swExtension     As SldWorks.ModelDocExtension
Dim swModel         As SldWorks.ModelDoc2
Dim swSheet         As SldWorks.Sheet
Dim swView          As SldWorks.View
Dim swNote          As SldWorks.Note
Dim regEx           As New RegExp

'local variable declarations'
Dim vSheetName      As Variant
Dim noteName        As String
Dim i               As Integer
Dim bool            As Boolean
Dim xDim            As Variant

'constant declarations, including sheet locations for cut and default templates'
Const cutTemplate      As String = _
    "X:\Engineering\Engineering Resources\SolidWorks Templates" + _
    "\Current Templates\DRAWING (IMPERIAL) CUT.slddrt"
Const defaultTemplate  As String = _
    "X:\Engineering\Engineering Resources\SolidWorks Templates" + _
    "\Current Templates\DRAWING (IMPERIAL).slddrt"
Const xOffset       As Double = -0.05

'initialize the modelextension'
Set swModel = swDrawing
Set swExtension = swModel.Extension

'set up the regex function with the default values'
With regEx
    .Global = True
    .Multiline = True
    .IgnoreCase = True
End With

'change all drawing sheet names to UPPERCASE'
vSheetName = swDrawing.GetSheetNames
For i = LBound(vSheetName) To UBound(vSheetName)
    swDrawing.Sheet(vSheetName(i)).SetName (UCase(vSheetName(i)))
Next i

'attempt to delete the inspection sheet
swModel.ClearSelection2 (True)
bool = swExtension.SelectByID2("INSPECTION", _
    "SHEET", _
    0, _
    0, _
    0, _
    False, _
    0, _
    Nothing, _
    0)
bool = swExtension.DeleteSelection2(0)

'reassign the sheet name array, since there may have been changes'
vSheetName = swDrawing.GetSheetNames

'main sheet inspection loop. will loop through each sheet in the drawing'
For i = LBound(vSheetName) To UBound(vSheetName)
    'activate the sheet for the current loop, and set the first view'
    bool = swDrawing.ActivateSheet(vSheetName(i))
    Set swSheet = swDrawing.Sheet(vSheetName(i))
    Set swView = swDrawing.GetFirstView
    'loop through views for as long as they exist'
    While Not swView Is Nothing
        'set the first note, and loop through notes for as long as they exist'
        Set swNote = swView.GetFirstNote
        While Not swNote Is Nothing
            'if the note matches the regex, then the cut sheet is unneeded'
            'delete the cut sheet, and set the sheet name for this loop to
            'DELETED
            regEx.Pattern = "THIS PART DOES NOT USE A CUT FILE"
            If regEx.test(swNote.GetText) Then
                Set swNote = swNote.GetNext
                swModel.ClearSelection2 (True)
                bool = swExtension.SelectByID2("CUT", _
                    "SHEET", _
                    0, _
                    0, _
                    0, _
                    False, _
                    0, _
                    Nothing, _
                    0)
                bool = swExtension.DeleteSelection2(0)
                vSheetName(i) = "DELETED"
            'if the note matches this regex, the cut sheet is needed, but since
            'the new sheet is set up with the note in the template, the cut
            'sheet note is not needed. Delete the note'
            Else
                regEx.Pattern = "dxf for cut file|" + _
                    "this sheet intentionally left blank"
                If regEx.test(swNote.GetText) Then
                    noteName = swNote.GetName + "@" + swView.GetName2
                    Set swNote = swNote.GetNext
                    swModel.ClearSelection2 (True)
                    bool = swExtension.SelectByID2(noteName, _
                        "NOTE", _
                        0, _
                        0, _
                        0, _
                        False, _
                        0, _
                        Nothing, _
                        0)
                    swModel.EditDelete
                Else
                'inspect the next note'
                    Set swNote = swNote.GetNext
                End If
            End If
        Wend
        'inspect the next view'
        Set swView = swView.GetNextView
    Wend

    'the sheets have been set up to have their templates changed now first,
    'we do the switch for sheets named "CUT"
    regEx.Pattern = "CUT"
    If regEx.test(vSheetName(i)) Then
        'clear the sheet template'
        bool = swDrawing.SetupSheet5(vSheetName(i), _
            0, _
            13, _
            swSheet.GetProperties(2), _
            swSheet.GetProperties(3), _
            False, _
            None, _
            0#, _
            0#, _
            "Default", _
            True)
        'apply the correct sheet template'
        bool = swDrawing.SetupSheet5(vSheetName(i), _
            0, _
            12, _
            swSheet.GetProperties(2), _
            swSheet.GetProperties(3), _
            False, _
            cutTemplate, _
            0#, _
            0#, _
            "Default", _
            True)
    Else
        'the sheet isn't named cut, so it gets the detault template, so long
        'as it wasn't deleted earlier in the loop
        If vSheetName(i) <> "DELETED" Then
            'clear the sheet template'
            bool = swDrawing.SetupSheet5(vSheetName(i), _
                0, _
                13, _
                swSheet.GetProperties(2), _
                swSheet.GetProperties(3), _
                False, _
                None, _
                0#, _
                0#, _
                "Default", _
                True)
            'apply the correct sheet template'
            bool = swDrawing.SetupSheet5(vSheetName(i), _
                0, _
                12, _
                swSheet.GetProperties(2), _
                swSheet.GetProperties(3), _
                False, _
                defaultTemplate, _
                0#, _
                0#, _
                "Default", _
                True)
        End If
    End If
'loop back to the next sheet in the drawing'
Next i

'return a list of drawing sheet names, with the cut sheet brought to front
'use this to reorder the drawing sheets
bool = swDrawing.ReorderSheets(bringToFront(swDrawing.GetSheetNames, "CUT"))

'if the drawing has a cuts sheet, loop through the views looking for model views
'if one is found, ensure it is pushed off to the left of the sheet
If swDrawing.ActivateSheet("CUT") Then
    Set swView = swDrawing.GetFirstView
    While Not swView Is Nothing
        If swView.Type = 7 Then
            xDim = swView.Position
            xDim(0) = swView.Position(0) - swView.GetOutline(2) + xOffset
            swView.Position = xDim
        End If
        Set swView = swView.GetNextView
    Wend
End If

End Sub
'------------------------------------------------------------------------------'
Private Function readData(filepath As String) As String()

'dimension the local variables'
Dim k As Integer
Dim records() As String

'open the passed file for input'
Open filepath For Input As #1

'loop through every line on the file'
Do Until EOF(1)
    'change the length of records to match the number of files read'
    ReDim Preserve records(k)
    'assign the line to records array'
    Line Input #1, records(k)
    k = k + 1
Loop

'close the data file and assign the output array of part numbers'
Close #1
readData = records()

End Function
'------------------------------------------------------------------------------'
Private Function bringToFront(inputArray As Variant, _
    stringToFind As String) As Variant

'dimension local variables'
Dim i               As Integer
Dim j               As Integer
Dim first           As Integer
Dim last            As Integer
Dim outputArray()   As String

'declare the extent of the input array'
first = LBound(inputArray)
last = UBound(inputArray)

'redim the output array to match the input array size'
ReDim outputArray(first To last)

'loop through every entry on the input array'
For i = first To last
    'if a match is found, move that entry to the front'
    If inputArray(i) = stringToFind Then
        For j = first To (i - 1)
            outputArray(j + 1) = inputArray(j)
        Next j
        For j = (i + 1) To last
            outputArray(j) = inputArray(j)
        Next j
        outputArray(first) = stringToFind
    End If
Next i

'set the function return to the output array'
bringToFront = outputArray

End Function
