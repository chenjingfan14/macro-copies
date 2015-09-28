Option Explicit

Dim swApp           As SldWorks.SldWorks

'------------------------------------------------------------------------------'
Sub modifyAndCheckin()

Dim swDrawing       As SldWorks.DrawingDoc
Dim swPart          As SldWorks.ModelDoc2
Dim swExtension     As SldWorks.ModelDocExtension
Dim swModel         As SldWorks.ModelDoc2
Dim PDMConnection   As IPDMWConnection
Dim PDMPart         As PDMWDocument
Dim PDMDrawing      As PDMWDocument
Dim checkInDocument As PDMWDocument
Dim swCustPropMgr   As SldWorks.CustomPropertyManager

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

Const inputFile     As String = "C:\Users\jpettit\Desktop\SCRIPTS\filesToChange.txt"
Const tempDir       As String = "X:\Engineering\TEMP\"
Const pdmName       As String = "jpettit"
Const pdmLogin      As String = "CDGshoxs!"
Const pdmServer     As String = "SHOXS1"

'Custom property values to be written to each file'
CP_Finish = "002"
CP_Change = "CHANGED FINISH SPECIFICATION"
CP_ChangeDate = UCase(Format(Now, "d-MMM-yy"))
CP_DrawnBy = "JP"
CP_DrawnDate = Format(Now, "mm/d/yy")
CP_Material = "6061-T6 ALLOY"

Set PDMConnection = CreateObject("PDMWorks.PDMWConnection")
Set swApp = Application.SldWorks

'function call which returns array of part numbers to change'
modelnumber() = readData(inputFile)
Debug.Print UBound(modelnumber()) + 1 & " PARTS TO CHANGE"

'initialize the pdmworks connection
PDMConnection.Login pdmName, pdmLogin, pdmServer

For j = LBound(modelnumber) To UBound(modelnumber)

    drawingName = modelnumber(j) + ".SLDDRW"
    modelName = modelnumber(j) + ".SLDPRT"

    Set PDMPart = PDMConnection.GetSpecificDocument(modelName)
    Set PDMDrawing = PDMConnection.GetSpecificDocument(drawingName)

    If PDMPart.Owner <> pdmName Then
        PDMPart.TakeOwnership
    End If

    If PDMDrawing.Owner <> pdmName Then
        PDMDrawing.TakeOwnership
    End If

    PDMDrawing.Save (tempDir)
    PDMPart.Save (tempDir)

    Set swPart = swApp.OpenDoc6(tempDir + modelName, _
        swDocPART, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)

    'do stuff with model here
    Set swExtension = swPart.Extension
    Set swCustPropMgr = swExtension.CustomPropertyManager("")

    bool = swCustPropMgr.Add2("Finish", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("Description of Change", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("Date of Change", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("DrawnBy", swCustomInfoType_e.swCustomInfoText, " ")
    bool = swCustPropMgr.Add2("DrawnDate", swCustomInfoType_e.swCustomInfoText, " ")

    bool = swCustPropMgr.Set("Finish", CP_Finish)
    bool = swCustPropMgr.Set("Description of Change", CP_Change)
    bool = swCustPropMgr.Set("Date of Change", CP_ChangeDate)
    bool = swCustPropMgr.Set("DrawnBy", CP_DrawnBy)
    bool = swCustPropMgr.Set("DrawnDate", CP_DrawnDate)
    bool = swCustPropMgr.Set("Material", CP_Material)

    bool = swPart.Save3(1, errors, warnings)

    Set swDrawing = swApp.OpenDoc6(tempDir + drawingName, _
        swDocDRAWING, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)

    changeDrawingSheet swDrawing

    bool = swDrawing.Save3(17, errors, warnings)

    swApp.QuitDoc swDrawing.GetTitle
    swApp.QuitDoc swPart.GetTitle

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

    errors = saveVendorFiles(modelnumber(j),PDMConnection)

    Debug.Print modelnumber(j) + " FINISHED"

Next j

PDMConnection.Logout

End Sub
'------------------------------------------------------------------------------'
Sub changeDrawingSheet(swDrawing As SldWorks.DrawingDoc)

Dim swExtension     As SldWorks.ModelDocExtension
Dim swModel         As SldWorks.ModelDoc2
Dim swSheet         As SldWorks.Sheet
Dim swView          As SldWorks.View
Dim swNote          As SldWorks.Note

Dim regEx           As New RegExp

Dim vSheetName      As Variant
Dim noteName        As String
Dim i               As Integer
Dim bool            As Boolean

Const cutTemplate      As String = _
    "X:\Engineering\Engineering Resources\SolidWorks Templates" + _
    "\Current Templates\DRAWING (IMPERIAL) CUT.slddrt"
Const defaultTemplate  As String = _
    "X:\Engineering\Engineering Resources\SolidWorks Templates" + _
    "\Current Templates\DRAWING (IMPERIAL).slddrt"

Set swModel = swDrawing
Set swExtension = swModel.Extension

With regEx
    .Global = True
    .Multiline = True
    .IgnoreCase = True
End With

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

vSheetName = swDrawing.GetSheetNames

for i = LBound(vSheetName) To UBound(vSheetName)
    swDrawing.Sheet(vSheetName(i)).SetName(UCase(vSheetName(i)))
Next i

vSheetName = swDrawing.GetSheetNames

For i = LBound(vSheetName) To UBound(vSheetName)
    bool = swDrawing.ActivateSheet(vSheetName(i))
    Set swView = swDrawing.GetFirstView
    While Not swView Is Nothing
        Set swNote = swView.GetFirstNote
        While Not swNote Is Nothing
            regEx.Pattern = "THIS PART DOES NOT USE A CUT FILE"
            If regEx.Test(swNote.GetText) Then
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
            Else
                regEx.Pattern = "dxf for cut file|" + _
                    "this sheet intentionally left blank"
                If regEx.Test(swNote.GetText) Then
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
                    Set swNote = swNote.GetNext
                End If
            End If
        Wend
        Set swView = swView.GetNextView
    Wend

    regEx.Pattern = "CUT"
    swDrawing.ActivateSheet (vSheetName(i))
    Set swSheet = swDrawing.Sheet(vSheetName(i))

    If regEx.Test(vSheetName(i)) Then
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
        If vSheetName(i) <> "DELETED" Then
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
Next i

bool = swDrawing.ReorderSheets(bringToFront(swDrawing.GetSheetNames, "CUT"))

End Sub
'------------------------------------------------------------------------------'
Private Function readData(filepath As String) As String()

Open filepath For Input As #1

Dim k As Integer
Dim records() As String

Do Until EOF(1)
    ReDim Preserve records(k)
    Line Input #1, records(k)
    k = k + 1
Loop

Close #1
readData = records()
End Function
'------------------------------------------------------------------------------'
Private Function bringToFront(inputArray As Variant, _
    stringToFind As String) As Variant

Dim i               As Integer
Dim j               As Integer
Dim first           As Integer
Dim last            As Integer
Dim outputArray()   As String

first = LBound(inputArray)
last = UBound(inputArray)

ReDim outputArray(first To last)

For i = first To last
    If inputArray(i) = stringToFind Then
        For j = first To (i - 1)
            outputArray(j + 1) = inputArray(j)
        Next j
        outputArray(first) = stringToFind
    End If
Next i

bringToFront = outputArray

End Function
