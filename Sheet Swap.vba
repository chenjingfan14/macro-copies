Option Explicit

Dim swApp           As SldWorks.SldWorks

'------------------------------------------------------------------------------'
Sub modifyAndCheckin()

Dim myDrawing       As SldWorks.DrawingDoc
Dim myPart          As SldWorks.ModelDoc2
Dim myExtension     As SldWorks.ModelDocExtension
Dim myModel         As SldWorks.ModelDoc2
Dim PDMConnection   As IPDMWConnection
Dim myPDMPart       As PDMWDocument
Dim myPDMDrawing    As PDMWDocument
Dim checkInDocument As PDMWDocument
Dim myCustPropMgr   As SldWorks.CustomPropertyManager

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

    Set myPDMPart = PDMConnection.GetSpecificDocument(modelName)
    Set myPDMDrawing = PDMConnection.GetSpecificDocument(drawingName)

    If myPDMPart.Owner <> pdmName Then
        myPDMPart.TakeOwnership
    End If

    If myPDMDrawing.Owner <> pdmName Then
        myPDMDrawing.TakeOwnership
    End If

    myPDMDrawing.Save (tempDir)
    myPDMPart.Save (tempDir)

    Set myPart = swApp.OpenDoc6(tempDir + modelName, _
        swDocPART, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)

    'do stuff with model here
    Set myExtension = myPart.Extension
    Set myCustPropMgr = myExtension.CustomPropertyManager("")

    bool = myCustPropMgr.Add2("Finish", swCustomInfoType_e.swCustomInfoText, " ")
    bool = myCustPropMgr.Add2("Description of Change", swCustomInfoType_e.swCustomInfoText, " ")
    bool = myCustPropMgr.Add2("Date of Change", swCustomInfoType_e.swCustomInfoText, " ")
    bool = myCustPropMgr.Add2("DrawnBy", swCustomInfoType_e.swCustomInfoText, " ")
    bool = myCustPropMgr.Add2("DrawnDate", swCustomInfoType_e.swCustomInfoText, " ")

    bool = myCustPropMgr.Set("Finish", CP_Finish)
    bool = myCustPropMgr.Set("Description of Change", CP_Change)
    bool = myCustPropMgr.Set("Date of Change", CP_ChangeDate)
    bool = myCustPropMgr.Set("DrawnBy", CP_DrawnBy)
    bool = myCustPropMgr.Set("DrawnDate", CP_DrawnDate)
    bool = myCustPropMgr.Set("Material", CP_Material)

    bool = myPart.Save3(1, errors, warnings)

    Set myDrawing = swApp.OpenDoc6(tempDir + drawingName, _
        swDocDRAWING, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)

    'pass an active drawing

    changeDrawingSheet(myDrawing)

    bool = myDrawing.Save3(17, errors, warnings)

    swApp.QuitDoc myDrawing.GetTitle
    swApp.QuitDoc myPart.GetTitle

    Set checkInDocument = PDMConnection.CheckIn( _
        tempDir + drawingName, _
        myPDMDrawing.project, _
        myPDMDrawing.Number, _
        myPDMDrawing.Description, _
        "", _
        Default, _
        "", _
        myPDMDrawing.GetStatus, _
        False, _
        "")
    Set checkInDocument = PDMConnection.CheckIn( _
        tempDir + modelName, _
        myPDMPart.project, _
        myPDMPart.Number, _
        myPDMPart.Description, _
        "", _
        Default, _
        "", _
        myPDMPart.GetStatus, _
        False, _
        "")

    errors = saveVendorFiles(modelnumber(j),PDMConnection)

    Debug.Print modelnumber(j) + " FINISHED"

Next j

PDMConnection.Logout

End Sub
'------------------------------------------------------------------------------'
Sub changeDrawingSheet(myDrawing as SldWorks.DrawingDoc)

Dim myExtension     As SldWorks.ModelDocExtension
Dim myModel         As SldWorks.ModelDoc2
Dim mySheet         As SldWorks.Sheet
Dim myView          As SldWorks.View
Dim myNote          As SldWorks.Note

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

set myModel = myDrawing
Set myExtension = myModel.Extension

With regEx
    .Global = True
    .Multiline = True
    .IgnoreCase = True
End With

vSheetName = myDrawing.GetSheetNames

For i = 0 To UBound(vSheetName)
    bool = myDrawing.ActivateSheet(vSheetName(i))
    Set myView = myDrawing.GetFirstView
    While Not myView Is Nothing
        Set myNote = myView.GetFirstNote
        While Not myNote Is Nothing
            regEx.Pattern = "THIS PART DOES NOT USE A CUT FILE"
            If regEx.Test(myNote.GetText) Then
                Set myNote = myNote.GetNext
                myModel.ClearSelection2 (True)
                bool = myExtension.SelectByID2("CUT", _
                    "SHEET", _
                    0, _
                    0, _
                    0, _
                    False, _
                    0, _
                    Nothing, _
                    0)
                bool = myExtension.DeleteSelection2(0)
                vSheetName(i) = "DELETED"
            Else
                regEx.Pattern = "dxf for cut file|" + _
                    "this sheet intentionally left blank"
                If regEx.Test(myNote.GetText) Then
                    noteName = myNote.GetName + "@" + myView.GetName2
                    Set myNote = myNote.GetNext
                    myModel.ClearSelection2 (True)
                    bool = myExtension.SelectByID2(noteName, _
                        "NOTE", _
                        0, _
                        0, _
                        0, _
                        False, _
                        0, _
                        Nothing, _
                        0)
                    myModel.EditDelete
                Else
                    Set myNote = myNote.GetNext
                End If
            End If
        Wend
        Set myView = myView.GetNextView
    Wend

    regEx.Pattern = "cut"
    myDrawing.ActivateSheet (vSheetName(i))
    Set mySheet = myDrawing.Sheet(vSheetName(i))

    If regEx.Test(vSheetName(i)) Then
        If vSheetName(i) <> "DELETED" Then
            bool = myDrawing.SetupSheet5(vSheetName(i), _
                0, _
                13, _
                mySheet.GetProperties(2), _
                mySheet.GetProperties(3), _
                False, _
                None, _
                0#, _
                0#, _
                "Default", _
                True)
            bool = myDrawing.SetupSheet5(vSheetName(i), _
                0, _
                12, _
                mySheet.GetProperties(2), _
                mySheet.GetProperties(3), _
                False, _
                cutTemplate, _
                0#, _
                0#, _
                "Default", _
                True)
        End If
    Else
        If vSheetName(i) <> "DELETED" Then
            bool = myDrawing.SetupSheet5(vSheetName(i), _
                0, _
                13, _
                mySheet.GetProperties(2), _
                mySheet.GetProperties(3), _
                False, _
                None, _
                0#, _
                0#, _
                "Default", _
                True)
            bool = myDrawing.SetupSheet5(vSheetName(i), _
                0, _
                12, _
                mySheet.GetProperties(2), _
                mySheet.GetProperties(3), _
                False, _
                defaultTemplate, _
                0#, _
                0#, _
                "Default", _
                True)
        End If
    End If
Next i
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
