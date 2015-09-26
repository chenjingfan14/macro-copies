Option Explicit

Dim swApp           As SldWorks.SldWorks
Dim myModel         As SldWorks.ModelDoc2
Dim myExtension     As SldWorks.ModelDocExtension
Dim myPart          As SldWorks.ModelDoc2
Dim myDrawing       As SldWorks.DrawingDoc
Dim mySheet         As SldWorks.Sheet
Dim myView          As SldWorks.View
Dim myNote          As SldWorks.Note
Dim myExportPDFData As SldWorks.ExportPdfData
Dim myCustPropMgr   As SldWorks.CustomPropertyManager
Dim fso             As Object
Dim PDMConnection   As IPDMWConnection
Dim documents       As IPDMWDocuments
Dim checkInDocument As PDMWDocument
Dim myPDMPart       As PDMWDocument
Dim myPDMDrawing    As PDMWDocument
Dim drawingName     As String
Dim modelName       As String
Dim Revision        As String
Dim PDMRevision     As String
Dim evalRev         As String
Dim revMessage      As String
Dim errors          As Long
Dim warnings        As Long
Dim boolstatus      As Boolean
Dim modelnumber()   As String
Dim j               As Integer

Sub main()'--------------------------------------------------------------------'

Const vendorDir     As String = "X:\Engineering\Vendor Files"
Const tempDir       As String = "X:\Engineering\TEMP"
Const pdmName       As String = "jpettit"
Const pdmLogin      As String = "CDGshoxs!"
Const pdmServer     As string = "SHOXS1"

'function call which returns via a global variable. should change this'
modelnumber() = readdata("C:\Users\jpettit\Desktop\SCRIPTS\filesToChange.txt")

Set fso = CreateObject("scripting.filesystemobject")
Set PDMConnection = CreateObject("PDMWorks.PDMWConnection")
Set swApp = Application.SldWorks

'initialize the pdmworks connection
PDMConnection.Login pdmName, pdmLogin, pdmName

For j = LBound(modelnumber) To UBound(modelnumber)

    drawingName = modelnumber(j) + ".SLDDRW"
    modelName = modelnumber(j) + ".SLDPRT"

    'Save the drawing in vendor files and determine the correct revision
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

    Set myPart = swApp.OpenDoc6(tempDir + "\" + modelName, _
        swDocPART, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)

    'do stuff with model here
    Set myExtension = myPart.Extension
    Set myCustPropMgr = myExtension.CustomPropertyManager("")

    boolstatus = myCustPropMgr.Add2("Finish", swCustomInfoType_e.swCustomInfoText, " ")
    boolstatus = myCustPropMgr.Add2("Description of Change", swCustomInfoType_e.swCustomInfoText, " ")
    boolstatus = myCustPropMgr.Add2("Date of Change", swCustomInfoType_e.swCustomInfoText, " ")
    boolstatus = myCustPropMgr.Add2("DrawnBy", swCustomInfoType_e.swCustomInfoText, " ")
    boolstatus = myCustPropMgr.Add2("DrawnDate", swCustomInfoType_e.swCustomInfoText, " ")

    boolstatus = myCustPropMgr.Set("Finish", "002")
    boolstatus = myCustPropMgr.Set("Description of Change", "CHANGED FINISH SPECIFICATION")
    boolstatus = myCustPropMgr.Set("Date of Change", "16-SEP-15")
    boolstatus = myCustPropMgr.Set("DrawnBy", "JP")
    boolstatus = myCustPropMgr.Set("DrawnDate", "09/16/15")
    boolstatus = myCustPropMgr.Set("Material", "6061-T6 ALLOY")

    boolstatus = myPart.Save3(1, errors, warnings)

    Set myDrawing = swApp.OpenDoc6(tempDir + "\" + drawingName, _
        swDocDRAWING, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)

    'pass an active drawing

    changeActiveDrawingSheet

    boolstatus = myDrawing.Save3(17, errors, warnings)

    swApp.QuitDoc myDrawing.GetTitle
    swApp.QuitDoc myPart.GetTitle

    Set checkInDocument = PDMConnection.CheckIn(
        tempDir + "\" + drawingName, _
        myPDMDrawing.project, _
        myPDMDrawing.Number, _
        myPDMDrawing.Description, _
        "", _
        Default, _
        "", _
        myPDMDrawing.GetStatus, _
        False, _
        "")
    Set checkInDocument = PDMConnection.CheckIn(
        tempDir + "\" + modelName, _
        myPDMPart.project, _
        myPDMPart.Number, _
        myPDMPart.Description, _
        "", _
        Default, _
        "", _
        myPDMPart.GetStatus, _
        False, _
        "")

    Set myPart = swApp.OpenDoc6(tempDir + "\" + modelName, _
        swDocPART, _
        swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)
    Set myExtension = myPart.Extension

    boolstatus = myExtension.SaveAs(vendorDir + "\" + left(myPart.GetTitle,6) + " " + checkInDocument.Revision + ".igs", _
        0, _
        0, _
        Nothing, _
        errors, _
        warnings)

    Set myDrawing = swApp.OpenDoc6(tempDir + "\" + drawingName, swDocDRAWING, swOpenDocOptions_Silent, _
        "", _
        errors, _
        warnings)
    Set myExtension = myDrawing.Extension
    Set myExportPDFData = swApp.GetExportFileData(1)

    boolstatus = myExportPDFData.SetSheets(1, Nothing)
    boolstatus = myExtension.SaveAs(vendorDir + "\" + left(myPart.GetTitle,6) + " " + checkInDocument.Revision + ".pdf", _
    0, _
    0, _
    myExportPDFData, _
    errors, _
    warnings)

    If myDrawing.ActivateSheet("CUT") Then
        boolstatus = myExtension.SaveAs(vendorDir + "\" + left(myPart.GetTitle,6) + " " + checkInDocument.Revision + ".dxf", _
        0, _
        0, _
        Nothing, _
        errors, _
        warnings)
    End If

    swApp.QuitDoc myPart.GetTitle
    swApp.QuitDoc myDrawing.GetTitle

    fso.DeleteFile (tempDir + "\" + modelnumber(j) + ".SLDDRW")
    fso.DeleteFile (tempDir + "\" + modelnumber(j) + ".SLDPRT")

    Debug.Print modelnumber(j) + " FINISHED"

Next j

PDMConnection.Logout

End Sub

Sub changeActiveDrawingSheet()'------------------------------------------------'

Dim regEx As New RegExp

Dim longstatus As Long
Dim longwarnings As Long
Dim vSheetName As Variant
Dim noteName As String
Dim i As Integer

Const cutTemplate      As String = _
    "X:\Engineering\Engineering Resources\SolidWorks Templates"_
    "\Current Templates\DRAWING (IMPERIAL) CUT.slddrt"
Const defaultTemplate  As String = _
    "X:\Engineering\Engineering Resources\SolidWorks Templates"_
    "\Current Templates\DRAWING (IMPERIAL).slddrt"

Set myModel = swApp.ActiveDoc
Set myExtension = myModel.Extension
Set myDrawing = myModel

With regEx
            .Global = True
            .Multiline = True
            .IgnoreCase = True
End With

vSheetName = myDrawing.GetSheetNames

For i = 0 To UBound(vSheetName)
    boolstatus = myDrawing.ActivateSheet(vSheetName(i))
    Set myView = myDrawing.GetFirstView
    While Not myView Is Nothing
        Set myNote = myView.GetFirstNote
        While Not myNote Is Nothing
            regEx.Pattern = "THIS PART DOES NOT USE A CUT FILE"
            If regEx.Test(myNote.GetText) Then
                Set myNote = myNote.GetNext
                myModel.ClearSelection2 (True)
                boolstatus = myExtension.SelectByID2("CUT", _
                    "SHEET", _
                    0, _
                    0, _
                    0, _
                    False, _
                    0, _
                    Nothing, _
                    0)
                boolstatus = myExtension.DeleteSelection2(0)
                vSheetName(i) = "DELETED"
            Else
                regEx.Pattern = "dxf for cut file|" _
                    "this sheet intentionally left blank"
                If regEx.Test(myNote.GetText) Then
                    noteName = myNote.GetName + "@" + myView.GetName2
                    Set myNote = myNote.GetNext
                    myModel.ClearSelection2 (True)
                    boolstatus = myExtension.SelectByID2(noteName, _
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
            boolstatus = myDrawing.SetupSheet5(vSheetName(i), _
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
            boolstatus = myDrawing.SetupSheet5(vSheetName(i), _
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
            boolstatus = myDrawing.SetupSheet5(vSheetName(i), _
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
            boolstatus = myDrawing.SetupSheet5(vSheetName(i), _
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

Function readdata(filepath as String) as string()'-----------------------------'

Open filepath For Input As #1

'declare the local loop variable'
Dim k As Integer = 0
Dim records() As String

Do Until EOF(1)
    ReDim Preserve records(k)
    Line Input #1, records(k)
    k = k + 1
Loop
Close #1

Debug.Print UBound(records()) + 1 & " PARTS TO CHANGE"

readdata() = records()

End Function
