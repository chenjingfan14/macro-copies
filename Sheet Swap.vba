Dim swApp           As SldWorks.SldWorks
Dim myModel         As SldWorks.ModelDoc2
Dim myPart          As SldWorks.ModelDoc2
Dim myDrawing       As SldWorks.DrawingDoc
Dim myView          As SldWorks.View
Dim myNote          As SldWorks.Note
Dim myExtension     As SldWorks.ModelDocExtension
Dim myExportPDFData As SldWorks.ExportPdfData
Dim myCustPropMgr   As SldWorks.CustomPropertyManager
Dim mySheet         As SldWorks.Sheet

Dim fso             As Object

Dim PDMConnection   As IPDMWConnection
Dim documents       As IPDMWDocuments
Dim checkInDocument As PDMWDocument
Dim myPDMPart       As PDMWDocument
Dim myPDMDrawing    As PDMWDocument

Dim drawingName     As String
Dim modelName       As String
Dim vendorDir       As String
Dim tempDir         As String
Dim Revision        As String
Dim PDMRevision     As String
Dim evalRev         As String
Dim revMessage      As String
Dim errors          As Long
Dim warnings        As Long
Dim boolstatus      As Boolean

Dim modelnumber()   As String
Dim j               As Integer


Sub main()

readdata

'vendor files and temp locations on x drive
vendorDir = "X:\Engineering\Vendor Files"
tempDir = "X:\Engineering\TEMP"
pdmdir = "C:\Users\jpettit\Documents\PDM Documents"

Set fso = CreateObject("scripting.filesystemobject")
Set PDMConnection = CreateObject("PDMWorks.PDMWConnection")
Set swApp = Application.SldWorks

'initialize the pdmworks connection
PDMConnection.Login "jpettit", "CDGshoxs!", "SHOXS1"

For j = LBound(modelnumber) To UBound(modelnumber)

    drawingName = modelnumber(j) + ".SLDDRW"
    modelName = modelnumber(j) + ".SLDPRT"

    'Save the drawing in vendor files and determine the correct revision
    Set myPDMPart = PDMConnection.GetSpecificDocument(modelName)
    Set myPDMDrawing = PDMConnection.GetSpecificDocument(drawingName)


    If myPDMPart.Owner <> "jpettit" Then
        myPDMPart.TakeOwnership
    End If
    If myPDMDrawing.Owner <> "jpettit" Then
        myPDMDrawing.TakeOwnership
    End If

    myPDMDrawing.Save (tempDir)
    myPDMPart.Save (tempDir)

    Set myPart = swApp.OpenDoc6(tempDir + "\" + modelName, swDocPART, swOpenDocOptions_Silent, "", errors, warnings)
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

    Set myDrawing = swApp.OpenDoc6(tempDir + "\" + drawingName, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
    'pass an active drawing

    changeActiveDrawingSheet

    boolstatus = myDrawing.Save3(17, errors, warnings)

    swApp.QuitDoc myDrawing.GetTitle
    swApp.QuitDoc myPart.GetTitle

    Set checkInDocument = PDMConnection.CheckIn(tempDir + "\" + drawingName, myPDMDrawing.project, myPDMDrawing.Number, myPDMDrawing.Description, "", Default, "", myPDMDrawing.GetStatus, False, "")
    Set checkInDocument = PDMConnection.CheckIn(tempDir + "\" + modelName, myPDMPart.project, myPDMPart.Number, myPDMPart.Description, "", Default, "", myPDMPart.GetStatus, False, "")

    Set myPart = swApp.OpenDoc6(tempDir + "\" + modelName, swDocPART, swOpenDocOptions_Silent, "", errors, warnings)

    Set myExtension = myPart.Extension

    boolstatus = myExtension.SaveAs(vendorDir + "\" + myPart.GetTitle + " " + checkInDocument.Revision + ".igs", 0, 0, Nothing, errors, warnings)


    Set myDrawing = swApp.OpenDoc6(tempDir + "\" + drawingName, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)

    Set myExtension = myDrawing.Extension
    Set myExportPDFData = swApp.GetExportFileData(1)

    boolstatus = myExportPDFData.SetSheets(1, Nothing)
    boolstatus = myExtension.SaveAs(vendorDir + "\" + myPart.GetTitle + " " + checkInDocument.Revision + ".pdf", 0, 0, myExportPDFData, errors, warnings)

    If myDrawing.ActivateSheet("CUT") Then
        boolstatus = myExtension.SaveAs(vendorDir + "\" + myPart.GetTitle + " " + checkInDocument.Revision + ".dxf", 0, 0, Nothing, errors, warnings)
    End If

    swApp.QuitDoc myPart.GetTitle
    swApp.QuitDoc myDrawing.GetTitle

    fso.DeleteFile (tempDir + "\" + modelnumber(j) + ".SLDDRW")
    fso.DeleteFile (tempDir + "\" + modelnumber(j) + ".SLDPRT")

    Debug.Print modelnumber(j) + " FINISHED"

Next j

PDMConnection.Logout

End Sub

Sub changeActiveDrawingSheet()

Dim regEx As New RegExp

Dim longstatus As Long
Dim longwarnings As Long
Dim vSheetName As Variant
Dim noteName As String
Dim i As Integer

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
                boolstatus = myExtension.SelectByID2("CUT", "SHEET", 0, 0, 0, False, 0, Nothing, 0)
                boolstatus = myExtension.DeleteSelection2(0)
                vSheetName(i) = "DELETED"

            Else

                regEx.Pattern = "dxf for cut file|this sheet intentionally left blank"

                If regEx.Test(myNote.GetText) Then

                    noteName = myNote.GetName + "@" + myView.GetName2

                    Set myNote = myNote.GetNext

                    myModel.ClearSelection2 (True)
                    boolstatus = myExtension.SelectByID2(noteName, "NOTE", 0, 0, 0, False, 0, Nothing, 0)
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
            boolstatus = myDrawing.SetupSheet5(vSheetName(i), 0, 13, mySheet.GetProperties(2), mySheet.GetProperties(3), False, None, 0#, 0#, "Default", True)
            boolstatus = myDrawing.SetupSheet5(vSheetName(i), 0, 12, mySheet.GetProperties(2), mySheet.GetProperties(3), False, "X:\Engineering\Engineering Resources\SolidWorks Templates\Current Templates\DRAWING (IMPERIAL) CUT.slddrt", 0#, 0#, "Default", True)
        End If
    Else
        If vSheetName(i) <> "DELETED" Then
            boolstatus = myDrawing.SetupSheet5(vSheetName(i), 0, 13, mySheet.GetProperties(2), mySheet.GetProperties(3), False, None, 0#, 0#, "Default", True)
            boolstatus = myDrawing.SetupSheet5(vSheetName(i), 0, 12, mySheet.GetProperties(2), mySheet.GetProperties(3), False, "X:\Engineering\Engineering Resources\SolidWorks Templates\Current Templates\DRAWING (IMPERIAL).slddrt", 0#, 0#, "Default", True)
        End If
    End If

Next i

End Sub

Sub readdata()

Dim k As Integer

Open "C:\Users\jpettit\Desktop\SCRIPTS\filesToChange.txt" For Input As #1

k = 0

Do Until EOF(1)
    ReDim Preserve modelnumber(k)
    Line Input #1, modelnumber(k)
    'Debug.Print modelnumber(k)

    k = k + 1

Loop
Close #1

Debug.Print UBound(modelnumber()) + 1 & " PARTS TO CHANGE"

End Sub
