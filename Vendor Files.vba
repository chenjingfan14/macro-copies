Option Explicit
Dim swApp As SldWorks.SldWorks

Sub main()

'object declarations
Dim fso             As Object
Dim PDMConnection   As IPDMWConnection
Dim documents       As IPDMWDocuments
Dim document        As PDMWDocument
Dim regEx           As New VBScript_RegExp_55.RegExp
Dim modelextension  As SldWorks.ModelDocExtension
Dim swModel         As SldWorks.ModelDoc2
Dim swDrawing       As SldWorks.DrawingDoc
Dim swModelDocExt   As SldWorks.ModelDocExtension
Dim swExportPDFData As SldWorks.ExportPdfData
Dim swCustPropMgr   As SldWorks.CustomPropertyManager

'data declarations
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

'vendor files and temp locations on x drive
vendorDir = "X:\Engineering\Vendor Files"
tempDir = "X:\Engineering\TEMP"
revMessage = "The revision of this file is not the latest revision checked into PDM. Continue saving Vendor Files?"

'set the objects required
Set fso = CreateObject("scripting.filesystemobject")
Set PDMConnection = CreateObject("PDMWorks.PDMWConnection")
Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc
Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")

'determine the names of the files in pdmworks by examining the active document name
drawingName = Left(swModel.GetTitle, 6) + ".SLDDRW"
modelName = Left(swModel.GetTitle, 6) + ".SLDPRT"

'initialize the pdmworks connection
PDMConnection.Login "jpettit", "CDGshoxs!", "SHOXS1"

'Save the drawing in vendor files and determine the correct revision
PDMConnection.GetSpecificDocument(drawingName).Save (tempDir)
boolstatus = swCustPropMgr.Get3("Revision", False, Revision, evalRev)
PDMRevision = PDMConnection.GetSpecificDocument(modelName).Revision

'If the revision in custom properties isn't equal to the latest revision in PDM, prompt for continue
If PDMRevision <> Revision Then
    If swApp.SendMsgToUser2(revMessage, 1, 5) <> 6 Then
        Exit Sub
    End If
End If

'save the active document (part) as an IGES file
swModel.SaveAs (vendorDir + "\" + Left(swModel.GetTitle, 6) + " " + Revision + ".IGS")

'open the drawing and save as PDF
Set swDrawing = swApp.OpenDoc6(tempDir + "\" + drawingName, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
Set swModelDocExt = swDrawing.Extension
Set swExportPDFData = swApp.GetExportFileData(1)

boolstatus = swExportPDFData.SetSheets(1, Nothing)
boolstatus = swModelDocExt.SaveAs(vendorDir + "\" + Left(swModel.GetTitle, 6) + " " + Revision + ".PDF", 0, 0, swExportPDFData, errors, warnings)

'if any drawing sheets are named CUT, switch to that sheet and save as a dxf
If swDrawing.ActivateSheet("CUT") Then
   boolstatus = swModelDocExt.SaveAs(vendorDir + "\" + Left(swModel.GetTitle, 6) + " " + Revision + ".DXF", 0, 0, Nothing, errors, warnings)
End If

'cleanup - close the drawing, and delete if from vendor files. close the pdm conneciton
swApp.QuitDoc swDrawing.GetTitle
fso.DeleteFile (tempDir + "\" + Left(swModel.GetTitle, 6) + ".SLDDRW")
PDMConnection.Logout

End Sub
