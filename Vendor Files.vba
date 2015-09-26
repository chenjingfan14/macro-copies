Option Explicit
Dim swApp As SldWorks.SldWorks

Sub main()

'object declarations
Dim swModel         As SldWorks.ModelDoc2

'data declarations
dim partToSave      as string
Dim errors          As Long
Dim warnings        As Long
Dim boolstatus      As Boolean

'set the objects required
Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

'determine the names of the files in pdmworks by examining the active document name

partToSave = Left(swModel.GetTitle, 6)
swApp.QuitDoc swModel.GetTitle
errors = saveVendorFiles(partToSave)

End Sub

Public function saveVendorFiles(partNumber as string) as Long

'object declarations
Dim fso             As Object
Dim PDMConnection   As IPDMWConnection
Dim document        As PDMWDocument
Dim swModel         As SldWorks.ModelDoc2
Dim swDrawing       As SldWorks.DrawingDoc
Dim swModelDocExt   As SldWorks.ModelDocExtension
Dim swExportPDFData As SldWorks.ExportPdfData
Dim swCustPropMgr   As SldWorks.CustomPropertyManager

'data declarations
Dim drawingName     As String
Dim modelName       As String
Dim Revision        As String
Dim errors          As Long
Dim warnings        As Long
Dim boolstatus      As Boolean

'vendor files and temp locations on x drive
Const vendorDir     As String = "X:\Engineering\Vendor Files\"
Const tempDir       As String = "X:\Engineering\TEMP\"
Const pdmName       As String = "jpettit"
Const pdmLogin      As String = "CDGshoxs!"
Const pdmServer     As String = "SHOXS1"

Set fso = CreateObject("scripting.filesystemobject")
Set PDMConnection = CreateObject("PDMWorks.PDMWConnection")
Set swApp = Application.SldWorks

drawingName = partNumber + ".SLDDRW"
modelName = partNumber + ".SLDPRT"

'initialize the pdmworks connection
PDMConnection.Login pdmName, pdmLogin, pdmServer

'Save the drawing in temp and determine the correct revision
PDMConnection.GetSpecificDocument(drawingName).Save (tempDir)
PDMConnection.GetSpecificDocument(modelName).Save (tempDir)
Revision = PDMConnection.GetSpecificDocument(modelName).Revision

'save the active document (part) as an IGES file
Set swModel = swApp.OpenDoc6(tempDir + modelName, _
    swDocPART, _
    swOpenDocOptions_Silent, _
    "", _
    errors, _
    warnings)
Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")
swModel.SaveAs (vendorDir + partNumber + " " + Revision + ".igs")
swApp.QuitDoc swModel.GetTitle

'open the drawing and save as PDF
Set swDrawing = swApp.OpenDoc6(tempDir + drawingName, _
    swDocDRAWING, _
    swOpenDocOptions_Silent, _
    "", _
    errors, _
    warnings)
Set swModelDocExt = swDrawing.Extension
Set swExportPDFData = swApp.GetExportFileData(1)

boolstatus = swExportPDFData.SetSheets(1, Nothing)
boolstatus = swModelDocExt.SaveAs(vendorDir + partNumber + " " + Revision + ".pdf", _
    0, _
    0, _
    swExportPDFData, _
    errors, _
    warnings)

'if any drawing sheets are named CUT, switch to that sheet and save as a dxf
If swDrawing.ActivateSheet("CUT") Then
    boolstatus = swModelDocExt.SaveAs(vendorDir + partNumber + " " + Revision + ".dxf", _
        0, _
        0, _
        Nothing, _
        errors, _
        warnings)
End If

'cleanup - close the drawing, and delete if from temp. close the pdm conneciton
swApp.QuitDoc swDrawing.GetTitle

fso.DeleteFile (tempDir + drawingName)
PDMConnection.Logout

End Function
