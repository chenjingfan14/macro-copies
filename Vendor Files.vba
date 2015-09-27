Option Explicit
Dim swApp As SldWorks.SldWorks
'------------------------------------------------------------------------------'
Sub main()

'object declarations
Dim swModel         As SldWorks.ModelDoc2
Dim PDMConnection   As IPDMWConnection
'data declarations
dim partToSave      as string
Dim errors          As Long

'set the objects required
Set PDMConnection = CreateObject("PDMWorks.PDMWConnection")
Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

Const pdmName       As String = "jpettit"
Const pdmLogin      As String = "CDGshoxs!"
Const pdmServer     As String = "SHOXS1"

PDMConnection.Login pdmName, pdmLogin, pdmServer

partToSave = Left(swModel.GetTitle, 6)
swApp.QuitDoc swModel.GetTitle
errors = saveVendorFiles(partToSave, PDMConnection)

PDMConnection.Logout

End Sub
'------------------------------------------------------------------------------'
Public function saveVendorFiles(partNumber as string, _
    passedPDMConnection as IPDMWConnection) as Long

'object declarations
Dim fso             As Object
Dim swModel         As SldWorks.ModelDoc2
Dim swDrawing       As SldWorks.DrawingDoc
Dim swModelDocExt   As SldWorks.ModelDocExtension
Dim swExportPDFData As SldWorks.ExportPdfData

'data declarations
Dim drawingName     As String
Dim modelName       As String
Dim Revision        As String
Dim saveName        As String
Dim errors          As Long
Dim warnings        As Long
Dim bool      As Boolean

'vendor files and temp locations on x drive
Const vendorDir     As String = "X:\Engineering\Vendor Files\"
Const tempDir       As String = "X:\Engineering\TEMP\"


Set fso = CreateObject("scripting.filesystemobject")
Set swApp = Application.SldWorks

drawingName = partNumber + ".SLDDRW"
modelName = partNumber + ".SLDPRT"

'Save the drawing in temp and determine the correct revision
passedPDMConnection.GetSpecificDocument(drawingName).Save (tempDir)
passedPDMConnection.GetSpecificDocument(modelName).Save (tempDir)
Revision = passedPDMConnection.GetSpecificDocument(modelName).Revision
saveName = vendorDir + partNumber + " " + Revision

'save the active document (part) as an IGES file
Set swModel = swApp.OpenDoc6(tempDir + modelName, _
    swDocPART, _
    swOpenDocOptions_Silent, _
    "", _
    errors, _
    warnings)
Set swModelDocExt = swModel.Extension
bool = swModelDocExt.SaveAs(saveName + ".igs", _
    swSaveAsCurrentVersion, _
    swSaveAsOptions_Silent, _
    Nothing, _
    errors, _
    warnings)


'open the drawing and save as PDF
Set swDrawing = swApp.OpenDoc6(tempDir + drawingName, _
    swDocDRAWING, _
    swOpenDocOptions_Silent, _
    "", _
    errors, _
    warnings)
Set swModelDocExt = swDrawing.Extension
Set swExportPDFData = swApp.GetExportFileData(1)
bool = swExportPDFData.SetSheets(swExportData_ExportAllSheets, Nothing)
bool = swModelDocExt.SaveAs(saveName + ".pdf", _
    swSaveAsCurrentVersion, _
    swSaveAsOptions_Silent, _
    swExportPDFData, _
    errors, _
    warnings)

'if any drawing sheets are named CUT, switch to that sheet and save as a dxf
If swDrawing.ActivateSheet("CUT") Then
    bool = swModelDocExt.SaveAs(saveName + ".dxf", _
        swSaveAsCurrentVersion, _
        swSaveAsOptions_Silent, _
        Nothing, _
        errors, _
        warnings)
End If
swApp.QuitDoc swModel.GetTitle
swApp.QuitDoc swDrawing.GetTitle
fso.DeleteFile (tempDir + drawingName)
fso.DeleteFile (tempDir + modelName)
End Function
