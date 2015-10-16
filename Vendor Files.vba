Option Explicit
Dim swApp As SldWorks.SldWorks
'------------------------------------------------------------------------------'
Sub saveOpenVendorFiles()

'object declarations
Dim swModel         As SldWorks.ModelDoc2
Dim PDMConnection   As IPDMWConnection
'data declarations
Dim partToSave      As String
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
Sub saveListedFiles()

'local object declarations'
Dim PDMConnection   As IPDMWConnection
Dim PDMPart         As PDMWDocument
Dim PDMDrawing      As PDMWDocument

'local data declarations'
Dim errors          As Long
Dim drawingName     As String
Dim modelName       As String
Dim modelnumber()   As String
Dim j               As Integer

'local constant declarations'
Const inputFile     As String = "C:\Users\jpettit\Desktop\SCRIPTS\filesToSave.txt"
Const outputFile    As String = "C:\Users\jpettit\Desktop\SCRIPTS\fileSaveOutput.txt"
Const tempDir       As String = "X:\Engineering\TEMP\"
Const pdmName       As String = "jpettit"
Const pdmLogin      As String = "CDGshoxs!"
Const pdmServer     As String = "SHOXS1"


'initialize objects and start the PDM connection'
Set swApp = Application.SldWorks
Open outputFile For Output As #2
Set PDMConnection = CreateObject("PDMWorks.PDMWConnection")
PDMConnection.Login pdmName, pdmLogin, pdmServer

'function call which returns array of part numbers to change'
'Part numbers are read from an external file'
modelnumber() = readData(inputFile)
Debug.Print UBound(modelnumber()) + 1 & " PARTS TO SAVE"

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

    errors = saveVendorFiles(modelnumber(j), PDMConnection)

    Debug.Print modelnumber(j) + " SAVED"
    Print #2, modelnumber(j) + ", SAVED"


'loop back to the next model number that was read from the input file the
'GOTO to eject from the loop points here.
nextLoop: Next j

'cleanup by logging out of pdm. the vendor files script saves over the
'files left in temp and then deletes them, but this is kind of a shoddy way
'to clean up the files in each loop...'
PDMConnection.Logout
Close #2
Debug.Print "DONE"

End Sub
'------------------------------------------------------------------------------'
Public Function saveVendorFiles(partNumber As String, _
    passedPDMConnection As IPDMWConnection) As Long

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
Dim bool            As Boolean

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

swDrawing.ForceRebuild3 True

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
