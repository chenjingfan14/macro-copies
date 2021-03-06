Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swBOMTable As SldWorks.BomTableAnnotation
Dim swTable As SldWorks.TableAnnotation
Dim swAnn As SldWorks.Annotation
Dim ActiveConfigname As String
Const BOMTemplate As String = "X:\Engineering\Engineering Resources\SolidWorks Templates\CDG BOM Template\SHOXS BOM R3.sldbomtbt"
Const OutputPath As String = "X:\Engineering\PDMWorks\BOM Exports\"

Sub main()
    Set swApp = Application.SldWorks

    Set swModel = swApp.ActiveDoc

    ActiveConfigname = swApp.GetActiveConfigurationName (swModel.GetPathName)

    Set swBOMTable = swModel.Extension.InsertBomTable (BOMTemplate, 0, 0, swBomType_e.swBomType_PartsOnly, ActiveConfigname)

    Set swTable = swBOMTable

    swTable.SaveAsText OutputPath & swModel.GetTitle() & " - " & ActiveConfigname & ".txt", ","

    Set swAnn = swTable.GetAnnotation

    swAnn.Select3 False, Nothing

    swModel.EditDelete

End Sub
