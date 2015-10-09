option Explicit

Dim swApp               As SldWorks.SldWorks
Dim CN                  As ADODB.Connection

'------------------------------------------------------------------------------'
Sub BOMCheck()

Dim swModel             As SldWorks.ModelDoc2
Dim swBOMTable          As SldWorks.BomTableAnnotation
Dim swTable             As SldWorks.TableAnnotation
Dim swComponent         As SldWorks.Component2
Dim componentName       As String
Dim swBOMComponent      As ModelDoc2
Dim NumbRow             As Integer
Dim match               As String
Dim SQL                 As String
Dim Connected           As Boolean
Dim RS                  As ADODB.Recordset
Dim Field               As ADODB.Field
Dim fails               As Integer
Dim failmatches()       As String
Dim Failsstring         As String
Dim swAnn               As SldWorks.Annotation
Dim ActiveConfigname    As String
Dim regEx               As New RegEx
Const BOMTemplate       As String = "X:\Engineering\Engineering Resources\SolidWorks Templates\CDG BOM Template\SHOXS BOM R3.sldbomtbt"

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

ActiveConfigname = swApp.GetActiveConfigurationName(swModel.GetPathName)
Set swBOMTable = swModel.Extension.InsertBomTable(BOMTemplate, 0, 0, swBomType_e.swBomType_PartsOnly, ActiveConfigname)
Set swTable = swBOMTable
Set swAnn = swTable.GetAnnotation
Set RS = New ADODB.Recordset

NumbRow = swTable.RowCount

With regEx
    .Global = True
    .Multiline = True
    .IgnoreCase = True
End With

fails = 0
Connected = Connect("shoxs2", "CDG_NAV2013_Prod")

For i = 1 To NumbRow - 1
    'Debug.Print swBOMTable.GetComponentsCount(i)
    Set swComponent = swBOMTable.GetComponents(i)(0)

    regEx.Pattern = "^.+\/(.+)-\d+$|^(.+)-\d+$"
    Set regexpmatches = regEx.Execute(swComponent.Name)

    If regexpmatches(0).SubMatches(0) <> "" Then
        match = regexpmatches(0).SubMatches(0)
    Else
        match = regexpmatches(0).SubMatches(1)
    End If

    Debug.Print match

    SQL = "SELECT [No_],[CAD Item No_] FROM [CDG_NAV2013_Prod].[dbo].[CDG$Item] WHERE [No_] LIKE '" & match & "%'" & " OR [CAD Item No_] = '" & match & "'"
    RS.Open SQL, CN, adOpenStatic, adLockReadOnly, adCmdText

    If Not RS.EOF Then
        Debug.Print RS.Fields(0)
    Else
        Debug.Print "----------------NO MATCH-----------------"
        fails = fails + 1
        ReDim Preserve failmatches(fails)
        failmatches(fails) = match

    End If

    RS.Close
Next i

If fails = 0 Then
    bool = swApp.SendMsgToUser2("All items match NAV", 2, 2)

Else
    For j = 0 To fails
        failstring = failstring & failmatches(j) & vbNewLine
    Next j

    bool = swApp.SendMsgToUser2(fails & " Items failed to match NAV" & vbNewLine & failstring, 2, 2)
End If

CN.Close

swAnn.Select3 False, Nothing

swModel.EditDelete

End Sub
'------------------------------------------------------------------------------'
Function Connect(Server As String, Database As String) As Boolean

    Set CN = New ADODB.Connection
    On Error Resume Next

    With CN
        ' Create connecting string
        .ConnectionString = "Provider=SQLOLEDB.1;" & _
                            "Integrated Security=SSPI;" & _
                            "Server=" & Server & ";" & _
                            "Database=" & Database & ";"
        ' Open connection
        .Open
    End With
    ' Check connection state
    If CN.State = 0 Then
        Connect = False
    Else
        Connect = True
    End If
 
End Function
