Option Explicit

Dim swApp               As SldWorks.SldWorks
Dim CN                  As ADODB.Connection

Dim regEx               As New RegExp

'------------------------------------------------------------------------------'
Sub BOMCheck()

Dim swModel             As SldWorks.ModelDoc2
Dim swBOMTable          As SldWorks.BomTableAnnotation
Dim swTable             As SldWorks.TableAnnotation
Dim swComponent         As SldWorks.Component2
Dim swAnn               As SldWorks.Annotation
Dim swBOMComponent      As SldWorks.ModelDoc2
Dim RS                  As ADODB.Recordset
Dim Field               As ADODB.Field
Dim regexpmatches       As MatchCollection

Dim ActiveConfigname    As String
Dim failmatches()       As String
Dim Failstring          As String
Dim match               As String
Dim SQL                 As String
Dim componentName       As String
Dim fails               As Integer
Dim i                   As Integer
Dim j                   As Integer
Dim bool                As Boolean

Const BOMTemplate       As String = "X:\Engineering\Engineering Resources\SolidWorks Templates\CDG BOM Template\SHOXS BOM R3.sldbomtbt"

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

ActiveConfigname = swApp.GetActiveConfigurationName(swModel.GetPathName)
Set swBOMTable = swModel.Extension.InsertBomTable(BOMTemplate, 0, 0, swBomType_e.swBomType_PartsOnly, ActiveConfigname)
Set swTable = swBOMTable
Set swAnn = swTable.GetAnnotation
Set RS = New ADODB.Recordset

With regEx
    .Global = True
    .Multiline = True
    .IgnoreCase = True
End With

fails = 0

bool = NAVConnect()

For i = 1 To swTable.RowCount - 1

    'Debug.Print swBOMTable.GetComponentsCount(i)

    Set swComponent = swBOMTable.GetComponents(i)(0)

    regEx.Pattern = "^.+\/(.+)-\d+$|^(.+)-\d+$"

    Set regexpmatches = regEx.Execute(swComponent.Name)

    If regexpmatches(0).SubMatches(0) <> "" Then
        match = regexpmatches(0).SubMatches(0)
    Else
        match = regexpmatches(0).SubMatches(1)
    End If

    'Debug.Print match

    SQL = "SELECT [No_],[CAD Item No_] FROM [CDG_NAV2013_Prod].[dbo].[CDG$Item] WHERE [No_] LIKE '" & match & "%'" & " OR [CAD Item No_] = '" & match & "'"

    RS.Open SQL, CN, adOpenStatic, adLockReadOnly, adCmdText

    If Not RS.EOF Then
        'Debug.Print RS.Fields(0)
    Else
        Debug.Print match
        fails = fails + 1
        ReDim Preserve failmatches(fails)
        failmatches(fails) = match

    End If

    RS.Close

Next i

CN.Close

swAnn.Select3 False, Nothing
swModel.EditDelete

If fails = 0 Then
    bool = swApp.SendMsgToUser2("All items match NAV", 2, 2)
Else
    For j = 0 To fails
        Failstring = Failstring & failmatches(j) & vbNewLine
    Next j

    bool = swApp.SendMsgToUser2(fails & " Items failed to match NAV" & vbNewLine & Failstring, 2, 2)

End If

End Sub
'------------------------------------------------------------------------------'
Function NAVConnect() As Boolean

    Set CN = New ADODB.Connection
    With CN
        ' Create connecting string
        .ConnectionString = "Provider=SQLOLEDB.1;" & _
                            "Integrated Security=SSPI;" & _
                            "Server=shoxs2;" & _
                            "Database=CDG_NAV2013_Prod;"
        ' Open connection
        .Open
    End With
    ' Check connection state
    If CN.State = 0 Then
        NAVConnect = False
    Else
        NAVConnect = True
    End If

End Function
