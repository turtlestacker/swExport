'==== USER SETTINGS ====
Const INCLUDE_TOP_LEVEL_ASSEMBLY As Boolean = True
Const SKIP_SUPPRESSED As Boolean = True
Const CHECK_DRAWINGS_SUBFOLDER As Boolean = True

' Set to "" if you want true concatenation: <base><Description>.pdf
Const NAME_SEPARATOR As String = ""
'=======================

' ...keep the rest of your macro, but REPLACE ExportDrawingToPdf
' and ADD the helper functions below.

Private Function ExportDrawingToPdf(ByVal drwPath As String, ByVal outFolder As String) As Boolean

    On Error GoTo EH

    Dim swDrw As SldWorks.ModelDoc2
    Dim errs As Long, warns As Long

    Set swDrw = swApp.OpenDoc6(drwPath, swDocDRAWING, swOpenDocOptions_Silent, "", errs, warns)
    If swDrw Is Nothing Then
        ExportDrawingToPdf = False
        Exit Function
    End If

    Dim desc As String
    desc = GetDescriptionForDrawing(swDrw) ' <-- NEW

    Dim base As String
    base = GetBaseNameNoExt(drwPath)

    Dim pdfName As String
    If Len(Trim$(desc)) > 0 Then
        pdfName = base & NAME_SEPARATOR & CleanFileName(desc) & ".pdf"
    Else
        pdfName = base & ".pdf"
    End If

    Dim pdfPath As String
    pdfPath = outFolder & "\" & pdfName

    Dim swExt As SldWorks.ModelDocExtension
    Set swExt = swDrw.Extension

    Dim ok As Boolean
    ok = swExt.SaveAs(pdfPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errs, warns)

    swApp.CloseDoc swDrw.GetTitle

    ExportDrawingToPdf = ok
    Exit Function

EH:
    ExportDrawingToPdf = False
End Function

'------------------ NEW HELPERS ------------------

Private Function GetDescriptionForDrawing(ByVal swDrw As SldWorks.ModelDoc2) As String
    ' Tries, in order:
    ' 1) Drawing custom property "Description"
    ' 2) Drawing custom property "SW-Description"
    ' 3) Referenced model custom property "Description" (config-specific then custom)
    ' 4) Referenced model custom property "SW-Description"

    On Error GoTo EH

    Dim s As String

    ' 1) Drawing property: Description
    s = GetCustomProp(swDrw, "", "Description")
    If Len(Trim$(s)) > 0 Then
        GetDescriptionForDrawing = s
        Exit Function
    End If

    ' 2) Drawing property: SW-Description
    s = GetCustomProp(swDrw, "", "SW-Description")
    If Len(Trim$(s)) > 0 Then
        GetDescriptionForDrawing = s
        Exit Function
    End If

    ' 3/4) Referenced model properties (use first model view on sheet)
    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDrw

    Dim vSheetView As SldWorks.View
    Set vSheetView = swDrawing.GetFirstView ' sheet view

    If vSheetView Is Nothing Then GoTo EH

    Dim vModelView As SldWorks.View
    Set vModelView = vSheetView.GetNextView ' first model view

    If vModelView Is Nothing Then GoTo EH

    Dim refModelPath As String
    refModelPath = vModelView.GetReferencedModelName2(False) ' full path
    If Len(refModelPath) = 0 Then GoTo EH

    Dim refCfg As String
    refCfg = vModelView.ReferencedConfiguration

    ' Open referenced model silently (read props)
    Dim refDoc As SldWorks.ModelDoc2
    Dim errs As Long, warns As Long
    Set refDoc = swApp.OpenDoc6(refModelPath, swDocUNKNOWN, swOpenDocOptions_Silent, "", errs, warns)

    If refDoc Is Nothing Then GoTo EH

    ' Try config-specific then custom-tab
    s = GetCustomProp(refDoc, refCfg, "Description")
    If Len(Trim$(s)) = 0 Then s = GetCustomProp(refDoc, "", "Description")
    If Len(Trim$(s)) > 0 Then
        GetDescriptionForDrawing = s
        ' Close model we opened
        swApp.CloseDoc refDoc.GetTitle
        Exit Function
    End If

    s = GetCustomProp(refDoc, refCfg, "SW-Description")
    If Len(Trim$(s)) = 0 Then s = GetCustomProp(refDoc, "", "SW-Description")
    If Len(Trim$(s)) > 0 Then
        GetDescriptionForDrawing = s
        swApp.CloseDoc refDoc.GetTitle
        Exit Function
    End If

    swApp.CloseDoc refDoc.GetTitle

EH:
    GetDescriptionForDrawing = ""
End Function

Private Function GetCustomProp(ByVal doc As SldWorks.ModelDoc2, ByVal cfgName As String, ByVal propName As String) As String
    On Error GoTo EH

    Dim ext As SldWorks.ModelDocExtension
    Set ext = doc.Extension

    Dim cpm As SldWorks.CustomPropertyManager
    Set cpm = ext.CustomPropertyManager(cfgName) ' "" = Custom tab, cfgName = config-specific

    Dim valOut As String, resolvedVal As String
    Dim wasResolved As Boolean

    ' Get4 is most reliable for resolved values
    cpm.Get4 propName, False, valOut, resolvedVal, wasResolved

    If Len(Trim$(resolvedVal)) > 0 Then
        GetCustomProp = resolvedVal
    Else
        GetCustomProp = valOut
    End If
    Exit Function

EH:
    GetCustomProp = ""
End Function

Private Function CleanFileName(ByVal s As String) As String
    ' Removes characters illegal in Windows filenames and trims whitespace.
    Dim t As String
    t = Trim$(s)

    t = Replace(t, "\", "-")
    t = Replace(t, "/", "-")
    t = Replace(t, ":", "-")
    t = Replace(t, "*", "")
    t = Replace(t, "?", "")
    t = Replace(t, """", "")
    t = Replace(t, "<", "")
    t = Replace(t, ">", "")
    t = Replace(t, "|", "-")

    ' Optional: collapse multiple spaces
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    CleanFileName = t
End Function
