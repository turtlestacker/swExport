'===============================================================================
' ExportAssemblyPackage.vb
' SolidWorks 2025 VBA Macro
'
' Exports all fabricated components (blank Vendor property) from the active
' assembly as PDFs + PNG screenshots and generates a self-contained HTML
' tree browser (index.html).
'===============================================================================

'==== USER SETTINGS ====
Const SKIP_SUPPRESSED As Boolean = True
Const NAME_SEPARATOR As String = " - "
Const SCREENSHOT_WIDTH As Long = 800
Const SCREENSHOT_HEIGHT As Long = 600
'=======================

Dim swApp As SldWorks.SldWorks
Dim gOutFolder As String
Dim gExported As Collection          ' dedup: keyed by UCase(modelPath)
Dim gHtml As String                  ' accumulates the HTML output
Dim gLogMessages As String           ' accumulated log

'===============================================================================
' Entry Point
'===============================================================================
Sub Main()
    On Error GoTo EH

    Set swApp = Application.SldWorks

    ' Validate active doc is an assembly
    Dim swDoc As SldWorks.ModelDoc2
    Set swDoc = swApp.ActiveDoc
    If swDoc Is Nothing Then
        MsgBox "No document is open.", vbExclamation, "Export Assembly Package"
        Exit Sub
    End If
    If swDoc.GetType <> swDocASSEMBLY Then
        MsgBox "Active document must be an assembly.", vbExclamation, "Export Assembly Package"
        Exit Sub
    End If

    ' Prompt for output folder
    gOutFolder = BrowseForFolder()
    If Len(gOutFolder) = 0 Then
        Exit Sub
    End If

    ' Initialise dedup collection
    Set gExported = New Collection

    ' Initialise HTML
    Dim assyName As String
    assyName = GetBaseNameNoExt(swDoc.GetPathName)
    Call InitHtml(assyName)

    ' Get root component and traverse
    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swDoc
    Dim swRootComp As SldWorks.Component2
    Set swRootComp = swAssy.GetRootComponent3(True)

    LogMessage "Starting export of assembly: " & assyName

    Call TraverseComponent(swRootComp, 0)

    ' Finalise and write HTML
    Dim htmlContent As String
    htmlContent = FinaliseHtml()
    Call WriteHtmlFile(gOutFolder, htmlContent)

    LogMessage "Export complete."
    MsgBox "Export complete!" & vbCrLf & vbCrLf & _
           "Output folder: " & gOutFolder & vbCrLf & _
           "Parts exported: " & gExported.Count, _
           vbInformation, "Export Assembly Package"
    Exit Sub

EH:
    MsgBox "Error in Main: " & Err.Description, vbCritical, "Export Assembly Package"
End Sub

'===============================================================================
' Assembly Traversal
'===============================================================================
Private Sub TraverseComponent(ByVal swComp As SldWorks.Component2, ByVal depth As Long)
    On Error GoTo EH

    Dim vChildren As Variant
    vChildren = swComp.GetChildren

    If IsEmpty(vChildren) Then Exit Sub

    Dim i As Long
    For i = 0 To UBound(vChildren)
        Dim swChild As SldWorks.Component2
        Set swChild = vChildren(i)

        ' Skip suppressed components
        If SKIP_SUPPRESSED Then
            Dim suppState As Long
            suppState = swChild.GetSuppression2
            If suppState = swComponentSuppressed Then
                LogMessage "Skipping suppressed: " & swChild.Name2
                GoTo NextChild
            ElseIf suppState = swComponentLightweight Then
                ' SolidWorks 2025 often loads components lightweight; resolve before access.
                Dim resolvedOk As Boolean
                resolvedOk = swChild.SetSuppression2(swComponentResolved)
                If Not resolvedOk Then
                    LogMessage "Could not resolve lightweight component: " & swChild.Name2
                    GoTo NextChild
                End If
            End If
        End If

        ' Get model doc
        Dim swChildModel As SldWorks.ModelDoc2
        Set swChildModel = swChild.GetModelDoc2
        If swChildModel Is Nothing Then
            LogMessage "Could not get model for: " & swChild.Name2
            GoTo NextChild
        End If

        ' Skip vendor parts
        If IsVendorPart(swChildModel) Then
            LogMessage "Skipping vendor part: " & swChild.Name2
            GoTo NextChild
        End If

        Dim modelPath As String
        modelPath = swChildModel.GetPathName
        Dim modelType As Long
        modelType = swChildModel.GetType
        Dim displayName As String
        displayName = swChild.Name2

        If modelType = swDocASSEMBLY Then
            ' Sub-assembly: add collapsible node, recurse, close
            Call AppendAssemblyNode(displayName, depth)
            Call TraverseComponent(swChild, depth + 1)
            Call CloseAssemblyNode
        Else
            ' Part: export if not already done, add leaf node
            Dim pdfFile As String
            Dim pngFile As String
            pdfFile = ""
            pngFile = ""

            Dim dedupKey As String
            dedupKey = UCase(modelPath)
            Dim alreadyExported As Boolean
            alreadyExported = False

            On Error Resume Next
            Dim dummy As String
            dummy = gExported(dedupKey)
            If Err.Number = 0 Then
                alreadyExported = True
            End If
            Err.Clear
            On Error GoTo EH

            If Not alreadyExported Then
                ' Mark as exported
                gExported.Add dedupKey, dedupKey

                Dim desc As String
                desc = GetDescription(swChildModel)
                Dim baseName As String
                baseName = GetBaseNameNoExt(modelPath)

                ' Find and export drawing PDF
                Dim drwPath As String
                drwPath = FindDrawingForPart(modelPath)
                If Len(drwPath) > 0 Then
                    pdfFile = ExportDrawingToPdf(drwPath, gOutFolder, desc)
                Else
                    LogMessage "No drawing found for: " & baseName
                End If

                ' Capture screenshot
                pngFile = CapturePartScreenshot(swChildModel, gOutFolder, baseName, desc)
            End If

            Call AppendPartNode(displayName, pngFile, pdfFile)
        End If

NextChild:
    Next i
    Exit Sub

EH:
    LogMessage "Error in TraverseComponent: " & Err.Description
    Resume NextChild
End Sub

'===============================================================================
' Part Processing
'===============================================================================
Private Sub ProcessPart(ByVal swComp As SldWorks.Component2, ByVal modelPath As String)
    ' Orchestration handled inline in TraverseComponent for simplicity
End Sub

'===============================================================================
' Drawing Lookup
'===============================================================================
Private Function FindDrawingForPart(ByVal partPath As String) As String
    On Error GoTo EH

    Dim partFolder As String
    partFolder = GetFolderFromPath(partPath)

    Dim drwFile As String
    drwFile = Dir(partFolder & "*.SLDDRW")

    Do While Len(drwFile) > 0
        Dim drwFullPath As String
        drwFullPath = partFolder & drwFile

        ' Check if this drawing references our part
        Dim vDeps As Variant
        vDeps = swApp.GetDocumentDependencies2(drwFullPath, False, False, False)

        If Not IsEmpty(vDeps) Then
            Dim j As Long
            ' Dependencies returned as pairs: [name, path, name, path, ...]
            For j = 1 To UBound(vDeps) Step 2
                If UCase(vDeps(j)) = UCase(partPath) Then
                    FindDrawingForPart = drwFullPath
                    Exit Function
                End If
            Next j
        End If

        drwFile = Dir()
    Loop

    FindDrawingForPart = ""
    Exit Function

EH:
    FindDrawingForPart = ""
End Function

'===============================================================================
' PDF Export
'===============================================================================
Private Function ExportDrawingToPdf(ByVal drwPath As String, ByVal outFolder As String, ByVal desc As String) As String
    On Error GoTo EH

    Dim swDrw As SldWorks.ModelDoc2
    Dim errs As Long, warns As Long

    Set swDrw = swApp.OpenDoc6(drwPath, swDocDRAWING, swOpenDocOptions_Silent, "", errs, warns)
    If swDrw Is Nothing Then
        LogMessage "Failed to open drawing: " & drwPath
        ExportDrawingToPdf = ""
        Exit Function
    End If

    Dim baseName As String
    baseName = GetBaseNameNoExt(drwPath)

    Dim pdfName As String
    pdfName = BuildExportName(baseName, desc) & ".pdf"

    Dim pdfPath As String
    pdfPath = outFolder & "\" & pdfName

    Dim swExt As SldWorks.ModelDocExtension
    Set swExt = swDrw.Extension

    Dim ok As Boolean
    ok = swExt.SaveAs(pdfPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errs, warns)

    swApp.CloseDoc swDrw.GetTitle

    If ok Then
        LogMessage "Exported PDF: " & pdfName
        ExportDrawingToPdf = pdfName
    Else
        LogMessage "PDF export failed for: " & pdfName
        ExportDrawingToPdf = ""
    End If
    Exit Function

EH:
    LogMessage "Error exporting PDF: " & Err.Description
    ExportDrawingToPdf = ""
End Function

'===============================================================================
' Screenshot Capture
'===============================================================================
Private Function CapturePartScreenshot(ByVal swPartDoc As SldWorks.ModelDoc2, ByVal outFolder As String, _
                                        ByVal baseName As String, ByVal desc As String) As String
    On Error GoTo EH

    ' Activate the part (use document title, not full path, for SolidWorks 2025)
    Dim errs As Long
    Dim warns As Long
    Dim activeDoc As SldWorks.ModelDoc2
    Dim partTitle As String
    Dim previousDoc As SldWorks.ModelDoc2
    Set previousDoc = swApp.ActiveDoc
    partTitle = swPartDoc.GetTitle
    Set activeDoc = swApp.ActivateDoc3(partTitle, False, swActivateDocError_e.swGenericActivateError, errs)

    If activeDoc Is Nothing Then
        LogMessage "Could not activate part for screenshot: " & baseName
        CapturePartScreenshot = ""
        Exit Function
    End If

    ' Set isometric view
    activeDoc.ShowNamedView2 "*Isometric", -1
    activeDoc.ViewZoomtofit2

    ' Build filename
    Dim pngName As String
    pngName = BuildExportName(baseName, desc) & ".png"

    Dim pngPath As String
    pngPath = outFolder & "\" & pngName

    ' Try SaveAs PNG via Extension
    Dim swExt As SldWorks.ModelDocExtension
    Set swExt = activeDoc.Extension

    Dim ok As Boolean
    ok = swExt.SaveAs(pngPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errs, warns)

    If Not ok Then
        ' Fallback: use SaveBMP
        LogMessage "PNG SaveAs failed, trying SaveBMP fallback for: " & baseName
        ok = activeDoc.SaveBMP(pngPath, SCREENSHOT_WIDTH, SCREENSHOT_HEIGHT)
    End If

    ' Re-activate the assembly
    If Not previousDoc Is Nothing Then
        swApp.ActivateDoc3 previousDoc.GetTitle, False, swActivateDocError_e.swGenericActivateError, errs
    End If

    If ok Then
        LogMessage "Captured screenshot: " & pngName
        CapturePartScreenshot = pngName
    Else
        LogMessage "Screenshot failed for: " & baseName
        CapturePartScreenshot = ""
    End If
    Exit Function

EH:
    LogMessage "Error capturing screenshot: " & Err.Description
    CapturePartScreenshot = ""
End Function

'===============================================================================
' Description / Property Helpers
'===============================================================================
Private Function GetDescription(ByVal swDoc As SldWorks.ModelDoc2) As String
    On Error GoTo EH

    Dim s As String

    ' Try active configuration first
    Dim cfgName As String
    cfgName = swDoc.ConfigurationManager.ActiveConfiguration.Name

    s = GetCustomProp(swDoc, cfgName, "Description")
    If Len(Trim$(s)) > 0 Then
        GetDescription = s
        Exit Function
    End If

    ' Try custom tab
    s = GetCustomProp(swDoc, "", "Description")
    If Len(Trim$(s)) > 0 Then
        GetDescription = s
        Exit Function
    End If

    ' Try SW-Description config-specific
    s = GetCustomProp(swDoc, cfgName, "SW-Description")
    If Len(Trim$(s)) > 0 Then
        GetDescription = s
        Exit Function
    End If

    ' Try SW-Description custom tab
    s = GetCustomProp(swDoc, "", "SW-Description")
    If Len(Trim$(s)) > 0 Then
        GetDescription = s
        Exit Function
    End If

    GetDescription = ""
    Exit Function

EH:
    GetDescription = ""
End Function

Private Function GetCustomProp(ByVal doc As SldWorks.ModelDoc2, ByVal cfgName As String, ByVal propName As String) As String
    On Error GoTo EH

    Dim ext As SldWorks.ModelDocExtension
    Set ext = doc.Extension

    Dim cpm As SldWorks.CustomPropertyManager
    Set cpm = ext.CustomPropertyManager(cfgName)

    Dim valOut As String, resolvedVal As String
    Dim wasResolved As Boolean

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

Private Function IsVendorPart(ByVal swDoc As SldWorks.ModelDoc2) As Boolean
    On Error GoTo EH

    Dim vendor As String

    ' Check config-specific first
    Dim cfgName As String
    cfgName = ""
    On Error Resume Next
    cfgName = swDoc.ConfigurationManager.ActiveConfiguration.Name
    On Error GoTo EH

    vendor = GetCustomProp(swDoc, cfgName, "Vendor")
    If Len(Trim$(vendor)) > 0 Then
        IsVendorPart = True
        Exit Function
    End If

    ' Check custom tab
    vendor = GetCustomProp(swDoc, "", "Vendor")
    If Len(Trim$(vendor)) > 0 Then
        IsVendorPart = True
        Exit Function
    End If

    IsVendorPart = False
    Exit Function

EH:
    IsVendorPart = False
End Function

'===============================================================================
' Filename Helpers
'===============================================================================
Private Function CleanFileName(ByVal s As String) As String
    On Error GoTo EH

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

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    CleanFileName = t
    Exit Function

EH:
    CleanFileName = s
End Function

Private Function GetBaseNameNoExt(ByVal filePath As String) As String
    On Error GoTo EH

    Dim fileName As String
    Dim pos As Long

    ' Get just the filename
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        fileName = Mid$(filePath, pos + 1)
    Else
        fileName = filePath
    End If

    ' Remove extension
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        GetBaseNameNoExt = Left$(fileName, pos - 1)
    Else
        GetBaseNameNoExt = fileName
    End If
    Exit Function

EH:
    GetBaseNameNoExt = filePath
End Function

Private Function BuildExportName(ByVal baseName As String, ByVal desc As String) As String
    On Error GoTo EH

    If Len(Trim$(desc)) > 0 Then
        BuildExportName = CleanFileName(baseName & NAME_SEPARATOR & desc)
    Else
        BuildExportName = CleanFileName(baseName)
    End If
    Exit Function

EH:
    BuildExportName = CleanFileName(baseName)
End Function

'===============================================================================
' HTML Generation
'===============================================================================
Private Sub InitHtml(ByVal assemblyName As String)
    gHtml = "<!DOCTYPE html>" & vbCrLf
    gHtml = gHtml & "<html lang=""en"">" & vbCrLf
    gHtml = gHtml & "<head>" & vbCrLf
    gHtml = gHtml & "<meta charset=""UTF-8"">" & vbCrLf
    gHtml = gHtml & "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf
    gHtml = gHtml & "<title>" & assemblyName & " - Assembly Package</title>" & vbCrLf
    gHtml = gHtml & "<style>" & vbCrLf
    gHtml = gHtml & "  * { box-sizing: border-box; margin: 0; padding: 0; }" & vbCrLf
    gHtml = gHtml & "  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;" & vbCrLf
    gHtml = gHtml & "         background: #f5f5f5; color: #333; padding: 20px; }" & vbCrLf
    gHtml = gHtml & "  h1 { margin-bottom: 16px; font-size: 1.4em; color: #1a1a1a; }" & vbCrLf
    gHtml = gHtml & "  ul { list-style: none; padding-left: 24px; }" & vbCrLf
    gHtml = gHtml & "  ul.root { padding-left: 0; }" & vbCrLf
    gHtml = gHtml & "  li { margin: 4px 0; }" & vbCrLf
    gHtml = gHtml & "  li.assy { cursor: pointer; }" & vbCrLf
    gHtml = gHtml & "  li.assy > .node-label { font-weight: 600; padding: 4px 8px;" & vbCrLf
    gHtml = gHtml & "    background: #e8e8e8; border-radius: 4px; display: inline-block; }" & vbCrLf
    gHtml = gHtml & "  li.assy > .node-label:hover { background: #d0d0d0; }" & vbCrLf
    gHtml = gHtml & "  .toggle-arrow { display: inline-block; width: 16px; text-align: center;" & vbCrLf
    gHtml = gHtml & "    transition: transform 0.15s; font-size: 0.8em; }" & vbCrLf
    gHtml = gHtml & "  li.assy.expanded > .node-label > .toggle-arrow { transform: rotate(90deg); }" & vbCrLf
    gHtml = gHtml & "  li.part { display: flex; align-items: center; gap: 10px;" & vbCrLf
    gHtml = gHtml & "    padding: 4px 0; border-bottom: 1px solid #eee; }" & vbCrLf
    gHtml = gHtml & "  .thumb { width: 150px; height: auto; border: 1px solid #ccc;" & vbCrLf
    gHtml = gHtml & "    border-radius: 4px; background: #fff; }" & vbCrLf
    gHtml = gHtml & "  .part-name { font-size: 0.95em; }" & vbCrLf
    gHtml = gHtml & "  .part-name a { color: #0066cc; text-decoration: none; }" & vbCrLf
    gHtml = gHtml & "  .part-name a:hover { text-decoration: underline; }" & vbCrLf
    gHtml = gHtml & "  .children { margin-top: 4px; }" & vbCrLf
    gHtml = gHtml & "</style>" & vbCrLf
    gHtml = gHtml & "</head>" & vbCrLf
    gHtml = gHtml & "<body>" & vbCrLf
    gHtml = gHtml & "<h1>&#128204; " & assemblyName & "</h1>" & vbCrLf
    gHtml = gHtml & "<ul class=""root"">" & vbCrLf
End Sub

Private Sub AppendAssemblyNode(ByVal Name As String, ByVal depth As Long)
    gHtml = gHtml & "<li class=""assy collapsed"">" & vbCrLf
    gHtml = gHtml & "  <span class=""node-label"" onclick=""toggle(this.parentElement)"">" & vbCrLf
    gHtml = gHtml & "    <span class=""toggle-arrow"">&#9654;</span> &#128193; " & HtmlEncode(Name) & vbCrLf
    gHtml = gHtml & "  </span>" & vbCrLf
    gHtml = gHtml & "  <ul class=""children"" style=""display:none"">" & vbCrLf
End Sub

Private Sub AppendPartNode(ByVal Name As String, ByVal pngFile As String, ByVal pdfFile As String)
    gHtml = gHtml & "<li class=""part"">" & vbCrLf

    ' Thumbnail
    If Len(pngFile) > 0 Then
        gHtml = gHtml & "  <img src=""" & HtmlEncode(pngFile) & """ class=""thumb"" alt=""" & HtmlEncode(Name) & """>" & vbCrLf
    End If

    ' Name with optional PDF link
    gHtml = gHtml & "  <span class=""part-name"">"
    If Len(pdfFile) > 0 Then
        gHtml = gHtml & "<a href=""" & HtmlEncode(pdfFile) & """>" & HtmlEncode(Name) & "</a>"
    Else
        gHtml = gHtml & HtmlEncode(Name)
    End If
    gHtml = gHtml & "</span>" & vbCrLf

    gHtml = gHtml & "</li>" & vbCrLf
End Sub

Private Sub CloseAssemblyNode()
    gHtml = gHtml & "  </ul>" & vbCrLf
    gHtml = gHtml & "</li>" & vbCrLf
End Sub

Private Function FinaliseHtml() As String
    Dim html As String
    html = gHtml

    ' Close root ul and add JavaScript
    html = html & "</ul>" & vbCrLf
    html = html & "<script>" & vbCrLf
    html = html & "function toggle(li) {" & vbCrLf
    html = html & "  var ul = li.querySelector('.children');" & vbCrLf
    html = html & "  if (!ul) return;" & vbCrLf
    html = html & "  if (ul.style.display === 'none') {" & vbCrLf
    html = html & "    ul.style.display = 'block';" & vbCrLf
    html = html & "    li.classList.remove('collapsed');" & vbCrLf
    html = html & "    li.classList.add('expanded');" & vbCrLf
    html = html & "  } else {" & vbCrLf
    html = html & "    ul.style.display = 'none';" & vbCrLf
    html = html & "    li.classList.remove('expanded');" & vbCrLf
    html = html & "    li.classList.add('collapsed');" & vbCrLf
    html = html & "  }" & vbCrLf
    html = html & "}" & vbCrLf
    html = html & "</script>" & vbCrLf
    html = html & "</body>" & vbCrLf
    html = html & "</html>" & vbCrLf

    FinaliseHtml = html
End Function

Private Sub WriteHtmlFile(ByVal outFolder As String, ByVal html As String)
    On Error GoTo EH

    Dim filePath As String
    filePath = outFolder & "\index.html"

    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, html
    Close #fNum

    LogMessage "Wrote index.html"
    Exit Sub

EH:
    LogMessage "Error writing HTML file: " & Err.Description
End Sub

'===============================================================================
' HTML Encoding Helper
'===============================================================================
Private Function HtmlEncode(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    HtmlEncode = t
End Function

'===============================================================================
' UI / Utility
'===============================================================================
Private Function BrowseForFolder() As String
    On Error GoTo EH

    Dim shell As Object
    Set shell = CreateObject("Shell.Application")

    Dim folder As Object
    Set folder = shell.BrowseForFolder(0, "Select output folder for Assembly Package export:", 0)

    If folder Is Nothing Then
        BrowseForFolder = ""
    Else
        BrowseForFolder = folder.Self.Path
    End If
    Exit Function

EH:
    BrowseForFolder = ""
End Function

Private Sub LogMessage(ByVal msg As String)
    Debug.Print msg
    gLogMessages = gLogMessages & msg & vbCrLf

    ' Update status bar if available
    On Error Resume Next
    swApp.ActiveDoc.ClearSelection2 True
    On Error GoTo 0
End Sub

Private Function GetFolderFromPath(ByVal filePath As String) As String
    On Error GoTo EH

    Dim pos As Long
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFolderFromPath = Left$(filePath, pos)
    Else
        GetFolderFromPath = ""
    End If
    Exit Function

EH:
    GetFolderFromPath = ""
End Function
