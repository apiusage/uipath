' ============================================================
' UiPath XAML Activity Renamer
' Usage: cscript rename_assigns.vbs "YourFile.xaml"
' Output: YourFile_out.xaml
' Idempotent: safe to run multiple times on the same file.
' ============================================================
' FEATURE TOGGLES
' Comment out any line below to disable that feature.
' ============================================================
Dim OPT_ACTIVITY_RENAME  : OPT_ACTIVITY_RENAME  = True  ' 1. Activity DisplayName renaming
Dim OPT_ARGUMENT_RENAME  : OPT_ARGUMENT_RENAME  = True  ' 2. Argument direction prefix (in_/out_/io_)
Dim OPT_VARIABLE_RENAME  : OPT_VARIABLE_RENAME  = True  ' 3. Variable type prefix (str_/bool_/dt_/...)
' ============================================================

Set x = CreateObject("MSXML2.DOMDocument")
x.async = False
x.setProperty "SelectionLanguage", "XPath"
x.load WScript.Arguments(0)

Function GetOutputFromChild(node, suffix)
    Dim child, txt
    For Each child In node.childNodes
        If InStr(child.nodeName, suffix) > 0 Then
            txt = Trim(child.Text)
            txt = Replace(Replace(txt, "[", ""), "]", "")
            If txt <> "" Then
                GetOutputFromChild = txt
                Exit Function
            End If
        End If
    Next
    GetOutputFromChild = ""
End Function

Function GetOutputFromAttr(node, attrName)
    Dim txt, i
    txt = node.getAttribute(attrName)
    If IsNull(txt) Then txt = ""
    If txt = "" Then
        For i = 0 To node.Attributes.Length - 1
            If LCase(node.Attributes(i).nodeName) = LCase(attrName) Then
                txt = node.Attributes(i).nodeValue
                Exit For
            End If
        Next
    End If
    If IsNull(txt) Then txt = ""
    txt = Trim(txt)
    txt = Replace(Replace(txt, "[", ""), "]", "")
    GetOutputFromAttr = txt
End Function

Function ExtractSmartLabel(raw)
    Dim txt, p1, p2
    txt = raw
    txt = Replace(Replace(txt, "[", ""), "]", "")
    txt = Replace(txt, Chr(34), """")
    p1 = InStr(txt, "(""")
    p2 = InStr(txt, """)")
    If p1 > 0 And p2 > p1 Then
        ExtractSmartLabel = Mid(txt, p1 + 2, p2 - p1 - 2)
        Exit Function
    End If
    If Left(txt, 1) = """" Then
        p2 = InStr(2, txt, """")
        If p2 > 1 Then
            ExtractSmartLabel = Mid(txt, 2, p2 - 2)
            Exit Function
        End If
    End If
    p1 = InStr(txt, "+")
    If p1 > 0 Then txt = Trim(Left(txt, p1 - 1))
    If Right(txt, 9) = ".ToString" Then txt = Left(txt, Len(txt) - 9)
    ExtractSmartLabel = txt
End Function

Function GetAanameFromSelector(node)
    Dim child, sel, p1, p2
    For Each child In node.childNodes
        If InStr(child.nodeName, "Target") > 0 Then
            sel = child.Text
            p1 = InStr(sel, "aaname='")
            If p1 > 0 Then
                p1 = p1 + 8
                p2 = InStr(p1, sel, "'")
                If p2 > p1 Then
                    GetAanameFromSelector = Mid(sel, p1, p2 - p1)
                    Exit Function
                End If
            End If
        End If
    Next
    GetAanameFromSelector = ""
End Function

Sub RenameNodes(localName, defaultLabel, varFunc, varArg)
    Dim nodes, n, v, newName
    Set nodes = x.SelectNodes("//*[local-name()='" & localName & "']")
    For Each n In nodes
        If varFunc = "child" Then
            v = GetOutputFromChild(n, varArg)
        ElseIf varFunc = "attr" Then
            v = GetOutputFromAttr(n, varArg)
        ElseIf varFunc = "smart" Then
            v = ExtractSmartLabel(GetOutputFromAttr(n, varArg))
        End If
        If v <> "" Then
            newName = defaultLabel & " - " & v
            n.setAttribute "DisplayName", newName
        End If
    Next
End Sub

Sub RenameCheckNodes()
    Dim nodes, n, d, action, aaname, chkLabel
    Set nodes = x.SelectNodes("//*[local-name()='Check']")
    For Each n In nodes
        d = n.getAttribute("DisplayName")
        If IsNull(d) Or d = "" Then d = "Check"
        action = n.getAttribute("Action")
        If IsNull(action) Then action = ""
        aaname = GetAanameFromSelector(n)
        chkLabel = ""
        If action <> "" And aaname <> "" Then
            chkLabel = action & " " & aaname
        ElseIf aaname <> "" Then
            chkLabel = aaname
        ElseIf action <> "" Then
            chkLabel = action
        End If
        If chkLabel <> "" Then n.setAttribute "DisplayName", "Check - " & chkLabel
    Next
End Sub

Sub RenameMultipleAssignNodes()
    Dim nodes, n, d, child, grandchild, txt, varList
    Set nodes = x.SelectNodes("//*[local-name()='MultipleAssign']")
    For Each n In nodes
        d = n.getAttribute("DisplayName")
        If IsNull(d) Or d = "" Then d = "Multiple Assign"
        varList = ""
        For Each child In n.childNodes
            If InStr(child.nodeName, "AssignOperations") > 0 Then
                For Each grandchild In child.childNodes
                    If InStr(grandchild.nodeName, "AssignOperation") > 0 Then
                        txt = GetOutputFromChild(grandchild, "AssignOperation.To")
                        If txt <> "" Then
                            If varList = "" Then varList = txt Else varList = varList & ", " & txt
                        End If
                    End If
                Next
            End If
        Next
        If varList <> "" Then n.setAttribute "DisplayName", "Multiple Assign - " & varList
    Next
End Sub

Sub RenameInvokeWorkflowNodes()
    Dim nodes, n, d, wfName, child, grandchild, outVar, txt, label
    Set nodes = x.SelectNodes("//*[local-name()='InvokeWorkflowFile']")
    For Each n In nodes
        d = n.getAttribute("DisplayName")
        If IsNull(d) Or d = "" Then d = "Invoke Workflow File"
        wfName = n.getAttribute("WorkflowFileName")
        If IsNull(wfName) Then wfName = ""
        outVar = ""
        For Each child In n.childNodes
            If InStr(child.nodeName, "Arguments") > 0 Then
                For Each grandchild In child.childNodes
                    If InStr(grandchild.nodeName, "OutArgument") > 0 Then
                        txt = Trim(grandchild.Text)
                        txt = Replace(Replace(txt, "[", ""), "]", "")
                        If txt <> "" And outVar = "" Then outVar = txt
                    End If
                Next
            End If
        Next
        label = wfName
        If outVar <> "" Then
            If label <> "" Then label = label & " -> " & outVar Else label = outVar
        End If
        If label <> "" Then n.setAttribute "DisplayName", "Invoke Workflow File - " & label
    Next
End Sub

Sub RenameExcelDeleteRangeNodes()
    Dim nodes, n, d, rng, sheet, label
    Set nodes = x.SelectNodes("//*[local-name()='ExcelDeleteRange']")
    For Each n In nodes
        d = n.getAttribute("DisplayName")
        If IsNull(d) Or d = "" Then d = "Delete Range"
        rng = n.getAttribute("Range")
        If IsNull(rng) Then rng = ""
        rng = Replace(Replace(rng, "[", ""), "]", "")
        sheet = n.getAttribute("SheetName")
        If IsNull(sheet) Then sheet = ""
        label = ""
        If rng <> "" Then label = rng
        If sheet <> "" Then
            If label <> "" Then label = label & " " & sheet Else label = sheet
        End If
        If label <> "" Then n.setAttribute "DisplayName", "Delete Range - " & label
    Next
End Sub

Sub RenameGetOutlookMailNodes()
    Dim nodes, n, d, folder, msgs, label
    Set nodes = x.SelectNodes("//*[local-name()='GetOutlookMailMessages']")
    For Each n In nodes
        d = n.getAttribute("DisplayName")
        If IsNull(d) Or d = "" Then d = "Get Outlook Mail Messages"
        folder = n.getAttribute("MailFolder")
        If IsNull(folder) Then folder = ""
        msgs = n.getAttribute("Messages")
        If IsNull(msgs) Then msgs = ""
        msgs = Replace(Replace(msgs, "[", ""), "]", "")
        label = ""
        If folder <> "" Then label = folder
        If msgs <> "" Then
            If label <> "" Then label = label & " -> " & msgs Else label = msgs
        End If
        If label <> "" Then n.setAttribute "DisplayName", "Get Outlook Mail Messages - " & label
    Next
End Sub

Sub RenameSendOutlookMailNodes()
    Dim nodes, n, d, toVal, subjVal, label
    Set nodes = x.SelectNodes("//*[local-name()='SendOutlookMail']")
    For Each n In nodes
        d = n.getAttribute("DisplayName")
        If IsNull(d) Or d = "" Then d = "Send Outlook Mail Message"
        toVal = ExtractSmartLabel(GetOutputFromAttr(n, "To"))
        subjVal = ExtractSmartLabel(GetOutputFromAttr(n, "Subject"))
        label = ""
        If toVal <> "" Then label = toVal
        If subjVal <> "" Then
            If label <> "" Then label = label & " -> " & subjVal Else label = subjVal
        End If
        If label <> "" Then n.setAttribute "DisplayName", "Send Outlook Mail Message - " & label
    Next
End Sub

Sub RenameWriteCellNodes()
    Dim nodes, n, cellVal, textVal
    Set nodes = x.SelectNodes("//*[local-name()='WriteCell']")
    For Each n In nodes
        cellVal = ExtractSmartLabel(GetOutputFromAttr(n, "Cell"))
        textVal = ExtractSmartLabel(GetOutputFromAttr(n, "Text"))
        If cellVal <> "" And textVal <> "" Then
            n.setAttribute "DisplayName", "Write Cell (" & cellVal & ") - " & textVal
        ElseIf cellVal <> "" Then
            n.setAttribute "DisplayName", "Write Cell (" & cellVal & ")"
        ElseIf textVal <> "" Then
            n.setAttribute "DisplayName", "Write Cell - " & textVal
        End If
    Next
End Sub

Sub RenameSwitchNodes()
    Dim nodes, n, child, keyAttr, caseList, expr, baseName, newName
    Set nodes = x.SelectNodes("//*[local-name()='Switch']")
    For Each n In nodes
        expr = n.getAttribute("Expression")
        If IsNull(expr) Then expr = ""
        expr = Replace(Replace(expr, "[", ""), "]", "")
        caseList = ""
        For Each child In n.childNodes
            If InStr(child.nodeName, "Switch.Default") = 0 Then
                keyAttr = child.getAttribute("x:Key")
                If IsNull(keyAttr) Then keyAttr = ""
                keyAttr = Trim(keyAttr)
                If keyAttr <> "" Then
                    If caseList = "" Then
                        caseList = keyAttr
                    Else
                        caseList = caseList & " / " & keyAttr
                    End If
                End If
            End If
        Next
        baseName = "Switch"
        If expr <> "" Then baseName = "Switch (" & expr & ")"
        If caseList <> "" Then
            newName = baseName & " - " & caseList
        ElseIf expr <> "" Then
            newName = baseName
        Else
            newName = ""
        End If
        If newName <> "" Then n.setAttribute "DisplayName", newName
    Next
End Sub

' Returns a type prefix (e.g. "str_", "bool_", "dt_") from a type string.
' Works for both bare types ("x:String") and wrapped argument types
' ("InArgument(x:String)", "OutArgument(sd:DataTable)"), by stripping
' the direction wrapper first.
Function GetTypePrefix(typeStr)
    Dim t, p1, p2
    t = typeStr
    ' Strip InOutArgument(...) / OutArgument(...) / InArgument(...) wrapper
    p1 = InStr(t, "(")
    p2 = InStr(t, ")")
    If p1 > 0 And p2 > p1 Then t = Mid(t, p1 + 1, p2 - p1 - 1)
    t = Trim(t)
    GetTypePrefix = ""
    If t = "x:String"                         Then GetTypePrefix = "str_"  : Exit Function
    If InStr(t, "String[]") > 0               Then GetTypePrefix = "arr_"  : Exit Function
    If t = "x:Boolean"                        Then GetTypePrefix = "bool_" : Exit Function
    If t = "x:Int32"                          Then GetTypePrefix = "int_"  : Exit Function
    If t = "x:Int64"                          Then GetTypePrefix = "lng_"  : Exit Function
    If t = "x:Double"                         Then GetTypePrefix = "dbl_"  : Exit Function
    If t = "x:Decimal"                        Then GetTypePrefix = "dec_"  : Exit Function
    If t = "x:DateTime"                       Then GetTypePrefix = "dt_"   : Exit Function
    If t = "sd:DataTable"                     Then GetTypePrefix = "dt_"   : Exit Function
    If InStr(t, "scg:List(") > 0              Then GetTypePrefix = "list_" : Exit Function
    If InStr(t, "scg:Dictionary(") > 0        Then GetTypePrefix = "dict_" : Exit Function
    If InStr(t, "scg:IEnumerable(") > 0       Then GetTypePrefix = "enum_" : Exit Function
    If t = "ui:GenericValue"                  Then GetTypePrefix = "gv_"   : Exit Function
    If InStr(t, "njl:JObject") > 0            Then GetTypePrefix = "jo_"   : Exit Function
    If InStr(t, "njl:JArray") > 0             Then GetTypePrefix = "ja_"   : Exit Function
    If t = "ui:QueueItem"                     Then GetTypePrefix = "qi_"   : Exit Function
    If InStr(t, "ui:Image") > 0               Then GetTypePrefix = "img_"  : Exit Function
    If InStr(t, "si:FileInfo") > 0            Then GetTypePrefix = "fi_"   : Exit Function
    If InStr(t, "si:DirectoryInfo") > 0       Then GetTypePrefix = "di_"   : Exit Function
End Function

' Returns True if varName already starts with any known direction or type prefix,
' INCLUDING merged direction+type combos like "in_str", "out_bool", "io_list" etc.
Function AlreadyPrefixed(varName)
    Dim kp, knownPrefixes, dirPfx, typePfx, dirPrefixes, typePrefixes
    ' Plain direction or type prefixes
    knownPrefixes = Array( _
        "in_","out_","io_", _
        "str_","arr_","bool_","int_","lng_","dbl_","dec_", _
        "dt_","list_","dict_","enum_","gv_","jo_","ja_", _
        "qi_","img_","fi_","di_")
    For Each kp In knownPrefixes
        If LCase(Left(varName, Len(kp))) = LCase(kp) Then
            AlreadyPrefixed = True
            Exit Function
        End If
    Next
    ' Also detect merged direction+type combos: out_str, in_bool, io_list, etc.
    ' These have no second underscore so won't match the plain checks above.
    dirPrefixes  = Array("in_", "out_", "io_")
    typePrefixes = Array("str","arr","bool","int","lng","dbl","dec", _
                         "dt","list","dict","enum","gv","jo","ja", _
                         "qi","img","fi","di")
    For Each dirPfx In dirPrefixes
        If LCase(Left(varName, Len(dirPfx))) = LCase(dirPfx) Then
            Dim rest
            rest = Mid(varName, Len(dirPfx) + 1)
            For Each typePfx In typePrefixes
                ' Merged form: dirPfx + typePfx + UppercaseLetter, e.g. out_strFoo
                If LCase(Left(rest, Len(typePfx))) = LCase(typePfx) Then
                    If Len(rest) > Len(typePfx) Then
                        Dim nextCh
                        nextCh = Mid(rest, Len(typePfx) + 1, 1)
                        If nextCh >= "A" And nextCh <= "Z" Then
                            AlreadyPrefixed = True
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
    Next
    AlreadyPrefixed = False
End Function

' Renames arguments with stacked direction+type prefix.
' e.g. OutArgument(x:String) "currentVoltage" -> "out_strCurrentVoltage"
' Direction prefix : in_ / out_ / io_
' Type prefix      : str_ / bool_ / dt_ / ... (merged, no second underscore)
' The base name is PascalCased (first letter uppercased) before merging so
' the boundary between the two prefixes stays readable:
'   out_ + str + CurrentVoltage  ->  out_strCurrentVoltage
Sub RenameArguments()
    Dim nodes, n, argName, argType, dirPrefix, typePrefix, typePart, baseName, newName
    Dim tp, typePrefixes
    Set nodes = x.SelectNodes("//*[local-name()='Property']")
    For Each n In nodes
        argName = n.getAttribute("Name")
        If IsNull(argName) Then argName = ""
        argType = n.getAttribute("Type")
        If IsNull(argType) Then argType = ""

        If argName <> "" And argType <> "" Then

            ' Determine direction prefix from the Type attribute
            dirPrefix = ""
            If Left(argType, 13) = "InOutArgument" Then
                dirPrefix = "io_"
            ElseIf Left(argType, 11) = "OutArgument" Then
                dirPrefix = "out_"
            ElseIf Left(argType, 10) = "InArgument" Then
                dirPrefix = "in_"
            End If

            If dirPrefix <> "" Then

                ' Strip existing direction prefix from name to get raw base
                baseName = argName
                If LCase(Left(baseName, 3)) = "in_"  Then baseName = Mid(baseName, 4)
                If LCase(Left(baseName, 4)) = "out_" Then baseName = Mid(baseName, 5)
                If LCase(Left(baseName, 3)) = "io_"  Then baseName = Mid(baseName, 4)

                ' Strip any type prefix already on the base (idempotency on re-runs).
                ' Must handle TWO forms:
                '   merged (no underscore): "listArgument4" -> strip "list" -> "Argument4"
                '   underscored form:       "list_Argument4" -> strip "list_" -> "Argument4"
                ' Check merged form first (abbrev + uppercase letter, no underscore).
                Dim typeAbbrevs, ta, nextCh2, stripped
                typeAbbrevs = Array("str","arr","bool","int","lng","dbl","dec", _
                                    "dt","list","dict","enum","gv","jo","ja", _
                                    "qi","img","fi","di")
                stripped = False
                For Each ta In typeAbbrevs
                    If LCase(Left(baseName, Len(ta))) = LCase(ta) And Len(baseName) > Len(ta) Then
                        nextCh2 = Mid(baseName, Len(ta) + 1, 1)
                        If nextCh2 >= "A" And nextCh2 <= "Z" Then
                            baseName = Mid(baseName, Len(ta) + 1)
                            stripped = True
                            Exit For
                        End If
                    End If
                Next
                ' Fall back to underscore form if merged form did not match
                If Not stripped Then
                    typePrefixes = Array("str_","arr_","bool_","int_","lng_","dbl_","dec_", _
                                         "dt_","list_","dict_","enum_","gv_","jo_","ja_", _
                                         "qi_","img_","fi_","di_")
                    For Each tp In typePrefixes
                        If LCase(Left(baseName, Len(tp))) = LCase(tp) Then
                            baseName = Mid(baseName, Len(tp) + 1)
                            Exit For
                        End If
                    Next
                End If

                ' PascalCase the clean base name so the seam is readable
                baseName = UCase(Left(baseName, 1)) & Mid(baseName, 2)

                ' Get type prefix e.g. "str_", strip trailing _ -> "str" for merging
                typePrefix = GetTypePrefix(argType)
                If typePrefix <> "" Then
                    typePart = Left(typePrefix, Len(typePrefix) - 1)
                Else
                    typePart = ""
                End If

                ' Build final name: out_ + str + CurrentVoltage = out_strCurrentVoltage
                If typePart <> "" Then
                    newName = dirPrefix & typePart & baseName
                Else
                    newName = dirPrefix & baseName
                End If

                ' Only update if something actually changed
                If newName <> argName Then
                    n.setAttribute "Name", newName
                    Call ReplaceVarRefs(argName, newName)
                End If

            End If
        End If
    Next
End Sub

Function ReplaceWordInStr(txt, oldName, newName)
    Dim pos, found, before, after, result, bBefore, bAfter
    result = ""
    pos = 1
    Do While pos <= Len(txt)
        found = InStr(pos, txt, oldName)
        If found = 0 Then
            result = result & Mid(txt, pos)
            Exit Do
        End If
        result = result & Mid(txt, pos, found - pos)
        If found > 1 Then
            before = Mid(txt, found - 1, 1)
        Else
            before = " "
        End If
        If found + Len(oldName) <= Len(txt) Then
            after = Mid(txt, found + Len(oldName), 1)
        Else
            after = " "
        End If
        bBefore = (before >= "A" And before <= "Z") Or _
                  (before >= "a" And before <= "z") Or _
                  (before >= "0" And before <= "9") Or _
                  (before = "_")
        bAfter  = (after >= "A" And after <= "Z") Or _
                  (after >= "a" And after <= "z") Or _
                  (after >= "0" And after <= "9") Or _
                  (after = "_")
        If Not bBefore And Not bAfter Then
            result = result & newName
        Else
            result = result & oldName
        End If
        pos = found + Len(oldName)
    Loop
    ReplaceWordInStr = result
End Function

Sub ReplaceVarRefs(oldName, newName)
    Dim allNodes, nd, attrIdx, attrVal, child
    Set allNodes = x.SelectNodes("//*")
    For Each nd In allNodes
        For attrIdx = 0 To nd.Attributes.Length - 1
            attrVal = nd.Attributes(attrIdx).nodeValue
            If InStr(attrVal, oldName) > 0 Then
                nd.Attributes(attrIdx).nodeValue = ReplaceWordInStr(attrVal, oldName, newName)
            End If
        Next
        For Each child In nd.childNodes
            If child.nodeType = 3 Then
                If InStr(child.nodeValue, oldName) > 0 Then
                    child.nodeValue = ReplaceWordInStr(child.nodeValue, oldName, newName)
                End If
            End If
        Next
    Next
End Sub

Sub RenameVariables()
    Dim nodes, n, typeArg, varName, prefix, newName
    Set nodes = x.SelectNodes("//*[local-name()='Variable']")
    For Each n In nodes
        typeArg = n.getAttribute("x:TypeArguments")
        If IsNull(typeArg) Then typeArg = ""
        varName = n.getAttribute("Name")
        If IsNull(varName) Then varName = ""

        If varName <> "" Then
            prefix = GetTypePrefix(typeArg)
            If prefix <> "" And Not AlreadyPrefixed(varName) Then
                newName = prefix & varName
                n.setAttribute "Name", newName
                Call ReplaceVarRefs(varName, newName)
            End If
        End If
    Next
End Sub

' ============================================================
' MAIN - Feature execution controlled by toggles at top of file
' ============================================================

' 1. ACTIVITY DISPLAY NAME RENAMING
If OPT_ACTIVITY_RENAME Then
    RenameNodes "Assign",               "Assign",                    "child", "Assign.To"
    RenameNodes "GetValue",             "Get Text",                  "child", "GetValue.Value"
    RenameNodes "GetOCRText",           "Get OCR Text",              "child", "GetOCRText.Text"
    RenameNodes "NGetText",             "Get Text",                  "child", "NGetText.Text"
    RenameNodes "InputDialog",          "Input Dialog",              "child", "InputDialog.Result"
    RenameNodes "GetRobotAsset",        "Get Orchestrator Asset",    "child", "GetRobotAsset.Value"
    RenameNodes "DeserializeJson",      "Deserialize JSON",          "smart", "JsonObject"
    RenameNodes "WriteLine",            "Write Line",                "smart", "Text"
    RenameNodes "LogMessage",           "Log Message",               "smart", "Message"
    RenameNodes "TypeInto",             "Type Into",                 "smart", "Text"
    RenameNodes "SelectItem",           "Select Item",               "smart", "Item"
    RenameNodes "MessageBox",           "Message Box",               "smart", "Text"
    RenameNodes "SelectFile",           "Select File",               "smart", "SelectedFile"
    RenameNodes "ExcelApplicationScope","Excel Application Scope",   "smart", "WorkbookPath"
    RenameNodes "ExcelReadRange",       "Read Range",                "smart", "DataTable"
    RenameNodes "ExcelWriteRange",      "Write Range",               "smart", "DataTable"
    RenameNodes "ExcelWriteCell",       "Write Cell",                "smart", "Text"
    RenameNodes "ReadRange",            "Read Range",                "smart", "DataTable"
    RenameNodes "ForEachRow",           "For Each Row in Data Table","smart", "DataTable"
    RenameNodes "ForEach",              "For Each",                  "smart", "Values"
    RenameNodes "ExtractFiles",         "Extract/Unzip Files",       "smart", "FileToExtract"
    RenameNodes "BuildDataTable",       "Build Data Table",          "smart", "DataTable"
    RenameNodes "AddDataRow",           "Add Data Row",              "smart", "DataTable"
    RenameNodes "AppendRange",          "Append Range",              "smart", "DataTable"
    RenameNodes "WriteRange",           "Write Range",               "smart", "DataTable"
    RenameNodes "ReadTextFile",         "Read Text File",            "smart", "Content"
    RenameNodes "ReadCsvFile",          "Read CSV",                  "smart", "DataTable"
    RenameNodes "WriteTextFile",        "Write Text File",           "smart", "Text"
    RenameNodes "FileExistsX",          "File Exists",               "smart", "Exists"
    RenameNodes "FolderExistsX",        "Folder Exists",             "smart", "Exists"
    RenameNodes "CreateDirectory",      "Create Folder",             "smart", "Path"
    RenameNodes "FormActivity",         "Create Form",               "smart", "FormFieldsOutputData"
    RenameNodes "GetQueueItem",         "Get Queue Item",            "smart", "TransactionItem"
    RenameNodes "TakeScreenshot",       "Take Screenshot",           "smart", "Screenshot"
    RenameNodes "SaveImage",            "Save Image",                "smart", "FileName"
    RenameNodes "IsMatch",              "Is Match",                  "attr",  "Result"
    RenameNodes "FindChildren",         "Find Children",             "attr",  "Children"
    RenameNodes "HttpClient",           "HTTP Request",              "attr",  "EndPoint"
    RenameCheckNodes
    RenameMultipleAssignNodes
    RenameInvokeWorkflowNodes
    RenameExcelDeleteRangeNodes
    RenameGetOutlookMailNodes
    RenameSendOutlookMailNodes
    RenameWriteCellNodes
    RenameSwitchNodes
End If

' 2. ARGUMENT DIRECTION PREFIX RENAMING
If OPT_ARGUMENT_RENAME Then
    RenameArguments
End If

' 3. VARIABLE TYPE PREFIX RENAMING
If OPT_VARIABLE_RENAME Then
    RenameVariables
End If

Dim outPath
outPath = Replace(WScript.Arguments(0), ".xaml", "_out.xaml")
x.Save outPath
WScript.Echo "Saved: " & outPath