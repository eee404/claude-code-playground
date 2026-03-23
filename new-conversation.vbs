Option Explicit

Dim fso, shell, scriptDir, envPath, conversationsDir, topic, slug, today, folderName, convPath, useGsd

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Resolve script directory
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Load .env or default to conversations/ next to script
conversationsDir = scriptDir & "\conversations"
envPath = scriptDir & "\.env"
If fso.FileExists(envPath) Then
    Dim envFile, line, parts
    Set envFile = fso.OpenTextFile(envPath, 1)
    Do While Not envFile.AtEndOfStream
        line = Trim(envFile.ReadLine)
        If Left(line, 18) = "CONVERSATIONS_DIR=" Then
            conversationsDir = Mid(line, 19)
            ' Resolve relative path
            If Mid(conversationsDir, 2, 1) <> ":" Then
                conversationsDir = scriptDir & "\" & conversationsDir
            End If
        End If
    Loop
    envFile.Close
End If

' Create conversations dir if needed
If Not fso.FolderExists(conversationsDir) Then
    fso.CreateFolder(conversationsDir)
End If

' Show dialog via temporary HTA
Dim htaPath, resultPath, htaFile, resultFile, resultContent, resultLines
htaPath = scriptDir & "\~new-conv-dialog.hta"
resultPath = scriptDir & "\~new-conv-result.tmp"

' Clean up any previous result file
If fso.FileExists(resultPath) Then fso.DeleteFile(resultPath)

Set htaFile = fso.CreateTextFile(htaPath, True)
htaFile.Write "<html>" & vbCrLf & _
    "<head>" & vbCrLf & _
    "<title>Nouvelle conversation Claude Code</title>" & vbCrLf & _
    "<HTA:APPLICATION ID=""oHTA"" BORDER=""thin"" BORDERSTYLE=""normal"" " & _
    "INNERBORDER=""no"" SCROLL=""no"" MAXIMIZEBUTTON=""no"" MINIMIZEBUTTON=""no"" " & _
    "SYSMENU=""yes"" SELECTION=""no"" SINGLEINSTANCE=""yes"" />" & vbCrLf & _
    "<style>" & vbCrLf & _
    "body { font-family: Segoe UI, sans-serif; font-size: 14px; padding: 20px; " & _
    "background: #f5f5f5; margin: 0; }" & vbCrLf & _
    "label { display: block; margin-bottom: 6px; font-weight: 600; }" & vbCrLf & _
    "input[type=text] { width: 100%; padding: 8px; font-size: 14px; " & _
    "border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }" & vbCrLf & _
    ".checkbox-row { margin-top: 14px; display: flex; align-items: center; gap: 8px; }" & vbCrLf & _
    ".checkbox-row input { width: 18px; height: 18px; margin: 0; }" & vbCrLf & _
    ".checkbox-row label { display: inline; font-weight: normal; margin: 0; }" & vbCrLf & _
    ".buttons { margin-top: 20px; text-align: right; }" & vbCrLf & _
    "button { padding: 8px 20px; font-size: 14px; border: 1px solid #999; " & _
    "border-radius: 4px; cursor: pointer; margin-left: 8px; }" & vbCrLf & _
    ".btn-ok { background: #0078d4; color: white; border-color: #0078d4; }" & vbCrLf & _
    ".btn-ok:hover { background: #106ebe; }" & vbCrLf & _
    ".btn-cancel { background: #e5e5e5; }" & vbCrLf & _
    ".btn-cancel:hover { background: #d0d0d0; }" & vbCrLf & _
    "</style>" & vbCrLf & _
    "<script language=""VBScript"">" & vbCrLf & _
    "Sub Window_OnLoad" & vbCrLf & _
    "  window.resizeTo 460, 250" & vbCrLf & _
    "  Dim sl, st" & vbCrLf & _
    "  sl = CInt((screen.width - 460) / 2)" & vbCrLf & _
    "  st = CInt((screen.height - 250) / 2)" & vbCrLf & _
    "  window.moveTo sl, st" & vbCrLf & _
    "  document.getElementById(""topic"").focus" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Sub DoOK" & vbCrLf & _
    "  Dim f, t, g" & vbCrLf & _
    "  t = document.getElementById(""topic"").value" & vbCrLf & _
    "  If document.getElementById(""gsd"").checked Then g = ""1"" Else g = ""0"" End If" & vbCrLf & _
    "  Set f = CreateObject(""Scripting.FileSystemObject"").CreateTextFile(""" & Replace(resultPath, "\", "\\") & """, True)" & vbCrLf & _
    "  f.WriteLine t" & vbCrLf & _
    "  f.WriteLine g" & vbCrLf & _
    "  f.Close" & vbCrLf & _
    "  self.close" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Sub DoCancel" & vbCrLf & _
    "  self.close" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Sub OnKeyPress" & vbCrLf & _
    "  If window.event.keyCode = 13 Then DoOK" & vbCrLf & _
    "  If window.event.keyCode = 27 Then DoCancel" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "</script>" & vbCrLf & _
    "</head>" & vbCrLf & _
    "<body onkeypress=""OnKeyPress"">" & vbCrLf & _
    "<label for=""topic"">Sujet de la conversation (laisser vide pour aucun) :</label>" & vbCrLf & _
    "<input type=""text"" id=""topic"" />" & vbCrLf & _
    "<div class=""checkbox-row"">" & vbCrLf & _
    "<input type=""checkbox"" id=""gsd"" />" & vbCrLf & _
    "<label for=""gsd"">Utiliser GSD (Get Shit Done)</label>" & vbCrLf & _
    "</div>" & vbCrLf & _
    "<div class=""buttons"">" & vbCrLf & _
    "<button class=""btn-cancel"" onclick=""DoCancel"">Annuler</button>" & vbCrLf & _
    "<button class=""btn-ok"" onclick=""DoOK"">OK</button>" & vbCrLf & _
    "</div>" & vbCrLf & _
    "</body></html>"
htaFile.Close

' Run HTA and wait for it to finish
shell.Run "mshta """ & htaPath & """", 1, True

' Clean up HTA file
If fso.FileExists(htaPath) Then fso.DeleteFile(htaPath)

' Read result (if user cancelled, no result file exists)
If Not fso.FileExists(resultPath) Then WScript.Quit

Set resultFile = fso.OpenTextFile(resultPath, 1)
topic = resultFile.ReadLine
useGsd = resultFile.ReadLine
resultFile.Close
fso.DeleteFile(resultPath)

' Build date
today = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)

' Find next index for today
Dim idx, prefix, folder
idx = 1
If fso.FolderExists(conversationsDir) Then
    Dim convFolder
    Set convFolder = fso.GetFolder(conversationsDir)
    For Each folder In convFolder.SubFolders
        If Left(folder.Name, Len(today)) = today Then
            ' Extract index from folder name (format: YYYY-MM-DD_NN or YYYY-MM-DD_NN_slug)
            Dim folderParts, folderIdx
            folderParts = Split(folder.Name, "_")
            ' folderParts(0)=YYYY-MM-DD, folderParts(1)=NN, folderParts(2...)=slug
            ' Date contains dashes not underscores, so folderParts(0) is the full date
            If UBound(folderParts) >= 1 Then
                If IsNumeric(folderParts(1)) Then
                    folderIdx = CInt(folderParts(1))
                    If folderIdx >= idx Then idx = folderIdx + 1
                End If
            End If
        End If
    Next
End If

Dim idxStr
idxStr = Right("00" & idx, 3)

' Slugify topic
If Len(Trim(topic)) = 0 Then
    folderName = today & "_" & idxStr
Else
    slug = LCase(Trim(topic))
    slug = Replace(slug, " ", "-")
    slug = Replace(slug, "'", "")
    slug = Replace(slug, ",", "")
    slug = Replace(slug, ".", "")
    slug = Replace(slug, "?", "")
    slug = Replace(slug, "!", "")
    slug = Replace(slug, "(", "")
    slug = Replace(slug, ")", "")
    slug = Replace(slug, "/", "-")
    slug = Replace(slug, "\", "-")
    slug = Replace(slug, """", "")
    ' Remove accents (basic)
    slug = Replace(slug, Chr(233), "e") ' é
    slug = Replace(slug, Chr(232), "e") ' è
    slug = Replace(slug, Chr(234), "e") ' ê
    slug = Replace(slug, Chr(224), "a") ' à
    slug = Replace(slug, Chr(226), "a") ' â
    slug = Replace(slug, Chr(231), "c") ' ç
    slug = Replace(slug, Chr(238), "i") ' î
    slug = Replace(slug, Chr(244), "o") ' ô
    slug = Replace(slug, Chr(249), "u") ' ù
    slug = Replace(slug, Chr(251), "u") ' û
    folderName = today & "_" & idxStr & "_" & slug
End If

convPath = conversationsDir & "\" & folderName

' Create folder
fso.CreateFolder(convPath)

' Copy CLAUDE.md template
If fso.FileExists(scriptDir & "\CLAUDE.md.template") Then
    fso.CopyFile scriptDir & "\CLAUDE.md.template", convPath & "\CLAUDE.md"
End If

' Create .vscode folder and tasks.json for auto-launching claude
Dim vscodePath
vscodePath = convPath & "\.vscode"
fso.CreateFolder(vscodePath)

Dim tasksFile
Set tasksFile = fso.CreateTextFile(vscodePath & "\tasks.json", True)
tasksFile.WriteLine "{"
tasksFile.WriteLine "  ""version"": ""2.0.0"","
tasksFile.WriteLine "  ""tasks"": ["
tasksFile.WriteLine "    {"
tasksFile.WriteLine "      ""label"": ""Start Claude Code"","
tasksFile.WriteLine "      ""type"": ""shell"","
tasksFile.WriteLine "      ""command"": ""claude"","
tasksFile.WriteLine "      ""isBackground"": true,"
tasksFile.WriteLine "      ""presentation"": {"
tasksFile.WriteLine "        ""reveal"": ""always"","
tasksFile.WriteLine "        ""panel"": ""dedicated"""
tasksFile.WriteLine "      },"
tasksFile.WriteLine "      ""problemMatcher"": [],"
tasksFile.WriteLine "      ""runOptions"": {"
tasksFile.WriteLine "        ""runOn"": ""folderOpen"""
tasksFile.WriteLine "      }"
tasksFile.WriteLine "    }"
tasksFile.WriteLine "  ]"
tasksFile.WriteLine "}"
tasksFile.Close

' Install GSD if requested
If useGsd = "1" Then
    shell.Run "cmd /c cd /d """ & convPath & """ && npx -y get-shit-done-cc@latest --claude --local", 1, True
End If

' Open VS Code
shell.Run """code"" """ & convPath & """", 0, False
