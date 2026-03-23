Option Explicit

Dim fso, shell, scriptDir, envPath, conversationsDir, topic, slug, today, folderName, convPath

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

' Ask for topic
topic = InputBox("Sujet de la conversation (laisser vide pour aucun) :", "Nouvelle conversation Claude Code", "")

' User cancelled
If IsEmpty(topic) Then WScript.Quit

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
            Dim parts, folderIdx
            parts = Split(folder.Name, "_")
            ' parts(0)=YYYY-MM-DD, parts(1)=NN, parts(2...)=slug
            ' Date contains dashes not underscores, so parts(0) is the full date
            If UBound(parts) >= 1 Then
                If IsNumeric(parts(1)) Then
                    folderIdx = CInt(parts(1))
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

' Open VS Code
shell.Run """code"" """ & convPath & """", 0, False
