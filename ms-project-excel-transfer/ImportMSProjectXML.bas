Attribute VB_Name = "ImportMSProjectXML"
Option Explicit

' ============================================================================
' Module : ImportMSProjectXML
' Objet  : Importer un export XML MS Project et reconstruire un Gantt dans
'          Excel avec le meme look que l'outil HTML fourni (memes couleurs,
'          memes regles). Fonctionne 100% hors ligne (MSXML est fourni en
'          standard avec Windows, aucune connexion ni add-in requis).
'
' Utilisation :
'   1. Alt+F11 > Fichier > Importer un fichier... > ImportMSProjectXML.bas
'   2. Alt+F8 > ImporterProjetMSProject > Executer
'   3. Choisir le fichier .xml exporte depuis MS Project
'      (MS Project : Fichier > Enregistrer sous > type "XML")
' ============================================================================

Private Type TTask
    id As String
    uid As String
    Name As String
    OutlineLevel As Long
    IsSummary As Boolean
    IsMilestone As Boolean
    Critical As Boolean
    HasDates As Boolean
    StartDate As Date
    FinishDate As Date
    DurationDays As Double
    pct As Double
    predText As String
End Type

Private Const COL_HEADER_BG As Long = &H64381F ' BGR de 1F3864
Private Const COL_HEADER2_BG As Long = &HF3E2D9 ' BGR de D9E2F3
Private Const COL_COMPLETE As Long = &H794E1F   ' BGR de 1F4E79
Private Const COL_REMAINING As Long = &HE6C39D  ' BGR de 9DC3E6
Private Const COL_SUMMARY As Long = &H404040
Private Const COL_MILESTONE As Long = &H000000
Private Const COL_CRITICAL As Long = &H0000C0   ' BGR de C00000
Private Const COL_WEEKEND As Long = &HF2F2F2
Private Const N_LEFT_COLS As Long = 7

Sub ImporterProjetMSProject()

    Dim filePathVar As Variant
    filePathVar = Application.GetOpenFilename("Export XML MS Project (*.xml), *.xml", , "Choisir l'export XML MS Project")
    If filePathVar = False Then Exit Sub
    Dim filePath As String
    filePath = CStr(filePathVar)

    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    If Not doc.Load(filePath) Then
        MsgBox "Fichier XML invalide : " & doc.parseError.reason, vbCritical
        Exit Sub
    End If

    Dim root As Object
    Set root = doc.documentElement
    If root Is Nothing Then
        MsgBox "Fichier XML illisible.", vbExclamation
        Exit Sub
    End If
    If InStr(1, root.nodeName, "Project", vbTextCompare) = 0 Then
        MsgBox "Ce fichier ne ressemble pas a un export XML MS Project.", vbExclamation
        Exit Sub
    End If

    Dim projectTitle As String
    projectTitle = DirectChildText(root, "Title")
    If projectTitle = "" Then projectTitle = DirectChildText(root, "Name")
    If projectTitle = "" Then projectTitle = "Projet"

    Dim tasksRoot As Object
    Set tasksRoot = FirstDirectChild(root, "Tasks")
    If tasksRoot Is Nothing Then
        MsgBox "Aucune tache trouvee (balise <Tasks> absente).", vbExclamation
        Exit Sub
    End If

    ' --- Lecture des taches ---
    Dim rawTasks As Object, taskNode As Object
    Dim uidToId As Object
    Set uidToId = CreateObject("Scripting.Dictionary")

    Dim n As Long
    Dim tmp() As TTask
    Dim count As Long: count = 0
    ReDim tmp(0 To tasksRoot.childNodes.length - 1)

    For n = 0 To tasksRoot.childNodes.length - 1
        Set taskNode = tasksRoot.childNodes(n)
        If taskNode.nodeType = 1 And InStr(taskNode.nodeName, "Task") > 0 Then
            If DirectChildText(taskNode, "IsNull") <> "1" And DirectChildText(taskNode, "Name") <> "" Then
                Dim t As TTask
                t.uid = DirectChildText(taskNode, "UID")
                t.id = DirectChildText(taskNode, "ID")
                t.Name = DirectChildText(taskNode, "Name")
                t.OutlineLevel = CLng(NzNum(DirectChildText(taskNode, "OutlineLevel"), 1))
                t.IsSummary = (DirectChildText(taskNode, "Summary") = "1")
                t.IsMilestone = (DirectChildText(taskNode, "Milestone") = "1")
                t.Critical = (DirectChildText(taskNode, "Critical") = "1")

                Dim startRaw As String, finishRaw As String
                startRaw = DirectChildText(taskNode, "Start")
                finishRaw = DirectChildText(taskNode, "Finish")
                t.HasDates = (startRaw <> "" And finishRaw <> "")
                If t.HasDates Then
                    t.StartDate = ParseIsoDate(startRaw)
                    t.FinishDate = ParseIsoDate(finishRaw)
                End If
                t.DurationDays = ParseDurationToDays(DirectChildText(taskNode, "Duration"))
                t.pct = NzNum(DirectChildText(taskNode, "PercentComplete"), 0)

                ' predecesseurs
                Dim preds As String, k As Long, predNode As Object
                preds = ""
                For k = 0 To taskNode.childNodes.length - 1
                    Set predNode = taskNode.childNodes(k)
                    If predNode.nodeType = 1 And InStr(predNode.nodeName, "PredecessorLink") > 0 Then
                        Dim pu As String
                        pu = DirectChildText(predNode, "PredecessorUID")
                        If pu <> "" Then
                            If preds <> "" Then preds = preds & ";"
                            preds = preds & pu
                        End If
                    End If
                Next k
                t.predText = preds ' UID->ID resolus dans une 2e passe

                If uidToId.Exists(t.uid) = False Then uidToId.Add t.uid, t.id
                tmp(count) = t
                count = count + 1
            End If
        End If
    Next n

    If count = 0 Then
        MsgBox "Aucune tache exploitable trouvee dans le fichier.", vbExclamation
        Exit Sub
    End If

    ReDim Preserve tmp(0 To count - 1)

    ' 2e passe : resoudre UID -> ID pour les predecesseurs (affichage lisible)
    Dim i As Long
    For i = 0 To count - 1
        If tmp(i).predText <> "" Then
            Dim parts() As String, p As Long, resolved As String
            parts = Split(tmp(i).predText, ";")
            resolved = ""
            For p = LBound(parts) To UBound(parts)
                If uidToId.Exists(parts(p)) Then
                    If resolved <> "" Then resolved = resolved & ";"
                    resolved = resolved & uidToId(parts(p))
                End If
            Next p
            tmp(i).predText = resolved
        End If
    Next i

    ' --- Bornes du projet ---
    Dim projStart As Date, projFinish As Date, first As Boolean
    first = True
    For i = 0 To count - 1
        If tmp(i).HasDates Then
            If first Then
                projStart = tmp(i).StartDate
                projFinish = tmp(i).FinishDate
                first = False
            Else
                If tmp(i).StartDate < projStart Then projStart = tmp(i).StartDate
                If tmp(i).FinishDate > projFinish Then projFinish = tmp(i).FinishDate
            End If
        End If
    Next i
    If first Then
        MsgBox "Aucune tache avec des dates exploitables.", vbExclamation
        Exit Sub
    End If

    ' --- Granularite (jour / semaine / mois), plafonnee a ~120-130 colonnes ---
    Dim totalDays As Long
    totalDays = DateDiff("d", projStart, projFinish) + 1
    Dim granularity As String
    If totalDays <= 120 Then
        granularity = "day"
    ElseIf totalDays <= 120 * 7 Then
        granularity = "week"
    Else
        granularity = "month"
    End If

    Dim highlightCritical As Boolean, shadeWeekend As Boolean
    highlightCritical = (MsgBox("Mettre en evidence le chemin critique en rouge ?", vbYesNo + vbQuestion, "Options") = vbYes)
    shadeWeekend = (MsgBox("Griser les week-ends ?", vbYesNo + vbQuestion, "Options") = vbYes)

    BuildGanttSheet tmp, count, projectTitle, projStart, projFinish, granularity, highlightCritical, shadeWeekend

    MsgBox "Import termine : " & count & " taches (" & Format(projStart, "dd/mm/yyyy") & " -> " & Format(projFinish, "dd/mm/yyyy") & ").", vbInformation

End Sub

' ============================================================================
' Construction de la feuille "Gantt"
' ============================================================================
Private Sub BuildGanttSheet(tasks() As TTask, taskCount As Long, projectTitle As String, _
                             projStart As Date, projFinish As Date, granularity As String, _
                             highlightCritical As Boolean, shadeWeekend As Boolean)

    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Gantt").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Gantt"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- Construction des periodes ---
    Dim periods() As Variant ' chaque element: Array(startDate, endDate, label1, label2, isWeekend)
    Dim periodCount As Long
    BuildPeriods projStart, projFinish, granularity, periods, periodCount

    Dim leftTitles(1 To N_LEFT_COLS) As String
    leftTitles(1) = "N°": leftTitles(2) = "Nom de la tache": leftTitles(3) = "Duree"
    leftTitles(4) = "Debut": leftTitles(5) = "Fin": leftTitles(6) = "% acheve": leftTitles(7) = "Pred."
    Dim leftWidths(1 To N_LEFT_COLS) As Double
    leftWidths(1) = 6: leftWidths(2) = 34: leftWidths(3) = 9
    leftWidths(4) = 11: leftWidths(5) = 11: leftWidths(6) = 9: leftWidths(7) = 10

    Dim c As Long
    For c = 1 To N_LEFT_COLS
        ws.Cells(1, c).Value = leftTitles(c)
        ws.Cells(2, c).Value = ""
        ws.Range(ws.Cells(1, c), ws.Cells(2, c)).Merge
        ws.Cells(1, c).Interior.Color = COL_HEADER_BG
        ws.Cells(1, c).Font.Color = RGB(255, 255, 255)
        ws.Cells(1, c).Font.Bold = True
        ws.Cells(1, c).HorizontalAlignment = xlCenter
        ws.Cells(1, c).VerticalAlignment = xlCenter
        ws.Columns(c).ColumnWidth = leftWidths(c)
        ' reecrit le titre sur la ligne 2 (fusion visuelle: le texte reste sur la cellule superieure gauche)
        ws.Cells(1, c).Value = leftTitles(c)
    Next c

    ' --- En-tetes de periode (mois/annee fusionnes + jour/semaine/mois) ---
    Dim tlWidth As Double
    Select Case granularity
        Case "day": tlWidth = 3.2
        Case "week": tlWidth = 6
        Case Else: tlWidth = 9
    End Select

    Dim startCol As Long: startCol = N_LEFT_COLS + 1
    Dim runStart As Long: runStart = 0
    Dim pIdx As Long
    For pIdx = 0 To periodCount - 1
        Dim col As Long: col = startCol + pIdx
        ws.Columns(col).ColumnWidth = tlWidth
        ws.Cells(2, col).Value = periods(pIdx)(3) ' label2
        ws.Cells(2, col).Interior.Color = COL_HEADER2_BG
        ws.Cells(2, col).Font.Bold = True
        ws.Cells(2, col).HorizontalAlignment = xlCenter

        If pIdx = 0 Or periods(pIdx)(2) <> periods(pIdx - 1)(2) Then
            If pIdx > 0 Then
                ws.Range(ws.Cells(1, startCol + runStart), ws.Cells(1, startCol + pIdx - 1)).Merge
            End If
            runStart = pIdx
        End If
        ws.Cells(1, col).Value = periods(pIdx)(2) ' label1 (repete, sera fusionne)
        ws.Cells(1, col).Interior.Color = COL_HEADER_BG
        ws.Cells(1, col).Font.Color = RGB(255, 255, 255)
        ws.Cells(1, col).Font.Bold = True
        ws.Cells(1, col).HorizontalAlignment = xlCenter
    Next pIdx
    If periodCount > 0 Then
        ws.Range(ws.Cells(1, startCol + runStart), ws.Cells(1, startCol + periodCount - 1)).Merge
    End If

    ' --- Lignes de taches ---
    Dim r As Long
    For r = 0 To taskCount - 1
        Dim rowIdx As Long: rowIdx = r + 3
        Dim t As TTask: t = tasks(r)

        ws.Cells(rowIdx, 1).Value = t.id
        ws.Cells(rowIdx, 1).HorizontalAlignment = xlCenter
        ws.Cells(rowIdx, 2).Value = t.Name
        ws.Cells(rowIdx, 2).IndentLevel = t.OutlineLevel
        ws.Cells(rowIdx, 3).Value = IIf(t.IsMilestone, "0 j", Format(t.DurationDays, "0.#") & " j")
        ws.Cells(rowIdx, 3).HorizontalAlignment = xlCenter
        If t.HasDates Then
            ws.Cells(rowIdx, 4).Value = Format(t.StartDate, "dd/mm/yyyy")
            ws.Cells(rowIdx, 5).Value = Format(t.FinishDate, "dd/mm/yyyy")
        End If
        ws.Cells(rowIdx, 4).HorizontalAlignment = xlCenter
        ws.Cells(rowIdx, 5).HorizontalAlignment = xlCenter
        ws.Cells(rowIdx, 6).Value = t.pct & "%"
        ws.Cells(rowIdx, 6).HorizontalAlignment = xlCenter
        ws.Cells(rowIdx, 7).Value = t.predText
        ws.Cells(rowIdx, 7).HorizontalAlignment = xlCenter

        If t.IsSummary Then
            ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, N_LEFT_COLS)).Font.Bold = True
            ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, N_LEFT_COLS)).Interior.Color = COL_HEADER2_BG
        End If

        Dim completedUntil As Date
        If t.HasDates Then
            completedUntil = t.StartDate + (t.FinishDate - t.StartDate) * (t.pct / 100)
        End If

        For pIdx = 0 To periodCount - 1
            Dim cell As Range
            Set cell = ws.Cells(rowIdx, startCol + pIdx)
            Dim pStart As Date, pEnd As Date, pIsWeekend As Boolean
            pStart = periods(pIdx)(0): pEnd = periods(pIdx)(1): pIsWeekend = periods(pIdx)(4)

            If Not t.HasDates Then
                ' rien
            ElseIf pEnd < t.StartDate Or pStart > t.FinishDate Then
                If pIsWeekend And shadeWeekend Then cell.Interior.Color = COL_WEEKEND
            ElseIf t.IsMilestone Then
                cell.Interior.Color = COL_MILESTONE
                cell.Value = ChrW(9670) ' losange
                cell.Font.Color = RGB(255, 255, 255)
                cell.HorizontalAlignment = xlCenter
            ElseIf t.IsSummary Then
                cell.Interior.Color = COL_SUMMARY
            ElseIf highlightCritical And t.Critical Then
                cell.Interior.Color = COL_CRITICAL
            ElseIf pEnd <= completedUntil Then
                cell.Interior.Color = COL_COMPLETE
            Else
                cell.Interior.Color = COL_REMAINING
            End If
        Next pIdx
    Next r

    ' --- Mise en forme generale ---
    Dim lastRow As Long: lastRow = taskCount + 2
    Dim lastCol As Long: lastCol = startCol + periodCount - 1
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(201, 205, 211)
        .Borders.Weight = xlThin
    End With
    ws.Rows(1).RowHeight = 18
    ws.Rows(2).RowHeight = 16

    ws.Activate
    ws.Cells(3, startCol).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

' ============================================================================
' Construction des periodes (jour / semaine / mois)
' ============================================================================
Private Sub BuildPeriods(projStart As Date, projFinish As Date, granularity As String, _
                          ByRef periods() As Variant, ByRef periodCount As Long)
    Dim moisNoms As Variant
    moisNoms = Array("Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre")
    Dim moisCourt As Variant
    moisCourt = Array("Jan", "Fev", "Mar", "Avr", "Mai", "Jun", "Jul", "Aou", "Sep", "Oct", "Nov", "Dec")

    Dim list As Collection
    Set list = New Collection

    Select Case granularity
        Case "day"
            Dim d As Date
            d = projStart
            Do While d <= projFinish
                Dim isWknd As Boolean
                isWknd = (Weekday(d, vbMonday) >= 6)
                list.Add Array(d, d, moisNoms(Month(d) - 1) & " " & Year(d), CStr(Day(d)), isWknd)
                d = d + 1
            Loop
        Case "week"
            Dim wStart As Date
            wStart = projStart - (Weekday(projStart, vbMonday) - 1)
            Do While wStart <= projFinish
                Dim wEnd As Date: wEnd = wStart + 6
                list.Add Array(wStart, wEnd, moisNoms(Month(wStart) - 1) & " " & Year(wStart), "S" & DatePart("ww", wStart, vbMonday, vbFirstFourDays), False)
                wStart = wStart + 7
            Loop
        Case Else
            Dim mStart As Date
            mStart = DateSerial(Year(projStart), Month(projStart), 1)
            Do While mStart <= projFinish
                Dim mEnd As Date: mEnd = DateSerial(Year(mStart), Month(mStart) + 1, 0)
                list.Add Array(mStart, mEnd, CStr(Year(mStart)), moisCourt(Month(mStart) - 1), False)
                mStart = DateSerial(Year(mStart), Month(mStart) + 1, 1)
            Loop
    End Select

    periodCount = list.count
    ReDim periods(0 To periodCount - 1)
    Dim i As Long
    For i = 1 To list.count
        periods(i - 1) = list(i)
    Next i
End Sub

' ============================================================================
' Utilitaires XML / parsing
' ============================================================================
Private Function FirstDirectChild(parent As Object, tag As String) As Object
    Dim c As Object
    For Each c In parent.childNodes
        If c.nodeType = 1 And InStr(c.nodeName, tag) > 0 Then
            Set FirstDirectChild = c
            Exit Function
        End If
    Next c
    Set FirstDirectChild = Nothing
End Function

Private Function DirectChildText(parent As Object, tag As String) As String
    Dim c As Object
    Set c = FirstDirectChild(parent, tag)
    If c Is Nothing Then
        DirectChildText = ""
    Else
        DirectChildText = Trim(c.Text)
    End If
End Function

Private Function NzNum(s As String, defaultVal As Double) As Double
    If s = "" Then
        NzNum = defaultVal
    Else
        On Error Resume Next
        NzNum = CDbl(s)
        On Error GoTo 0
    End If
End Function

' Convertit une date ISO "2026-01-05T08:00:00" (independante des parametres regionaux)
Private Function ParseIsoDate(s As String) As Date
    On Error GoTo fallback
    Dim y As Integer, mo As Integer, da As Integer, h As Integer, mi As Integer, se As Integer
    y = CInt(Mid(s, 1, 4))
    mo = CInt(Mid(s, 6, 2))
    da = CInt(Mid(s, 9, 2))
    h = 0: mi = 0: se = 0
    If Len(s) >= 19 Then
        h = CInt(Mid(s, 12, 2))
        mi = CInt(Mid(s, 15, 2))
        se = CInt(Mid(s, 18, 2))
    End If
    ParseIsoDate = DateSerial(y, mo, da) + TimeSerial(h, mi, se)
    Exit Function
fallback:
    ParseIsoDate = CDate(s)
End Function

' Convertit une duree MS Project ("PT40H0M0S", "P2DT8H0M0S", ...) en jours (8h = 1 jour)
Private Function ParseDurationToDays(s As String) As Double
    If s = "" Then
        ParseDurationToDays = 0
        Exit Function
    End If
    Dim days As Double, hours As Double, minutes As Double, seconds As Double
    days = 0: hours = 0: minutes = 0: seconds = 0

    Dim posT As Long
    posT = InStr(s, "T")
    Dim datePart As String, timePart As String
    If posT > 0 Then
        datePart = Mid(s, 2, posT - 2) ' entre "P" et "T"
        timePart = Mid(s, posT + 1)
    Else
        datePart = Mid(s, 2)
        timePart = ""
    End If

    Dim posD As Long
    posD = InStr(datePart, "D")
    If posD > 0 Then days = Val(Left(datePart, posD - 1))

    Dim rest As String: rest = timePart
    Dim posH As Long, posM As Long, posS As Long
    posH = InStr(rest, "H")
    If posH > 0 Then
        hours = Val(Left(rest, posH - 1))
        rest = Mid(rest, posH + 1)
    End If
    posM = InStr(rest, "M")
    If posM > 0 Then
        minutes = Val(Left(rest, posM - 1))
        rest = Mid(rest, posM + 1)
    End If
    posS = InStr(rest, "S")
    If posS > 0 Then
        seconds = Val(Left(rest, posS - 1))
    End If

    Dim totalHours As Double
    totalHours = days * 24 + hours + minutes / 60 + seconds / 3600
    ParseDurationToDays = Int((totalHours / 8) * 10 + 0.5) / 10
End Function
