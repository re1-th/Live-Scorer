Attribute VB_Name = "LiveScorer"
'==============================================================================
' KC Live Scorer v3.0
' This one module replaces the four v2 modules:
' Scoreboard, Validator, SyncTeamNames and Duplicates.
'
' HOW IT WORKS
'   * Scores are stored once, inside the presentation file (as hidden tags).
'     The text boxes only show them, so copies can never drift apart.
'   * Any shape named C1..C10 shows that team's score. Any shape named AA..JJ
'     shows that team's name. The shape can be on any slide or in the slide
'     master / a layout, inside a group or not, and a slide can have any
'     number of copies, or none at all.
'   * Team names come from the Setup slide: the slide that has a shape named
'     "LS_Setup". If no slide has one, slide 2 is used.
'   * Score buttons: name the shape "Score_<team number>" (for example Score_3)
'     and set Insert > Action > Run macro > ScoreButton. The points are read from
'     the button's own text, so a button labelled -5, 5, +10 or 20 just works.
'   * Every click refreshes every copy on every slide. Nothing needs to be
'     validated, synced or de-duplicated by hand.
'
' BUTTON MACROS (Insert > Action > Run macro)
'   ScoreButton     add the points on the button to its team
'   StartNewQuiz    asks first, then sets every score to 0 and updates names
'   ResetScores     same as StartNewQuiz (handy to copy onto round slides)
'   SyncTeamNames   update team names everywhere (also happens on every click)
'   UndoLastScore   take back the last score change
'   EditScore       put this on a C1..C10 box: click it to type a score in
'==============================================================================
Option Explicit

Private Const MAX_TEAMS As Long = 10
Private Const SETUP_MARKER As String = "LS_Setup"
Private Const TAG_SCORE As String = "LS_SCORE_"
Private Const TAG_HISTORY As String = "LS_HISTORY"
Private Const MAX_HISTORY As Long = 100
Private Const APP_TITLE As String = "Live Scorer"
Private Const REDRAW_IN_SHOW As Boolean = True

'==============================================================================
' Button macros
'==============================================================================

Public Sub ScoreButton(oShp As Shape)
    Dim team As Long
    Dim label As String

    team = TeamFromButtonName(oShp.Name)
    If team = 0 Then
        MsgBox "This button isn't linked to a team." & vbCrLf & vbCrLf & _
               "Open the Selection Pane and rename it Score_1 to Score_" & MAX_TEAMS & ".", _
               vbExclamation, APP_TITLE
        Exit Sub
    End If

    If oShp.HasTextFrame Then label = oShp.TextFrame.TextRange.Text
    label = Replace(Replace(label, " ", ""), ChrW(8722), "-")   ' allow a typographic minus sign
    If Len(label) = 0 Then
        MsgBox "This button has no points on it. Type a number such as 10 or -5 on it.", _
               vbExclamation, APP_TITLE
        Exit Sub
    End If

    AddPoints team, Val(label), True
End Sub

Public Sub StartNewQuiz()
    If MsgBox("Start a new quiz?" & vbCrLf & vbCrLf & "All scores will be set to 0.", _
              vbYesNo + vbQuestion + vbDefaultButton2, APP_TITLE) <> vbYes Then Exit Sub
    ClearScores
End Sub

Public Sub ResetScores()
    If MsgBox("Set every team's score back to 0?", _
              vbYesNo + vbQuestion + vbDefaultButton2, APP_TITLE) <> vbYes Then Exit Sub
    ClearScores
End Sub

Public Sub SyncTeamNames()
    RefreshAll
End Sub

Public Sub UndoLastScore()
    Dim hist As String, lastEntry As String
    Dim cut As Long, parts() As String

    hist = ActivePresentation.Tags(TAG_HISTORY)
    If Len(hist) = 0 Then
        MsgBox "There is nothing to undo.", vbInformation, APP_TITLE
        Exit Sub
    End If

    cut = InStrRev(hist, "|")
    lastEntry = Mid$(hist, cut + 1)
    If cut > 0 Then hist = Left$(hist, cut - 1) Else hist = ""
    ActivePresentation.Tags.Add TAG_HISTORY, hist

    parts = Split(lastEntry, ":")
    If UBound(parts) = 1 Then AddPoints CLng(parts(0)), -Val(parts(1)), False
End Sub

Public Sub EditScore(oShp As Shape)
    Dim team As Long, answer As String

    team = ScoreBoxTeam(oShp.Name)
    If team = 0 Then Exit Sub

    answer = InputBox("New score for " & TeamName(team) & ":", APP_TITLE, FormatScore(GetScore(team)))
    If StrPtr(answer) = 0 Or Len(Trim$(answer)) = 0 Then Exit Sub   ' Cancel or empty

    SetScore team, Val(Replace(answer, ",", "."))
    RefreshAll
End Sub

'==============================================================================
' Score storage (presentation tags)
'==============================================================================

Private Sub AddPoints(ByVal team As Long, ByVal pts As Double, ByVal remember As Boolean)
    If team < 1 Or team > MAX_TEAMS Then Exit Sub
    SetScore team, GetScore(team) + pts
    If remember Then PushHistory team, pts
    RefreshAll
End Sub

Private Sub ClearScores()
    Dim t As Long
    For t = 1 To MAX_TEAMS
        SetScore t, 0
    Next t
    ActivePresentation.Tags.Add TAG_HISTORY, ""
    RefreshAll
End Sub

Private Function GetScore(ByVal team As Long) As Double
    Dim v As String
    v = ActivePresentation.Tags(TAG_SCORE & team)
    If Len(v) = 0 Then
        ' First run in a deck made with v2: carry on from the score on screen.
        GetScore = Val(ReadSetupText("C" & team))
    Else
        GetScore = Val(v)
    End If
End Function

Private Sub SetScore(ByVal team As Long, ByVal value As Double)
    ActivePresentation.Tags.Add TAG_SCORE & team, Trim$(Str$(value))   ' Str$ always uses "."
End Sub

Private Sub PushHistory(ByVal team As Long, ByVal pts As Double)
    Dim hist As String, entry As String
    entry = team & ":" & Trim$(Str$(pts))
    hist = ActivePresentation.Tags(TAG_HISTORY)
    If Len(hist) = 0 Then hist = entry Else hist = hist & "|" & entry
    Do While UBound(Split(hist, "|")) >= MAX_HISTORY
        hist = Mid$(hist, InStr(hist, "|") + 1)
    Loop
    ActivePresentation.Tags.Add TAG_HISTORY, hist
End Sub

Private Function FormatScore(ByVal value As Double) As String
    FormatScore = CStr(value)
End Function

'==============================================================================
' Refreshing every scoreboard copy
'==============================================================================

Public Sub RefreshAll()
    Dim scores(1 To MAX_TEAMS) As String
    Dim names(1 To MAX_TEAMS) As String
    Dim haveName(1 To MAX_TEAMS) As Boolean
    Dim setup As Slide, sld As Slide, shp As Shape
    Dim t As Long

    Set setup = GetSetupSlide()
    For t = 1 To MAX_TEAMS
        scores(t) = FormatScore(GetScore(t))
        If Not setup Is Nothing Then
            haveName(t) = FindText(setup.Shapes, NameBoxName(t), names(t))
        End If
    Next t

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            ApplyToShape shp, scores, names, haveName
        Next shp
    Next sld

    ' Scoreboards that live in the slide master / layouts (View > Slide Master)
    Dim des As Design, lay As CustomLayout
    For Each des In ActivePresentation.Designs
        For Each shp In des.SlideMaster.Shapes
            ApplyToShape shp, scores, names, haveName
        Next shp
        For Each lay In des.SlideMaster.CustomLayouts
            For Each shp In lay.Shapes
                ApplyToShape shp, scores, names, haveName
            Next shp
        Next lay
    Next des

    RedrawShow
End Sub

' PowerPoint doesn't always repaint layout shapes mid-show. If the slide on
' screen gets its scoreboard from its layout, re-show the same slide without
' resetting its animations. Set REDRAW_IN_SHOW to False if this causes a flicker.
Private Sub RedrawShow()
    If Not REDRAW_IN_SHOW Then Exit Sub
    If SlideShowWindows.Count = 0 Then Exit Sub
    On Error Resume Next
    With SlideShowWindows(1).View
        If .State <> ppSlideShowRunning Then Exit Sub
        If Not LayoutHasScoreboard(.Slide.CustomLayout) Then Exit Sub
        .GotoSlide .CurrentShowPosition, msoFalse
    End With
End Sub

Private Function LayoutHasScoreboard(lay As CustomLayout) As Boolean
    Dim shp As Shape
    For Each shp In lay.Shapes
        If HasScoreBox(shp) Then LayoutHasScoreboard = True: Exit Function
    Next shp
End Function

Private Function HasScoreBox(shp As Shape) As Boolean
    Dim child As Shape
    If shp.Type = msoGroup Then
        For Each child In shp.GroupItems
            If ScoreBoxTeam(child.Name) > 0 Then HasScoreBox = True: Exit Function
        Next child
    Else
        HasScoreBox = ScoreBoxTeam(shp.Name) > 0
    End If
End Function

Private Sub ApplyToShape(shp As Shape, scores() As String, names() As String, haveName() As Boolean)
    Dim child As Shape, t As Long

    If shp.Type = msoGroup Then
        For Each child In shp.GroupItems
            ApplyToShape child, scores, names, haveName
        Next child
        Exit Sub
    End If

    t = ScoreBoxTeam(shp.Name)
    If t > 0 Then
        SetTextIfChanged shp, scores(t)
        Exit Sub
    End If

    t = NameBoxTeam(shp.Name)
    If t > 0 Then
        If haveName(t) Then SetTextIfChanged shp, names(t)
    End If
End Sub

Private Sub SetTextIfChanged(shp As Shape, ByVal newText As String)
    If Not shp.HasTextFrame Then Exit Sub
    With shp.TextFrame.TextRange
        If .Text <> newText Then .Text = newText
    End With
End Sub

'==============================================================================
' Lookups
'==============================================================================

Private Function GetSetupSlide() As Slide
    Dim sld As Slide, shp As Shape
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Name = SETUP_MARKER Then
                Set GetSetupSlide = sld
                Exit Function
            End If
        Next shp
    Next sld
    If ActivePresentation.Slides.Count >= 2 Then Set GetSetupSlide = ActivePresentation.Slides(2)
End Function

' Finds the first shape called shapeName (also inside groups) and returns its text.
Private Function FindText(shps As Object, ByVal shapeName As String, ByRef result As String) As Boolean
    Dim shp As Shape
    For Each shp In shps
        If shp.Type = msoGroup Then
            If FindText(shp.GroupItems, shapeName, result) Then FindText = True: Exit Function
        ElseIf shp.Name = shapeName Then
            If shp.HasTextFrame Then
                result = shp.TextFrame.TextRange.Text
                FindText = True
                Exit Function
            End If
        End If
    Next shp
End Function

Private Function ReadSetupText(ByVal shapeName As String) As String
    Dim setup As Slide, txt As String
    Set setup = GetSetupSlide()
    If setup Is Nothing Then Exit Function
    If FindText(setup.Shapes, shapeName, txt) Then ReadSetupText = txt
End Function

Private Function TeamName(ByVal team As Long) As String
    TeamName = ReadSetupText(NameBoxName(team))
    If Len(TeamName) = 0 Then TeamName = "Team " & team
End Function

' "C1".."C10" -> 1..10, anything else -> 0
Private Function ScoreBoxTeam(ByVal nm As String) As Long
    Dim digits As String
    If Len(nm) < 2 Or Len(nm) > 3 Then Exit Function
    If Left$(nm, 1) <> "C" Then Exit Function
    digits = Mid$(nm, 2)
    If digits Like "#" Or digits Like "##" Then
        If CLng(digits) >= 1 And CLng(digits) <= MAX_TEAMS Then ScoreBoxTeam = CLng(digits)
    End If
End Function

' "AA".."JJ" -> 1..10, anything else -> 0
Private Function NameBoxTeam(ByVal nm As String) As Long
    Dim n As Long
    If Len(nm) <> 2 Then Exit Function
    If Left$(nm, 1) <> Right$(nm, 1) Then Exit Function
    n = Asc(Left$(nm, 1)) - 64
    If n >= 1 And n <= MAX_TEAMS Then NameBoxTeam = n
End Function

Private Function NameBoxName(ByVal team As Long) As String
    NameBoxName = String$(2, Chr$(64 + team))
End Function

' "Score_3" -> 3, anything else -> 0
Private Function TeamFromButtonName(ByVal nm As String) As Long
    Dim n As Long
    If LCase$(Left$(nm, 6)) <> "score_" Then Exit Function
    n = Val(Mid$(nm, 7))
    If n >= 1 And n <= MAX_TEAMS Then TeamFromButtonName = n
End Function

'==============================================================================
' v2 compatibility
' Buttons in decks built with v2 still call these names, so importing this
' module into an older quiz keeps every old button working.
'==============================================================================

Public Sub counterReset(): ResetScores: End Sub

Public Sub AddScoreboardToSlides()
    RefreshAll
    MsgBox "This step isn't needed any more. Scoreboards are updated automatically.", vbInformation, APP_TITLE
End Sub

Public Sub CheckDuplicateScoreboards()
    RefreshAll
    MsgBox "This step isn't needed any more. A slide can have any number of scoreboards.", vbInformation, APP_TITLE
End Sub

Public Sub counter1NEG(): AddPoints 1, -5, True: End Sub
Public Sub counter1PART(): AddPoints 1, 5, True: End Sub
Public Sub counter1CORRECT(): AddPoints 1, 10, True: End Sub
Public Sub counter2NEG(): AddPoints 2, -5, True: End Sub
Public Sub counter2PART(): AddPoints 2, 5, True: End Sub
Public Sub counter2CORRECT(): AddPoints 2, 10, True: End Sub
Public Sub counter3NEG(): AddPoints 3, -5, True: End Sub
Public Sub counter3PART(): AddPoints 3, 5, True: End Sub
Public Sub counter3CORRECT(): AddPoints 3, 10, True: End Sub
Public Sub counter4NEG(): AddPoints 4, -5, True: End Sub
Public Sub counter4PART(): AddPoints 4, 5, True: End Sub
Public Sub counter4CORRECT(): AddPoints 4, 10, True: End Sub
Public Sub counter5NEG(): AddPoints 5, -5, True: End Sub
Public Sub counter5PART(): AddPoints 5, 5, True: End Sub
Public Sub counter5CORRECT(): AddPoints 5, 10, True: End Sub
Public Sub counter6NEG(): AddPoints 6, -5, True: End Sub
Public Sub counter6PART(): AddPoints 6, 5, True: End Sub
Public Sub counter6CORRECT(): AddPoints 6, 10, True: End Sub
Public Sub counter7NEG(): AddPoints 7, -5, True: End Sub
Public Sub counter7PART(): AddPoints 7, 5, True: End Sub
Public Sub counter7CORRECT(): AddPoints 7, 10, True: End Sub
Public Sub counter8NEG(): AddPoints 8, -5, True: End Sub
Public Sub counter8PART(): AddPoints 8, 5, True: End Sub
Public Sub counter8CORRECT(): AddPoints 8, 10, True: End Sub
Public Sub counter9NEG(): AddPoints 9, -5, True: End Sub
Public Sub counter9PART(): AddPoints 9, 5, True: End Sub
Public Sub counter9CORRECT(): AddPoints 9, 10, True: End Sub
Public Sub counter10NEG(): AddPoints 10, -5, True: End Sub
Public Sub counter10PART(): AddPoints 10, 5, True: End Sub
Public Sub counter10CORRECT(): AddPoints 10, 10, True: End Sub
