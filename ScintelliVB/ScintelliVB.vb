Imports ScintillaNET
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Public Class ScintelliVB
    Private mTextArea As Scintilla

    Private keywords_vb As String = "AddHandler AddressOf Alias And AndAlso Ansi Append As Assembly Auto Binary ByRef ByVal Call Case Catch CBool CByte CChar CDate CDbl CInt Class CLng CObj Const Continue CSByte CShort CSng CStr CType CUInt CULng CUShort Declare Default Delegate Dim DirectCast Do Each Else ElseIf End EndIf Enum Erase Error Event Exit False Finally For Friend Function Get GetType GetXMLNamespace Global GoSub GoTo Handles If Implements Imports In Inherits Interface Is IsNot Let Lib Like Loop Me Mod Module MustInherit MustOverride MyBase MyClass Narrowing New Next Not Nothing NotInheritable NotOverridable Of On Operator Option Optional Or OrElse Out Overloads Overridable Overrides ParamArray Partial Private Property Protected Public RaiseEvent ReadOnly ReDim REM RemoveHandler Resume Return Select Set Shadows Shared Static Step Stop Structure Sub SyncLock Then Throw To True Try TryCast TypeOf Using Wend When While Widening With WithEvents WriteOnly Xor #Const #Else #ElseIf #End #If"
    Private keywords_As As String = "Boolean Byte Char Date Decimal Double Integer Long Matrix Object Plane SByte short Single String UInteger ULong UShort Variant Vector2 Vector3"
    Private Alphabet As String = "abcdefghijklmnopqrstuvwxyz"
    Private keywords_vb_List() As String

    Private Word_Separator As String = "  ."
    'Temp variable for test
    Public Property Text0 As String
    Public Property Text1 As String
    Public Property Text2 As String
    Public Property Text3 As String
    Public Property Text4 As String

    Private mCanCheckLine As Boolean = True
    Public Sub New(TextArea As Scintilla)
        mTextArea = TextArea
        keywords_vb_List = Split(keywords_vb & " " & keywords_As, " ")

        Init_Handle()
        Config()
    End Sub

    Private Sub Init_Handle()
        AddHandler mTextArea.AutoCCancelled, AddressOf AutoCCancelled
        AddHandler mTextArea.AutoCCharDeleted, AddressOf AutoCCharDeleted
        AddHandler mTextArea.AutoCCompleted, AddressOf AutoCCompleted
        AddHandler mTextArea.AutoCSelection, AddressOf AutoCSelection
        AddHandler mTextArea.AutoSizeChanged, AddressOf AutoSizeChanged
        AddHandler mTextArea.BackColorChanged, AddressOf BackColorChanged
        AddHandler mTextArea.BackgroundImageChanged, AddressOf BackgroundImageChanged
        AddHandler mTextArea.BackgroundImageLayoutChanged, AddressOf BackgroundImageLayoutChanged
        AddHandler mTextArea.BeforeDelete, AddressOf BeforeDelete
        AddHandler mTextArea.BeforeInsert, AddressOf BeforeInsert
        AddHandler mTextArea.BindingContextChanged, AddressOf BindingContextChanged
        AddHandler mTextArea.BorderStyleChanged, AddressOf BorderStyleChanged
        AddHandler mTextArea.CausesValidationChanged, AddressOf CausesValidationChanged
        AddHandler mTextArea.ChangeAnnotation, AddressOf ChangeAnnotation
        AddHandler mTextArea.ChangeUICues, AddressOf ChangeUICues
        AddHandler mTextArea.CharAdded, AddressOf CharAdded
        AddHandler mTextArea.Click, AddressOf Click
        AddHandler mTextArea.ClientSizeChanged, AddressOf ClientSizeChanged
        AddHandler mTextArea.ContextMenuChanged, AddressOf ContextMenuChanged
        AddHandler mTextArea.ContextMenuStripChanged, AddressOf ContextMenuStripChanged
        AddHandler mTextArea.ControlAdded, AddressOf ControlAdded
        AddHandler mTextArea.ControlRemoved, AddressOf ControlRemoved
        AddHandler mTextArea.CursorChanged, AddressOf CursorChanged
        AddHandler mTextArea.Delete, AddressOf Delete
        AddHandler mTextArea.Disposed, AddressOf Disposed
        AddHandler mTextArea.DockChanged, AddressOf DockChanged
        AddHandler mTextArea.DoubleClick, AddressOf DoubleClick
        AddHandler mTextArea.DragDrop, AddressOf DragDrop
        AddHandler mTextArea.DragEnter, AddressOf DragEnter
        AddHandler mTextArea.DragLeave, AddressOf DragLeave
        AddHandler mTextArea.DragOver, AddressOf DragOver
        AddHandler mTextArea.DwellEnd, AddressOf DwellEnd
        AddHandler mTextArea.DwellStart, AddressOf DwellStart
        AddHandler mTextArea.EnabledChanged, AddressOf EnabledChanged
        AddHandler mTextArea.Enter, AddressOf Enter
        AddHandler mTextArea.FontChanged, AddressOf FontChanged
        AddHandler mTextArea.ForeColorChanged, AddressOf ForeColorChanged
        AddHandler mTextArea.GiveFeedback, AddressOf GiveFeedback
        AddHandler mTextArea.GotFocus, AddressOf GotFocus
        AddHandler mTextArea.HandleCreated, AddressOf HandleCreated
        AddHandler mTextArea.HandleDestroyed, AddressOf HandleDestroyed
        AddHandler mTextArea.HelpRequested, AddressOf HelpRequested
        AddHandler mTextArea.HotspotClick, AddressOf HotspotClick
        AddHandler mTextArea.HotspotDoubleClick, AddressOf HotspotDoubleClick
        AddHandler mTextArea.HotspotReleaseClick, AddressOf HotspotReleaseClick
        AddHandler mTextArea.ImeModeChanged, AddressOf ImeModeChanged
        AddHandler mTextArea.IndicatorClick, AddressOf IndicatorClick
        AddHandler mTextArea.IndicatorRelease, AddressOf IndicatorRelease
        AddHandler mTextArea.Insert, AddressOf Insert
        AddHandler mTextArea.InsertCheck, AddressOf InsertCheck
        AddHandler mTextArea.Invalidated, AddressOf Invalidated
        AddHandler mTextArea.KeyDown, AddressOf KeyDown
        AddHandler mTextArea.KeyPress, AddressOf KeyPress
        AddHandler mTextArea.KeyUp, AddressOf KeyUp
        AddHandler mTextArea.Layout, AddressOf Layout
        AddHandler mTextArea.Leave, AddressOf Leave
        AddHandler mTextArea.LocationChanged, AddressOf LocationChanged
        AddHandler mTextArea.LostFocus, AddressOf LostFocus
        AddHandler mTextArea.MarginChanged, AddressOf MarginChanged
        AddHandler mTextArea.MarginClick, AddressOf MarginClick
        AddHandler mTextArea.MarginRightClick, AddressOf MarginRightClick
        AddHandler mTextArea.ModifyAttempt, AddressOf ModifyAttempt
        AddHandler mTextArea.MouseCaptureChanged, AddressOf MouseCaptureChanged
        AddHandler mTextArea.MouseClick, AddressOf MouseClick
        AddHandler mTextArea.MouseDoubleClick, AddressOf MouseDoubleClick
        AddHandler mTextArea.MouseDown, AddressOf MouseDown
        AddHandler mTextArea.MouseEnter, AddressOf MouseEnter
        AddHandler mTextArea.MouseHover, AddressOf MouseHover
        AddHandler mTextArea.MouseLeave, AddressOf MouseLeave
        AddHandler mTextArea.MouseMove, AddressOf MouseMove
        AddHandler mTextArea.MouseUp, AddressOf MouseUp
        AddHandler mTextArea.MouseWheel, AddressOf MouseWheel
        AddHandler mTextArea.Move, AddressOf Move
        AddHandler mTextArea.NeedShown, AddressOf NeedShown
        AddHandler mTextArea.PaddingChanged, AddressOf PaddingChanged
        AddHandler mTextArea.Paint, AddressOf Paint
        AddHandler mTextArea.Painted, AddressOf Painted
        AddHandler mTextArea.ParentChanged, AddressOf ParentChanged
        AddHandler mTextArea.PreviewKeyDown, AddressOf PreviewKeyDown
        AddHandler mTextArea.QueryAccessibilityHelp, AddressOf QueryAccessibilityHelp
        AddHandler mTextArea.QueryContinueDrag, AddressOf QueryContinueDrag
        AddHandler mTextArea.RegionChanged, AddressOf RegionChanged
        AddHandler mTextArea.Resize, AddressOf Resize
        AddHandler mTextArea.RightToLeftChanged, AddressOf RightToLeftChanged
        AddHandler mTextArea.SavePointLeft, AddressOf SavePointLeft
        AddHandler mTextArea.SavePointReached, AddressOf SavePointReached
        AddHandler mTextArea.SizeChanged, AddressOf SizeChanged
        AddHandler mTextArea.StyleChanged, AddressOf StyleChanged
        AddHandler mTextArea.StyleNeeded, AddressOf StyleNeeded
        AddHandler mTextArea.SystemColorsChanged, AddressOf SystemColorsChanged
        AddHandler mTextArea.TabIndexChanged, AddressOf TabIndexChanged
        AddHandler mTextArea.TabStopChanged, AddressOf TabStopChanged
        AddHandler mTextArea.TextChanged, AddressOf TextChanged
        AddHandler mTextArea.UpdateUI, AddressOf UpdateUI
        AddHandler mTextArea.Validated, AddressOf Validated
        AddHandler mTextArea.Validating, AddressOf Validating
        AddHandler mTextArea.VisibleChanged, AddressOf VisibleChanged
        AddHandler mTextArea.ZoomChanged, AddressOf ZoomChanged
    End Sub

    Public Sub Config()
        mTextArea.WrapMode = ScintillaNET.WrapMode.None
        mTextArea.IndentationGuides = ScintillaNET.IndentView.LookBoth

        mTextArea.SetSelectionBackColor(True, Color.FromArgb(31, 98, 169))
        mTextArea.StyleResetDefault()
        mTextArea.CaretForeColor = Color.White
        mTextArea.CaretLineBackColor = Color.FromArgb(31, 98, 169)
        With mTextArea.Styles(Style.Default)
            .Font = "Consolas"
            .Size = 10
            .BackColor = Color.FromArgb(30, 30, 30)
            .ForeColor = Color.White
        End With

        mTextArea.StyleClearAll()
        mTextArea.Styles(Style.Vb.Default).ForeColor = Color.White
        mTextArea.Styles(Style.Vb.Comment).ForeColor = Color.FromArgb(87, 166, 74)
        mTextArea.Styles(Style.Vb.CommentBlock).ForeColor = Color.FromArgb(87, 166, 74)
        mTextArea.Styles(Style.Vb.Constant).ForeColor = Color.FromArgb(86, 156, 214)
        mTextArea.Styles(Style.Vb.Date).ForeColor = Color.DarkRed
        mTextArea.Styles(Style.Vb.Keyword).ForeColor = Color.FromArgb(86, 156, 214)
        mTextArea.Styles(Style.Vb.Keyword2).ForeColor = Color.FromArgb(86, 156, 214)
        mTextArea.Styles(Style.Vb.Number).ForeColor = Color.FromArgb(181, 206, 168)
        mTextArea.Styles(Style.Vb.String).ForeColor = Color.FromArgb(217, 157, 133)
        mTextArea.Styles(Style.Vb.Operator).ForeColor = Color.FromArgb(180, 180, 180)

        mTextArea.Styles(Style.LineNumber).BackColor = Color.FromArgb(30, 30, 30)
        mTextArea.Styles(Style.LineNumber).ForeColor = Color.FromArgb(43, 145, 175)

        mTextArea.SetFoldMarginColor(True, Color.FromArgb(30, 30, 30))
        mTextArea.SetFoldMarginHighlightColor(True, Color.FromArgb(30, 30, 30))
        With mTextArea.Margins(1)
            .Width = 30
            .Type = MarginType.Number
            .Sensitive = True
            .Mask = 0
        End With

        mTextArea.IndentWidth = 4
        mTextArea.TabWidth = 4
        mTextArea.UseTabs = True
        mTextArea.IndentationGuides = IndentView.LookBoth
        'TextArea.IndentW
        mTextArea.Lexer = Lexer.Vb
        mTextArea.LexerLanguage = "vb"

        mTextArea.SetKeywords(0, keywords_vb.ToLower & " " & keywords_As.ToLower)

        mTextArea.Markers(0).Symbol = MarkerSymbol.Background
        mTextArea.Markers(0).SetBackColor(Color.DarkRed)

        mTextArea.Markers(1).Symbol = MarkerSymbol.Background
        mTextArea.Markers(1).SetBackColor(Color.FromArgb(255, 127, 106, 0))

        mTextArea.Indicators(8).Style = IndicatorStyle.StraightBox
        mTextArea.Indicators(8).Under = True
        mTextArea.Indicators(8).ForeColor = Color.FromArgb(255, 101, 51, 6)
        mTextArea.Indicators(8).OutlineAlpha = 50
        mTextArea.Indicators(8).Alpha = 255

        mTextArea.AutoCIgnoreCase = True

        'Folding
        mTextArea.SetProperty("fold", "1")
        mTextArea.SetProperty("fold.compact", "1")
        With mTextArea.Margins(2)
            .Width = 30
            .Type = MarginType.Symbol
            .Sensitive = True
            .Mask = Marker.MaskFolders
        End With
        For i As Integer = 25 To 31
            mTextArea.Markers(i).SetForeColor(Color.FromArgb(30, 30, 30))
            mTextArea.Markers(i).SetBackColor(Color.FromArgb(155, 155, 155))
            '   TextArea.Markers(i).SetAlpha(0)
        Next
        mTextArea.Markers(Marker.Folder).Symbol = MarkerSymbol.BoxPlus
        mTextArea.Markers(Marker.FolderOpen).Symbol = MarkerSymbol.BoxMinus
        mTextArea.Markers(Marker.FolderEnd).Symbol = MarkerSymbol.BoxPlusConnected
        mTextArea.Markers(Marker.FolderMidTail).Symbol = MarkerSymbol.VLine
        mTextArea.Markers(Marker.FolderOpenMid).Symbol = MarkerSymbol.BoxMinusConnected
        mTextArea.Markers(Marker.FolderSub).Symbol = MarkerSymbol.VLine
        mTextArea.Markers(Marker.FolderTail).Symbol = MarkerSymbol.LCorner
        mTextArea.AutomaticFold = (AutomaticFold.Show Or AutomaticFold.Click Or AutomaticFold.Change)

        'image AutoComplete
        'mTextArea.RegisterRgbaImage(0, New Bitmap(App.ImageList_Autocomplete.Images(0)))
        'mTextArea.RegisterRgbaImage(1, New Bitmap(App.ImageList_Autocomplete.Images(1)))

        mTextArea.Styles(Style.CallTip).BackColor = Color.FromArgb(66, 66, 69)
        mTextArea.Styles(Style.CallTip).ForeColor = Color.FromArgb(255, 255, 255)
        mTextArea.Styles(Style.CallTip).Bold = True
        mTextArea.CallTipSetForeHlt(Color.FromArgb(86, 156, 214))
    End Sub

    Public Sub AutoCCancelled(sender As Object, e As EventArgs)

    End Sub

    Public Sub AutoCCharDeleted(sender As Object, e As EventArgs)

    End Sub

    Public Sub AutoCCompleted(sender As Object, e As AutoCSelectionEventArgs)

    End Sub

    Public Sub AutoCSelection(sender As Object, e As AutoCSelectionEventArgs)

    End Sub

    Public Sub AutoSizeChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub BackColorChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub BackgroundImageChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub BackgroundImageLayoutChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub BeforeDelete(sender As Object, e As BeforeModificationEventArgs)

    End Sub

    Public Sub BeforeInsert(sender As Object, e As BeforeModificationEventArgs)

    End Sub

    Public Sub BindingContextChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub BorderStyleChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub CausesValidationChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub ChangeAnnotation(sender As Object, e As ChangeAnnotationEventArgs)

    End Sub

    Public Sub ChangeUICues(sender As Object, e As UICuesEventArgs)

    End Sub

    Public Sub CharAdded(sender As Object, e As CharAddedEventArgs)
        'UpperCase first letter of keywords
        Replace_Line(mTextArea.CurrentLine, UpperKeyWord(mTextArea.CurrentLine))

        'if return is pressed check previous line
        If e.Char = 13 Then
            Replace_Line(mTextArea.CurrentLine - 1, UpperKeyWord(mTextArea.CurrentLine - 1))
        End If
    End Sub

    Public Sub Click(sender As Object, e As EventArgs)

    End Sub

    Public Sub ClientSizeChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub ContextMenuChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub ContextMenuStripChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub ControlAdded(sender As Object, e As ControlEventArgs)

    End Sub

    Public Sub ControlRemoved(sender As Object, e As ControlEventArgs)

    End Sub

    Public Sub CursorChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub Delete(sender As Object, e As ModificationEventArgs)

    End Sub

    Public Sub Disposed(sender As Object, e As EventArgs)

    End Sub

    Public Sub DockChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub DoubleClick(sender As Object, e As DoubleClickEventArgs)

    End Sub

    Public Sub DragDrop(sender As Object, e As DragEventArgs)

    End Sub

    Public Sub DragEnter(sender As Object, e As DragEventArgs)

    End Sub

    Public Sub DragLeave(sender As Object, e As EventArgs)

    End Sub

    Public Sub DragOver(sender As Object, e As DragEventArgs)

    End Sub

    Public Sub DwellEnd(sender As Object, e As DwellEventArgs)

    End Sub

    Public Sub DwellStart(sender As Object, e As DwellEventArgs)

    End Sub

    Public Sub EnabledChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub Enter(sender As Object, e As EventArgs)

    End Sub

    Public Sub FontChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub ForeColorChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub GiveFeedback(sender As Object, e As GiveFeedbackEventArgs)

    End Sub

    Public Sub GotFocus(sender As Object, e As EventArgs)

    End Sub

    Public Sub HandleCreated(sender As Object, e As EventArgs)

    End Sub

    Public Sub HandleDestroyed(sender As Object, e As EventArgs)

    End Sub

    Public Sub HelpRequested(sender As Object, hlpevent As HelpEventArgs)

    End Sub

    Public Sub HotspotClick(sender As Object, e As HotspotClickEventArgs)

    End Sub

    Public Sub HotspotDoubleClick(sender As Object, e As HotspotClickEventArgs)

    End Sub

    Public Sub HotspotReleaseClick(sender As Object, e As HotspotClickEventArgs)

    End Sub

    Public Sub ImeModeChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub IndicatorClick(sender As Object, e As IndicatorClickEventArgs)

    End Sub

    Public Sub IndicatorRelease(sender As Object, e As IndicatorReleaseEventArgs)

    End Sub

    Public Sub Insert(sender As Object, e As ModificationEventArgs)

    End Sub

    Public Sub InsertCheck(sender As Object, e As InsertCheckEventArgs)

    End Sub

    Public Sub Invalidated(sender As Object, e As InvalidateEventArgs)

    End Sub

    Public Sub KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Left OrElse e.KeyCode = Keys.Right OrElse e.KeyCode = Keys.Up OrElse e.KeyCode = Keys.Down Then
            Replace_Line(mTextArea.CurrentLine, UpperKeyWord(mTextArea.CurrentLine))
        End If
    End Sub

    Public Sub KeyPress(sender As Object, e As KeyPressEventArgs)
        'Remove special character
        'If (AscW(e.KeyChar) < 32 AndAlso AscW(e.KeyChar) > 13 OrElse AscW(e.KeyChar) < 8) Then
        '    e.Handled = True
        '    Exit Sub
        'End If
    End Sub

    Public Sub KeyUp(sender As Object, e As KeyEventArgs)
        mCanCheckLine = True
    End Sub

    Public Sub Layout(sender As Object, e As LayoutEventArgs)

    End Sub

    Public Sub Leave(sender As Object, e As EventArgs)

    End Sub

    Public Sub LocationChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub LostFocus(sender As Object, e As EventArgs)

    End Sub

    Public Sub MarginChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub MarginClick(sender As Object, e As MarginClickEventArgs)

    End Sub

    Public Sub MarginRightClick(sender As Object, e As MarginClickEventArgs)

    End Sub

    Public Sub ModifyAttempt(sender As Object, e As EventArgs)

    End Sub

    Public Sub MouseCaptureChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub MouseClick(sender As Object, e As MouseEventArgs)

    End Sub

    Public Sub MouseDoubleClick(sender As Object, e As MouseEventArgs)

    End Sub

    Public Sub MouseDown(sender As Object, e As MouseEventArgs)

    End Sub

    Public Sub MouseEnter(sender As Object, e As EventArgs)

    End Sub

    Public Sub MouseHover(sender As Object, e As EventArgs)

    End Sub

    Public Sub MouseLeave(sender As Object, e As EventArgs)

    End Sub

    Public Sub MouseMove(sender As Object, e As MouseEventArgs)

    End Sub

    Public Sub MouseUp(sender As Object, e As MouseEventArgs)

    End Sub

    Public Sub MouseWheel(sender As Object, e As MouseEventArgs)

    End Sub

    Public Sub Move(sender As Object, e As EventArgs)

    End Sub

    Public Sub NeedShown(sender As Object, e As NeedShownEventArgs)

    End Sub

    Public Sub PaddingChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Public Sub Painted(sender As Object, e As EventArgs)

    End Sub

    Public Sub ParentChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs)

    End Sub

    Public Sub QueryAccessibilityHelp(sender As Object, e As QueryAccessibilityHelpEventArgs)

    End Sub

    Public Sub QueryContinueDrag(sender As Object, e As QueryContinueDragEventArgs)

    End Sub

    Public Sub RegionChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub Resize(sender As Object, e As EventArgs)

    End Sub

    Public Sub RightToLeftChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub SavePointLeft(sender As Object, e As EventArgs)

    End Sub

    Public Sub SavePointReached(sender As Object, e As EventArgs)

    End Sub

    Public Sub SizeChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub StyleChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub StyleNeeded(sender As Object, e As StyleNeededEventArgs)

    End Sub

    Public Sub SystemColorsChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub TabIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub TabStopChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub TextChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub UpdateUI(sender As Object, e As UpdateUIEventArgs)

    End Sub

    Public Sub Validated(sender As Object, e As EventArgs)

    End Sub

    Public Sub Validating(sender As Object, e As CancelEventArgs)

    End Sub

    Public Sub VisibleChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub ZoomChanged(sender As Object, e As EventArgs)

    End Sub

    Private Function UpperKeyWord(Line As Integer) As String
        Dim Text As String = mTextArea.Lines(Line).Text
        If String.IsNullOrWhiteSpace(Text) Then Return Text

        Dim Word_Start As Integer = -1
        Dim Word_End As Integer = -1
        For i As Integer = 0 To Text.Length - 1
            'get position of first letter
            If Word_Start = -1 AndAlso Alphabet.ToLower.Contains(Text(i).ToString.ToLower) Then
                Word_Start = i
            End If
            'get position of last letter
            If Word_Start > -1 AndAlso Word_End = -1 AndAlso i > Word_Start Then

                'if it's a word separator
                If Word_Separator.Contains(Text(i)) Then
                    Word_End = i
                End If

                'special case
                If i = Text.Length - 1 Then
                    Word_End = i - 1
                End If
            End If

            'if a word is found
            If Word_Start > -1 AndAlso Word_End > -1 Then

                'set first letter Uppercase if it's a keyword
                Dim current_Word As String = Text.Substring(Word_Start, Word_End - Word_Start)
                If IsKeyWord(current_Word) Then
                    current_Word = current_Word(0).ToString.ToUpper & current_Word.Substring(1)
                End If

                'reconstruct text
                Dim finaltext As String = Text.Substring(0, Word_Start) & current_Word & Text.Substring(Word_End)
                Text = finaltext

                Word_Start = -1
                Word_End = -1
            End If
        Next

        Return Text
    End Function

    Private Function IsKeyWord(Word As String) As Boolean
        For i As Integer = 0 To keywords_vb_List.Count - 1
            If keywords_vb_List(i).ToLower = Word.ToLower Then Return True
        Next
        'If keywords_vb.ToLower.Contains(Word.ToLower) Then Return True
        'If keywords_As.ToLower.Contains(Word.ToLower) Then Return True

        Return False
    End Function

    Private Sub Replace_Line(Line As Integer, Text As String)
        Dim CursorPosition As Integer = mTextArea.CurrentPosition
        mTextArea.TargetStart = mTextArea.Lines(Line).Position
        mTextArea.TargetEnd = mTextArea.Lines(Line).EndPosition
        mTextArea.ReplaceTarget(Text)
        mTextArea.GotoPosition(CursorPosition)
    End Sub
End Class
