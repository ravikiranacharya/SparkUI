Imports MintUI
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

''' <summary>
''' Spark Theme for VB .net
''' Created by: Ravi
''' Date: 04-Feb-2016
''' Credits: Aeonhack
''' 
''' For using this, you will have to add MintUI.dll in your project references. 
''' You can get the required DLL from here: https://github.com/CodeRail/MintUI/releases
''' 
''' 
''' </summary>
''' <remarks></remarks>
Module Help
    Friend G As Graphics
    Public Function GetBrush(ByVal Color As Color) As SolidBrush
        Return New SolidBrush(Color)
    End Function
    Public Function GetPen(ByVal Color As Color) As Pen
        Return New Pen(Color)
    End Function
End Module

Public Class SparkForm
    Inherits ContainerControl
    Private G As Graphics
    Private Path1, Path2 As GraphicsPath
    Private _BorderRadius As Integer
    Private _BorderColor As Color
    Private FormRectangle As Rectangle
    Private _TopLeftRectangle, _TopRightRectangle, _IconRectangle As Rectangle
    Private _TopLeftColor, _TopRightColor As Color
    Private _TextLocation As Point
    Private _DrawSeparator As Boolean
    Private _SeparatorColor As Color
    Property TopLeftColor As Color
        Get
            Return _TopLeftColor
        End Get
        Set(value As Color)
            _TopLeftColor = value
            Invalidate()
        End Set
    End Property
    Property TopRightColor As Color
        Get
            Return _TopRightColor
        End Get
        Set(value As Color)
            _TopRightColor = value
            Invalidate()
        End Set
    End Property
    Property TopLeftRectangle As Rectangle
        Get
            Return _TopLeftRectangle
        End Get
        Set(value As Rectangle)
            _TopLeftRectangle = value
            Invalidate()
        End Set
    End Property
    Property IconRectangle As Rectangle
        Get
            Return _IconRectangle
        End Get
        Set(value As Rectangle)
            _IconRectangle = value
            Invalidate()
        End Set
    End Property
    Property TopRightRectangle As Rectangle
        Get
            Return _TopRightRectangle
        End Get
        Set(value As Rectangle)
            _TopRightRectangle = value
            Invalidate()
        End Set
    End Property
    Property TextLocation As Point
        Get
            Return _TextLocation
        End Get
        Set(value As Point)
            _TextLocation = value
            Invalidate()
        End Set
    End Property
    Property BorderRadius As Integer
        Get
            Return _BorderRadius
        End Get
        Set(value As Integer)
            _BorderRadius = value
            Invalidate()
        End Set
    End Property
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Property DrawSeparator As Boolean
        Get
            Return _DrawSeparator
        End Get
        Set(value As Boolean)
            _DrawSeparator = value
            Invalidate()
        End Set
    End Property
    Property SeparatorColor As Color
        Get
            Return _SeparatorColor
        End Get
        Set(value As Color)
            _SeparatorColor = value
            Invalidate()
        End Set
    End Property

    Private WM_NCLBUTTONDOWN As Integer = 52
    Private Movable As Boolean = False
    Private Pt As Point
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        Movable = False
    End Sub
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If New Rectangle(0, 0, Width, WM_NCLBUTTONDOWN).Contains(e.Location) AndAlso e.Button = Windows.Forms.MouseButtons.Left Then
            Pt = e.Location
            Movable = True
        End If
    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If Movable = True Then
            ParentForm.Location = MousePosition - Pt
        End If
    End Sub
    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()
        BackColor = BackColor
        Dock = DockStyle.Fill
        FindForm.AllowTransparency = True
        FindForm.TransparencyKey = Color.Fuchsia
        FindForm.FormBorderStyle = FormBorderStyle.None
        Invalidate()
    End Sub
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        Dock = DockStyle.Fill
        DoubleBuffered = True
        BorderRadius = 0
        BackColor = Color.Teal
        BorderColor = Color.FromArgb(80, 80, 80)
        Font = New Font("Candara", 12)
        ForeColor = Color.WhiteSmoke
        TopLeftRectangle = New Rectangle(0, 0, Width * 0.75F, 40)
        TopRightRectangle = New Rectangle(Width * 0.75F, 0, Width * 0.25F, 40)
        IconRectangle = New Rectangle(15, 3, 32, 32)
        TopLeftColor = Color.Teal
        TopRightColor = Color.White
        TextLocation = New Point(55, 10)
        DrawSeparator = True
        SeparatorColor = Color.White
    End Sub
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        G = e.Graphics
        G.Clear(Color.Fuchsia)
        FormRectangle = New Rectangle(0, 0, Width - 1, Height - 1)
        G.SmoothingMode = SmoothingMode.HighSpeed
        Path1 = PathHelper.FilletRectangle(FormRectangle, BorderRadius, CornerAlignment.All)
        G.PixelOffsetMode = PixelOffsetMode.HighSpeed
        G.FillPath(GetBrush(BackColor), Path1)
        Path2 = PathHelper.FilletRectangle(TopLeftRectangle, BorderRadius, CornerAlignment.All)
        G.FillPath(GetBrush(TopLeftColor), Path2)
        G.FillRectangle(GetBrush(TopRightColor), TopRightRectangle)
        If DrawSeparator Then
            G.FillRectangle(GetBrush(SeparatorColor), New Rectangle(0, TopLeftRectangle.Height, Width, 5))
        End If
        G.DrawPath(GetPen(BorderColor), Path1)
        G.DrawString(Text, Font, GetBrush(ForeColor), New Point(TextLocation))
        If Not ParentForm.Icon Is Nothing Then
            G.DrawIcon(ParentForm.Icon, IconRectangle)
            'G.DrawRectangle(GetPen(BorderColor), IconRectangle)
        End If
        Path1.Dispose()
        Path2.Dispose()
    End Sub
End Class

Class SparkControlBox
    Inherits Control
    Private _MinimizeColor, _MaximizeColor, _CloseColor As Color
    Property MinimizeColor As Color
        Get
            Return _MinimizeColor
        End Get
        Set(value As Color)
            _MinimizeColor = value
            Invalidate()
        End Set
    End Property
    Property MaximizeColor As Color
        Get
            Return _MaximizeColor
        End Get
        Set(value As Color)
            _MaximizeColor = value
            Invalidate()
        End Set
    End Property
    Property CloseColor As Color
        Get
            Return _CloseColor
        End Get
        Set(value As Color)
            _CloseColor = value
            Invalidate()
        End Set
    End Property
    Private X As Integer
    Private S As MouseState = MouseState.Normal
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        S = MouseState.MouseOver
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        S = MouseState.Normal
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        S = MouseState.MouseDown
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        S = MouseState.MouseOver
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        X = e.X
        Invalidate()
    End Sub

    Protected Overrides Sub OnClick(e As EventArgs)
        MyBase.OnClick(e)
        If X <= Width / 3.0F Then
            FindForm.WindowState = FormWindowState.Minimized
        ElseIf X > Width / 3.0F And (2 * Width / 3.0F) > X Then
            If Not FindForm.WindowState = FormWindowState.Normal Then
                FindForm.WindowState = FormWindowState.Normal
            Else
                FindForm.WindowState = FormWindowState.Maximized
            End If
        Else
            Application.Exit()
        End If
    End Sub
    Sub New()
        Size = New Size(72, 24)
        DoubleBuffered = True
        MinimizeColor = Color.Teal
        MaximizeColor = Color.Aqua
        CloseColor = Color.DarkCyan
        BackColor = Color.Teal
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        'Size = New Size(52, 24)
    End Sub
    Private MinimizeRectangle, MaximizeRectangle, CloseRectangle As Rectangle

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        G = e.Graphics
        G.Clear(BackColor)
        MinimizeRectangle = New Rectangle(0, 0, Width / 3.0F, Height)
        MaximizeRectangle = New Rectangle(Width / 3.0F, 0, Width / 3.0F, Height)
        CloseRectangle = New Rectangle(2 * Width / 3.0F, 0, Width / 3.0F, Height)
        G.FillRectangle(GetBrush(MinimizeColor), MinimizeRectangle)
        G.FillRectangle(GetBrush(MaximizeColor), MaximizeRectangle)
        G.FillRectangle(GetBrush(CloseColor), CloseRectangle)
        GlyphRenderer.DrawWindowGlyph(G, MinimizeRectangle, WindowGlyph.Minimize, ForeColor)
        GlyphRenderer.DrawWindowGlyph(G, MaximizeRectangle, WindowGlyph.Maximize, ForeColor)
        GlyphRenderer.DrawWindowGlyph(G, CloseRectangle, WindowGlyph.Close, ForeColor)

        If MouseState.MouseOver Then
            If X <= Width / 3.0F Then
                G.FillRectangle(GetBrush(Color.FromArgb(15, MinimizeColor)), MinimizeRectangle)
            ElseIf (2 * Width / 3.0F) > X >= Width / 3.0F Then
                G.FillRectangle(GetBrush(Color.FromArgb(15, MaximizeColor)), MaximizeRectangle)
            Else
                G.FillRectangle(GetBrush(Color.FromArgb(15, CloseColor)), CloseRectangle)
            End If
        ElseIf MouseState.MouseDown Then
            If X <= Width / 3.0F Then
                G.FillRectangle(GetBrush(Color.FromArgb(0, MinimizeColor)), MinimizeRectangle)
            ElseIf (2 * Width / 3.0F) > X >= Width / 3.0F Then
                G.FillRectangle(GetBrush(Color.FromArgb(0, MaximizeColor)), MaximizeRectangle)
            Else
                G.FillRectangle(GetBrush(Color.FromArgb(0, CloseColor)), CloseRectangle)
            End If
        End If

    End Sub
End Class

Public Class SparkTextBox
    Inherits MintTextBox
    Private G As Graphics
    Private TextRectangle As Rectangle
    Private _BaseColor, _BorderColor As Color
    Private _BorderRadius As Integer
    Private Path1 As GraphicsPath
    Property BaseColor As Color
        Get
            Return _BaseColor
        End Get
        Set(value As Color)
            _BaseColor = value
            Invalidate()
        End Set
    End Property
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Property BorderRadius As Integer
        Get
            Return _BorderRadius
        End Get
        Set(value As Integer)
            _BorderRadius = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        'SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        'SetStyle(ControlStyles.ResizeRedraw, True)
        ForeColor = Color.White
        Font = New Font("Lucida Sans Unicode", 9)
        BorderColor = Color.FromArgb(0, 130, 124)
        BackColor = Color.Teal
        BaseColor = Color.Teal
        BorderRadius = 0
        Padding = New Padding(4, 4, 4, 4)
        Size = New Size(120, 25)
        Multiline = False
    End Sub
    Public Overrides Sub OnPaintFrame(e As TextBoxPaintFrameEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        TextRectangle = New Rectangle(0, 0, Width - 1, Height - 1)
        Path1 = PathHelper.FilletRectangle(TextRectangle, BorderRadius, CornerAlignment.All)
        G.FillPath(GetBrush(BaseColor), Path1)
        G.SmoothingMode = SmoothingMode.HighQuality
        G.DrawPath(GetPen(BorderColor), Path1)
        Path1.Dispose()
    End Sub
End Class

Public Class SparkGroupBox
    Inherits MintGroupBox
    Private G As Graphics
    Private GroupRectangle As Rectangle
    Private _BorderColor As Color
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        BackColor = Color.FromArgb(0, 167, 157)
        ForeColor = Color.White
        Font = New Font("Candara", 10)
        'BaseColor = Color.FromArgb(0, 167, 157)
        BorderColor = Color.FromArgb(35, 209, 198)
    End Sub
    Public Overloads Overrides Sub OnPaint(e As GroupBoxPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        G.SmoothingMode = SmoothingMode.HighQuality
        GroupRectangle = New Rectangle(1, 1, Width - 3, Height - 3)
        G.FillRectangle(GetBrush(Color.FromArgb(80, 80, 80)), ClientRectangle)
        G.DrawRectangle(GetPen(Color.FromArgb(80, 80, 80)), ClientRectangle)
        G.FillRectangle(GetBrush(BackColor), GroupRectangle)
        G.DrawRectangle(GetPen(BorderColor), GroupRectangle)
        G.DrawString(Text, Font, New SolidBrush(ForeColor), 20, 10)
    End Sub
End Class

Public Class SparkButton
    Inherits MintButton
    Private G As Graphics
    Private _BaseColor, _BorderColor, _ButtonClickColor, _HoverColor As Color
    Private _BorderRadius As Integer
    Private Path1, Path2 As GraphicsPath
    Private GroupRectangle, BorderRectangle As Rectangle
    Property BaseColor As Color
        Get
            Return _BaseColor
        End Get
        Set(value As Color)
            _BaseColor = value
            Invalidate()
        End Set
    End Property
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Property BorderRadius As Integer
        Get
            Return _BorderRadius
        End Get
        Set(value As Integer)
            _BorderRadius = value
            Invalidate()
        End Set
    End Property
    Property ButtonClickColor As Color
        Get
            Return _ButtonClickColor
        End Get
        Set(value As Color)
            _ButtonClickColor = value
            Invalidate()
        End Set
    End Property
    Property HoverColor As Color
        Get
            Return _HoverColor
        End Get
        Set(value As Color)
            _HoverColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        BackColor = Color.FromArgb(0, 167, 157)
        ForeColor = Color.White
        Font = New Font("Segoe UI Semibold", 9)
        BorderColor = Color.FromArgb(35, 209, 198)
        TextAlign = ContentAlignment.MiddleCenter
        ButtonClickColor = Color.DarkSlateGray
        HoverColor = Color.Teal
    End Sub
    Public Overloads Overrides Sub OnPaint(e As ButtonPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        G.SmoothingMode = SmoothingMode.HighQuality
        GroupRectangle = New Rectangle(1, 1, Width - 3, Height - 3)
        BorderRectangle = New Rectangle(1, 1, Width - 2, Height - 2)
        Path1 = PathHelper.FilletRectangle(GroupRectangle, BorderRadius, CornerAlignment.All)
        Path2 = PathHelper.FilletRectangle(BorderRectangle, BorderRadius, CornerAlignment.All)
        Select Case e.MouseState
            Case MouseState.MouseDown
                G.FillPath(GetBrush(ButtonClickColor), Path1)
                G.DrawPath(GetPen(Color.FromArgb(80, 80, 80)), Path2)
                G.DrawPath(GetPen(BorderColor), Path2)
            Case MouseState.MouseOver
                G.FillPath(GetBrush(HoverColor), Path1)
                G.DrawPath(GetPen(Color.FromArgb(80, 80, 80)), Path2)
                G.DrawPath(GetPen(BorderColor), Path2)
            Case Else
                G.FillPath(GetBrush(BackColor), Path2)
                G.DrawPath(GetPen(Color.FromArgb(80, 80, 80)), Path2)
                G.FillPath(GetBrush(BackColor), Path1)
                G.DrawPath(GetPen(BorderColor), Path2)
        End Select
        TextRenderer.DrawText(G, Text, Font, e.TextBounds, ForeColor, e.TextFormatFlags)
        Path1.Dispose()
        Path2.Dispose()
    End Sub
End Class

Public Class SparkCheckBox
    Inherits MintCheckBox
    Private G As Graphics
    Private Path1, Path2 As GraphicsPath
    Private _BaseColor, _CheckColor, _BorderColor As Color
    Private _Box As Boolean
    Property Box As Boolean
        Get
            Return _Box
        End Get
        Set(value As Boolean)
            _Box = value
            Invalidate()
        End Set
    End Property
    Property BaseColor As Color
        Get
            Return _BaseColor
        End Get
        Set(value As Color)
            _BaseColor = value
            Invalidate()
        End Set
    End Property
    Property CheckColor As Color
        Get
            Return _CheckColor
        End Get
        Set(value As Color)
            _CheckColor = value
            Invalidate()
        End Set
    End Property
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        BackColor = Color.FromArgb(0, 167, 157)
        ForeColor = Color.White
        Font = New Font("Segoe UI Semibold", 9)
        TextAlign = ContentAlignment.MiddleCenter
        BaseColor = Color.White
        CheckColor = Color.Teal
        BorderColor = Color.FromArgb(80, 80, 80)
    End Sub
    Public Overloads Overrides Sub OnPaint(e As CheckBoxPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        Path1 = PathHelper.FilletRectangle(New Rectangle(0, 0, 18, 18), 0, CornerAlignment.All)
        G.FillPath(GetBrush(BaseColor), Path1)
        If Checked Then
            If Box Then
                G.DrawPath(New Pen(BorderColor), Path1)
                G.FillRectangle(GetBrush(CheckColor), New Rectangle(2, 2, 14, 14))
                G.DrawRectangle(GetPen(BorderColor), New Rectangle(2, 2, 14, 14))
            Else
                G.DrawString("ü", New Font("Wingdings", 16, FontStyle.Regular), GetBrush(CheckColor), -1, -1)
            End If
        End If
        TextRenderer.DrawText(G, Text, Font, New Rectangle(e.TextBounds.X + 6, e.TextBounds.Y, e.TextBounds.Width, e.TextBounds.Height), ForeColor, e.TextFormatFlags)
    End Sub
End Class

Public Class SparkRadioButton
    Inherits MintRadioButton
    Private G As Graphics
    Private Path1, Path2 As New GraphicsPath
    Private _BaseColor, _CheckColor, _BorderColor As Color
    Private _BorderRadius As Integer
    Private Circle1, Circle2 As Rectangle
    Property BaseColor As Color
        Get
            Return _BaseColor
        End Get
        Set(value As Color)
            _BaseColor = value
        End Set
    End Property
    Property CheckColor As Color
        Get
            Return _CheckColor
        End Get
        Set(value As Color)
            _CheckColor = value
        End Set
    End Property
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Property BorderRadius As Integer
        Get
            Return _BorderRadius
        End Get
        Set(value As Integer)
            _BorderRadius = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        BackColor = Color.FromArgb(0, 167, 157)
        ForeColor = Color.White
        Font = New Font("Segoe UI Semibold", 9)
        TextAlign = ContentAlignment.MiddleCenter
        BaseColor = Color.White
        CheckColor = Color.Teal
        BorderColor = Color.FromArgb(80, 80, 80)
        Size = New Size(100, 25)
        BorderRadius = Height - 2
    End Sub
    Public Overloads Overrides Sub OnPaint(e As CheckBoxPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        G.SmoothingMode = SmoothingMode.AntiAlias
        Circle1 = New Rectangle(0, 0, Height - 1, Height - 1)
        Circle2 = New Rectangle(4, 4, Height - 9, Height - 9)
        G.FillEllipse(GetBrush(BaseColor), Circle1)
        'RadioButtonRenderer.GetGlyphSize(G, VisualStyles.RadioButtonState.UncheckedNormal)
        'RadioButtonRenderer.DrawRadioButton(G, New Point(4, 4), VisualStyles.RadioButtonState.UncheckedNormal)
        G.DrawEllipse(New Pen(BorderColor), Circle1)
        If Checked Then
            G.FillEllipse(GetBrush(CheckColor), Circle2)
            G.DrawEllipse(New Pen(CheckColor), Circle2)
            'RadioButtonRenderer.DrawRadioButton(G, New Point(4, 4), VisualStyles.RadioButtonState.CheckedNormal)
        End If
        TextRenderer.DrawText(G, Text, Font, New Rectangle(e.TextBounds.X + 6, e.TextBounds.Y, e.TextBounds.Width, e.TextBounds.Height), ForeColor, e.TextFormatFlags)
        Path1.Dispose()
        Path2.Dispose()
    End Sub
End Class

Public Class SparkSeparator
    Inherits MintSeparator
    Private G As Graphics
    Private Path1, Path2 As New GraphicsPath
    Private _Orientation As Boolean
    Private _Shadow As Boolean
    Private _ShadowColor As Color

    Enum Style
        Horizontal
        Vertical
    End Enum
    Property Orientation As Style
        Get
            Return _Orientation
        End Get
        Set(value As Style)
            _Orientation = value
            Invalidate()
        End Set
    End Property
    Property Shadow As Boolean
        Get
            Return _Shadow
        End Get
        Set(value As Boolean)
            _Shadow = value
            Invalidate()
        End Set
    End Property
    Property ShadowColor As Color
        Get
            Return _ShadowColor
        End Get
        Set(value As Color)
            _ShadowColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        BackColor = Color.FromArgb(0, 167, 157)
        Shadow = True
        ShadowColor = Color.FromArgb(80, 80, 80)
        Size = New Size(100, 4)
        Font = New Font("Segoe UI Semibold", 4)
    End Sub
    Public Overloads Overrides Sub OnPaint(e As SeparatorPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        If Orientation = Style.Vertical Then
            G.DrawLine(GetPen(ForeColor), Width / 2.0F, 0, Width / 2.0F, 0)
        Else
            G.DrawLine(GetPen(ForeColor), 0, Height / 2.0F, Width, Height / 2.0F)
            If Shadow Then
                G.DrawLine(GetPen(ShadowColor), 0, Height / 2.0F + 1, Width, Height / 2.0F + 1)
            End If
        End If
    End Sub
End Class

Public Class SparkTabControl
    Inherits TabControl
    Private G As Graphics
    Private ItemBounds, ItemRectangle As Rectangle
    Private _TabPageColor, _TabItemColor, _SelectedTabColor As Color
    Property TabPageColor As Color
        Get
            Return _TabPageColor
        End Get
        Set(value As Color)
            _TabPageColor = value
            Invalidate()
        End Set
    End Property
    Property TabItemColor As Color
        Get
            Return _TabItemColor
        End Get
        Set(value As Color)
            _TabItemColor = value
        End Set
    End Property
    Property SelectedTabColor As Color
        Get
            Return _SelectedTabColor
        End Get
        Set(value As Color)
            _SelectedTabColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        TabPageColor = Color.FromArgb(0, 167, 157)
        TabItemColor = Color.FromArgb(0, 167, 157)
        SelectedTabColor = Color.FromArgb(3, 138, 131)
        TabPageColor = Color.FromArgb(54, 55, 60)
        ItemSize = New Size(100, 40)
        Font = New Font("Segoe UI", 12)
        DrawMode = TabDrawMode.Normal
        SizeMode = TabSizeMode.Fixed
    End Sub
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        Try
            For Each T As TabPage In TabPages
                T.BorderStyle = BorderStyle.None
                T.BackColor = TabPageColor
            Next
        Catch ex As Exception
        End Try
        For TabItemIndex As Integer = 0 To TabPages.Count - 1
            ItemBounds = GetTabRect(TabItemIndex)
            ItemRectangle = New Rectangle(ItemBounds.X, ItemBounds.Y, ItemBounds.Width - 1, ItemBounds.Height - 1)
            G.FillRectangle(GetBrush(TabItemColor), ItemRectangle)
            'G.DrawRectangle(GetPen(Color.FromArgb(80, 80, 80)), ItemRectangle)
            If SelectedIndex = TabItemIndex Then
                G.FillRectangle(GetBrush(SelectedTabColor), ItemRectangle)
            End If
            Dim StringFormat As New StringFormat
            StringFormat.Alignment = StringAlignment.Center
            StringFormat.LineAlignment = StringAlignment.Center
            G.DrawString(TabPages(TabItemIndex).Text, Font, GetBrush(TabPages(TabItemIndex).ForeColor), GetTabRect(TabItemIndex), StringFormat)
        Next
    End Sub
End Class

Public Class SparkComboBox
    Inherits MintComboBox
    Private FieldRectangle, DropRectangle, TextRectangle As Rectangle
    Private _BorderColor, _DropColor, _ItemColor As Color
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Property DropColor As Color
        Get
            Return _DropColor
        End Get
        Set(value As Color)
            _DropColor = value
            Invalidate()
        End Set
    End Property
    Property ItemColor As Color
        Get
            Return _ItemColor
        End Get
        Set(value As Color)
            _ItemColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        DoubleBuffered = True
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Teal
        ItemHeight = 22
        BorderColor = Color.FromArgb(80, 80, 80)
        DropColor = Color.Teal
        ItemColor = Color.DarkCyan
    End Sub
    Public Overloads Overrides Sub OnPaint(e As ComboBoxPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        FieldRectangle = New Rectangle(0, 0, Width - 1, Height - 1)
        G.FillRectangle(GetBrush(BackColor), FieldRectangle)
        G.DrawRectangle(GetPen(BorderColor), FieldRectangle)
        DropRectangle = New Rectangle(e.ArrowBounds.X - 5, e.ArrowBounds.Y, e.ArrowBounds.Width, e.ArrowBounds.Height)
        DropRectangle.Inflate(2, 2)
        TextRectangle = e.TextBounds
        GlyphRenderer.DrawChevronGlyph(G, DropRectangle, ArrowDirection.Down, ForeColor)
        TextRenderer.DrawText(G, Text, Font, e.TextBounds, ForeColor)
    End Sub

    Public Overrides Sub OnPaintItem(e As ComboBoxDrawItemEventArgs)
        G = e.Graphics
        If e.Selected Then
            G.FillRectangle(GetBrush(DropColor), e.Bounds)
            G.DrawRectangle(GetPen(BorderColor), e.Bounds)
        Else
            G.FillRectangle(GetBrush(ItemColor), e.Bounds)
            G.DrawRectangle(GetPen(BorderColor), e.Bounds)
        End If
        TextRenderer.DrawText(G, e.Text, Font, e.Bounds, ForeColor, e.TextFormatFlags)
    End Sub
End Class

Public Class SparkListBox
    Inherits ListBox
    Private _BorderColor, _ItemColor, _SelectedItemColor As Color
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Property ItemColor As Color
        Get
            Return _ItemColor
        End Get
        Set(value As Color)
            _ItemColor = value
            Invalidate()
        End Set
    End Property
    Property SelectedItemColor As Color
        Get
            Return _SelectedItemColor
        End Get
        Set(value As Color)
            _SelectedItemColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        BackColor = Color.Teal
        ForeColor = Color.White
        DoubleBuffered = True
        BorderStyle = Windows.Forms.BorderStyle.None
        BorderColor = Color.FromArgb(64, 64, 64)
        ItemColor = Color.FromArgb(50, 50, 50)
        SelectedItemColor = Color.FromArgb(0, 190, 190)
        ItemHeight = 20
        DrawMode = Windows.Forms.DrawMode.OwnerDrawFixed
    End Sub
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        G.FillRectangle(GetBrush(BackColor), ClientRectangle)
        G.DrawRectangle(GetPen(BorderColor), 1, 1, Width - 2, Height - 2)
        ControlPaint.DrawBorder(G, Bounds, BorderColor, ButtonBorderStyle.Solid)
    End Sub
    Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
        MyBase.OnDrawItem(e)
        G = e.Graphics
        G.SmoothingMode = SmoothingMode.HighQuality
        G.FillRectangle(New SolidBrush(BackColor), New Rectangle(e.Bounds.X, e.Bounds.Y - 1, e.Bounds.Width, e.Bounds.Height + 4))
        If e.State.ToString().Contains("Selected,") Then
            G.FillRectangle(GetBrush(SelectedItemColor), New Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1))
        Else
            G.FillRectangle(New SolidBrush(BackColor), New Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1))
        End If
        G.DrawRectangle(GetPen(BorderColor), New Rectangle(0, 0, Width - 1, Height - 1))
        Try
            TextRenderer.DrawText(G, Items(e.Index).ToString(), Font, New Point(e.Bounds.X, e.Bounds.Y), ForeColor)
        Catch ex As Exception
        End Try
    End Sub
End Class

Public NotInheritable Class SparkNumericUpDown
    Inherits MintNumericUpDown
    Private G As Graphics
    Private _UpperRectangleColor, _LowerRectangleColor, _HighlightColor, _BorderColor As Color
    Private _BorderRadius As Integer
    Property BorderRadius As Integer
        Get
            Return _BorderRadius
        End Get
        Set(value As Integer)
            _BorderRadius = value
            Invalidate()
        End Set
    End Property
    Property UpperRectangleColor As Color
        Get
            Return _UpperRectangleColor
        End Get
        Set(value As Color)
            _UpperRectangleColor = value
            Invalidate()
        End Set
    End Property
    Property LowerRectangleColor As Color
        Get
            Return _LowerRectangleColor
        End Get
        Set(value As Color)
            _LowerRectangleColor = value
            Invalidate()
        End Set
    End Property
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Property HighlightColor As Color
        Get
            Return _HighlightColor
        End Get
        Set(value As Color)
            _HighlightColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Teal
        ForeColor = Color.White
        BorderRadius = 0
        BorderStyle = Windows.Forms.BorderStyle.None
        UpperRectangleColor = Color.Teal
        LowerRectangleColor = Color.DarkCyan
        HighlightColor = Color.DarkSlateGray
        BorderColor = Color.FromArgb(80, 80, 80)
        Padding = New Padding(4, 3, 4, 3)
    End Sub

    Public Overloads Overrides Sub OnPaint(e As NumericUpDownPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        Select Case e.UpButtonMouseState
            Case MouseState.MouseDown
                G.FillRectangle(GetBrush(HighlightColor), e.UpButtonBounds)
            Case MouseState.MouseOver
                G.FillRectangle(GetBrush(HighlightColor), e.UpButtonBounds)
            Case Else
                G.FillRectangle(GetBrush(UpperRectangleColor), e.UpButtonBounds)
        End Select
        Select Case e.DownButtonMouseState
            Case MouseState.MouseDown
                G.FillRectangle(GetBrush(HighlightColor), e.DownButtonBounds)
            Case MouseState.MouseOver
                G.FillRectangle(GetBrush(HighlightColor), e.UpButtonBounds)
            Case Else
                G.FillRectangle(GetBrush(LowerRectangleColor), e.DownButtonBounds)
        End Select
        ControlPaint.DrawBorder(e.Graphics, e.UpButtonBounds, BorderColor, ButtonBorderStyle.Solid)
        ControlPaint.DrawBorder(e.Graphics, e.DownButtonBounds, BorderColor, ButtonBorderStyle.Solid)
        If Enabled Then
            PaintArrowButtons(e)
        Else
            PaintArrowButtonsDisabled(e)
        End If
    End Sub
    Private Path1 As GraphicsPath
    Public Overrides Sub OnPaintFrame(e As NumericUpDownPaintFrameEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        Dim Bounds As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)
        Path1 = PathHelper.FilletRectangle(Bounds, BorderRadius, CornerAlignment.All)
        G.FillPath(GetBrush(BackColor), Path1)
        G.SmoothingMode = SmoothingMode.AntiAlias
        If Focused Then
            G.DrawPath(GetPen(HighlightColor), Path1)
        Else
            G.DrawPath(GetPen(BorderColor), Path1)
        End If
        Path1.Dispose()
    End Sub
    Private Sub PaintArrowButtons(e As NumericUpDownPaintEventArgs)
        Dim UpButtonBounds As Rectangle = e.UpButtonBounds
        UpButtonBounds.Inflate(-2, -2)

        Dim DownButtonBounds As Rectangle = e.DownButtonBounds
        DownButtonBounds.Inflate(-2, -2)

        GlyphRenderer.DrawArrowGlyph(e.Graphics, UpButtonBounds, ArrowDirection.Up, ForeColor)
        GlyphRenderer.DrawArrowGlyph(e.Graphics, DownButtonBounds, ArrowDirection.Down, ForeColor)
    End Sub
    Private Sub PaintArrowButtonsDisabled(e As NumericUpDownPaintEventArgs)
        Dim UpButtonBounds As Rectangle = e.UpButtonBounds
        UpButtonBounds.Inflate(-2, -2)

        Dim DownButtonBounds As Rectangle = e.DownButtonBounds
        DownButtonBounds.Inflate(-2, -2)

        GlyphRenderer.DrawChevronGlyph(e.Graphics, UpButtonBounds, ArrowDirection.Up, ForeColor)
        GlyphRenderer.DrawChevronGlyph(e.Graphics, DownButtonBounds, ArrowDirection.Down, ForeColor)
    End Sub
End Class

Public Class SparkProgressBar
    Inherits MintProgressBar
    Private _BorderColor, _ProgressColor As Color
    Private _BorderRadius As Integer
    Private _Font As Font
    Property TextFont As Font
        Get
            Return _Font
        End Get
        Set(value As Font)
            _Font = value
            Invalidate()
        End Set
    End Property
    Property BorderRadius As Integer
        Get
            Return _BorderRadius
        End Get
        Set(value As Integer)
            _BorderRadius = value
            Invalidate()
        End Set
    End Property
    Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    Property ProgressColor As Color
        Get
            Return _ProgressColor
        End Get
        Set(value As Color)
            _ProgressColor = value
            Invalidate()
        End Set
    End Property
    Sub New()
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        BackColor = Color.Teal
        BorderColor = Color.DarkCyan
        Size = New Size(100, 15)
        BorderRadius = 1
        ProgressColor = Color.WhiteSmoke
        TextFont = New Font("Segoe UI", 9)
    End Sub
    Private Path1, Path2 As GraphicsPath
    Private BoundsRectangle As Rectangle
    Public Overloads Overrides Sub OnPaint(e As ProgressBarPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        BoundsRectangle = New Rectangle(e.ProgressBounds.X, e.ProgressBounds.Y, e.ProgressBounds.Width - 1, e.ProgressBounds.Height - 1)
        ClientRectangle.Inflate(-1, -2)
        Path1 = PathHelper.FilletRectangle(ClientRectangle, BorderRadius, CornerAlignment.All)
        G.SmoothingMode = SmoothingMode.HighQuality
        G.FillPath(GetBrush(BackColor), Path1)
        G.DrawPath(GetPen(BorderColor), Path1)
        Path2 = PathHelper.FilletRectangle(BoundsRectangle, BorderRadius, CornerAlignment.All)
        If Enabled Then
            G.FillPath(GetBrush(ProgressColor), Path2)
            G.DrawPath(GetPen(BorderColor), Path2)
        Else
            G.FillPath(GetBrush(Color.DimGray), Path2)
            G.DrawPath(GetPen(BorderColor), Path2)
        End If
        TextRenderer.DrawText(G, Value.ToString(), TextFont, e.ProgressBounds, ForeColor)
        Path1.Dispose()
        Path2.Dispose()
    End Sub
End Class

Public Class SparkLabel
    Inherits MintLabel
    Sub New()
        Font = New Font("Candara", 9)
        ForeColor = Color.White
        BackColor = Color.Teal
    End Sub
    Public Overloads Overrides Sub OnPaint(e As LabelPaintEventArgs)
        G = e.Graphics
        G.Clear(Parent.BackColor)
        TextRenderer.DrawText(G, Text, Font, New Point(0, 0), ForeColor)
    End Sub
End Class