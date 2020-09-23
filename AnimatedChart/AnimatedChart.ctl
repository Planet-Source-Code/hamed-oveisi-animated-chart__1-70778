VERSION 5.00
Begin VB.UserControl AnimatedChart 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   5580
   ScaleWidth      =   8400
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   2565
   End
   Begin VB.PictureBox picLegend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F5F5&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFF0F0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF7040&
      Height          =   5430
      Left            =   3360
      ScaleHeight     =   5430
      ScaleWidth      =   2130
      TabIndex        =   1
      Top             =   0
      Width           =   2130
      Begin VB.VScrollBar vsbContainer 
         Height          =   5445
         LargeChange     =   5
         Left            =   1905
         Max             =   100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   225
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   5205
         Left            =   150
         ScaleHeight     =   5205
         ScaleWidth      =   1710
         TabIndex        =   2
         Top             =   0
         Width           =   1710
         Begin VB.PictureBox Box 
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000C0&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   45
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   6
            Top             =   135
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   0
            Left            =   315
            TabIndex        =   3
            Top             =   135
            Visible         =   0   'False
            Width           =   1290
         End
      End
      Begin VB.Label lblSlider 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   5430
         Left            =   15
         TabIndex        =   5
         ToolTipText     =   "Display Legend"
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"AnimatedChart.ctx":0000
      ForeColor       =   &H80000017&
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectionInfo 
         Caption         =   "Selection &Information"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoMoveInfo 
         Caption         =   "Auto &Move Information"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLegend 
         Caption         =   "&Display Legend"
      End
   End
   Begin VB.Menu mnuLegend 
      Caption         =   "&Legend"
      Begin VB.Menu mnuLegendHide 
         Caption         =   "&Hide"
      End
   End
End
Attribute VB_Name = "AnimatedChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
' AnimatedChart      By Hamed Oveisi
'                    Based On ActiveChart UPDT (Bar)
'                    A Very Nice Submission from Mirage
'                    http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=32583&lngWId=1
'
' What's New         Adding Gradient Color Theme
'                    Scrolling Chart Items to create Animated Chart
'
' Limitation         I create this to use in my project
'                    so use some fix colors due to time limitation in my project
'                    Maybe at a better time I'll put some props to customize this
'***********************************************************************************

'Gradient Constants
Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V  As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2
Private Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Private Type TRIVERTEX          'For gradient Drawing
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UPPERLEFT As Long
    LOWERRIGHT As Long
End Type

Enum GRADIENT_FILL_RECT
    FillHor = GRADIENT_FILL_RECT_H
    FillVer = GRADIENT_FILL_RECT_V
End Enum

Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long


Private uColumns()        As Double       'Array of column height values
                                          'used to determine hittest feature.

Private uColWidth         As Double       'The calculated width of each column.
Private uRowHeight        As Double       'The calculated height of each column.
Private uTopMargin        As Double         '--------------------------------------
Private uBottomMargin     As Double         'Margins used around the chart content.
Private uLeftMargin       As Double         '
Private uRightMargin      As Double         '--------------------------------------
Private uContentBorder    As Boolean      'Border around the chart content?
Private uSelectable       As Boolean      'Marker indicating whether user can select a column.
Private uHotTracking      As Boolean      'Marker indicating use of hot tracking.
Private uSelectedColumn   As Double         'Marker indicating the selected column.
Private uOldSelection     As Double
Private uDisplayDescript  As Boolean      'Display description when selectable
Private uChartTitle       As String       'Chart title
Private uChartSubTitle    As String       'Chart sub title
Private uDisplayXAxis     As Boolean      'Marker indicating display of x axis
Private uDisplayYAxis     As Boolean      'Marker indicating display of y axis
Private uColorBars        As Boolean      'Marker indicating use of different coloured bars
Private uIntersectMajor   As Double       'Major intersect value
Private uIntersectMinor   As Double       'Minor intersect value
Private uMaxYValue        As Double       'Default maximum y value
Private uXAxisLabel       As String       'Label to be displayed below the X-Axis
Private uYAxisLabel       As String       'Label to be displayed left of the Y-Axis
Private cItems            As Collection   'Collection of chart items

Private offsetX           As Double
Private offsetY           As Double

Private bLegendAdded      As Boolean
Private bLegendClicked    As Boolean
Private bDisplayLegend    As Boolean
Private bResize           As Boolean


Private bProcessingOver   As Boolean      'Marker to speed up mouse over effects.

Public Enum Theme
   [ThemePersianGulf] = 0
   [ThemeSky] = 1
   [ThemeNeon] = 2
   [ThemeNormal] = 3
End Enum

Private m_ActiveTheme      As Theme

Private IsDrawedOnce       As Boolean
Private IsInDrawMode       As Boolean

Private Colors(15, 1)      As Long
Private cItem()            As String

Public Event ItemClick(cItem As clsChartItem)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


Public Function AddItem(cItem As clsChartItem) As Boolean
    cItems.Add cItem
    If cItem.Value > uMaxYValue Then
      uMaxYValue = cItem.Value
    End If
End Function

Public Function EditCopy() As Boolean
    Clipboard.SetData UserControl.Image
End Function

Public Property Let MarginTop(lMargin As Double)
    uTopMargin = lMargin * Screen.TwipsPerPixelY
    DrawChart
End Property
Public Property Get MarginTop() As Double
    MarginTop = uTopMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginBottom(lMargin As Double)
    uBottomMargin = lMargin * Screen.TwipsPerPixelY
    DrawChart
End Property
Public Property Get MarginBottom() As Double
    MarginBottom = uBottomMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginLeft(lMargin As Double)
    uLeftMargin = lMargin * Screen.TwipsPerPixelX
    DrawChart
End Property
Public Property Get MarginLeft() As Double
    MarginLeft = uLeftMargin / Screen.TwipsPerPixelX
End Property

Public Property Let MarginRight(lMargin As Double)
    uRightMargin = lMargin * Screen.TwipsPerPixelX
    DrawChart
End Property
Public Property Get MarginRight() As Double
    MarginRight = uRightMargin / Screen.TwipsPerPixelX
End Property

Public Property Let ContentBorder(DisplayBorder As Boolean)
    uContentBorder = DisplayBorder
    DrawChart
End Property
Public Property Get ContentBorder() As Boolean
    ContentBorder = uContentBorder
End Property

Public Property Let Selectable(EnableSelection As Boolean)
    uSelectable = EnableSelection
    DrawChart
End Property
Public Property Get Selectable() As Boolean
    Selectable = uSelectable
End Property

Public Property Let HotTracking(UseHotTracking As Boolean)
    uHotTracking = UseHotTracking
    DrawChart
End Property
Public Property Get HotTracking() As Boolean
    HotTracking = uHotTracking
End Property

Public Property Let SelectedColumn(ColNumber As Long)
    Dim ret As Double
    Dim oItem As clsChartItem
    On Error Resume Next
    
    uSelectedColumn = ColNumber
    DrawChart
    
    ret = uColumns(ColNumber)
    If Err.Number Then
        uSelectedColumn = -1
    Else
        Set oItem = cItems(ColNumber + 1)
        RaiseEvent ItemClick(oItem)
    End If

End Property
Public Property Get SelectedColumn() As Long
    SelectedColumn = uSelectedColumn
End Property

Public Property Let ChartTitle(sTitle As String)
    uChartTitle = sTitle
    DrawChart
End Property
Public Property Get ChartTitle() As String
    ChartTitle = uChartTitle
End Property

Public Property Let ChartSubTitle(sTitle As String)
    uChartSubTitle = sTitle
    DrawChart
End Property
Public Property Get ChartSubTitle() As String
    ChartSubTitle = uChartSubTitle
End Property

Public Property Let IntersectMajor(ISValue As Double)
    uIntersectMajor = ISValue
    DrawChart
End Property
Public Property Get IntersectMajor() As Double
    IntersectMajor = uIntersectMajor
End Property

Public Property Let IntersectMinor(ISValue As Double)
    uIntersectMinor = ISValue
    DrawChart
End Property
Public Property Get IntersectMinor() As Double
    IntersectMinor = uIntersectMinor
End Property

Public Property Let DisplayYAxis(DisplayAxis As Boolean)
    uDisplayYAxis = DisplayAxis
    DrawChart
End Property
Public Property Get DisplayYAxis() As Boolean
    DisplayYAxis = uDisplayYAxis
End Property

Public Property Let DisplayXAxis(DisplayAxis As Boolean)
    uDisplayXAxis = DisplayAxis
    DrawChart
End Property
Public Property Get DisplayXAxis() As Boolean
    DisplayXAxis = uDisplayXAxis
End Property

Public Property Let MaxY(dMax As Double)
    uMaxYValue = dMax
    DrawChart
End Property
Public Property Get MaxY() As Double
    MaxY = uMaxYValue
End Property

Public Property Let SelectionInformation(DisplayInfo As Boolean)
    uDisplayDescript = DisplayInfo
    DrawChart
End Property
Public Property Get SelectionInformation() As Boolean
    SelectionInformation = uDisplayDescript
End Property

Public Property Let AxisLabelY(sCaption As String)
    uYAxisLabel = sCaption
    DrawChart
End Property
Public Property Get AxisLabelY() As String
    AxisLabelY = uYAxisLabel
End Property

Public Property Let AxisLabelX(sCaption As String)
    uXAxisLabel = sCaption
    DrawChart
End Property
Public Property Get AxisLabelX() As String
    AxisLabelX = uXAxisLabel
End Property

Public Property Let BackColor(hColor As OLE_COLOR)
    UserControl.BackColor = hColor
    DrawChart
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let ForeColor(hColor As OLE_COLOR)
    UserControl.ForeColor = hColor
    DrawChart
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ColorBars(bUseColor As Boolean)
    uColorBars = bUseColor
    DrawChart
End Property
Public Property Get ColorBars() As Boolean
    ColorBars = uColorBars
End Property

Private Sub lblDescription_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If uSelectable Then

            uSelectedColumn = Index
            uOldSelection = uSelectedColumn
            
            lScrollvalue = vsbContainer.Value
            
            bLegendClicked = True
            
            DrawChart
            
            bLegendClicked = False
        
            vsbContainer.Value = lScrollvalue
        End If
    End If
End Sub

Private Sub lblInfo_DblClick()
   lblInfo.Visible = False
   lblInfo.Tag = vbNullString
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        offsetX = X
        offsetY = Y
        lblInfo.Drag
        lblInfo.Tag = "Fix"
        mnuAutoMoveInfo.Checked = False
    Else
        PopupMenu mnuMain
    End If
End Sub

Private Sub mnuRefresh_Click()
    DrawChart
End Sub

Private Sub lblSlider_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend Not (bDisplayLegend)
    DrawChart
End Sub

Private Sub mnuAutoMoveInfo_Click()
   mnuAutoMoveInfo.Checked = Not mnuAutoMoveInfo.Checked
   lblInfo.Tag = IIf(mnuAutoMoveInfo.Checked, "", "Fix")
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.SetData UserControl.Image
End Sub

Private Sub mnuLegendHide_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend True
    DrawChart
End Sub



Private Sub mnuSaveAs_Click()
   Dim blnReturn As Long
   Dim strBuffer As String
   strBuffer = Space(255)
   blnReturn = SHGetSpecialFolderPath(0, _
      strBuffer, _
      CSIDL_MYPICTURES, _
      False)
      
   strBuffer = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
   
   
   
   Dim sFilters As String
   Dim OFN As OPENFILENAME
   Dim lret As Long
   
  'used after call
   Dim buff As String
   Dim sLname As String
   Dim sSname As String

  'create string of filters for the dialog
   sFilters = "Windows Bitmap" & vbNullChar & _
              "*.bmp" & vbNullChar & vbNullChar
  
   With OFN
      .nStructSize = Len(OFN)
      .hWndOwner = UserControl.hWnd
      .sFilter = sFilters
      .nFilterIndex = 0
      .sFile = "ActiveChart.bmp" & Space$(1024) & _
               vbNullChar & vbNullChar
      .nMaxFile = Len(.sFile)
      .sDefFileExt = "bmp" & vbNullChar & vbNullChar
      .sFileTitle = vbNullChar & Space$(512) & _
                    vbNullChar & vbNullChar
      .nMaxTitle = Len(OFN.sFileTitle)
      .sInitialDir = strBuffer & vbNullChar & vbNullChar
      .sDialogTitle = "VBnet GetSaveFileName Demo"
      .flags = OFS_FILE_SAVE_FLAGS

   End With
   
   
  'call the API
   blnReturn = GetSaveFileName(OFN)
   
   If blnReturn Then
      SavePicture UserControl.Image, OFN.sFile
   End If
End Sub

Private Sub mnuSelectionInfo_Click()
    mnuSelectionInfo.Checked = Not mnuSelectionInfo.Checked
    uDisplayDescript = mnuSelectionInfo.Checked
    DrawChart
End Sub

Private Sub mnuViewLegend_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend Not (bDisplayLegend)
    DrawChart
End Sub


Private Sub picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLegend
    End If
End Sub

Private Sub DrawContainer()
   Dim lColor As Long
   lColor = GetPixel(picLegend.hDC, 1, picContainer.Height / 15)
   
   picContainer.Cls
   
   Select Case m_ActiveTheme
      Case ThemePersianGulf
         DoGradient RGB(0, 100, 202), lColor, FillVer, 0, 0, picContainer.ScaleWidth / 15, picContainer.ScaleHeight / 15, picContainer.hDC
      Case ThemeNeon
         DoGradient RGB(75, 75, 75), lColor, FillVer, 0, 0, picContainer.ScaleWidth / 15, picContainer.ScaleHeight / 15, picContainer.hDC
      Case ThemeSky
         DoGradient RGB(185, 210, 239), lColor, FillVer, 0, 0, picContainer.ScaleWidth / 15, picContainer.ScaleHeight / 15, picContainer.hDC
   End Select
End Sub

Private Sub picLegend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLegend
    End If
End Sub

Private Sub picLegend_Resize()
  Call DrawLegend
End Sub

Private Sub tmrStart_Timer()
   IsDrawedOnce = False
   tmrStart.Enabled = False
   Call SetColors
   Call DrawChart
End Sub

Private Sub UserControl_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Left = X - offsetX
    Source.Top = Y - offsetY
End Sub

Private Sub UserControl_Initialize()
    Set cItems = New Collection
End Sub

Private Sub UserControl_InitProperties()
    Dim X As Integer
    Dim oChartItem As clsChartItem
    
    uTopMargin = 50 * Screen.TwipsPerPixelY
    uBottomMargin = 55 * Screen.TwipsPerPixelY
    uLeftMargin = 55 * Screen.TwipsPerPixelX
    uRightMargin = 55 * Screen.TwipsPerPixelX
    uContentBorder = True
    uSelectable = False
    uHotTracking = False
    uSelectedColumn = -1
    uOldSelection = -1
    uChartTitle = UserControl.Name
    uChartSubTitle = "Animated Chart"
    uDisplayYAxis = True
    uDisplayXAxis = True
    uColorBars = False
    uIntersectMajor = 10
    uIntersectMinor = 2
    uMaxYValue = 100
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X1 As Single
    Dim oItem As clsChartItem
    
    If IsInDrawMode Then GoTo TrackExit
    
    If Button = vbLeftButton Then
        X1 = (uColWidth)
        
        On Error GoTo TrackExit
        
        If (Y <= UserControl.ScaleHeight - uBottomMargin) And (uColumns((X - uLeftMargin) \ (X1)) <= Y) And uSelectable Then
            If Not bProcessingOver Then
                bProcessingOver = True
                uSelectedColumn = (X - uLeftMargin) \ (X1)
                If Not uSelectedColumn = uOldSelection Then
                    Cls
                    DrawChart
                    uOldSelection = uSelectedColumn
                    oItem = cItems(uSelectedColumn + 1)
                    RaiseEvent ItemClick(oItem)
                End If
    
                bProcessingOver = False
             End If
        End If
    ElseIf Button = vbRightButton Then
        mnuSelectionInfo.Visible = False
        If uSelectable Then
            mnuSelectionInfo.Visible = True
            mnuSeperator.Visible = True
        End If
        PopupMenu mnuMain
    End If
        
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
TrackExit:
    Exit Sub
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X1 As Long
    Dim oItem As clsChartItem
    X1 = (uColWidth)
    
    On Error GoTo TrackExit
    
    If IsInDrawMode Then GoTo TrackExit
    
    If uHotTracking Then
        If (Y <= UserControl.ScaleHeight - uBottomMargin) And (uColumns((X - uLeftMargin) \ (X1)) <= Y) And uSelectable Then
           If Not bProcessingOver Then
               bProcessingOver = True
               uSelectedColumn = (X - uLeftMargin) \ (X1)
               If Not uSelectedColumn = uOldSelection Then
                   Cls
                   DrawChart
                   uOldSelection = uSelectedColumn
               End If
               bProcessingOver = False
           End If
        Else
            If Not bProcessingOver Then
               bProcessingOver = True
               uSelectedColumn = -1
               If Not uSelectedColumn = uOldSelection Then
                  Cls
                  DrawChart
                  uOldSelection = uSelectedColumn
               End If
                bProcessingOver = False
           End If
        End If
    ElseIf Button = vbLeftButton Then
        If (Y <= UserControl.ScaleHeight - uBottomMargin) And (uColumns((X - uLeftMargin) \ (X1)) <= Y) And uSelectable Then
           If Not bProcessingOver Then
               bProcessingOver = True
               uSelectedColumn = (X - uLeftMargin) \ (X1)
               If Not uSelectedColumn = uOldSelection Then
                   Cls
                   DrawChart
                   uOldSelection = uSelectedColumn
                   oItem = cItems(uSelectedColumn + 1)
                   RaiseEvent ItemClick(oItem)
               End If
   

       
               bProcessingOver = False
           End If
        End If
    End If

TrackExit:

    Exit Sub
End Sub

Public Sub Refresh()
    DrawChart
End Sub

Public Sub Clear()
    Dim X As Integer
    
    Set cItems = Nothing
    Set cItems = New Collection
    If bLegendAdded Then
        ClearLegendItems
    End If
    DrawChart
End Sub

Public Function ShowLegend(Optional bHidden As Boolean = False)
    lblSlider.Height = picLegend.ScaleHeight
    'picLegend.Line (0, 0)-(picLegend.ScaleWidth - Screen.TwipsPerPixelX, picLegend.ScaleHeight - Screen.TwipsPerPixelY), &HFFE0E0, B
    
    If bHidden Then bDisplayLegend = False Else bDisplayLegend = True
    
    If bDisplayLegend Then
        uRightMargin = uRightMargin + picLegend.ScaleWidth
        picLegend.Move UserControl.ScaleWidth - picLegend.Width + Screen.TwipsPerPixelX, 0, picLegend.Width, UserControl.ScaleHeight
        DrawContainer
        lblSlider = Chr(187)
    Else
        uRightMargin = uRightMargin - picLegend.Width
        picLegend.Move UserControl.ScaleWidth - lblSlider.Width
        lblSlider = Chr(171)
    End If
End Function

Private Sub AddLegendItem(sDescription As String, ColorIndex As Long)
    Dim X As Integer
    Dim ShortDescript As String
    
    ShortDescript = sDescription
    If Len(ShortDescript) > 17 Then ShortDescript = Left(ShortDescript, 15) & ".."
    
    If bLegendAdded Then
        X = Box.Count
        Load Box(X)
        Load lblDescription(X)
        
        Box(X).BackColor = Colors(ColorIndex, 0)
        
        Box(X).Top = Box(X - 1).Top + Box(X - 1).Height + 10 * Screen.TwipsPerPixelY
        lblDescription(X).Top = Box(X).Top
                
        lblDescription(X) = ShortDescript
        lblDescription(X).ToolTipText = sDescription
    Else
        X = 0
        Box(X).BackColor = Colors(ColorIndex, 0)
        
        lblDescription(X) = ShortDescript
        lblDescription(X).ToolTipText = sDescription
        bLegendAdded = True
    End If
    
    DoGradient Colors(ColorIndex, 1), Colors(ColorIndex, 0), FillVer, 0, 0, Box(X).Width / 15, Box(X).Height / 15, Box(X).hDC
    DoGradient Colors(ColorIndex, 0), Colors(ColorIndex, 1), FillVer, 1, 1, Box(X).Width / 15 - 2, Box(X).Height / 15 - 2, Box(X).hDC
    Box(X).Visible = True
    lblDescription(X).Visible = True
            
    picContainer.Height = ((Box(0).Height + (10 * Screen.TwipsPerPixelY)) * Box.Count - 1) + 10 * Screen.TwipsPerPixelY
    If picContainer.ScaleHeight > picLegend.ScaleHeight Then
        vsbContainer.Max = (picContainer.ScaleHeight / Screen.TwipsPerPixelY) - (picLegend.ScaleHeight / Screen.TwipsPerPixelY)
        If Not vsbContainer.Visible Then vsbContainer.Visible = True
    Else
        vsbContainer.Visible = False
    End If
    
    
   
End Sub


Private Sub ClearLegendItems()
    Dim X As Integer
    
    On Error Resume Next    'we are expecting an error for item 1
    
    If bLegendAdded Then
        bLegendAdded = False
        
        For X = 1 To Box.Count
            Unload Box(X)
            Unload lblDescription(X)
            vsbContainer.Value = 0
            Box(0).Visible = False
            lblDescription(0).Visible = False
        Next X
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        uTopMargin = .ReadProperty("uTopMargin")
        uBottomMargin = .ReadProperty("uBottomMargin")
        uLeftMargin = .ReadProperty("uLeftMargin")
        uRightMargin = .ReadProperty("uRightMargin")
        uContentBorder = .ReadProperty("uContentBorder")
        uSelectable = .ReadProperty("uSelectable", False)
        uHotTracking = .ReadProperty("uHotTracking", False)
        uSelectedColumn = .ReadProperty("uSelectedColumn", -1)
        uChartTitle = .ReadProperty("uChartTitle", UserControl.Name)
        uChartSubTitle = .ReadProperty("uChartSubTitle", uChartSubTitle)
        uDisplayYAxis = .ReadProperty("uDisplayXAxis", uDisplayXAxis)
        uDisplayXAxis = .ReadProperty("uDisplayYAxis", uDisplayYAxis)
        uColorBars = .ReadProperty("uColorBars", False)
        uIntersectMajor = .ReadProperty("uIntersectMajor", 10)
        uIntersectMinor = .ReadProperty("uIntersectMinor", 2)
        uMaxYValue = .ReadProperty("uMaxYValue", 100)
        uDisplayDescript = .ReadProperty("uDisplayDescript", False)
        uXAxisLabel = .ReadProperty("uXAxisLabel")
        uYAxisLabel = .ReadProperty("uYAxisLabel")
        UserControl.BackColor = .ReadProperty("BackColor")
        UserControl.ForeColor = .ReadProperty("ForeColor")
        uOldSelection = -1
        m_ActiveTheme = .ReadProperty("ActiveTheme", ThemePersianGulf)
    End With
End Sub

Private Sub UserControl_Resize()
    If bDisplayLegend Then
        picLegend.Left = UserControl.ScaleWidth - picLegend.Width
    Else
        picLegend.Left = UserControl.ScaleWidth - lblSlider.Width
    End If
    picLegend.Height = UserControl.ScaleHeight
    vsbContainer.Height = picLegend.ScaleHeight
    lblSlider.Height = picLegend.ScaleHeight

    If IsDrawedOnce Then
      bResize = True
      DrawChart
      bResize = False
    End If


End Sub

Private Sub UserControl_Show()
    'DrawChart
    Call SetStyle
    
    UserControl.Cls
    DrawBackTheme

    tmrStart.Enabled = True
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    Dim oChartItm  As clsChartItem
    
    For Each oChartItm In cItems
      Set oChartItm = Nothing
    Next
    Set cItems = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "uTopMargin", uTopMargin
        .WriteProperty "uBottomMargin", uBottomMargin
        .WriteProperty "uLeftMargin", uLeftMargin
        .WriteProperty "uRightMargin", uRightMargin
        .WriteProperty "uContentBorder", uContentBorder
        .WriteProperty "uSelectable", uSelectable
        .WriteProperty "uHotTracking", uHotTracking
        .WriteProperty "uSelectedColumn", uSelectedColumn
        .WriteProperty "uChartTitle", uChartTitle
        .WriteProperty "uChartSubTitle", uChartSubTitle
        .WriteProperty "uDisplayXAxis", uDisplayXAxis
        .WriteProperty "uDisplayYAxis", uDisplayYAxis
        .WriteProperty "uColorBars", uColorBars
        .WriteProperty "uIntersectMajor", uIntersectMajor
        .WriteProperty "uIntersectMinor", uIntersectMinor
        .WriteProperty "uMaxYValue", uMaxYValue
        .WriteProperty "uDisplayDescript", uDisplayDescript
        .WriteProperty "uXAxisLabel", uXAxisLabel
        .WriteProperty "uYAxislabel", uYAxisLabel
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "ForeColor", UserControl.ForeColor
        .WriteProperty "ActiveTheme", m_ActiveTheme
    End With
End Sub

Private Sub vsbContainer_Change()
    picContainer.Top = -vsbContainer.Value * Screen.TwipsPerPixelY
End Sub

Private Sub vsbContainer_Scroll()
    picContainer.Top = -vsbContainer.Value * Screen.TwipsPerPixelY
End Sub

Private Function DoGradient(FromColor As Long, ToColor As Long, _
                     Optional DrawHorVer As GRADIENT_FILL_RECT = FillHor, _
                     Optional Left As Long = 0, Optional Top As Long = 0, _
                     Optional Width As Long = -1, _
                     Optional Height As Long = -1, _
                     Optional ByVal Drawhdc As Long = -1) As Boolean
    Dim Vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    Dim r As Byte, G As Byte, B As Byte
       
    Long2RGB FromColor, r, G, B
    With Vert(0)
        .X = Left
        .Y = Top
        .Red = Val("&h" & Hex(r) & "00")
        .Green = Val("&h" & Hex(G) & "00")
        .Blue = Val("&h" & Hex(B) & "00")
        .Alpha = 0&
    End With
    
    Long2RGB ToColor, r, G, B
    With Vert(1)
        .X = Left + Width
        .Y = Top + Height
        .Red = Val("&h" & Hex(r) & "00")
        .Green = Val("&h" & Hex(G) & "00")
        .Blue = Val("&h" & Hex(B) & "00")
        .Alpha = 0&
    End With

    gRect.UPPERLEFT = 0
    gRect.LOWERRIGHT = 1

    DoGradient = GradientFillRect(IIf(Drawhdc = -1, UserControl.hDC, Drawhdc), Vert(0), 2, gRect, 1, DrawHorVer)
    
End Function

Private Function Long2RGB(nColor As Long, Red As Byte, Green As Byte, Blue As Byte)
    Red = (nColor And &HFF&)
    Green = (nColor And &HFF00&) / &H100
    Blue = (nColor And &HFF0000) / &H10000
End Function


Private Sub DrawItem(ByVal ColorOne As Long, ByVal ColorTwo As Long, _
                     ByVal Left As Long, ByVal Top As Long, _
                     ByVal Width As Long, ByVal Height As Long, _
                     Optional ByVal Animated As Boolean = False)

   Select Case m_ActiveTheme
      Case ThemePersianGulf
         DoGradient ColorTwo, ColorOne, FillVer, Left, Top, Width, Height
         DoGradient ColorOne, ColorTwo, FillVer, Left + 2, Top + 2, Width - 4, Height - 4
      Case Else
         DoGradient ColorTwo, ColorOne, FillHor, Left, Top, Width, Height
         
         DoGradient ColorOne, ColorTwo, FillHor, Left + 2, Top + 2, (Width / 3) * 2 - 4, Height - 4
         DoGradient ColorTwo, ColorOne, FillHor, Left + (Width / 3) * 2 - 2, Top + 2, (Width / 3) - 1, Height - 4
   End Select
End Sub

Private Sub DrawAllItems()
   Dim i          As Double
   Dim Down       As Long
   Dim ColorOne   As Long
   Dim ColorTwo   As Long
   Dim Left       As Long
   Dim Top        As Long
   Dim Width      As Long
   Dim Height     As Long
   Dim Item       As Variant
   Dim lStep      As Long
   
   On Error GoTo Er
   For i = 1 To 10
      For j = 0 To UBound(cItem)
         Item = Split(cItem(j), "|")
         Left = Item(0): Top = Item(1): Width = Item(2): Height = Item(3)
         ColorOne = Item(4): ColorTwo = Item(5)
         
         Down = Top + Height
         
         If Height >= (Height / 10) * i Then _
            DrawItem ColorOne, ColorTwo, Left, Down - ((Height / 10) * i), Width, (Height / 10) * i
         
      Next j
      
      Tim = Timer
      Do While Timer - Tim < 0.07: Loop
      
      UserControl.Refresh
      
   Next i
   
Er:
   IsDrawedOnce = True
   
End Sub


Private Sub SetColors()
   Colors(0, 0) = RGB(185, 239, 255): Colors(0, 1) = RGB(30, 155, 230)
   Colors(1, 0) = RGB(255, 125, 79): Colors(1, 1) = RGB(129, 0, 0)
   Colors(2, 0) = RGB(0, 254, 0): Colors(2, 1) = RGB(0, 122, 0)
   Colors(3, 0) = RGB(233, 131, 255): Colors(3, 1) = RGB(214, 23, 255)
   Colors(4, 0) = RGB(95, 206, 255): Colors(4, 1) = RGB(0, 116, 210)
   Colors(5, 0) = RGB(255, 193, 66): Colors(5, 1) = RGB(185, 0, 0)
   Colors(6, 0) = RGB(215, 255, 168): Colors(6, 1) = RGB(99, 163, 23)
   Colors(7, 0) = RGB(201, 61, 154): Colors(7, 1) = RGB(153, 13, 106)
   Colors(8, 0) = RGB(0, 0, 254): Colors(8, 1) = RGB(0, 0, 122)
   Colors(9, 0) = RGB(255, 255, 160): Colors(9, 1) = RGB(250, 197, 12)
   
End Sub


Public Function GetYTopLegend(ByVal MaxChartValue As Long) As Long
   Dim Text    As String
   Dim MyStr   As String
   Dim Num     As Long
   
   Text = CStr(MaxChartValue)
   
   If Val(Text) > 10 Then
      MyStr = String(Len(Text) - 2, "0")
      
      Num = Val(Left(Text, 2))
      
      If Num Mod 10 = 0 Then
         MyStr = Num & MyStr
      ElseIf Num Mod 10 > 5 Then
         MyStr = CStr(Int(Num / 10) + 1) & "0" & MyStr
      Else
         MyStr = CStr(Int(Num / 10)) & "5" & MyStr
      End If
   Else
      MyStr = 10
   End If
   
   GetYTopLegend = CLng(MyStr)
End Function

Public Property Get ActiveTheme() As Theme
   ActiveTheme = m_ActiveTheme
End Property

Public Property Let ActiveTheme(ByVal NewTheme As Theme)
   m_ActiveTheme = NewTheme
   
   SetStyle

   PropertyChanged "ActiveTheme"
   
   DrawChart
End Property

Public Sub DrawChart()
    Dim CurrentColor    As Integer
    Dim iCols           As Integer
    Dim X               As Double
    Dim X1              As Double
    Dim X2              As Double
    Dim Y1              As Double
    Dim y2              As Double
    Dim xTemp           As Double
    Dim yTemp           As Double
    Dim sDescription    As String
    Dim oChartItem      As clsChartItem
    Dim lTopYValue      As Double
    
    If IsInDrawMode Then Exit Sub
    
    IsInDrawMode = True
    
    'If uIntersectMajor = 0 Then uIntersectMajor = 10
    'If uIntersectMinor = 0 Then uIntersectMinor = 2
    
    lTopYValue = GetYTopLegend(uMaxYValue)
    
    uIntersectMajor = lTopYValue / 10

    lblInfo.Visible = False
    lblDescription(0).ForeColor = UserControl.ForeColor
    
    iCols = cItems.Count
    
    mnuSelectionInfo.Checked = uDisplayDescript
    lblInfo.Visible = False
    If uDisplayDescript And uSelectedColumn > -1 And IsDrawedOnce Then lblInfo.Visible = True
    
    
    'Kill existing legend
    If bDisplayLegend Then
        vsbContainer.Visible = False
        picContainer.Visible = False
    End If
    
    If Not bResize Then ClearLegendItems
    
    uRowHeight = lTopYValue
    For X = 1 To cItems.Count
        Set oChartItem = cItems(X)
        If uRowHeight - CDbl(oChartItem.Value) < 0 Then uRowHeight = CDbl(oChartItem.Value)
    Next X
    
    If uRowHeight = 0 Then uRowHeight = 0.001
    
    If uMaxYValue < uRowHeight Then uMaxYValue = uRowHeight
    
    uRowHeight = ((UserControl.ScaleHeight - (uTopMargin + uBottomMargin)) / uRowHeight)
    If iCols Then uColWidth = ((UserControl.ScaleWidth - (uLeftMargin + uRightMargin)) / iCols)
    
    'UserControl.AutoRedraw = True
   
    DrawBackTheme
    
    
    If iCols Then ReDim uColumns(iCols - 1)

    On Error Resume Next
    'Intersect lines
    
    UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uChartTitle) / 2)
    UserControl.CurrentY = 0
    UserControl.FontBold = True
    UserControl.Print uChartTitle
    UserControl.FontBold = False
        
    UserControl.FontSize = UserControl.FontSize - 2
    UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uChartSubTitle) / 2)
    UserControl.Print uChartSubTitle
    UserControl.FontSize = UserControl.FontSize + 2
    
    If uDisplayYAxis Then
        Dim Counter  As Double
        Dim LastLine As Double
        For X = 0 To lTopYValue Step lTopYValue * 0.1
            X1 = uLeftMargin + (2 * Screen.TwipsPerPixelX): X2 = UserControl.ScaleWidth - uRightMargin
            Y1 = (UserControl.ScaleHeight - uBottomMargin) - (X * uRowHeight)
            If (X) Mod uIntersectMajor = 0 Then
                Counter = Counter + 1
                
                If Counter Mod 2 = 0 Then
                  DrawIntersect X1, Y1, X2, LastLine
                Else
                  LastLine = Y1
                End If

                UserControl.Line (X1, Y1)-(X2 + 1, Y1 + 15), GetThemeLineColor, BF
                UserControl.FontSize = UserControl.FontSize - 2
                UserControl.CurrentX = uLeftMargin - UserControl.TextWidth(X) - (5 * Screen.TwipsPerPixelX)
                UserControl.CurrentY = Y1 - (UserControl.TextHeight("0") / 2)
                UserControl.Print (X)
                UserControl.FontSize = UserControl.FontSize + 2
            End If
        Next X
    End If
   
    ReDim cItem(cItems.Count - 1)
    
    'On Error GoTo 0
    If uContentBorder Then
      UserControl.Line (uLeftMargin - 15, uTopMargin)-(uLeftMargin, UserControl.ScaleHeight - uBottomMargin), GetThemeLineColor, BF
    End If
    

    For X = 0 To cItems.Count - 1
        Set oChartItem = cItems(X + 1)
        
        X1 = (X * uColWidth) + uLeftMargin + (2 * Screen.TwipsPerPixelX)
        X2 = X1 + uColWidth - (2 * Screen.TwipsPerPixelX)
        Y1 = (UserControl.ScaleHeight - uBottomMargin) - (CDbl(oChartItem.Value) * uRowHeight)
        y2 = UserControl.ScaleHeight - uBottomMargin
                
        uColumns(X) = Y1
                     
        'Selected bar outline
        If X = uSelectedColumn And uSelectable And IsDrawedOnce Then
            'DrawItem RGB(254, 0, 0), RGB(122, 0, 0), (X1 + 1) / 15, Y1 / 15, (X2 - X1 - 1) / 15, (y2 - Y1) / 15, False
            DrawItem RGB(252, 233, 179), RGB(244, 192, 51), (X1 + 1) / 15, Y1 / 15, (X2 - X1 - 1) / 15, (y2 - Y1) / 15, False
            'Add Legend item
            If Not bResize Then AddLegendItem oChartItem.SelectedDescription, 9
                                                      
            If uDisplayDescript Then
                lblInfo.Visible = False
                lblInfo = "Item: " & oChartItem.XAxisDescription & vbCr & "Value: " & Format(oChartItem.Value, "#,0") & vbCr & oChartItem.SelectedDescription
                If lblInfo.Tag <> "Fix" Then lblInfo.Move X1 + ((X2 - X1) - lblInfo.Width) / 2, y2 + 20
                If IsDrawedOnce Then lblInfo.Visible = True
            End If
        Else
            CurrentColor = (oChartItem.ItemID - 1) Mod 10
            
            If Not IsDrawedOnce Then
               cItem(X) = (X1 + 1) / 15 & "|" & Y1 / 15 & "|" & (X2 - X1 - 1) / 15 & "|" & (y2 - Y1) / 15 & "|" & IIf(uColorBars, Colors(CurrentColor, 0), Colors(2, 0)) & "|" & IIf(uColorBars, Colors(CurrentColor, 1), Colors(2, 1))
            Else
               DrawItem IIf(uColorBars, Colors(CurrentColor, 0), Colors(2, 0)), IIf(uColorBars, Colors(CurrentColor, 1), Colors(2, 1)), (X1 + 1) / 15, Y1 / 15, (X2 - X1 - 1) / 15, (y2 - Y1) / 15, Not IsDrawedOnce
            End If
            
            
            'Add Legend item
            If Not bResize Then AddLegendItem oChartItem.SelectedDescription, IIf(uColorBars, CurrentColor, 1)
            
            'CurrentColor = CurrentColor + 1
            'If CurrentColor >= 10 Then CurrentColor = 0
        End If
        
        If uDisplayXAxis Then
            UserControl.FontSize = UserControl.FontSize - 1
            
            xTemp = (((X2 - X1) / 2) + X1) / Screen.TwipsPerPixelX
            yTemp = (UserControl.ScaleHeight - uBottomMargin + UserControl.TextWidth(oChartItem.XAxisDescription) / 1.25) / Screen.TwipsPerPixelY
            
            PrintRotText UserControl.hDC, oChartItem.XAxisDescription, xTemp, yTemp, 270
            
            UserControl.Line (X1 - 1 * Screen.TwipsPerPixelX, y2)-(X1 - 1 * Screen.TwipsPerPixelX, y2 + UserControl.TextHeight(oChartItem.XAxisDescription) / 2), GetThemeLineColor
            UserControl.FontSize = UserControl.FontSize + 1
        End If
        
    Next X
    
    If Not IsDrawedOnce Then DrawAllItems
    
    'Print the x axis label
    If Len(uXAxisLabel) Then
        UserControl.FontSize = UserControl.FontSize - 1
        UserControl.CurrentY = UserControl.ScaleHeight - UserControl.TextHeight(uXAxisLabel) * 1.5
        UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uXAxisLabel) / 2)
        UserControl.Print uXAxisLabel
        UserControl.FontSize = UserControl.FontSize + 1
    End If
    
    'Print the y axis label
    If Len(uYAxisLabel) Then
        UserControl.FontSize = UserControl.FontSize - 1
        PrintRotText UserControl.hDC, uYAxisLabel, UserControl.TextHeight(uYAxisLabel) / Screen.TwipsPerPixelX, UserControl.ScaleHeight / 2 / Screen.TwipsPerPixelY, 90
        UserControl.FontSize = UserControl.FontSize + 1
    End If

    If bDisplayLegend Then
        If uSelectable And uSelectedColumn > -1 Then
            Dim perScreen As Integer
            Dim scrollValue As Integer
                        
            perScreen = Abs((picLegend.ScaleHeight / ((Box(0).Height + (10 * Screen.TwipsPerPixelY)))) - 1)
                        
            If (uSelectedColumn + 1) > perScreen Then
                scrollValue = ((uSelectedColumn + 1) * ((Box(0).Height / Screen.TwipsPerPixelY) + 10)) - (Box(perScreen).Top / Screen.TwipsPerPixelY)
                If scrollValue > vsbContainer.Max Then scrollValue = vsbContainer.Max
                vsbContainer.Value = scrollValue
            Else
                vsbContainer.Value = 0
            End If
                        
            DrawContainer
            picContainer.Line ((Box(uSelectedColumn).Left - 3 * Screen.TwipsPerPixelX), (Box(uSelectedColumn).Top - 3 * Screen.TwipsPerPixelY))-(lblDescription(uSelectedColumn).Left + lblDescription(uSelectedColumn).Width + 2 * Screen.TwipsPerPixelX, Box(uSelectedColumn).Top + Box(uSelectedColumn).Height + 2 * Screen.TwipsPerPixelY), vbWhite, B
            
        End If
        picContainer.Visible = True
    End If
    
    IsInDrawMode = False
    
End Sub
Public Sub DrawBackTheme()
   Dim lWidth     As Long
   Dim lHeight    As Long
   lWidth = (UserControl.ScaleWidth) / Screen.TwipsPerPixelX
   lHeight = (UserControl.ScaleHeight / Screen.TwipsPerPixelY)
   
    UserControl.Cls
    Select Case m_ActiveTheme
      Case ThemePersianGulf
         DoGradient RGB(0, 3, 102), RGB(0, 100, 202), FillVer, 0, 0, lWidth, lHeight, UserControl.hDC
      Case ThemeSky
         DoGradient RGB(158, 190, 230), RGB(185, 210, 239), FillVer, 0, 0, lWidth, lHeight
      Case ThemeNeon
         DoGradient RGB(0, 0, 0), RGB(75, 75, 75), FillVer, 0, 0, lWidth, lHeight
   End Select
End Sub

Public Sub DrawIntersect(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal LastLine As Long)
   Dim lHeight As Long
   Dim lWidth  As Long
   Dim lLeft   As Long
   Dim lTop    As Long
   
   lHeight = ((Y1 - LastLine) / 15) - 1
   lWidth = (X2 - X1) / 15 + 1
   lLeft = X1 / 15
   lTop = (Y1 / 15) + 1
   Select Case m_ActiveTheme
      Case ThemePersianGulf
         DoGradient RGB(0, 54, 144), RGB(0, 59, 149), FillVer, lLeft, lTop, lWidth, lHeight
         UserControl.Line (X1, Y1)-(X2 + 1, Y1 + 30), RGB(0, 129, 199), BF
      Case ThemeSky
         DoGradient RGB(227, 239, 255), RGB(201, 224, 255), FillVer, lLeft, lTop, lWidth, (lHeight / 2)
         DoGradient RGB(183, 214, 255), RGB(190, 218, 255), FillVer, lLeft, lTop + (lHeight / 2), lWidth, lHeight - (lHeight / 2) + 1
      Case ThemeNeon
         DoGradient RGB(66, 70, 81), RGB(58, 61, 69), FillVer, lLeft, lTop, lWidth, (lHeight / 8) * 3
         DoGradient RGB(46, 47, 47), RGB(59, 59, 59), FillVer, lLeft, lTop + (lHeight / 8) * 3, lWidth, (lHeight / 8) * 4
         DoGradient RGB(68, 68, 68), RGB(75, 75, 75), FillVer, lLeft, lTop + ((lHeight / 8) * 7) - 1, lWidth, lHeight - (lHeight / 8) * 7 + 1
   End Select
End Sub

Public Function GetThemeLineColor() As Long
   Select Case m_ActiveTheme
      Case ThemePersianGulf
         GetThemeLineColor = RGB(0, 129, 199)
      Case ThemeNeon
         GetThemeLineColor = RGB(40, 40, 40)
      Case ThemeSky
         GetThemeLineColor = RGB(141, 178, 227) 'RGB(173, 209, 255) '
   End Select
End Function

Private Sub DrawLegend()
   picLegend.Cls
   
   Select Case m_ActiveTheme
      Case ThemePersianGulf
         DoGradient RGB(0, 100, 202), RGB(0, 3, 102), FillVer, 0, 0, picLegend.ScaleWidth / 15, picLegend.ScaleHeight / 15, picLegend.hDC
      Case ThemeNeon
         DoGradient RGB(75, 75, 75), RGB(0, 0, 0), FillVer, 0, 0, picLegend.ScaleWidth / 15, picLegend.ScaleHeight / 15, picLegend.hDC
      Case ThemeSky
         DoGradient RGB(185, 210, 239), RGB(158, 190, 230), FillVer, 0, 0, picLegend.ScaleWidth / 15, picLegend.ScaleHeight / 15, picLegend.hDC
   End Select
End Sub

Private Sub SetStyle()
    Select Case m_ActiveTheme
      Case ThemePersianGulf
         lblSlider.BackColor = &H400000
         lblSlider.ForeColor = vbWhite
         UserControl.ForeColor = vbWhite
      Case ThemeSky
         lblSlider.BackColor = RGB(158, 190, 230)
         lblSlider.ForeColor = vbBlack
         UserControl.ForeColor = vbBlack 'RGB(131, 200, 240)
      Case ThemeNeon
         lblSlider.BackColor = vbBlack
         lblSlider.ForeColor = vbWhite
         UserControl.ForeColor = vbWhite
   End Select
End Sub

