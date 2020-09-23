VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "AnimatedChart Test"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin ChartTest.AnimatedChart ActiveChart1 
      Height          =   5085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8969
      uTopMargin      =   750
      uBottomMargin   =   825
      uLeftMargin     =   825
      uRightMargin    =   825
      uContentBorder  =   -1  'True
      uSelectable     =   -1  'True
      uHotTracking    =   -1  'True
      uSelectedColumn =   -1
      uChartTitle     =   "AnimatedChart"
      uChartSubTitle  =   "Animated Chart"
      uDisplayXAxis   =   -1  'True
      uDisplayYAxis   =   -1  'True
      uColorBars      =   -1  'True
      uIntersectMajor =   10
      uIntersectMinor =   2
      uMaxYValue      =   100
      uDisplayDescript=   0   'False
      uXAxisLabel     =   ""
      uYAxislabel     =   ""
      BackColor       =   -2147483643
      ForeColor       =   16777215
      ActiveTheme     =   0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim X As Integer, oChartItem As clsChartItem
   
    Randomize
    
    'ActiveChart1.IntersectMajor = 1000

    For X = 1 To 4
         Set oChartItem = New clsChartItem
        oChartItem.ItemID = X
        oChartItem.SelectedDescription = "Total Sale Of 200" & X
        oChartItem.Value = CLng(Rnd * 2000)
        oChartItem.XAxisDescription = "Bar " & X
        ActiveChart1.AddItem oChartItem
            

    Next X
   

End Sub

Private Sub Form_Resize()
    
   ActiveChart1.Width = Me.ScaleWidth
   ActiveChart1.Height = Me.ScaleHeight
    
End Sub

Private Sub grd_Click()
    DoEvents
    ActiveChart1.SelectedColumn = grd.Row - 1
End Sub

