VERSION 5.00
Begin VB.Form frmIconSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Settings"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmIconSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHorzSpc 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtVertSpc 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox chkTitleWrap 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblVertSpc 
      Caption         =   "Vertical Spacing"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblHorzSpc 
      Caption         =   "Horizontal Spacing"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblTitleWrap 
      Caption         =   "Title Wrap"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmIconSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim ICONMETRICS As ICONMETRICS
    ICONMETRICS.cbSize = Len(ICONMETRICS)
    
    If SystemParametersInfo(SPI_GETICONMETRICS, Len(ICONMETRICS), ICONMETRICS, 0) = 0 Then Failed "SystemParametersInfo"
    
    With ICONMETRICS
        .iHorzSpacing = Val(txtHorzSpc.Text)
        .iVertSpacing = Val(txtVertSpc.Text)
        .iTitleWrap = chkTitleWrap.value
    End With
    
    If SystemParametersInfo(SPI_SETICONMETRICS, Len(ICONMETRICS), ICONMETRICS, SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    Dim ICONMETRICS As ICONMETRICS
    ICONMETRICS.cbSize = Len(ICONMETRICS)
    
    If SystemParametersInfo(SPI_GETICONMETRICS, Len(ICONMETRICS), ICONMETRICS, 0) = 0 Then Failed "SystemParametersInfo"
    
    With ICONMETRICS
        txtHorzSpc.Text = CStr(.iHorzSpacing)
        txtVertSpc.Text = CStr(.iVertSpacing)
        chkTitleWrap.value = CStr(.iTitleWrap)
    End With
End Sub

Private Sub txtHorzSpc_Change()
    txtHorzSpc.Text = CStr(Val(Rem_NonNumeric_Chr(txtHorzSpc.Text)))
    If Val(txtHorzSpc.Text) < 1 Then txtHorzSpc.Text = "1"
    If Val(txtHorzSpc.Text) > 65535 Then txtHorzSpc.Text = "65535"
End Sub

Private Sub txtVertSpc_Change()
    txtVertSpc.Text = CStr(Val(Rem_NonNumeric_Chr(txtVertSpc.Text)))
    If Val(txtVertSpc.Text) < 1 Then txtVertSpc.Text = "1"
    If Val(txtVertSpc.Text) > 65535 Then txtVertSpc.Text = "65535"
End Sub
