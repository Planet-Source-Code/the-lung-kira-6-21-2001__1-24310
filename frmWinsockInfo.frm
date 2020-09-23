VERSION 5.00
Begin VB.Form frmWinsockInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Info"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmWinsockInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   350
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.ListBox lstProtocols 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.ListBox lstServices 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblProtocols 
      Caption         =   "Protocols"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblServices 
      Caption         =   "Services"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmWinsockInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetData_Click()
    cmdGetData.Enabled = False
    
    Dim lngIncrement As Long
    Dim lngReturn As Long
    Dim strName As String
    
    Dim protoent As protoent
    Dim servent As servent
    For lngIncrement = 0 To 65535
        lngReturn = getprotobynumber(lngIncrement)
        If lngReturn > 0 Then
            CopyMemory protoent, ByVal lngReturn, Len(protoent)
            
            strName = String$(255, &H0)
            CopyMemory ByVal strName, ByVal protoent.p_name, 255
            
            lstProtocols.AddItem Left$(CStr(lngIncrement) & Space$(7), 7) & strName
        End If
        
        lngReturn = getservbyport(lngIncrement, &H0)
        If lngReturn > 0 Then
            CopyMemory servent, ByVal lngReturn, Len(servent)
            
            strName = String$(255, &H0)
            CopyMemory ByVal strName, ByVal servent.s_name, 255
            
            lstServices.AddItem Left$(CStr(lngIncrement) & Space$(7), 7) & strName
        End If
        
        DoEvents
    Next lngIncrement
End Sub

Private Sub Form_Load()
    If WS2 = False Then
        lblProtocols.Enabled = False
        lstProtocols.Enabled = False
        lblServices.Enabled = False
        lstServices.Enabled = False
        cmdGetData.Enabled = False
        Exit Sub
    End If
End Sub
