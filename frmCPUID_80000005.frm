VERSION 5.00
Begin VB.Form frmCPUID_80000005 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID 80000005"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmCPUID_80000005.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCacheTLB 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   6615
   End
   Begin VB.TextBox txtEDX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtECX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtEBX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtEAX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblCacheTLB 
      Caption         =   "Cache - TLB"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblEBX 
      Caption         =   "EBX"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblEAX 
      Caption         =   "EAX"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblEDX 
      Caption         =   "EDX"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblECX 
      Caption         =   "ECX"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmCPUID_80000005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If MaxExtCPUIDLevel > strtoul_("80000004", 16) Then
        'EAX = 80000005
        
        Dim strRegister As String
        
        Dim outEAX As Long
        Dim outEBX As Long
        Dim outECX As Long
        Dim outEDX As Long
        
        
        cpuid_ strtoul_("80000005", 16), outEAX, outEBX, outECX, outEDX
        
        
        With lstCacheTLB
            strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outEAX, 2), 32)))
            
            .AddItem "L1 Cache And TLB Configuration Descriptors "
            .AddItem Left$("Code TLB Entries" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 1, 12)), 2)
            .AddItem Left$("Code TLB Associativity" & String$(45, " "), 45) & StrReverse(Mid$(strRegister, 13, 4))
            .AddItem Left$("Data TLB Entries" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 17, 12)), 2)
            .AddItem Left$("Data TLB Associativity" & String$(45, " "), 45) & StrReverse(Mid$(strRegister, 29, 4))
            .AddItem ""
            
            
            strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outEBX, 2), 32)))
            
            .AddItem "4 KB L1 TLB Configuration Descriptor"
            .AddItem Left$("Code TLB Entries" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 1, 12)), 2)
            .AddItem Left$("Code TLB Associativity" & String$(45, " "), 45) & StrReverse(Mid$(strRegister, 13, 4))
            .AddItem Left$("Data TLB Entries" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 17, 12)), 2)
            .AddItem Left$("Data TLB Associativity" & String$(45, " "), 45) & StrReverse(Mid$(strRegister, 29, 4))
            .AddItem ""
            
            
            strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outECX, 2), 32)))
            
            .AddItem "Data L1 Cache Configuration Descriptor"
            .AddItem Left$("Data L1 Cache Line Size In Bytes" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 1, 8)), 2)
            .AddItem Left$("Data L1 Cache Lines Per Tag" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 9, 4)), 2)
            .AddItem Left$("Data L1 Cache Associativity" & String$(45, " "), 45) & StrReverse(Mid$(strRegister, 13, 4))
            .AddItem Left$("Data L1 Cache Size In KBs" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 17, 16)), 2)
            .AddItem ""
            
            
            strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outEDX, 2), 32)))
            
            .AddItem "Code L1 Cache Configuration Descriptor"
            .AddItem Left$("Code L1 Cache Line Size In Bytes" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 1, 8)), 2)
            .AddItem Left$("Code L1 Cache Lines Per Tag" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 9, 4)), 2)
            .AddItem Left$("Code L1 Cache Associativity" & String$(45, " "), 45) & StrReverse(Mid$(strRegister, 13, 4))
            .AddItem Left$("Code L1 Cache Size In KBs" & String$(45, " "), 45) & strtol_(StrReverse(Mid$(strRegister, 17, 16)), 2)
        End With
        
        
        txtEAX.Text = Right$("00000000" & ltoa_(outEAX, 16), 8)
        txtEBX.Text = Right$("00000000" & ltoa_(outEBX, 16), 8)
        txtECX.Text = Right$("00000000" & ltoa_(outECX, 16), 8)
        txtEDX.Text = Right$("00000000" & ltoa_(outEDX, 16), 8)
    End If
End Sub
