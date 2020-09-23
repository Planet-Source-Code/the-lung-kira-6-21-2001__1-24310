VERSION 5.00
Begin VB.Form frmCPUID_00000002 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID 00000002"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmCPUID_00000002.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQueriesRequired 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtEDX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
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
      Left            =   4560
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox lstCacheTLB 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   6975
   End
   Begin VB.TextBox txtEAX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblQueriesRequired 
      Caption         =   "Queries Required"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblCacheTLB 
      Caption         =   "Cache - TLB"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
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
   Begin VB.Label lblEDX 
      Caption         =   "EDX"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
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
   Begin VB.Label lblEBX 
      Caption         =   "EBX"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmCPUID_00000002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If MaxCPUIDLevel > &H1 Then
        'EAX = 2
        
        Dim strRegister As String
        Dim lngIncrement As Long
        Dim lngQuery As Long
        
        Dim outEAX As Long
        Dim outEBX As Long
        Dim outECX As Long
        Dim outEDX As Long
        
        
        cpuid_ 2, outEAX, outEBX, outECX, outEDX
        
        strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outEAX, 2), 32)))
        
        
        lngQuery = strtol_(StrReverse(Mid$(strRegister, 1, 8)), 2)
        
        For lngIncrement = 1 To lngQuery
            cpuid_ 2, outEAX, outEBX, outECX, outEDX
            DoEvents
        Next lngIncrement
        
        
        txtQueriesRequired.Text = CStr(lngQuery)
        
        With lstCacheTLB
            .AddItem StrReverse(Mid$(strRegister, 9, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 9, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 17, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 17, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 25, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 25, 8)), 2))
            
            strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outEBX, 2), 32)))
            
            .AddItem StrReverse(Mid$(strRegister, 1, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 1, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 9, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 9, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 17, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 17, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 25, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 25, 8)), 2))
            
            strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outECX, 2), 32)))
            
            .AddItem StrReverse(Mid$(strRegister, 1, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 1, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 9, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 9, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 17, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 17, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 25, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 25, 8)), 2))
            
            strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outEDX, 2), 32)))
            
            .AddItem StrReverse(Mid$(strRegister, 1, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 1, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 9, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 9, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 17, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 17, 8)), 2))
            .AddItem StrReverse(Mid$(strRegister, 25, 8)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(strRegister, 25, 8)), 2))
        End With
        
        
        txtEAX.Text = Right$("00000000" & ltoa_(outEAX, 16), 8)
        txtEBX.Text = Right$("00000000" & ltoa_(outEBX, 16), 8)
        txtECX.Text = Right$("00000000" & ltoa_(outECX, 16), 8)
        txtEDX.Text = Right$("00000000" & ltoa_(outEDX, 16), 8)
    End If
End Sub

Private Function CacheTLB_Select(ByVal lngValue As Long) As String
    Dim strDescriptor As String
    
    Select Case lngValue
        Case &H0: strDescriptor = "Null Descriptor"
        Case &H1: strDescriptor = "code TLB, 4K pages, 4 ways, 32 entries"
        Case &H2: strDescriptor = "code TLB, 4M pages, fully, 2 entries"
        Case &H3: strDescriptor = "data TLB, 4K pages, 4 ways, 64 entries"
        Case &H4: strDescriptor = "data TLB, 4M pages, 4 ways, 8 entries"
        Case &H6: strDescriptor = "code L1 cache, 8KB, 4 ways, 32 byte lines"
        Case &H8: strDescriptor = "code L1 cache, 16KB, 4 ways, 32 byte lines"
        Case &HA: strDescriptor = "data L1 cache, 8KB, 2 ways, 32 byte lines"
        Case &HC: strDescriptor = "data L1 cache, 16KB, 4 ways, 32 byte lines"
        Case &H22: strDescriptor = "code & data L3 cache, 512KB, 4 ways (!), 64 byte lines, sectored"
        Case &H23: strDescriptor = "code & data L3 cache, 1024KB, 8 ways, 64 byte lines, sectored"
        Case &H25: strDescriptor = "code & data L3 cache, 2048KB, 8 ways, 64 byte lines, sectored"
        Case &H29: strDescriptor = "code & data L3 cache, 4096KB, 8 ways, 64 byte lines, sectored"
        Case &H40: strDescriptor = "no integrated L2 cache (P6 core) or L3 cache (P4 core)"
        Case &H41: strDescriptor = "code & data L2 cache, 128KB, 4 ways, 32 byte lines"
        Case &H42: strDescriptor = "code & data L2 cache, 256KB, 4 ways, 32 byte lines"
        Case &H43: strDescriptor = "code & data L2 cache, 512KB, 4 ways, 32 byte lines"
        Case &H44: strDescriptor = "code & data L2 cache, 1024KB, 4 ways, 32 byte lines"
        Case &H45: strDescriptor = "code & data L2 cache, 2048KB, 4 ways, 32 byte lines"
        Case &H50: strDescriptor = "code TLB, 4K/4M/2M pages, fully, 64 entries"
        Case &H51: strDescriptor = "code TLB, 4K/4M/2M pages, fully, 128 entries"
        Case &H52: strDescriptor = "code TLB, 4K/4M/2M pages, fully, 256 entries"
        Case &H5B: strDescriptor = "data TLB, 4K/4M pages, fully, 64 entries"
        Case &H5C: strDescriptor = "data TLB, 4K/4M pages, fully, 128 entries"
        Case &H5D: strDescriptor = "data TLB, 4K/4M pages, fully, 256 entries"
        Case &H66: strDescriptor = "data L1 cache, 8KB, 4 ways, 64 byte lines, sectored"
        Case &H67: strDescriptor = "data L1 cache, 16KB, 4 ways, 64 byte lines, sectored"
        Case &H68: strDescriptor = "data L1 cache, 32KB, 4 ways, 64 byte lines, sectored"
        Case &H70: strDescriptor = "trace L1 cache, 12 KµOPs, 4 ways"
        Case &H71: strDescriptor = "trace L1 cache, 16 KµOPs, 4 ways"
        Case &H72: strDescriptor = "trace L1 cache, 32 KµOPs, 4 ways"
        Case &H79: strDescriptor = "code & data L2 cache, 128KB, 8 ways, 64 byte lines, sectored"
        Case &H7A: strDescriptor = "code & data L2 cache, 256KB, 8 ways, 64 byte lines, sectored"
        Case &H7B: strDescriptor = "code & data L2 cache, 512KB, 8 ways, 64 byte lines, sectored"
        Case &H7C: strDescriptor = "code & data L2 cache, 1024KB, 8 ways, 64 byte lines, sectored"
        Case &H81: strDescriptor = "code & data L2 cache, 128KB, 8 ways, 32 byte lines"
        Case &H82: strDescriptor = "code & data L2 cache, 256KB, 8 ways, 32 byte lines"
        Case &H83: strDescriptor = "code & data L2 cache, 512KB, 8 ways, 32 byte lines"
        Case &H84: strDescriptor = "code & data L2 cache, 1024KB, 8 ways, 32 byte lines"
        Case &H85: strDescriptor = "code & data L2 cache, 2048KB, 8 ways, 32 byte lines"
        Case Else: strDescriptor = "Unknown"
    End Select
    
    CacheTLB_Select = strDescriptor
End Function
