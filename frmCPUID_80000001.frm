VERSION 5.00
Begin VB.Form frmCPUID_80000001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID 80000001"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmCPUID_80000001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFeatures 
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
      TabIndex        =   15
      Top             =   2280
      Width           =   6135
   End
   Begin VB.TextBox txtFamily 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtModel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtStepping 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
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
      Left            =   3720
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
      Left            =   3720
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
      Left            =   3720
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblFeatures 
      Caption         =   "Features"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblFamily 
      Caption         =   "Family"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblModel 
      Caption         =   "Model"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblStepping 
      Caption         =   "Stepping"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
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
Attribute VB_Name = "frmCPUID_80000001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If MaxExtCPUIDLevel > strtoul_("80000000", 16) Then
        'EAX = 80000001
        
        Dim strRegister As String
        
        Dim outEAX As Long
        Dim outEBX As Long
        Dim outECX As Long
        Dim outEDX As Long
        
        
        cpuid_ strtoul_("80000001", 16), outEAX, outEBX, outECX, outEDX
        
        
        strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outEAX, 2), 32)))
        
        txtStepping.Text = Right$("0000" & StrReverse(Mid$(strRegister, 1, 4)), 4)
        txtModel.Text = Right$("0000" & StrReverse(Mid$(strRegister, 5, 4)), 4)
        txtFamily.Text = Right$("0000" & StrReverse(Mid$(strRegister, 9, 4)), 4)
        
        
        strRegister = StrReverse(CStr(Right$(String$(32, "0") & ltoa_(outEDX, 2), 32)))
        
        With lstFeatures
            .AddItem "0   " & Left$("Floating Point Unit on chip" & Space$(45), 45) & CBool(Mid$(strRegister, 1, 1))
            .AddItem "1   " & Left$("Virtual 8086 Mode Extension" & Space$(45), 45) & CBool(Mid$(strRegister, 2, 1))
            .AddItem "2   " & Left$("Debugging Extension" & Space$(45), 45) & CBool(Mid$(strRegister, 3, 1))
            .AddItem "3   " & Left$("Page Size Extension" & Space$(45), 45) & CBool(Mid$(strRegister, 4, 1))
            .AddItem "4   " & Left$("Time Stamp Counter" & Space$(45), 45) & CBool(Mid$(strRegister, 5, 1))
            .AddItem "5   " & Left$("Model Specific Registers" & Space$(45), 45) & CBool(Mid$(strRegister, 6, 1))
            .AddItem "6   " & Left$("Physical Address Extension" & Space$(45), 45) & CBool(Mid$(strRegister, 7, 1))
            .AddItem "7   " & Left$("Machine Check Exception" & Space$(45), 45) & CBool(Mid$(strRegister, 8, 1))
            .AddItem "8   " & Left$("CMPXCHG8 Instruction" & Space$(45), 45) & CBool(Mid$(strRegister, 9, 1))
            .AddItem "9   " & Left$("On Chip APIC" & Space$(45), 45) & CBool(Mid$(strRegister, 10, 1))
            .AddItem "10  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 11, 1))
            .AddItem "11  " & Left$("Fast System Call (SEP)" & Space$(45), 45) & CBool(Mid$(strRegister, 12, 1))
            .AddItem "12  " & Left$("Memory Type Range Registers" & Space$(45), 45) & CBool(Mid$(strRegister, 13, 1))
            .AddItem "13  " & Left$("Page Global Enable" & Space$(45), 45) & CBool(Mid$(strRegister, 14, 1))
            .AddItem "14  " & Left$("Machine Check Architecture" & Space$(45), 45) & CBool(Mid$(strRegister, 15, 1))
            .AddItem "15  " & Left$("Conditional Move and Compare Instructions" & Space$(45), 45) & CBool(Mid$(strRegister, 16, 1))
            .AddItem "16  " & Left$("Page Attribute Table" & Space$(45), 45) & CBool(Mid$(strRegister, 17, 1))
            .AddItem "17  " & Left$("36bit Page Size Extension" & Space$(45), 45) & CBool(Mid$(strRegister, 18, 1))
            .AddItem "18  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 19, 1))
            .AddItem "19  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 20, 1))
            .AddItem "20  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 21, 1))
            .AddItem "21  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 22, 1))
            .AddItem "22  " & Left$("MMX+ Technology" & Space$(45), 45) & CBool(Mid$(strRegister, 23, 1))
            .AddItem "23  " & Left$("MMX Technology" & Space$(45), 45) & CBool(Mid$(strRegister, 24, 1))
            .AddItem "24  " & Left$("Fast Save and Restor Instructions" & Space$(45), 45) & CBool(Mid$(strRegister, 25, 1))
            .AddItem "25  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 26, 1))
            .AddItem "26  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 27, 1))
            .AddItem "27  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 28, 1))
            .AddItem "28  " & Left$("Reserved" & Space$(45), 45) & CBool(Mid$(strRegister, 29, 1))
            .AddItem "29  " & Left$("AA-64 Architecture" & Space$(45), 45) & CBool(Mid$(strRegister, 30, 1))
            .AddItem "30  " & Left$("3DNow!+ Technology" & Space$(45), 45) & CBool(Mid$(strRegister, 31, 1))
            .AddItem "31  " & Left$("3DNow! Technology" & Space$(45), 45) & CBool(Mid$(strRegister, 32, 1))
        End With
        
        
        txtEAX.Text = Right$("00000000" & ltoa_(outEAX, 16), 8)
        txtEBX.Text = Right$("00000000" & ltoa_(outEBX, 16), 8)
        txtECX.Text = Right$("00000000" & ltoa_(outECX, 16), 8)
        txtEDX.Text = Right$("00000000" & ltoa_(outEDX, 16), 8)
    End If
End Sub
