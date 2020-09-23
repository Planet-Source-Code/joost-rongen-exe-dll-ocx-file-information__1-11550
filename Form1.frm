VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FileInfo   EXE / DLL / OCX"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Get FileInfo"
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   3000
      Width           =   6375
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   2685
      Width           =   6375
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2370
      Width           =   6375
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   2055
      Width           =   6375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1740
      Width           =   6375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1425
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1110
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   795
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "example: ""c:\temp\test.exe"""
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "ProductVersion"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "ProductName"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2685
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "OrigionalFileName"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2370
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "LegalCopyright"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2055
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "InternalName"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1740
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "FileVersion"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1425
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "FileDescription"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1110
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "CompanyName"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   795
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 
 If Len(Dir$(Trim$(Me.Text1.Text), 6)) > 0 Then
    ' file found
    Me.Text2.Text = FileInfo(Me.Text1.Text).CompanyName
    ' to avoid running the api's all the time the info has been cashed
    Me.Text3.Text = FileInfo.FileDescription
    Me.Text4.Text = FileInfo.FileVersion
    Me.Text5.Text = FileInfo.InternalName
    Me.Text6.Text = FileInfo.LegalCopyright
    Me.Text7.Text = FileInfo.OrigionalFileName
    Me.Text8.Text = FileInfo.ProductName
    Me.Text9.Text = FileInfo.ProductVersion
  Else
   ' file not found
    MsgBox "File not found", vbOKOnly, "Problem"
 End If

End Sub

