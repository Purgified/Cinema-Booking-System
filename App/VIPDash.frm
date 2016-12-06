VERSION 5.00
Begin VB.Form VIPDash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIP Dashboard"
   ClientHeight    =   3555
   ClientLeft      =   7245
   ClientTop       =   3060
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4575
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   2
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdSearchMember 
      Caption         =   "Search member"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CommandButton cmdAddMember 
      Caption         =   "Add member"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "VIPDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Unload Me
    CinemaSettings.Show
End Sub

Private Sub cmdAddMember_Click()
    Unload Me
    AddVIP.Show
End Sub

Private Sub cmdSearchMember_Click()
    Unload Me
    SearchVIP.Show
End Sub
