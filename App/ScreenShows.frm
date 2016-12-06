VERSION 5.00
Begin VB.Form ScreenShows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cinema Screens"
   ClientHeight    =   2220
   ClientLeft      =   6690
   ClientTop       =   4185
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5280
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
      Left            =   4800
      TabIndex        =   2
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Screen 2"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2640
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.CommandButton cmdScreen1 
      Caption         =   "Screen 1"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "ScreenShows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Unload Me
    ShowsDash.Show
End Sub

Private Sub cmdScreen1_Click()
    Unload Me
    BookShow.Show
    BookShow.lblScreen.Caption = 1
End Sub

Private Sub Command1_Click()
    Unload Me
    BookShow.Show
    BookShow.lblScreen.Caption = 2
End Sub

