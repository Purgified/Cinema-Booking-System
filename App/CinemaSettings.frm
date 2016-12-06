VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CinemaSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cinema Settings"
   ClientHeight    =   2865
   ClientLeft      =   7245
   ClientTop       =   4890
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4350
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Cinema.mdb"
      Filter          =   "MS Access (*.mdb)|*.mdb"
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
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
      Left            =   3960
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdSeat 
      Caption         =   "Seats"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdVIP 
      Caption         =   "VIP"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdPricing 
      Caption         =   "Pricing"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "CinemaSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conf As Integer

Private Sub cmdBack_Click()
    Unload Me
    Dashboard.Show
End Sub

Private Sub cmdBackup_Click()
    conf = MsgBox("Are you sure you want to make a backup?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
    If conf = vbYes Then
        Dim dbpath As String
        Dim savepath As String
        dbpath = App.Path & "\Cinema.mdb"
        
        CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
        CommonDialog1.InitDir = App.Path
        CommonDialog1.ShowSave
        
        If Len(CommonDialog1.FileName) <> 0 And CommonDialog1.FileName <> "Cinema.mdb" Then
            'Saving directory location assigned to variable
            savepath = CommonDialog1.FileName
            
            'File system object declared to deal with file management processes
            Dim fso As New FileSystemObject
            'Copies file to current program's Movies folder
            fso.CopyFile dbpath, savepath, True
                
            MsgBox "File has been saved to: " & savepath
            CommonDialog1.FileName = "Cinema.mdb"
        End If
    End If
End Sub

Private Sub cmdPricing_Click()
    Unload Me
    Pricing.Show
End Sub

Private Sub cmdSeat_Click()
    Unload Me
    SeatsSettings.Show
End Sub

Private Sub cmdVIP_Click()
    Unload Me
    VIPDash.Show
End Sub
