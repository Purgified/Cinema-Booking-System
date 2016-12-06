VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ShowMovies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show Movies"
   ClientHeight    =   4530
   ClientLeft      =   6825
   ClientTop       =   3480
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6870
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
      Left            =   6480
      TabIndex        =   14
      Top             =   4320
      Width           =   375
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   11
      Top             =   3360
      Width           =   2055
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMark 
         Caption         =   "Bookmark"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Navigate"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   3360
      Width           =   2775
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2040
      Top             =   5160
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "2d/3d"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   2280
      Width           =   2535
      Begin VB.TextBox txtCamera 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rating"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
      Begin VB.TextBox txtRating 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movie name"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   2535
      Begin VB.TextBox txtMovie 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   3015
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "ShowMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Make a variable to store the bookmarked record
'Set as variant to not worry about data types (will just be used for bookmark)
Dim mark As Variant
'Variable to make sure the actual first record in the database is shown
'It will only be given to values, 0 or 1
Dim check As Integer

Private Sub cmdFirst_Click()
    'Checks to see if the record is not already at the first one
    'Calls the checkBOF function to check for the first record
    If checkBOF() = True Then
        MsgBox "You are already at the first record"
    Else
        'Moves to the first record
        Adodc1.Recordset.MoveFirst
        'Sets the appropraite fields to their given values in the table
        txtMovie.Text = Adodc1.Recordset.Fields("MovieName")
        txtRating.Text = Adodc1.Recordset.Fields("Rating")
        txtCamera.Text = Adodc1.Recordset.Fields("Dimension")
        'Loads the image location from the database into the imagebox
        imgPreview.Picture = LoadPicture(App.Path & Adodc1.Recordset.Fields("Image"))
    End If
End Sub

Private Sub cmdGo_Click()
    Adodc1.Recordset.Bookmark = mark
    
    'Sets the appropraite fields to their given values in the table
    txtMovie.Text = Adodc1.Recordset.Fields("MovieName")
    txtRating.Text = Adodc1.Recordset.Fields("Rating")
    txtCamera.Text = Adodc1.Recordset.Fields("Dimension")
    'Loads the image location from the database into the imagebox
    imgPreview.Picture = LoadPicture(App.Path & Adodc1.Recordset.Fields("Image"))
    
    'Enable bookmark button again for new bookmarks and disable go button again
    cmdGo.Enabled = False
    cmdMark.Enabled = True
End Sub

Private Sub cmdMark_Click()
    'Set the mark variable to the current record (bookmarks it)
    mark = Adodc1.Recordset.Bookmark
    
    'Accordingly enables/disables the go/mark buttons
    cmdGo.Enabled = True
    cmdMark.Enabled = False
End Sub

Private Sub cmdNext_Click()
    'Checks if the button has been clicked at least once
    If check <> 0 Then
        Adodc1.Recordset.MoveNext
    End If
    
    'Check to see if we have reached the last row of the table
    If Adodc1.Recordset.EOF = True Then
        MsgBox "No more movies in next records."
        'Moves back the ADODC object to the last record as it would
        'Be in a non-existing recordset if it moved next on an imaginary record
        Adodc1.Recordset.MovePrevious
    Else
        'Sets the appropraite fields to their given values in the table
        txtMovie.Text = Adodc1.Recordset.Fields("MovieName")
        txtRating.Text = Adodc1.Recordset.Fields("Rating")
        txtCamera.Text = Adodc1.Recordset.Fields("Dimension")
        'Loads the image location from the database into the imagebox
        imgPreview.Picture = LoadPicture(App.Path & Adodc1.Recordset.Fields("Image"))
        'Sets the variable to 1 to show that the button has been clicked at least once.
        check = 1
    End If
End Sub

Private Sub cmdLast_Click()
    'Checks to see if the record is not already at the last one
    'Calls the checkEOF function to check for the last record
    If checkEOF() = True Then
        MsgBox "You are already at the last record"
    Else
        'Moves to the last record
        Adodc1.Recordset.MoveLast
        'Sets the appropraite fields to their given values in the table
        txtMovie.Text = Adodc1.Recordset.Fields("MovieName")
        txtRating.Text = Adodc1.Recordset.Fields("Rating")
        txtCamera.Text = Adodc1.Recordset.Fields("Dimension")
        'Loads the image location from the database into the imagebox
        imgPreview.Picture = LoadPicture(App.Path & Adodc1.Recordset.Fields("Image"))
    End If
End Sub

Private Function checkEOF() As Boolean
    'Moves to next record to check if EOF has been reached or not
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF = True Then
        Adodc1.Recordset.MoveLast
        'Sets the checkEOF variable declared as the function to false
        'Which means EOF has been reached
        checkEOF = True
    Else
        'Sets the checkEOF variable declared as the function to false
        'Which means EOF has not been reached
        checkEOF = False
    End If
End Function

Private Function checkBOF() As Boolean
    'Moves to next record to check if EOF has been reached or not
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF = True Then
        Adodc1.Recordset.MoveFirst
        'Sets the checkEOF variable declared as the function to false
        'Which means EOF has been reached
        checkBOF = True
    Else
        'Sets the checkEOF variable declared as the function to false
        'Which means EOF has not been reached
        checkBOF = False
    End If
End Function

Private Sub cmdPrevious_Click()
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF = True Then
        MsgBox "No more Records"
        Adodc1.Recordset.MoveNext
    Else
        'Sets the appropraite fields to their given values in the table
        txtMovie.Text = Adodc1.Recordset.Fields("MovieName")
        txtRating.Text = Adodc1.Recordset.Fields("Rating")
        txtCamera.Text = Adodc1.Recordset.Fields("Dimension")
        'Loads the image location from the database into the imagebox
        imgPreview.Picture = LoadPicture(App.Path & Adodc1.Recordset.Fields("Image"))
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
    MoviesDash.Show
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Login
    Adodc1.RecordSource = "Movies"
    Adodc1.Refresh
    
    'Make sure the form starts with the first record of the table
    Adodc1.Recordset.MoveFirst
    
    'Initializes the value of the check variable
    check = 0
End Sub

Private Sub txtMovie_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "This field cannot be altered."
End Sub

Private Sub txtRating_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "This field cannot be altered."
End Sub

Private Sub txtCamera_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "This field cannot be altered."
End Sub

