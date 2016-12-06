VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AddShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add show"
   ClientHeight    =   4440
   ClientLeft      =   5130
   ClientTop       =   3480
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8730
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
      Left            =   8280
      TabIndex        =   19
      Top             =   4200
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2160
      Top             =   5520
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3720
      Top             =   6480
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
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Show"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   16
      Top             =   3240
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   6480
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
   Begin VB.Frame Frame8 
      Caption         =   "Ending"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   3240
      Width           =   2535
      Begin VB.ComboBox cmboEnd 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Starting"
      Height          =   735
      Left            =   6720
      TabIndex        =   14
      Top             =   2280
      Width           =   1935
      Begin VB.ComboBox cmboStart 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Date"
      Height          =   735
      Left            =   6720
      TabIndex        =   12
      Top             =   1320
      Width           =   1935
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   129826817
         CurrentDate     =   41686
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Screen"
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
      Left            =   6720
      TabIndex        =   11
      Top             =   120
      Width           =   1935
      Begin VB.ComboBox cmboScreen 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1695
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
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtMovie 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1935
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
      Left            =   4080
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
      Begin VB.TextBox txtRating 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
      Begin VB.TextBox txtCamera 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
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
      Left            =   720
      TabIndex        =   0
      Top             =   3240
      Width           =   2775
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
         TabIndex        =   4
         Top             =   360
         Width           =   495
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
         TabIndex        =   3
         Top             =   360
         Width           =   375
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
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
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
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   3015
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "AddShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variable to make sure the actual first record in the database is shown
'It will only be given to values, 0 or 1
Dim check, a, d, conf As Integer

Private Sub cmboSCreen_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "To alter this field, select a value from the list."
End Sub

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

Private Sub cmdNext_Click()
    'Checks if the button has been clicked at least once
    If check <> 0 Then
        Adodc1.Recordset.MoveNext
    End If
    
    'Check to see if we have reached the last row of the table
    If Adodc1.Recordset.EOF = True Then
        MsgBox "No more Records"
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

Private Sub cmdAdd_Click()
    If Not (txtMovie.Text = "" Or txtRating.Text = "" Or txtCamera.Text = "" _
    Or cmboScreen.Text = "" Or cmboStart.Text = "" Or cmboEnd.Text = "") Then
        'Prompts confirmation
        conf = MsgBox("Are you sure you want to add this record?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
        If conf = vbYes Then
            Dim exists As Boolean
            While Adodc2.Recordset.EOF = False
                If (Adodc2.Recordset.Fields("ShowDate") = dtDate.Value And _
                Adodc2.Recordset.Fields("ScreenID") = Val(cmboScreen.Text) And _
                Adodc2.Recordset.Fields("TimeID") = d) Then
                    exists = True
                End If
                Adodc2.Recordset.MoveNext
            Wend
            Adodc2.Recordset.MoveFirst
            If exists = False Then
                Adodc2.Recordset.MoveLast
                Adodc2.Recordset.AddNew
                Adodc2.Recordset.Fields("MovieID") = Adodc1.Recordset.Fields("ID")
                Adodc2.Recordset.Fields("ScreenID") = Val(cmboScreen.Text)
                Adodc2.Recordset.Fields("ShowDate") = dtDate.Value
                Adodc2.Recordset.Fields("TimeID") = d
                Adodc2.Recordset.Update
                Adodc2.Refresh
                Adodc2.Recordset.MoveFirst
                MsgBox "Added!"
            Else
                MsgBox "This screen is currently booked with the given date and time."
            End If
        End If
    Else
        MsgBox "Please make sure all fields are filled in."
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
    ShowsDash.Show
End Sub

Private Sub dtDate_LostFocus()
    If DateValue(dtDate) < DateValue(Now) Then
        MsgBox "You cannot add a show to a date that has already passed."
        dtDate.Value = DateValue(Now)
    End If
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Movies
    Adodc1.RecordSource = "Movies"
    Adodc1.Refresh
    
    'Establish a connection with the database
    Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Shows
    Adodc2.RecordSource = "Shows"
    Adodc2.Refresh
    
    'Establish a connection with the database
    Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Timings
    Adodc3.RecordSource = "Timings"
    Adodc3.Refresh
    
    'Grabs the timings from the appropriate fields from Shows table
    While Adodc3.Recordset.EOF = False
        cmboStart.AddItem (Adodc3.Recordset.Fields("StartTime"))
        cmboEnd.AddItem (Adodc3.Recordset.Fields("EndTime"))
        Adodc3.Recordset.MoveNext
    Wend
    'Make sure the form starts with the first record of the table
    Adodc1.Recordset.MoveFirst
    
    cmboScreen.List(0) = 1
    cmboScreen.List(1) = 2
    cmboScreen.Text = 1
    'Initializes the value of the check variable
    check = 0
    
    dtDate.Value = DateValue(Now)
End Sub

Private Sub txtCamera_KeyPress(Keyascii As Integer)
    'Disallows modifying a field which requires navigation from the buttons
    Keyascii = 0
    MsgBox "You cannot modify this field"
End Sub

Private Sub txtMovie_KeyPress(Keyascii As Integer)
    'Disallows modifying a field which requires navigation from the buttons
    Keyascii = 0
    MsgBox "You cannot modify this field"
End Sub

Private Sub txtRating_KeyPress(Keyascii As Integer)
    'Disallows modifying a field which requires navigation from the buttons
    Keyascii = 0
    MsgBox "You cannot modify this field"
End Sub

Private Sub cmboStart_KeyPress(Keyascii As Integer)
    'Disallows modifying the list elements in the combobox
    Keyascii = 0
    MsgBox "You can only select from the given fields"
End Sub

Private Sub cmboEnd_KeyPress(Keyascii As Integer)
    'Disallows modifying the list elements in the combobox
    Keyascii = 0
    MsgBox "You can only select from the given fields"
End Sub

Private Sub cmboStart_Click()
    Dim b, c As Integer
    'Retrieve the number of elements in the combobox for looping later
    a = cmboStart.ListCount
    'Looping variable
    b = 0
    c = 0
    'Sets the ending time according to the appropriate starting time
    While b < a
        If cmboStart.Text = cmboStart.List(b) Then
            cmboEnd.Text = cmboEnd.List(b)
        End If
        b = b + 1
    Wend
    
    'Sets the TimeID field in the Shows table to the appropriate ID
    'from the Timings table
    While c < a
        If cmboStart.Text = cmboStart.List(c) Then
            d = c + 1
        End If
        c = c + 1
    Wend
End Sub



Private Sub cmboEnd_GotFocus()
    'Focuses back to the starting time combobox as
    'ending time is determined through its value
    MsgBox "You have to select a starting time to alter this field"
    cmboStart.SetFocus
End Sub
