VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BookShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Show"
   ClientHeight    =   4200
   ClientLeft      =   7530
   ClientTop       =   3900
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   3855
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
      Left            =   3480
      TabIndex        =   16
      Top             =   3960
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2880
      Top             =   4560
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
      Left            =   1560
      Top             =   4560
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   4560
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
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   3000
      Width           =   375
   End
   Begin VB.Frame frameNav 
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   3615
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdBook 
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame frameEnds 
      Caption         =   "Ends"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
      Begin VB.TextBox txtEnd 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frameStart 
      Caption         =   "Starts"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
      Begin VB.TextBox txtStart 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame framStyle 
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
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1695
      Begin VB.TextBox txtStyle 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frameMovie 
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame frameDate 
      Caption         =   "Show date"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Format          =   131137537
         CurrentDate     =   41688
      End
   End
   Begin VB.Label lblScreen 
      Height          =   135
      Left            =   960
      TabIndex        =   15
      Top             =   5160
      Width           =   255
   End
End
Attribute VB_Name = "BookShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim check As Integer
Dim conf As Integer
Dim find As Boolean

Private Sub cmdBack_Click()
    Unload Me
    ScreenShows.Show
End Sub

Private Sub cmdBook_Click()
    If Not (txtName.Text = "" Or txtStyle.Text = "") Then
        'Prompts confirmation
        conf = MsgBox("Are you sure you want to book this show?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
        If conf = vbYes Then
            Me.Hide
            BookSeat.Show
            'Passes the show selected information to the booking form
            BookSeat.Adodc2.Recordset.Filter = "ShowID=" & Adodc1.Recordset.Fields("ShowID")
            If lblScreen.Caption = "1" Then
                BookSeat.cmboScreen.Text = "One"
            ElseIf lblScreen.Caption = "2" Then
                BookSeat.cmboScreen.Text = "Two"
            End If
        End If
    Else
        MsgBox "Please make sure all fields are filled in."
    End If
End Sub

Private Sub cmdFind_Click()
    'Shows all records in selected date
    Adodc1.Recordset.Filter = "ShowDate='" & dtDate.Value & "' And ScreenID='" & lblScreen.Caption & "'"
    If Adodc1.Recordset.RecordCount <> 0 Then
        Adodc2.Recordset.Filter = "ID=" & Adodc1.Recordset.Fields("MovieID")
        Adodc3.Recordset.Filter = "TimeID=" & Adodc1.Recordset.Fields("TimeID")
        
        txtName.Text = Adodc2.Recordset.Fields("MovieName")
        txtStyle.Text = Adodc2.Recordset.Fields("Dimension")
        txtStart.Text = Adodc3.Recordset.Fields("StartTime")
        txtEnd.Text = Adodc3.Recordset.Fields("EndTime")
        find = True
    Else
        MsgBox "No show found on that date."
    End If
End Sub

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

Private Sub cmdFirst_Click()
    If find = True Then
        If Adodc1.Recordset.RecordCount <> 0 Then
            'Checks to see if the record is not already at the first one
            'Calls the checkBOF function to check for the first record
            If checkBOF() = True Then
                MsgBox "You are already at the first record"
            Else
                'Moves to the first record
                Adodc1.Recordset.MoveFirst
                
                Adodc2.Recordset.Filter = "ID=" & Adodc1.Recordset.Fields("MovieID")
                Adodc3.Recordset.Filter = "TimeID=" & Adodc1.Recordset.Fields("TimeID")
                
                'Sets the appropraite fields to their given values in the table
                txtName.Text = Adodc2.Recordset.Fields("MovieName")
                txtStyle.Text = Adodc2.Recordset.Fields("Dimension")
                txtStart.Text = Adodc3.Recordset.Fields("StartTime")
                txtEnd.Text = Adodc3.Recordset.Fields("EndTime")
            End If
        Else
            MsgBox "Please select a date first."
        End If
    Else
        MsgBox "Please search something with a selected date first!"
    End If
End Sub

Private Sub cmdLast_Click()
    If find = True Then
        If Adodc1.Recordset.RecordCount <> 0 Then
            'Checks to see if the record is not already at the last one
            'Calls the checkEOF function to check for the last record
            If checkEOF() = True Then
                MsgBox "You are already at the last record"
            Else
                'Moves to the last record
                Adodc1.Recordset.MoveLast
                Adodc2.Recordset.Filter = "ID=" & Adodc1.Recordset.Fields("MovieID")
                Adodc3.Recordset.Filter = "TimeID=" & Adodc1.Recordset.Fields("TimeID")
                
                'Sets the appropraite fields to their given values in the table
                txtName.Text = Adodc2.Recordset.Fields("MovieName")
                txtStyle.Text = Adodc2.Recordset.Fields("Dimension")
                txtStart.Text = Adodc3.Recordset.Fields("StartTime")
                txtEnd.Text = Adodc3.Recordset.Fields("EndTime")
            End If
        Else
            MsgBox "Please select a date first."
        End If
    Else
        MsgBox "Please search something with a selected date first!"
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


Private Sub cmdNext_Click()
    If find = True Then
        If Adodc1.Recordset.RecordCount <> 0 Then
            'Checks if the button has been clicked at least once
            If check <> 0 Then
                Adodc1.Recordset.MoveNext
            End If
            
            'Check to see if we have reached the last row of the table
            If Adodc1.Recordset.EOF = True Then
                MsgBox "No more further shows."
                'Moves back the ADODC object to the last record as it would
                'Be in a non-existing recordset if it moved next on an imaginary record
                Adodc1.Recordset.MovePrevious
            Else
                'Locks onto a ID field in the Movies table
                Adodc2.Recordset.Filter = "ID=" & Adodc1.Recordset.Fields("MovieID")
                'Locks onto a TimeID field in the Timings table
                Adodc3.Recordset.Filter = "TimeID=" & Adodc1.Recordset.Fields("TimeID")
                
                txtName.Text = Adodc2.Recordset.Fields("MovieName")
                txtStyle.Text = Adodc2.Recordset.Fields("Dimension")
                txtStart.Text = Adodc3.Recordset.Fields("StartTime")
                txtEnd.Text = Adodc3.Recordset.Fields("EndTime")
                'Sets the variable to 1 to show that the button has been clicked at least once.
                check = 1
            End If
        Else
            MsgBox "Please select a date first."
        End If
    Else
        MsgBox "Please search something with a selected date first!"
    End If
End Sub

Private Sub cmdPrevious_Click()
    If find = True Then
        If Adodc1.Recordset.RecordCount <> 0 Then
            Adodc1.Recordset.MovePrevious
            If Adodc1.Recordset.BOF = True Then
                MsgBox "No more further shows."
                Adodc1.Recordset.MoveNext
            Else
                'Locks onto a ID field in the Movies table
                Adodc2.Recordset.Filter = "ID=" & Adodc1.Recordset.Fields("MovieID")
                'Locks onto a TimeID field in the Timings table
                Adodc3.Recordset.Filter = "TimeID=" & Adodc1.Recordset.Fields("TimeID")
                
                txtName.Text = Adodc2.Recordset.Fields("MovieName")
                txtStyle.Text = Adodc2.Recordset.Fields("Dimension")
                txtStart.Text = Adodc3.Recordset.Fields("StartTime")
                txtEnd.Text = Adodc3.Recordset.Fields("EndTime")
            End If
        Else
            MsgBox "Please select a date first."
        End If
    Else
        MsgBox "Please search something with a selected date first!"
    End If
End Sub

Private Sub dtDate_LostFocus()
    'Past date check
    If DateValue(dtDate) < DateValue(Now) Then
    
        MsgBox "You cannot book a show to a date that has already passed."
        dtDate.Value = DateValue(Now)
        dtDate.SetFocus
    
    End If
End Sub

Private Sub txtStart_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "You cannot alter this field."
    cmdNext.SetFocus
End Sub

Private Sub txtEnd_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "You cannot alter this field."
    cmdNext.SetFocus
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Shows
    Adodc1.RecordSource = "Shows"
    Adodc1.Refresh
    
    'Establish a connection with the database
    Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Movies
    Adodc2.RecordSource = "Movies"
    Adodc2.Refresh

    'Establish a connection with the database
    Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Timings
    Adodc3.RecordSource = "Timings"
    Adodc3.Refresh
    
    'Initializes some starting variables
    check = 0
    find = False
    
    dtDate.Value = DateValue(Now)
End Sub

Private Sub txtName_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "You cannot modify this field. Use the navigation buttons to choose movie."
    cmdNext.SetFocus
End Sub

Private Sub txtStyle_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "You cannot modify this field. Use the navigation buttons to choose movie."
    cmdNext.SetFocus
End Sub
