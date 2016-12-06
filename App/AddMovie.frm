VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AddMovie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Movie"
   ClientHeight    =   4080
   ClientLeft      =   6690
   ClientTop       =   4050
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7215
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
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2880
      Top             =   4920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add movie"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   14
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Caption         =   "Image preview"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdImage 
         Caption         =   "Add Image"
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
         Left            =   480
         TabIndex        =   13
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Image imgPreview 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame2 
         Caption         =   "2D/3D"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1560
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
         Begin VB.OptionButton op3D 
            Caption         =   "Option2"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   255
         End
         Begin VB.OptionButton op2D 
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lbl3D 
            Caption         =   "3D"
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
            Left            =   480
            TabIndex        =   11
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lbl2D 
            Caption         =   "2D"
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
            Left            =   480
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.ComboBox cmboRating 
         Height          =   315
         ItemData        =   "AddMovie.frx":0000
         Left            =   1920
         List            =   "AddMovie.frx":0002
         TabIndex        =   6
         Top             =   1920
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtRelease 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   114360321
         CurrentDate     =   41678
      End
      Begin VB.TextBox txtMovie 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblRating 
         Caption         =   "Rating"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblRelase 
         Caption         =   "Release Date"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblMovie 
         Caption         =   "Movie name"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "JPG (*.jpg)|*.jpg|GIF (*.gif)|*.gifj|Bitmap (*.bmp)|*.bmp"
   End
End
Attribute VB_Name = "AddMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare a variable to store the location of a selected image
Dim img As String
Dim conf As Integer
'Path to store image path based on path of the program
Dim currentpath As String

Private Sub cmboRating_Keypress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "Select a value from the list to change this field."
End Sub

Private Sub cmdBack_Click()
    Unload Me
    MoviesDash.Show
End Sub

Private Sub cmdImage_Click()
    Dim loc As Integer
    Dim title As String

    'Opens a dialog box which asks the user to navigate to a picture to add
    CommonDialog1.ShowOpen
    
    'Checks if any image was selected
    If Len(CommonDialog1.FileName) <> 0 Then
        'Sets variable to the name of file (including extension)to be used in title variable
        title = CommonDialog1.FileTitle
        'Set the location of the image file to the img variable
        img = CommonDialog1.FileName
        'Sets path that file will be copied to based on current programs' path
        currentpath = App.Path & "\Movies\" & title
        
        'File system object declared to deal with file management processes
        Dim fso As New FileSystemObject
        'Copies file to current program's Movies folder
        fso.CopyFile img, currentpath, True
        
        'Sets the loc variable to the starting position of the string \Movies
        'This is done to later on slice the string of all the characters before
        'The \Movies string.
        loc = InStr(1, currentpath, "\Movies")
        
        'Sets the picture of the imagebox to the selected picture by loading it through the img variable
        imgPreview.Picture = LoadPicture(img)
        
        'Slices the string so that the image loads on any computer with the App.Path property
        currentpath = Mid$(currentpath, loc)
    End If
End Sub

Private Sub cmdAdd_Click()
    'Checks for all fields if they are filled
    If txtMovie.Text <> "" And cmboRating.Text <> "" And currentpath <> "" And _
    ((op2D = True And op3D = False) Or (op2D = False And op3D = True)) Then
        'Prompts confirmation
        conf = MsgBox("Are you sure you want to add this record?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
        If conf = vbYes Then
            Dim exists As Boolean
            While Adodc1.Recordset.EOF = False
                If Adodc1.Recordset.Fields("MovieName") = txtMovie.Text Then
                    exists = True
                End If
                Adodc1.Recordset.MoveNext
            Wend
            Adodc1.Recordset.MoveFirst
            If exists = False Then
                'Create new row for new records
                Adodc1.Recordset.AddNew
                
                Adodc1.Recordset.Fields("MovieName") = txtMovie.Text
                Adodc1.Recordset.Fields("ReleaseDate") = dtRelease.Value
                Adodc1.Recordset.Fields("Rating") = cmboRating.Text
                
                'Checks to see which optionbox was selected
                'And then adds the corresponding camera view catagory to the field in the database
                If op2D.Value = True Then
                    Adodc1.Recordset.Fields("Dimension") = "2D"
                ElseIf op3D.Value = True Then
                    Adodc1.Recordset.Fields("Dimension") = "3D"
                End If
                
                'Sets the image field in the table to be the location inside the img String Variable
                Adodc1.Recordset.Fields("Image") = currentpath
                
                'Update record
                Adodc1.Recordset.Update
                'Refresh to make sure everything is working properly
                Adodc1.Refresh
                
                MsgBox "Movie added to the database."
                ShowMovies.Adodc1.Refresh
            Else
                MsgBox "This movie already exists in the database."
            End If
        End If
    ElseIf currentpath = Null Then
        MsgBox "You must select a preview image to the movie in order to add the record."
    Else
        MsgBox "Please make sure all fields are filled"
    End If
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Movies
    Adodc1.RecordSource = "Movies"
    Adodc1.Refresh
    Adodc1.Recordset.MoveFirst
    
    'Initializes combobox list values
    cmboRating.List(0) = "Everyone"
    cmboRating.List(1) = "7+"
    cmboRating.List(2) = "10+"
    cmboRating.List(3) = "PG 13"
    cmboRating.List(4) = "17+"
    cmboRating.List(5) = "Mature"
    cmboRating.Text = "Everyone"
End Sub

Private Sub txtMovie_KeyPress(Keyascii As Integer)
    'Checks for letters (capital and small), backspace, space, shift and shift keypress
    If Not ((Keyascii >= 97 And Keyascii <= 122) Or _
    (Keyascii >= 65 And Keyascii <= 90) Or Keyascii = 8 Or _
    Keyascii = 16 Or Keyascii = 32 Or Keyascii = 20) Then
        Keyascii = 0
        MsgBox "Please enter letters only"
    End If
End Sub

Private Sub txtMovie_LostFocus()
    'Converts first letter of every word to its capital letter
    txtMovie = StrConv(txtMovie.Text, vbProperCase)
End Sub
