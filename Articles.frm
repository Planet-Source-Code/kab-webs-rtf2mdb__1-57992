VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmArticles 
   Caption         =   "Save RTF to Access Database"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbArticle 
      Height          =   5655
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9975
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Articles.frx":0000
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   8535
   End
   Begin MSDataGridLib.DataGrid DG1 
      Bindings        =   "Articles.frx":0082
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Title"
         Caption         =   "Title"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            WrapText        =   -1  'True
            ColumnWidth     =   2160
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoArticles 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Articles"
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
End
Attribute VB_Name = "frmArticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##################################################################
'Author: KAB WEBs
'Date: 12/29/2004
'admin@kabwebs.com
'http://www.kabwebs.com
'##################################################################
'This example takes an RTF Text Box and saves it as an object in
'a database. The example uses access, but you can switch the
'database by changing the sCmdSource variable in the load event,
'and making other necessary changes to table and field names.
'
'The example program's purpose is only to show how this works. The
'example is not really a complete application, but it will run and
'will show you how this works.
'##################################################################

Private Sub adoArticles_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
Dim varImageData As Variant
Dim oGetImageData As New clsReadData
Tbl = "Articles"
Cmd = "Select * From " + Tbl + " Where Title = '" & adoArticles.Recordset("Title") & "'"
varImageData = oGetImageData.ShowRTF(ReturnCode, Tbl, Cmd)
  If ReturnCode = False Then
     MsgBox "Article not found in DB", vbOKOnly, "Display Article"
     Exit Sub
  End If
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
txtTitle.Text = ""
rtbArticle.Text = ""
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim varImageData As Variant
Dim oGetImageData As New clsReadData

Tbl = "Articles"
Cmd = "Select * from " + Tbl
varImageData = oGetImageData.SaveData(ReturnCode, Tbl, Cmd)
  If ReturnCode = False Then
     MsgBox "There are no articles to save.", vbOKOnly, "Save Article"
     Exit Sub
  End If

adoArticles.Refresh

MsgBox "Article saved in database", vbInformation
Exit Sub
ShowError:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    'Load ado control
    sCmdSource = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\RTF2MDB.mdb" + ";Persist Security Info=False"
    adoArticles.ConnectionString = sCmdSource
    adoArticles.CommandType = adCmdText
    adoArticles.RecordSource = "Select RecNo, Title from Articles"
    adoArticles.Refresh
sPath = App.Path
End Sub
Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
