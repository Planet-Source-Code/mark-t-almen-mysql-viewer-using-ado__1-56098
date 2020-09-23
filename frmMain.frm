VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "mySQL Viewer"
   ClientHeight    =   12150
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   ScaleHeight     =   12150
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkQuery 
      BackColor       =   &H00E0E0E0&
      Caption         =   "QUERY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cboTable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Text            =   "dbTFServiceMonitoring"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Text            =   "spsdat"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "+"
      TabIndex        =   4
      Text            =   "agu"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "agu"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtSql 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1200
      Width           =   16935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   10095
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   17806
      _Version        =   393216
      BackColorFixed  =   16777215
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdExecute 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   8400
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "TABLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "DATABASE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2640
      TabIndex        =   10
      Top             =   480
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "SERVER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1050
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private cn As New ADODB.Connection
  Private rs As New ADODB.Recordset
  Private cnString As String
  Private sSQL As String
  
  
  
Private Sub connect()
  
    cnString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & txtServer & ";PORT=3306;" & _
                    "DATABASE=" & txtDatabase & ";USER=" & txtUsername & ";PASSWORD=" & txtPassword & _
                    ";OPTION=3;"
    
    cn.ConnectionString = cnString
    cn.Open
    
    'sSQL = "insert into tblTrans values(1,1,'x','x','x','x','2004-08-08 01:02:56',1,'x',1,1,'x','x','x','x')"
    'cn.Execute sSQL
End Sub

Private Sub cmdExecute_Click()

On Error GoTo ErrHandler

  Dim sTable As String
  
    sTable = cboTable.Text
    
    connect
    loadTables
    
    If sTable = "" Then sTable = cboTable.List(0)
    
    cboTable.Text = sTable
    
    If chkQuery.Value <> 1 Then txtSql = "Select * from " & sTable
    
    sSQL = txtSql
    rs.CursorLocation = adUseClient
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
    Set MSHFlexGrid1.Recordset = rs
    Set cn = Nothing
    Set rs = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub loadTables()
  
  Dim rsSchema As New ADODB.Recordset
  Dim mtblName As String
  
  Set rsSchema = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
  
  cboTable.Clear
  Do Until rsSchema.EOF
         ' Since MS current schema returns tables named "MSys...."
         ' as well as their TABLE_TYPE is also "TABLE", we exclude them.
        If UCase(Left(rsSchema!Table_name, 4)) <> "MSYS" Then
           If UCase(Left(rsSchema!Table_name, 11)) <> "SWITCHBOARD" Then
               mtblName = rsSchema!Table_name
               cboTable.AddItem mtblName
           End If
        End If
        rsSchema.MoveNext
   Loop
   
End Sub

