VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exporting ADO Recordset"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdExportNow 
      Caption         =   "Export &Now!"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdOpenDBase 
      Caption         =   "Open Database"
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CheckBox chkIgnore 
      Caption         =   "Ignore any error that occur"
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "True"
         FalseValue      =   "False"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   7
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export To..."
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdExport 
         Caption         =   "&HTML (html)"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4815
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Excel (xls)"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   4815
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Comma Separated Values (csv)"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4815
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   3915
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   4680
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Export Recordset To..."
      Filter          =   "Comma Separated Values (*.csv)|*.csv|Text Files (*.txt)|*.txt|All files (*.*)|*.*"
      InitDir         =   "C:\Windows\Desktop\"
   End
   Begin VB.Label Label3 
      Caption         =   "CSV"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Currently Selected Option:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "File Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'ADO Recordset Export Class Demo                                '
'Refer Documentation for more info.                             '

'This demo is pretty weak, nonetheless, it works and demonstrates
'the ability of the class.
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'

Private rs As ADODB.Recordset
Private cn As ADODB.Connection
Private WithEvents ex As DatabaseExport
Attribute ex.VB_VarHelpID = -1

Private Sub cmdOpenDBase_Click()

    Shell "explorer " & App.Path & "\Biblio.mdb"

End Sub

Private Sub Form_Load()

    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset                'Create References and
    Set ex = New DatabaseExport                 'open database file

    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BIBLIO.mdb" & ";Persist Security Info=False"
    ex.OpenRecordset rs, "Titles", cn

    Set ex.ADORecordset = rs
    Set ex.ProgressBar = ProgressBar1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set cn = Nothing                            'Remove all references to
    Set rs = Nothing                            'Clear up memory
    Set ex = Nothing

End Sub

Private Sub Text1_DblClick()
Text1.Text = ""
End Sub

Private Sub cmdBrowse_Click()

    Select Case Label3.Caption                'Browse to save a file...
    Case "CSV"
        If ex.ShowSave(cmdlg, CSV) Then
            Text1.Text = ex.FilePath
        End If
    Case "HTML"
        If ex.ShowSave(cmdlg, HTML) Then
            Text1.Text = ex.FilePath
        End If
    Case "Excel"
        If ex.ShowSave(cmdlg, 2) Then
            Text1.Text = ex.FilePath
        End If
    End Select

End Sub

Private Sub cmdExport_Click(Index As Integer)

    Text1.Text = ""
    
    If Index = 0 Then
        Label3.Caption = "CSV"
    ElseIf Index = 1 Then
        Label3.Caption = "HTML"
    Else
        Label3.Caption = "Excel"
    End If

End Sub

Private Function isClean() As Boolean

    If Text1.Text = "" Then
        MsgBox "Please specify a file name to export to!" & vbCrLf & "FOR CSV AND HTML, IT IS A MUST BUT FOR EXCEL, IT IS NOT", vbExclamation, Label3.Caption
        isClean = True
    Else
        isClean = False
    End If

End Function

Private Sub cmdExportNow_Click()

    Select Case Label3.Caption
    Case "CSV"
        ex.FilePath = Text1.Text
        If isClean = False Then ex.ExportToCSV
    Case "HTML"
        ex.FilePath = Text1.Text
        If isClean = False Then ex.ExportToHTML "Biblio Titles (stripped)", "Tahoma,Verdena,Times New Roman", "Tahoma,Verdena,'Times New Roman'", , 2, , , , "3366CC", , , , , "maroon"
    Case "Excel"
        Dim j As Boolean
        j = IIf(Text1.Text = "", True, False)
        If Not j Then MsgBox "If you do not specify file path, I can still Export to Excel."
        ex.ExportToExcel Not j, j
    End Select

End Sub

Private Sub ex_ExportStarted(ByVal ExportingFormat As DatabaseExportEnum)

    Me.MousePointer = 11
    Debug.Print IIf(ExportingFormat = CSV, "CSV", IIf(ExportingFormat = HTML, "HTML", "Excel"))

End Sub

Private Sub ex_ExportError(Error As ErrObject, ByVal ExportingFormat As DatabaseExportEnum)

    If chkIgnore.Value = vbChecked Then
        Error.Clear 'Ignore Error
    Else
        If Error.Number = 424 Then
            MsgBox "Excel Abruptly closed!", vbCritical
        Else
            MsgBox "Error Number: " & Error.Number & vbCrLf & Error.Description, vbCritical, "Unexpected Error"
        End If
        Me.MousePointer = 0
        ProgressBar1.Value = 0
    End If

End Sub

Private Sub ex_ExportComplete(ByVal Success As Boolean, ByVal ExportingFormat As DatabaseExportEnum)
    If Success = True Then
        MsgBox "Export Completed!"
    Else
        MsgBox "Export Completed! However, errors occured but where ignored"
    End If
    
    If ExportingFormat <> 2 Then
        If vbYes = MsgBox("Would you like to open the file just exported?", vbYesNo + vbQuestion) Then
            Shell "explorer " & ex.FilePath
        End If
    End If
    Me.MousePointer = 0
End Sub
