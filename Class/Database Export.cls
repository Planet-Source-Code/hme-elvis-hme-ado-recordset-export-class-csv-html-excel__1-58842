VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'  ___________________________________                       _______
' / DatabaseExport                    \_____________________/ v1.50 |
' |                                                                 |
' |       Description:  Make your recordset universal! Use this     |
' |                     class to export whatever recordset into the |
' |                     industrial standered [csv], commercial      |
' |                     standered [html] and the power user         |
' |                     standered [Excel]. This is an updated version
' |                     of my previous work, now complete Events,   |
' |                     Properties, Documentation and even more     |
' |                     options! Can export anything that ADO can   |
' |                     connect to! Fuly optimized and commented.   |
' |                     A must have for any database devoloper and  |
' |                     a must add to any code library!             |
' |                                                                 |
' |    Notes on usage:  If you do not have the Excel 10.0 Library   |
' |                     used to originally write the class then     |
' |                     you can simply remove the reference from    |
' |                     Excel 10.0 and add a tick to the version you|
' |                     have. And if you do not have Excel at all,  |
' |                     just do not use the subroutines related to  |
' |                     Excel.                                      |
' |                                                                 |
' Program/Module Name:  ADO Recordset Export Class Module           |
' |                     © 2005 by Hussain Mohd. Elius               |
' |   Original Author:  Hussain Mohd. Elius                         |
' |           Webpage:  www.elvista.cjb.net                         |
' |                                                                 |
' |      Date Created:  February 12, 2005                           |
' |      OS Tested On:  Windows ME, XP SP 2                         |
' |                  _____________________________                  |
' |_________________/      Licence Agreement      \_________________|
' | °         °     \_____________________________/  °           °  |
' |                                                                 |
' |         This file is intended for and was uploaded to:          |
' |         www.elvista.cjb.net & www.planetsourcecode.com          |
' |                                                                 |
' |Redistribution of this code, whole or in part, as source code or |
' |in binary form, alone or as part of a larger distribution or     |
' |product, is forbidden for any commercial or for-profit use       |
' |without the author's explicit permission. (written or through    |
' |electronic mail)                                                 |
' |                                                                 |
' |Redistribution of this code, as source code or in binary form,   |
' |with or without modification, is permitted provided that the     |
' |following conditions are met:                                    |
' |                                                                 |
' |Redistributions of source code must include this list of         |
' |conditions, and the following acknowledgment:                    |
' |                                                                 |
' |This code was developed by Hussain Mohd. Elius.                  |
' |The source code, written in Visual Basic, is freely available for|
' |non-commercial, non-profit use at www.elvista.cjb.net            |
' |                                                                 |
' |Redistributions in binary form, as part of a larger project, must|
' |include the above acknowledgment in the end-user documentation.  |
' |Alternatively, the above acknowledgment may appear in the        |
' |software itself, if and wherever such third-party acknowledgments|
' |normally appear.                                                 |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.00     H.M.Elius   Feb 12, 2005  Original Version           |
'  | 1.50     H.M.Elius   Feb 27, 2005  Fixed All bugs. Added      |
'  |                                    Events,  Properties and    |
'  |                                    extended the HTML, CSV and |
'  |                                    Excel options. Added       |
'  |                                    documentaion file.         |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Public Event ExportStarted(ByVal ExportingFormat As DatabaseExportEnum)
Public Event ExportError(Error As ErrObject, ByVal ExportingFormat As DatabaseExportEnum)
Public Event ExportComplete(ByVal Success As Boolean, ByVal ExportingFormat As DatabaseExportEnum)

Private Progress As ProgressBar
Private ExportFilePath As String
Private ADODBRecordset As ADODB.Recordset

Public Enum DatabaseExportEnum
    [CSV] = 0
    [HTML] = 1
    [Excel] = 2
End Enum

'API used for custom DoEvents Function
Private Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long

Public Property Set ADORecordset(ByVal vADORec As ADODB.Recordset)
    Set ADODBRecordset = vADORec
End Property

Public Property Set ProgressBar(ByVal vProgress As ProgressBar)
    Set Progress = vProgress
End Property

Public Property Get FilePath() As String
    FilePath = ExportFilePath
End Property

Public Property Let FilePath(ByVal sNewValue As String)
    ExportFilePath = sNewValue
End Property

Private Function DoEventsEx() As Long

'DoEventsEx is tested and is believed to be 100% to 800% faster then
'DoEvents. For more info, search planetsourcecode.com in the VBWorld for:
'DoEvents evolution

    On Local Error Resume Next
        DoEventsEx = GetQueueStatus(&H80 Or &H1 Or &H4 Or &H20 Or &H10)
        If DoEventsEx <> 0 Then
            DoEvents
        End If

End Function

Public Sub ExportToCSV(Optional ByVal PrintHeader As Boolean = True)

Dim TotalRecords As Long, i As Long, NumberOfFields As Integer
Dim ErrorOccured As Boolean
Const Quote As String = """"            'Faster then Chr$(34)

    On Error GoTo hell
    RaiseEvent ExportStarted(CSV)

    With Progress
        .Min = 0
        .Max = ADODBRecordset.RecordCount
        .Value = 0
    End With

    Open ExportFilePath For Output Access Write As #1

    With ADODBRecordset
        .MoveFirst
        NumberOfFields = .Fields.Count - 1

        If PrintHeader Then
            For i = 0 To NumberOfFields - 1      'Now add the field names
                Print #1, .Fields(i).Name & ","; 'similar to the ones below
            Next i
            Print #1, .Fields(NumberOfFields).Name
        End If

        Do While Not .EOF
            TotalRecords = TotalRecords + 1
            For i = 0 To NumberOfFields         'If there is an emty field,
                If (IsNull(.Fields(i))) Then    'add a , to indicate it is
                    Print #1, ",";              'empty
                Else
                    If i = NumberOfFields Then
                        Print #1, Quote & Trim$(CStr(.Fields(i))) & Quote;
                    Else
                        Print #1, Quote & Trim$(CStr(.Fields(i))) & Quote & ",";
                    End If
                End If                  'Putting data under "" will not
            Next i                      'confuse the reader of the file
            DoEventsEx                  'between Dhaka, Bangladesh as two
            Print #1,                   'fields or as one field.
            .MoveNext

            Progress.Value = Progress.Value + 1

        Loop
    End With
    Close #1

    'So ExportComplete will always occur, if an error is ignored
    'Success=True   'No Error Occured
    'Success=False  'Error Occured and were ignored
    RaiseEvent ExportComplete(Not ErrorOccured, CSV)

Exit Sub

    'ExportError is generated whenever there is an error. Error can be
    'ignored by using the code Error.Clear in the ExportError Event

hell:
    RaiseEvent ExportError(Err, CSV)
    If Err.Number = 0 Then
        Resume Next
        ErrorOccured = True
    End If

End Sub

Public Sub ExportToHTML(ByVal TitleOfHTML As String, _
                        Optional TitleFont As String = "Tahoma", _
                        Optional HeaderFont As String = "Tahoma", _
                        Optional TitleFontSize As Byte = 5, _
                        Optional HeaderFontSize As Byte = 3, _
                        Optional TableBorder As Integer = 0, _
                        Optional CellPadding As Integer = 0, _
                        Optional CellSpacing As Integer = 5, _
                        Optional hexBodyBackground As String = "FFFFFF", _
                        Optional hexTitleBackground As String = "800000", _
                        Optional hexTitleForeground As String = "FFFFFF", _
                        Optional hexHeaderBackground As String = "FFFFEF", _
                        Optional hexHeaderForeground As String = "111111", _
                        Optional hexRecordsForeground As String = "111111", _
                        Optional hexTableBackground = "FFFFEF", _
                        Optional hexTableForeground = "111111", _
                        Optional hexBorderColor As String = "111111")

Dim TotalRecords As Long, i As Integer, NumberOfFields As Integer
Dim ErrorOccured As Boolean
Const Quote As String = """"

    On Error GoTo hell
    RaiseEvent ExportStarted(HTML)

    With Progress
        .Min = 0
        .Max = ADODBRecordset.RecordCount
        .Value = 0
    End With

    Open ExportFilePath For Output Access Write As #1

    With ADODBRecordset
        .MoveFirst
        NumberOfFields = .Fields.Count - 1

        Print #1, "<HTML><HEAD><TITLE>" & TitleOfHTML & "</TITLE></HEAD>"
        Print #1, "<meta name=""GENERATEDBY"" content="" [HME] ADO Recordset Export Class "">"
        Print #1, "<meta name=""GENERATEDINFO"" content = "" www.elvista.cjb.net "">"
        Print #1, "<BODY BGCOLOR= " & Quote & hexBodyBackground & Quote & " Text = " & Quote & hexRecordsForeground & Quote & ">"
        Print #1, "<TABLE BORDER= " & Quote & TableBorder & Quote & " CellPadding = " & Quote & CellPadding & Quote & " CellSpacing = " & Quote & CellSpacing & Quote & " BODERCOLOR = " & hexBorderColor & " BGCOLOR = " & Quote & hexTableBackground & Quote & " Width = " & Quote & "100%" & Quote & ">"
        Print #1, "<TR><TD WIDTH=""100%"" COLSPAN=" & Quote & NumberOfFields + 1 & Quote & " BGCOLOR=" & Quote & hexTitleBackground & Quote & ">"
        Print #1, "<FONT COLOR = " & Quote & hexTitleForeground & Quote & "FACE=" & TitleFont & " SIZE=" & Quote & TitleFontSize & Quote & "><B>" & TitleOfHTML & "</B></FONT></TD></TR>"

        Print #1,
        Print #1, "<!-- Database Headers are are listed below -->"
        Print #1,

        Print #1, "     <TR>"        'First, add the Usual HTML Tags ^^^
        For i = 0 To NumberOfFields  'Now, add the titles to the file
            Print #1, "          <TD BGCOLOR=" & hexHeaderBackground & "><B>"
            Print #1, "          <FONT COLOR=" & hexTableForeground & Quote & " FACE=" & Quote & HeaderFont & Quote & " SIZE=" & Quote & HeaderFontSize & Quote & ">" & .Fields(i).Name & "</FONT></B></TD>"
        Next i
        Print #1, "     </TR>"

        Print #1,
        Print #1, "<!-- Database Records are are listed below -->"
        Print #1,

        Do While Not .EOF
            Print #1, "  <TR>"  'Add database records in HTML Format
            For i = 0 To NumberOfFields
                Print #1, "    <TD>" & .Fields(i) & "</TD>"
            Next i
            Print #1, "  </TR>"
            Progress.Value = Progress.Value + 1
            .MoveNext
            DoEventsEx
        Loop

    End With

    Print #1, "</TABLE></BODY></HTML>" 'Complete and close the HTML file
    Close #1

    RaiseEvent ExportComplete(Not ErrorOccured, HTML)

Exit Sub

hell:
    RaiseEvent ExportError(Err, HTML)
    If Err.Number = 0 Then
        Resume Next
        ErrorOccured = True
    End If

End Sub

Public Sub ExportToExcel(Optional SaveFile As Boolean = False, _
                         Optional VisibleInstance As Boolean = True, _
                         Optional Password As String = "", _
                         Optional WriteResPassword As String = "", _
                         Optional ReadOnlyRecommended As Boolean = False, _
                         Optional HeaderFont As String = "Tahoma", _
                         Optional HeaderFontSize As Integer = 9)

Dim iRowIndex As Integer, avRows As Variant, ErrorOccured As Boolean
Dim iFieldCount As Integer, objExcel As Object, objTemp As Object
Dim iColIndex As Integer, iRecordCount As Integer

'I know that some tweaking are left here... I gave priority to compability here

    On Error GoTo hell

    RaiseEvent ExportStarted(Excel)

    With ADODBRecordset
        .MoveFirst
        avRows = .GetRows()                  'Read all the records in an
        iRecordCount = UBound(avRows, 2) + 1 'array and determine how
        iFieldCount = UBound(avRows, 1) + 1  'many fields in an array
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = VisibleInstance
        objExcel.Workbooks.Add

        Set objTemp = objExcel           'Ensure excel remains visible

        If Val(objExcel.Application.Version) >= 8 Then
            Set objExcel = objExcel.ActiveSheet
        End If

        iRowIndex = 1

        'Place Name of the fields
        For iColIndex = 1 To iFieldCount
            With objExcel.Cells(iRowIndex, iColIndex)
                .Value = ADODBRecordset.Fields(iColIndex - 1).Name
                With .Font
                    .Name = HeaderFont      'Make the headers stand out
                    .Size = HeaderFontSize
                    .Bold = True
                End With
            End With
        Next iColIndex

    End With

    With Progress
        .Min = 0
        .Max = ADODBRecordset.RecordCount
        .Value = 0
    End With

    With objExcel
        
        For iRowIndex = 2 To iRecordCount + 1
            For iColIndex = 1 To iFieldCount
                .Cells(iRowIndex, iColIndex) = avRows(iColIndex - 1, iRowIndex - 2)
            Next iColIndex
            Progress.Value = Progress.Value + 1
            DoEventsEx
        Next iRowIndex
        .Cells(1, 1).CurrentRegion.EntireColumn.AutoFit
        If SaveFile Then
            .SaveAs ExportFilePath, , Password, WriteResPassword, ReadOnlyRecommended
        End If
    End With
    
    If Not VisibleInstance Then objExcel.Application.Quit
    Set objTemp = Nothing
    Set objExcel = Nothing
    
    RaiseEvent ExportComplete(Not ErrorOccured, Excel)

Exit Sub

hell:
    RaiseEvent ExportError(Err, Excel)
    If Err.Number = 0 Then
        Resume Next
        ErrorOccured = True
    End If

End Sub

Public Sub OpenRecordset(ByRef rs As ADODB.Recordset, ByVal TableName As String, ByVal Connection As ADODB.Connection, Optional ByVal QueryParameter As String = "")
'Supporting routine. Can be deleted if you will.
    With rs
        .CursorLocation = adUseClient
        .Open Trim$("SELECT * FROM [" & TableName & "] " & QueryParameter), Connection, adOpenKeyset, adLockOptimistic
    End With

End Sub

Public Function ShowSave(ByVal cmdlg As CommonDialog, ByVal ExportMode As DatabaseExportEnum, Optional DialogTitle As String = "Export Recordset to...") As Boolean

Dim Extention As String

    If ExportMode = CSV Then
        Extention = "Comma Separated Values (*.csv)|*.csv|Text Files (*.txt)|*.txt|"
    ElseIf ExportMode = HTML Then
        Extention = "Hyper Text Markup Language Files (*.htm,*.html)|*.htm;*.html|"
    ElseIf ExportMode = Excel Then
        Extention = "Microsoft Excel Workbook (*.xls)|*.xls|"
    End If
        Extention = Extention & "All Files (*.*)|*.*"

    On Error GoTo hell
    With cmdlg
        .CancelError = True
        .DialogTitle = DialogTitle
        .FileName = ""
        .Filter = Extention
        .FilterIndex = 0
        .ShowSave
        ExportFilePath = .FileName
    End With

    If ExportFilePath <> "" Then
        ShowSave = True
    End If

hell:

End Function
