VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Function SaveData(ReturnCode, Tbl, Cmd)
On Error GoTo ErrorSection
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
' open ado connection to sql server
cn.Open sCmdSource
' open recordset with custom command or if no command passed open all data in table
If Cmd = "" Then
    rs.Open "Select * from " + Tbl, cn, adOpenForwardOnly, adLockOptimistic
Else
    rs.Open Cmd, cn, adOpenForwardOnly, adLockOptimistic
End If
'Function succeeded, Data retrieved set Return Value
ReturnCode = True
If rs.RecordCount > 0 Then
rs.MoveLast
rs.MoveFirst
End If
  With rs
    .AddNew
    Dim file_num As String
    Dim FileName As String
    Dim file_length As Long
    Dim bytes() As Byte
    Dim num_blocks As Long
    Dim left_over As Long
    Dim block_num As Long
   
    file_num = FreeFile
    FileName = sPath + "\tmp.txt"
    'save rtf file here
    frmArticles.rtbArticle.SaveFile FileName, rtfRTF
    'Open the file and read each block into the recordset
    Open FileName For Binary Access Read As file_num
    file_length = LOF(file_num)
    If file_length > 0 Then
        num_blocks = file_length / BLOCK_SIZE
        left_over = file_length Mod BLOCK_SIZE

        rs("Title") = frmArticles.txtTitle.Text

        ReDim bytes(BLOCK_SIZE)
        For block_num = 1 To num_blocks
            Get file_num, , bytes()
            rs("Article").AppendChunk bytes()
        Next block_num

        If left_over > 0 Then
            ReDim bytes(left_over)
            Get file_num, , bytes()
            rs("Article").AppendChunk bytes()
        End If

        Close file_num
    End If
    .Update
  End With
'Close connection
rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing
Cmd = ""
Exit Function

ErrorSection:
'Close connection
Set rs = Nothing
Set cn = Nothing
ReturnCode = False
End Function

Public Function ShowRTF(ReturnCode, Tbl, Cmd)
On Error GoTo ErrorSection
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
' open ado connection to sql server
cn.Open sCmdSource
' open recordset with custom command or if no command passed open all data in table
If Cmd = "" Then
    rs.Open "Select * from " + Tbl, cn, adOpenForwardOnly, adLockOptimistic
Else
    rs.Open Cmd, cn, adOpenForwardOnly, adLockOptimistic
End If
'Function succeeded, Data retrieved set Return Value
ReturnCode = True
Dim bytes() As Byte
Dim file_name As String
Dim file_num As Integer
Dim file_length As Long
Dim num_blocks As Long
Dim left_over As Long
Dim block_num As Long
Dim hgt As Single
DoEvents
file_name = TemporaryFileName()
file_num = FreeFile
Open file_name For Binary As file_num
file_length = rs("Article").ActualSize
    num_blocks = file_length / BLOCK_SIZE
    left_over = file_length Mod BLOCK_SIZE

    For block_num = 1 To num_blocks
        bytes() = rs("Article").GetChunk(BLOCK_SIZE)
        Put file_num, , bytes()
    Next block_num

    If left_over > 0 Then
        On Error Resume Next
        bytes() = rs("Article").GetChunk(left_over)
        Put file_num, , bytes()
    End If

    Close file_num

    frmArticles.rtbArticle.LoadFile file_name, rtfRTF
    frmArticles.txtTitle = rs("Title")
 
'Close connection
rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing
Cmd = ""
Exit Function

ErrorSection:
'Close connection
Set rs = Nothing
Set cn = Nothing
'Function failed, If Error or no data set Return Value
ReturnCode = False
End Function


