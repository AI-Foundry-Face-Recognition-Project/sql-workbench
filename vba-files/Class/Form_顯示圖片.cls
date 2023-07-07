VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_顯示圖片"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim db As DAO.database
Dim rs As DAO.Recordset
Dim imgBytes() As Byte
Dim imgPath As String
Option Explicit

Private Declare PtrSafe Function CryptStringToBinaryA Lib "Crypt32.dll" ( _
    ByVal pszString As LongPtr, _
    ByVal cchString As Long, _
    ByVal dwFlags As Long, _
    ByVal pbBinary As LongPtr, _
    ByRef pcbBinary As Long _
) As Long

Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32.dll" ( _
    ByVal hGlobal As LongPtr, _
    ByVal fDeleteOnRelease As Boolean, _
    ByRef ppstm As Object _
) As Long

Private Declare PtrSafe Function OleLoadPicture Lib "oleaut32.dll" ( _
    ByVal pStream As Object, _
    ByVal lSize As Long, _
    ByVal fRunMode As Boolean, _
    ByRef riid As Any, _
    ByRef ppvObj As Object _
) As Long

Private Sub Command0_Click()
    save_img
    'imgPath = "C:\Users\a7771\Desktop\proj\Face-collection-system-Desktop\temp_image.jpg"
    'SaveBytesAsImageFile imgBytes, imgPath

End Sub

Sub save_img()
    Dim fileDialog As Object
    Dim selectedFile As Variant
    
    Dim defaultExtension As String
    
    Set fileDialog = Application.fileDialog(msoFileDialogSaveAs)
    With fileDialog
        .Title = "Save As"
        defaultExtension = "jpg"
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            If Not Right$(selectedFile, Len(defaultExtension) + 1) Like "*." & defaultExtension Then
                selectedFile = selectedFile & "." & defaultExtension
            End If
            SaveBytesAsImageFile img_box.PictureData, CStr(selectedFile)
            MsgBox "Selected File: " & selectedFile
        Else
            MsgBox "Operation Canceled"
        End If
    End With
    
    Set fileDialog = Nothing
End Sub

Private Sub run_Click()
    Dim table As String
    Dim id As Long
    
    table = table_name
    id = img_id
    
    If table = "人臉圖片" Then
        show_face id
    ElseIf table = "原始圖片" Then
        show_origin_img id
    Else
        msgbox_show "請輸入正確內容"
    End If
End Sub

Sub msgbox_show(text As String)
    MsgBox text
End Sub

Sub show_origin_img(id As Long)
    Dim strSQL As String
    Dim imgBase64 As String

    
    Set db = CurrentDb
    strSQL = "SELECT img FROM origin_img WHERE origin_img_id = " & Chr(34) & CStr(id) & Chr(34)
    Set rs = db.OpenRecordset(strSQL)
    
    If Not rs.EOF Then
        imgBase64 = rs("img")
        rs.MoveNext
    Else
        msgbox_show "Not found. Please check your input"
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    If Len(imgBase64) > 0 Then
        imgBytes = Base64Decode(imgBase64)
        img_box.PictureData = imgBytes
    End If
End Sub

Sub show_face(id As Long)
    Dim strSQL As String
    Dim imgBase64 As String
    Dim imgBytes() As Byte
    Dim imgPath As String
    
    Set db = CurrentDb
    strSQL = "SELECT face_img FROM face WHERE face_id = " & Chr(34) & CStr(id) & Chr(34)
    Set rs = db.OpenRecordset(strSQL)
    
    If Not rs.EOF Then
        imgBase64 = rs("face_img")
        rs.MoveNext
    Else
        msgbox_show "Not found. Please check your input"
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    If Len(imgBase64) > 0 Then
        imgBytes = Base64Decode(imgBase64)
        img_box.PictureData = imgBytes
    End If
End Sub

Function Base64Decode(base64String As String) As Byte()
    Dim xmlDoc As Object
    Dim base64Node As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set base64Node = xmlDoc.createElement("b64")
    base64Node.DataType = "bin.base64"
    base64Node.text = base64String
    
    Base64Decode = base64Node.nodeTypedValue
End Function


Sub SaveBytesAsImageFile(bytes() As Byte, filePath As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    stream.Type = 1 'adTypeBinary
    stream.Open
    stream.Write bytes
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite

    stream.Flush
    stream.Close
    Set stream = Nothing

End Sub

