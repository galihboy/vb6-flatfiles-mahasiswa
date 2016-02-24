Attribute VB_Name = "Module1"
Option Explicit

Const JumlahData = 3
Const pemisah = "#"

Public Function CekFileAda(strAlamatFile As String) As Boolean
    If Dir(strAlamatFile) <> "" Then
        CekFileAda = True
    Else
        CekFileAda = False
    End If
End Function

Public Sub LoadSampleDataMahasiswa(strAlamatFile As String)
    Dim iFileNo As Integer
    iFileNo = FreeFile
    
    Open strAlamatFile For Output As #iFileNo
        Print #iFileNo, "10110001#Adi Sudrajat#l"
        Print #iFileNo, "10110002#Adinda#p"
        Print #iFileNo, "10110003#Ayu Indah#p"
        Print #iFileNo, "10110004#Budi Cahya#l"
        Print #iFileNo, "10110005#Cecep Gorbachev#l"
    Close #iFileNo
End Sub

Public Sub HapusData(strAlamatFile As String, strHapus As String)
    Dim iFileNo As Integer, i As Integer, j As Integer
    Dim strData As String
    Dim boolAda As Boolean
    Dim arrData() As String
    boolAda = False
    iFileNo = FreeFile
    i = 0
    
    Open strAlamatFile For Input As #iFileNo
        Do While Not EOF(iFileNo)
            Input #iFileNo, strData
            If strData = strHapus Then
                'Data yang dicari dilewati (tidak disimpan dalam array)
                boolAda = True
            Else
                'Data yang tidak dicari dan panjang karakter > 0 disimpan di array
                If Len(Trim(strData)) > 0 Then
                    ReDim Preserve arrData(i + 1)
                    arrData(i) = strData
                    i = i + 1
                End If
            End If
        Loop
    Close #iFileNo
    
    'Timpa isi file dengan yang ada di array
    Open strAlamatFile For Output As #iFileNo
        For j = 0 To i - 1
            Print #iFileNo, arrData(j)
        Next
    Close #iFileNo
End Sub

Public Sub UbahData(strAlamatFile As String, strBaruUbah As String, strDataDiubah As String)
    Dim iFileNo As Integer, i As Integer, j As Integer
    Dim strData As String
    Dim boolAda As Boolean
    Dim arrData() As String
    iFileNo = FreeFile
    i = 0
    
    Open strAlamatFile For Input As #iFileNo
        Do While Not EOF(iFileNo)
            ReDim Preserve arrData(i + 1)
            Input #iFileNo, strData
            If strData = strDataDiubah Then
                'Data yang dicari ditemukan, isi dengan yang baru
                arrData(i) = strBaruUbah
                i = i + 1
            Else
                'Data yang tidak dicari dan panjang karakter > 0 disimpan di array
                If Len(Trim(strData)) > 0 Then
                    'ReDim Preserve arrData(i + 1)
                    arrData(i) = strData
                    i = i + 1
                End If
            End If
        Loop
    Close #iFileNo
    
    'Timpa isi file dengan yang ada di array
    Open strAlamatFile For Output As #iFileNo
        For j = 0 To i - 1
            Print #iFileNo, arrData(j)
        Next
    Close #iFileNo
End Sub

Public Sub CariData(strAlamatFile As String, strJenis As String, strKata As String, ByRef boolKetemu As Boolean, ByRef arrData() As String)
    Dim iFileNo As Integer, i As Integer, j As Integer
    Dim strData As String
    Dim boolAda As Boolean
    'Dim arrData() As String
    iFileNo = FreeFile
    i = 0
    boolKetemu = False
    
    Open strAlamatFile For Input As #iFileNo
        Do While Not EOF(iFileNo)
            
            Input #iFileNo, strData
            
            If AdaDalamTeks(strJenis, strKata, strData) Then
                'MsgBox strKata & " - " & strData & " = " & AdaDalamTeks(strJenis, strKata, strData)
                ReDim Preserve arrData(i + 1)
                boolKetemu = True
                arrData(i) = strData
                MsgBox arrData(i)
                i = i + 1
            End If
        Loop
    Close #iFileNo
End Sub

Public Sub TambahData(strAlamatFile As String, strData As String)
    Dim iFileNo As Integer
    iFileNo = FreeFile
    
    Open strAlamatFile For Append As #iFileNo
        Print #iFileNo, strData
    Close #iFileNo
End Sub

Public Sub BacaIsiFile(strAlamatFile As String, lstPenampung As ListBox)
    Dim iFileNo As Integer
    Dim strIsi As String
    iFileNo = FreeFile
    
    Open strAlamatFile For Input As #iFileNo

    Do While Not EOF(iFileNo)
        Input #iFileNo, strIsi
        lstPenampung.AddItem strIsi
        'MsgBox strIsi
    Loop
    
    Close #iFileNo
End Sub

Public Function CekFormatTeks(strTeks As String, ByRef arrData() As String) As Boolean
    Dim arrDataTeks() As String
    arrDataTeks = Split(strTeks, pemisah)
    If UBound(arrDataTeks) + 1 = JumlahData Then
        CekFormatTeks = True
        arrData = arrDataTeks
    Else
        CekFormatTeks = False
    End If
End Function

Public Function AdaDalamTeks(strJenis As String, strTeks As String, strData As String) As Boolean
    Dim arrDataTeks() As String
    
    arrDataTeks = Split(strData, pemisah)
    AdaDalamTeks = False
    
    If strJenis = "nim" Then
        If LCase(strTeks) = LCase(arrDataTeks(0)) Then
            AdaDalamTeks = True
        End If
    ElseIf strJenis = "nama" Then
        'Mencari strTeks dalam array (teks sudah difilter)
        If InStr(LCase(arrDataTeks(1)), LCase(strTeks)) > 0 Then
            AdaDalamTeks = True
        End If
    End If
End Function







