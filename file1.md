# Catatan Tentang Pemrograman VBA di MS Excel

Dalam catatan ini kita akan membahas tentang bagaimana membangun aplikasi menggunakan VBA  di  MS Excel. Untuk penyimpanan data
juga menggunakan MariaDB. jadi di MS Excel cuma ditaruh data-data yang ingin ditampilkan saja. Di sini MS Excel juga terkoneksi
dengan VPS yang memiliki IP Publik. Sehingga data-data bisa diakses dari file MS Excel dari mana saja di seluruh dunia. Untuk
keamanan koneksi maka diterapkan penggunaan SSH Tunnel.


Source code pertama yang ingin di share adalah source code untuk meload data dari remote MariaDB server di internet ke sheet di
MS Excel:

```vba

Public koneksi As New ADODB.Connection

Public catatan As New ADODB.Recordset


Private Sub CommandButton1_Click()

' Tombol untuk load data dari server di internet


    koneksi = New ADODB.Connection
    
    koneksi.ConnectionString = "DSN=sahamInternet"
    
    koneksi.Open
    
    
    
    catatan.Open "SELECT * FROM datakomputer1", koneksi, adOpenKeyset
    
    Sheet4.Range("A1").CopyFromRecordset catatan
    
    koneksi.Close
    

End Sub


```


Source code berikut ini untuk menghapus seluruh data yang sudah di load ke sheet di MS Excel :

```vba

Private Sub CommandButton2_Click()

' Tombol hapus data di sheet

    barisTerakhir = Application.ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    If barisTerakhir > 2 Then
    
        Application.ActiveSheet.Range("A:AA" & CStr(lastRow)).ClearContents
    
    End If
    

End Sub


```





