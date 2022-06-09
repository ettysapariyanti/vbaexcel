# vbaexcel
membahas tentang pemrograman VBA di software MS Excel


Source Code Untuk Menghapus angka setelah tanda titik (.), termasuk juga tanda titik nya :

```vba

Private Sub CommandButton1_Click()

'Memilih workbook dan sheet yg akan dipakai melakukan pekerjaan

ThisWorkbook.Activate

Worksheets("PWON").Select


'Kode untuk pemurnian data Open

Range("C1").Select

ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],SEARCH(""."",RC[-1])-1)"

ActiveCell.Select

Selection.AutoFill Destination:=ActiveCell.Range("A1:A757"), Type:=xlFillDefault

ActiveCell.Range("A1:A757").Select





'Kode untuk pemurnian data High

Range("E1").Select

ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],SEARCH(""."",RC[-1])-1)"

ActiveCell.Select

Selection.AutoFill Destination:=ActiveCell.Range("A1:A757"), Type:=xlFillDefault

ActiveCell.Range("A1:A757").Select



End Sub




```


Versi lebih lengkap dari source code di atas:

```vb

Private Sub CommandButton1_Click()

'Memilih workbook dan sheet yg akan dipakai melakukan pekerjaan

ThisWorkbook.Activate

Worksheets("PWON").Select


'Kode untuk pemurnian data Open

Range("C2").Select

ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],SEARCH(""."",RC[-1])-1)"

ActiveCell.Select

Selection.AutoFill Destination:=ActiveCell.Range("A1:A757"), Type:=xlFillDefault

ActiveCell.Range("A1:A757").Select





'Kode untuk pemurnian data High

Range("E2").Select

ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],SEARCH(""."",RC[-1])-1)"

ActiveCell.Select

Selection.AutoFill Destination:=ActiveCell.Range("A1:A757"), Type:=xlFillDefault

ActiveCell.Range("A1:A757").Select



'Kode untuk pemurnian data Low

Range("G2").Select

ActiveCell.FormulaR1C1 = "=left(rc[-1],search(""."",rc[-1])-1)"

ActiveCell.Select

Selection.AutoFill Destination:=ActiveCell.Range("A1:A757"), Type:=xlFillDefault

ActiveCell.Range("A1:A757").Select





'Kode untuk pemurnian data Close

Range("I2").Select

ActiveCell.FormulaR1C1 = "=left(rc[-1],search(""."",rc[-1])-1)"

ActiveCell.Select

Selection.AutoFill Destination:=ActiveCell.Range("A1:A757"), Type:=xlFillDefault

ActiveCell.Range("A1:A757").Select


End Sub


```



Source code awal untuk copy paste angka yg ada di dalam cell tanpa mengikutsertakan formulanya :

```vba

Private Sub CommandButton1_Click()

' Copy paste angka yg ada formulanya dari C3 di sheet PWONCSV ke A2 di Sheet4

Sheets("PWONCSV").Select

Range("C3").Select

Selection.Copy

Sheets("Sheet4").Select

Range("A2").Select

Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks:=False, Transpose:=False


End Sub


```


source code yg sudah bisa copy paste data tanpa mengikutsertakan volume nya untuk open,high,low :

```vba

Private Sub CommandButton1_Click()


'Copy Paste Untuk Data Open, tanpa perlu mengikutsertakan rumus


'Membuat Caption Cell Open

Sheets("Sheet4").Select

Range("A1").Select

ActiveCell.FormulaR1C1 = "Open"


' Copy Paste data harga open

Sheets("PWONCSV").Select

Range("C2:C758").Select

Selection.Copy

Sheets("Sheet4").Select

Range("A2").Select

Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks:=False, Transpose:=False



'Copy Paste Untuk Data High, tanpa perlu mengikutsertakan rumus


'Membuat Caption Cell High

Sheets("Sheet4").Select

Range("B1").Select

ActiveCell.FormulaR1C1 = "High"




' Copy Paste data harga High

Sheets("PWONCSV").Select

Range("E2:E758").Select

Selection.Copy


Sheets("Sheet4").Select

Range("B2").Select

Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks:=False, Transpose:=False




'Copy Paste Untuk Data Low, tanpa perlu mengikutsertakan rumus

' Membuat caption cell Low

Sheets("Sheet4").Select

Range("C1").Select

ActiveCell.FormulaR1C1 = "Low"




' Copy Paste data harga low

Sheets("PWONCSV").Select

Range("G2:G758").Select

Selection.Copy



Sheets("Sheet4").Select

Range("C2").Select

Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks:=False, Transpose:=False


End Sub



```



Membuat format tanggal di VBA :

```vba

Sub FormatTanggal1()
'
' FormatTanggal1 Macro
'

'
    Application.CutCopyMode = False
    Selection.NumberFormat = "m/d/yyyy"
End Sub


```


source code untuk tarik data ke bawah dengan tujuan copy paste di cell yg sama:

```vba

Sub copypastecell()
'
' copypastecell Macro
'

'
    Selection.AutoFill Destination:=Range("B2:B758"), Type:=xlFillDefault
    Range("B2:B758").Select
End Sub


```


source code sementara untuk isi data harga saham, lengkap dengan tanggal, ticker, nama lengkap:

```vba

Private Sub CommandButton1_Click()


Workbooks("MacroSaham.xlsm").Activate




' Membuat Caption Cell Date

Sheets("PWONRapi").Select

Range("A1").Select

ActiveCell.FormulaR1C1 = "Date"




' Mengcopy data tanggal


Sheets("PWONCSV").Select

Range("A2:A758").Select

Selection.Copy

Sheets("PWONRapi").Select

Range("A2").Select

Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks:=False, Transpose:=False

Selection.NumberFormat = "m/d/yyyy"




' Membuat Caption Cell Ticker

Sheets("PWONRapi").Select

Range("B1").Select

ActiveCell.FormulaR1C1 = "Ticker"




' Membuat data Ticker : PWON

Sheets("PWONRapi").Select

Range("B2").Select

ActiveCell.FormulaR1C1 = "PWON"



' Mengcopy paste tulisan PWON di cell yg sama, ke arah bawah

Range("B2").Select

Selection.AutoFill Destination:=Range("B2:B758"), Type:=xlFillDefault

Range("B2:B758").Select



End Sub

```


Rumus penting:

=LEFT(A6,SEARCH(".",A6)-1)

rumus menghapus titik dan angka setelahnya.


=IF(B2=TRUE,A2,LEFT(A2,SEARCH(""."",A2)-1))

rumus untuk menentukan apakah perlu dilakukan penghapusan . (titik) dan angka setelah titik, atau cuma perlu menyamakan isi cell dengan cell sumber. agar semua harga sesuai dengan format Integer.





Source Code untuk pengujian integer dan perubahan harga saham agar sesuai dengan kriteria integer:

```vba

Sub Halaman4()



' Menyelesaikan masalah Pengecekan apakah angka bisa termasuk integer atau tidak

Sheets("Sheet4").Select

Range("B2").Select

ActiveCell.Formula = "=Int(A2)=A2"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4"), Type:=xlFillDefault



' Melakukan pemurnian agar semua angka bisa sesuai dengan integer

Sheets("Sheet4").Select

Range("C2").Select

ActiveCell.Formula = "=IF(B2=TRUE,A2,LEFT(A2,SEARCH(""."",A2)-1))"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4"), Type:=xlFillDefault


End Sub



```


source code untuk mencoba melakukan 2 perintah yg berbeda menggunakan IF ... Else :

```vba

Sub CekInteger()

'Membuat kolom untuk cek integer nilai Open

Columns("C:C").Select

Selection.Insert shift:=xlToLeft, copyorigin:=xlFormatFromLeftOrAbove

Range("C1").Select

ActiveCell.FormulaR1C1 = "Integer Open"



' Melakukan pengecekan nilai integer dan menuliskan hasilnya di kolom Integer Open

Range("C2").Select

ActiveCell.Formula = "=Int(B2) = B2"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4103"), Type:=xlFillDefault



' Membuat kolom hasil akhir copy paste data Open1

Columns("D:D").Select

Selection.Insert shift:=xlToLeft, copyorigin:=xlFormatFromLeftOrAbove

Range("D1").Select

ActiveCell.Formula = "Open1"



' melakukan copy paste & penghapusan angka di belakang titik (.) berdasarkan pada True atau False di kolom Integer Open

Range("D2").Select

If Range("C2").Value = "FALSE" Then

ActiveCell.Formula = "Berhasil"


Else


ActiveCell.Formula = "GAGAL"

End If





End Sub


```


source code yang akan menjadi cikal bakal untuk pemurnian data harga saham PWON, ada fitur otomatis deteksi huruf di belakang tanda titik (.), apabila tidak ada tanda titik (.), maka akan langsung copy paste nilainya:

```vba

Sub Halaman5()

Sheets("Sheet5").Select


' Pembuatan kolom Harga Saham

Range("A1").Select

ActiveCell.Formula = "Harga Saham"

Range("A2").Select

ActiveCell.Formula = "15.569972"

Range("A3").Select

ActiveCell.Formula = "18.535681"

Range("A4").Select

ActiveCell.Formula = "100"

Range("A5").Select

ActiveCell.Formula = "200"




' Pembuatan kolom Pengecekan Integer


Range("B1").Select

ActiveCell.Formula = "Periksa Kondisi Integer"

Range("B2").Select

ActiveCell.Formula = "=INT(A2)=A2"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4"), Type:=xlFillDefault


'Lakukan pemurnian harga saham

Range("C1").Select

ActiveCell.Formula = "Harga Pemurnian"

Range("C2").Select

ActiveCell.Formula = "=IF(B2=TRUE,A2,LEFT(A2,SEARCH(""."",A2)-1))"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4"), Type:=xlFillDefault




End Sub




```


Source code untuk copy data dan normalisasi angka untuk harga saham PWON :


```vba

Sub halaman6()

' Harga yg akan siap dikirim ke Amibroker ada di Sheet6

' Pilih sheet6 untuk membuat kolom tanggal

Sheets("Sheet6").Select

Range("A1").Select

ActiveCell.Formula = "Tanggal"

Range("A2:A4104").Select

Selection.NumberFormat = "m/d/yyyy"



' Pilih sheet PWONHarga untuk mengcopy paste tanggal

Sheets("PWONJK").Select

Range("A2:A4104").Select

Selection.Copy




' Piih lagi Sheet6 untuk mempaste nilai yg sudah di copy dari sheet PWONJK

Sheets("Sheet6").Select

Range("A2").Select

Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks:=False, Transpose:=False

Columns("A:A").EntireColumn.AutoFit





' Pilih sheet6 untuk membuat kolom HargaOpen

Sheets("Sheet6").Select

Range("B1").Select

ActiveCell.Formula = "HargaOpen"




' Pilih sheet PWONHarga untuk mengcopy paste Harga Open

Sheets("PWONJK").Select

Range("B2:B4104").Select

Selection.Copy



' Piih lagi Sheet6 untuk mempaste nilai yg sudah di copy dari sheet PWONJK

Sheets("Sheet6").Select

Range("B2").Select

Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks:=False, Transpose:=False

Columns("B:B").EntireColumn.AutoFit


' Pilih Sheet6

' Buat kolom HargaOpenTrueFalse

' Kemudian lakukan penyortiran data, untuk mengecek mana yg integer.

' Kalau memang integer maka nilainya akan menjadi TRUE kalau bukan integer maka nilainya akan FALSE


Sheets("Sheet6").Select

Range("C1").Select

ActiveCell.Formula = "HargaOpenTrueFalse"



Range("C2").Select

ActiveCell.Formula = "=INT(B2)=B2"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4103"), Type:=xlFillDefault

Columns("C:C").EntireColumn.AutoFit




Sheets("Sheet6").Select

Range("D1").Select

ActiveCell.Formula = "HargaOpen2"

Columns("D:D").EntireColumn.AutoFit


Sheets("Sheet6").Select

Range("D2").Select

ActiveCell.Formula = "=IF(C2=TRUE,B2,LEFT(B2,SEARCH(""."",B2)-1))"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4103"), Type:=xlFillDefault


' Pilih sheet sheet6 kemudian buat kolom untuk harga High

Sheets("Sheet6").Select

Range("E1").Select

ActiveCell.Formula = "HargaHigh"




' Pilih sheet PWONJK untuk mengcopy paste harga High

Sheets("PWONJK").Select

Range("C2:C4104").Select

Selection.Copy


' Pilih lagi Sheet6 untuk mempaste nilai yg sudah di copy dari sheet PWONJK (Harga High)

Sheets("Sheet6").Select

Range("E2").Select

Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks:=False, Transpose:=False

Columns("E:E").EntireColumn.AutoFit




' Pilih sheet sheet6 kemudian buat kolom untuk Harga High True False

Sheets("Sheet6").Select

Range("F1").Select

ActiveCell.Formula = "HargaHighTrueFalse"


Range("F2").Select

ActiveCell.Formula = "=INT(E2)=E2"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4103"), Type:=xlFillDefault

Columns("F:F").EntireColumn.AutoFit



' Pilih sheet sheet6 kemudian buat kolom untuk Harga High 2

Sheets("Sheet6").Select

Range("G1").Select

ActiveCell.Formula = "HargaHigh2"

Columns("G:G").EntireColumn.AutoFit



Sheets("Sheet6").Select

Range("G2").Select

ActiveCell.Formula = "=IF(F2=TRUE,E2,LEFT(E2,SEARCH(""."",E2)-1))"

Selection.AutoFill Destination:=ActiveCell.Range("A1:A4103"), Type:=xlFillDefault



End Sub


```





Source Code untuk test koneksi VBA ke server MariaDB di Cloud menggunakan SSH Tunnel:

```vba

Sub testKoneksi()

    Dim koneksi As ADODB.Connection

    Dim rekordset As ADODB.Recordset

    Set koneksi = New ADODB.Connection
    
    koneksi.ConnectionString = "Driver={MariaDB ODBC 3.1 Driver};Server=127.0.0.1;Port=3306;Database=namabasisdata;User=pengguna1;Password=password1;Option=3"
    
    koneksi.Open
    
    koneksi.Close
    
    MsgBox "Berhasil Terkoneksi"


End Sub



```



source code yg berhasil untuk menarik data dari database mariadb yg ada di server IDCloudHost di internet dan memasukannya ke dalam sheet di MS Excel:

```vba

    Sub testKoneksi()

    ' Berhasil terkoneksi ke server mariadb di IDCloudHost

    Dim koneksi As ADODB.Connection

    Dim rekordset As ADODB.Recordset

    Set koneksi = New ADODB.Connection
    
    koneksi.ConnectionString = "Driver={MariaDB ODBC 3.1 Driver};Server=127.0.0.1;Port=3306;Database=namadatabase;User=userdatabase;Password=passworddatabase;Option=3"
    
    koneksi.Open


    ' Berhasil loading data ke sheet MS Excel dari Mariadb
    
    Set rekordset = New ADODB.Recordset
    
    rekordset.ActiveConnection = koneksi
    
    rekordset.Source = "antinasi1"
    
    rekordset.Open
    
    
    Sheets("Sheet1").Select
    
    Range("A1").CopyFromRecordset rekordset
    
    
    rekordset.Close
    
    koneksi.Close


End Sub




```



source code untuk menampilkan nama kolom yg ada di tabel MariaDB ke worksheet MS Excel :

```vba

Sub tulisDataKeSheet(KumpulanHasil As ADODB.Recordset)

' Membuat sub tersendiri khusus untuk menuliskan data


Dim lembarKerja As Worksheet

Dim namaKolom As ADODB.field

Dim i As Integer



Set lembarKerja = Worksheets("Sheet1")

lembarKerja.Select


' Meload nama kolom yg ada di tabel database ke Worksheet

For Each namaKolom In KumpulanHasil.Fields

    i = i + 1

    lembarKerja.Cells(1, i).Value = namaKolom.Name
    
Next namaKolom



Range("A2").CopyFromRecordset KumpulanHasil




End Sub





Sub testKoneksi2()


' Parameter koneksi ke server mariadb di IDCloudHost

Dim koneksi As ADODB.Connection

Dim rekordset As ADODB.Recordset


Set koneksi = New ADODB.Connection

koneksi.ConnectionString = "Driver={MariaDB ODBC 3.1 Driver};Server=127.0.0.1;Port=3306;Database=namaDatabase;User=namaUser;Password=isiPassword;Option=3"

koneksi.Open



Set rekordset = New ADODB.Recordset

rekordset.ActiveConnection = koneksi

rekordset.Source = "makannasi1"

rekordset.CursorType = adOpenForwardOnly

rekordset.LockType = adLockReadOnly

rekordset.Open



tulisDataKeSheet rekordset





rekordset.Close

koneksi.Close




End Sub





```






