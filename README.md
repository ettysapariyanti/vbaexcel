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








