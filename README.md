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
















