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


