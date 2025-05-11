VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Homepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Detail_Click()

End Sub

Private Sub Form_Load()
    Me.optMenu = 1 ' Default to Button 1
    
    Call optMenu_AfterUpdate
    Me.tahunText.Value = Year(Date)
    Call LoadReportPenerimaan
    Call LoadReportPembayaran
    Call LoadJumlah
End Sub
Private Sub LoadJumlah()
    Dim jumlah1 As Currency
    Dim jumlah2 As Currency
    Dim jumlah3 As Currency
    
    jumlah1 = Me.jumlahPenerimaanText.Value
    jumlah2 = Me.jumlahPembayaranText.Value
    jumlah3 = jumlah1 - jumlah2
    
    Me.jumlah1.Value = jumlah1
    Me.jumlah2.Value = jumlah2
    Me.jumlah3.Value = jumlah3
End Sub

Private Sub optMenu_AfterUpdate()
    
    Me.rekodTab.Visible = True
    Me.semakTab.Visible = False
    Me.printTab.Visible = False
    Select Case Me.optMenu.Value
        Case 1
            Me.rekodTab.Visible = True
            Me.semakTab.Visible = False
            Me.printTab.Visible = False
            Call ResetAddForm
        Case 2
            Me.rekodTab.Visible = False
            Me.semakTab.Visible = True
            Me.printTab.Visible = False
            Me.subform1.Form.Requery
            Me.subform2.Form.Requery
        Case 3
            Me.rekodTab.Visible = False
            Me.semakTab.Visible = False
            Me.printTab.Visible = True
            Select Case Me.printTab.Value
                Case 0
                    Call LoadReportPenerimaan
                Case 1
                    Call LoadReportPembayaran
            End Select
            
    End Select
    
End Sub
Private Sub ResetAddForm()
    ' Reset controls for new data entry
    Me.penerimaanCombo = Null        ' Combo box for subkategori
    Me.amaunPenerimaan = Null             ' Text box for amount
    Me.tarikhPenerimaan = Null               ' Set current date as default
    
    Me.pembayaranCombo = Null        ' Combo box for subkategori
    Me.amaunPembayaran = Null             ' Text box for amount
    Me.tarikhPembayaran = Null               ' Set current date as default
End Sub
Private Sub LoadReportPenerimaan()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim totalPersepuluhan As Currency
    Dim totalPersembahan As Currency
    Dim totalSumbanganAnggota As Currency
    Dim totalUnit As Currency
    Dim totalSumbanganKhas As Currency
    Dim totalSumbanganLain As Currency
    Dim totalPenerimaan As Currency
    Dim selectedMonth As Integer
    Dim selectedYear As Integer
    selectedMonth = Me.bulanCombo.Value
    selectedYear = Me.tahunText.Value
    
    Set db = CurrentDb
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_1 FROM Expenses WHERE kategori_id = 1 AND subkategori_id = 1 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalPersepuluhan = Nz(rs!total_1, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_2 FROM Expenses WHERE kategori_id = 1 AND subkategori_id = 2 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalPersembahan = Nz(rs!total_2, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_3 FROM Expenses WHERE kategori_id = 1 AND subkategori_id = 3 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalSumbanganAnggota = Nz(rs!total_3, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_4 FROM Expenses WHERE kategori_id = 1 AND subkategori_id = 4 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalSumbanganKhas = Nz(rs!total_4, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_5 FROM Expenses WHERE kategori_id = 1 AND subkategori_id = 5 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalSumbanganLain = Nz(rs!total_5, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_6 FROM Expenses WHERE kategori_id = 1 AND subkategori_id = 6 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalUnit = Nz(rs!total_6, 0)
    End If
    rs.Close
    
    totalPenerimaan = totalPersepuluhan + totalPersembahan + totalSumbanganAnggota + totalUnit + totalSumbanganKhas + totalSumbanganLain
    
    Me.persepuluhanText.Value = totalPersepuluhan
    Me.persembahanText.Value = totalPersembahan
    Me.sumbanganText.Value = totalSumbanganAnggota
    Me.unitText.Value = totalUnit
    Me.khasText.Value = totalSumbanganKhas
    Me.lainText.Value = totalSumbanganLain
    Me.jumlahPenerimaanText.Value = totalPenerimaan
    Debug.Print "Function started"
End Sub
Private Sub LoadReportPembayaran()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim totalStaf As Currency
    Dim totalEpf As Currency
    Dim totalUtiliti As Currency
    Dim totalMesyuarat As Currency
    Dim totalInsurans As Currency
    Dim totalPenggajian As Currency
    Dim totalKebajikan As Currency
    Dim totalPenginjilan As Currency
    Dim totalLawatan As Currency
    Dim totalSumbangan As Currency
    Dim totalPerayaan As Currency
    Dim totalPersepuluhan As Currency
    Dim totalSumbangan2 As Currency
    Dim totalSumbangan3 As Currency
    Dim totalSumbangan4 As Currency
    Dim totalPembayaran As Currency
    Dim selectedMonth As Integer
    Dim selectedYear As Integer
    selectedMonth = Me.bulanCombo.Value
    selectedYear = Me.tahunText.Value
    
    Set db = CurrentDb
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_1 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 1 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalStaf = Nz(rs!total_1, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_2 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 2 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalUtiliti = Nz(rs!total_2, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_3 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 3 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalEpf = Nz(rs!total_3, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_4 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 4 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalMesyuarat = Nz(rs!total_4, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_5 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 5 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalInsurans = Nz(rs!total_5, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_6 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 6 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalPenggajian = Nz(rs!total_6, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_7 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 7 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalKebajikan = Nz(rs!total_7, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_8 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 8 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalPenginjilan = Nz(rs!total_8, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_9 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 9 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalLawatan = Nz(rs!total_9, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_10 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 10 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalSumbangan = Nz(rs!total_10, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_11 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 11 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalPerayaan = Nz(rs!total_11, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_12 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 12 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalPersepuluhan = Nz(rs!total_12, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_13 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 13 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalSumbangan2 = Nz(rs!total_13, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_14 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 14 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalSumbangan3 = Nz(rs!total_14, 0)
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT Sum(amount) AS total_15 FROM Expenses WHERE kategori_id = 2 AND subkategori_id = 15 AND Month(tarikh) = " & selectedMonth & " AND Year(tarikh) = " & selectedYear)
    If Not rs.EOF Then
        totalSumbangan4 = Nz(rs!total_15, 0)
    End If
    rs.Close
    
    totalPembayaran = totalStaf + totalUtiliti + totalEpf + totalMesyuarat + totalInsurans + totalPenggajian + totalKebajikan + totalPenginjilan + totalLawatan + totalSumbangan + totalPerayaan + totalPersepuluhan + totalSumbangan2 + totalSumbangan3 + totalSumbangan4
    Me.stafText.Value = totalStaf
    Me.utilitiText.Value = totalUtiliti
    Me.epfText.Value = totalEpf
    Me.mesyuaratText.Value = totalMesyuarat
    Me.insuransText.Value = totalInsurans
    Me.pastorText.Value = totalPenggajian
    Me.kebajikanText.Value = totalKebajikan
    Me.penginjilanText.Value = totalPenginjilan
    Me.lawatanText.Value = totalLawatan
    Me.sumbanganText2.Value = totalSumbangan
    Me.perayaanText.Value = totalPerayaan
    Me.persepuluhantext2.Value = totalPersepuluhan
    Me.sumbanganText3.Value = totalSumbangan2
    Me.sumbanganText4.Value = totalSumbangan3
    Me.sumbanganText5.Value = totalSumbangan4
    Me.jumlahPembayaranText.Value = totalPembayaran
    Debug.Print "Function2 started"
End Sub

Private Sub paparBtn_Click()
    Call LoadReportPenerimaan
    Call LoadReportPembayaran
    Call LoadJumlah
End Sub

Private Sub pembayaran_add_Click()

    Dim db As DAO.Database
    Dim subkategori_id As Long
    Dim kategori_id As Long
    Dim tarikh As Date
    Dim amount As Currency
    Dim strSQL As String
    
     On Error GoTo ErrorHandler

    ' Set database object
    Set db = CurrentDb
    
    ' Basic validation
    If IsNull(Me.pembayaranCombo) Or IsNull(Me.amaunPembayaran) Then
        MsgBox "Sila pilih jenis Transaksi Pembayaran dan isi amaun wang", vbExclamation
        Exit Sub
    End If
    
    subkategori_id = Me.pembayaranCombo.Value
    amount = CCur(Me.amaunPembayaran.Value)
    kategori_id = 2
    tarikh = Me.tarikhPembayaran.Value
    
    ' Create SQL statement with kategori hardcoded as 1
    strSQL = "INSERT INTO Expenses (kategori_id, subkategori_id, tarikh, amount) " & _
          "VALUES (2, " & subkategori_id & ", #" & Format(tarikh, "yyyy-mm-dd") & "#, " & amount & ");"

   
    ' Execute SQL
    db.Execute strSQL, dbFailOnError

    MsgBox "Berjaya Direkodkan.", vbInformation

    ' Optional: Clear the form
    Me.pembayaranCombo = Null
    Me.tarikhPembayaran = Null
    Me.amaunPembayaran = Null

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub peneriman_add_Click()
    Dim db As DAO.Database
    Dim subkategori_id As Long
    Dim kategori_id As Long
    Dim tarikh As Date
    Dim amount As Currency
    Dim strSQL As String
    
     On Error GoTo ErrorHandler

    ' Set database object
    Set db = CurrentDb
    
    ' Basic validation
    If IsNull(Me.penerimaanCombo) Or IsNull(Me.amaunPenerimaan) Then
        MsgBox "Sila pilih jenis Transaksi Penerimaan dan isi amaun wang", vbExclamation
        Exit Sub
    End If
    
    subkategori_id = Me.penerimaanCombo.Value
    amount = CCur(Me.amaunPenerimaan.Value)
    kategori_id = 1
    tarikh = Me.tarikhPenerimaan.Value
    
    ' Create SQL statement with kategori hardcoded as 1
    strSQL = "INSERT INTO Expenses (kategori_id, subkategori_id, tarikh, amount) " & _
          "VALUES (1, " & subkategori_id & ", #" & Format(tarikh, "yyyy-mm-dd") & "#, " & amount & ");"

   
    ' Execute SQL
    db.Execute strSQL, dbFailOnError

    MsgBox "Berjaya Direkodkan.", vbInformation

    ' Optional: Clear the form
    Me.penerimaanCombo = Null
    Me.tarikhPenerimaan = Null
    Me.amaunPenerimaan = Null

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
