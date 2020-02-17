Attribute VB_Name = "MainModule"
Option Explicit
Public UserID As String
Public sPassword As String
Public sWarnaLabel As iWarna
Public sWarnaText As iWarna
Public sWarnaBackcolour As iWarna
Public sWarnaTextFrame As iWarna
Public sWarnaTextOption As iWarna
Public sWarnaTextCheck As iWarna

Public Enum typemenu
        setting = 1
        entrian = 2
        report = 3
End Enum

Public Enum iWarna
        hitam = 1
        abuabu = 2
        merahmuda = 3
        merahtua = 4
        kuningnyala = 5
        kuningbiasa = 6
        hijaubiasa = 7
        hijaumenyala = 8
        background = 9
        birumuda = 10
        birumenyala = 11
        birutua = 12
End Enum

 ' hijaumenyala = 8 &H0000FF00& background = 9 birumuda=10 &H00C0C000&
        ' birumenyala=11 &H00FFFF00& birutua=12 &H00808000&

Public Enum toolsBottom
    btm_new = 1
    btm_Undo = 2
    btm_Save = 3
    btm_del = 4
    btm_First = 5
    btm_prev = 6
    btm_next = 7
    btm_Last = 8
    btm_expexcel = 9
    btm_impexcel = 10
    btm_btmclose = 11
    btm_execut = 12
End Enum

Public Enum PssText
    ratakanan = flexAlignRightCenter
    ratatengah = flexAlignCenterCenter
    ratakiri = flexAlignLeftCenter
End Enum
Public Enum TranTypeSts
    Tranin = 0
    TranOut = 1
    InvNo = 2
End Enum
Public Enum DocumentNo
    master_siswa = 1
    transaksi_pendaftaran = 2
    transaksi_kwitansi = 3
    transaksi_kelas = 4
    Transaksi_Cuti = 5
    transaksi_pembelian = 6
    transaksi_masuklain = 7
    transaksi_penjualan = 8
    transaksi_keluarlain = 9
    transaksi_siswakeluar = 10
    transaksi_transfer = 11
    transaksi_trnrab = 12
    transaksi_trngl = 13
    transaksi_lhpp = 14
    transaksi_lhpp_entry = 15
    transaksi_lhpp_tf = 16
End Enum

Public Enum Modul

mnbar_sett_label = 1
mnSetting_DocumentFrm = 2
mnPreferenceFrm = 3
mnPreference_special_Frm = 4
mnModulFrm = 5
mnmriAccessModuleFrm = 6
mnBackupDatabaseFrm = 7
mnUserFrm = 8
mnUserGroup = 9
mnmst_company_frm = 10
mnBrandFrm = 11
mnCategoryFrm = 12
mnFunctionFrm = 13
mnHargaFrm = 14
mnDiskonFrm = 15
mnFeeFrm = 16
mnSatuanProduk = 17
mnTipeGudangFrm = 18
mnGudangFrm = 19
mnProdukFrm = 20
mnSalesmanFrm = 21
mnKolektorFrm = 22
mnSupplierFrm = 23
mnCustomerFrm = 24
mnGLMasterAkun = 25
mnGroupEntryDataFrm = 26
mnBeliFrm = 27
mnMasukLainFrm = 28
mnTransferFrm = 29
mnKeluarLainFrm = 30
mnJualFrm = 31
mnMonitoringTukarFaktur = 32
mnKwitansiFrm = 33
mnGLTransGL = 34
mnLhppFromTF = 35
mnLhppForm = 36
mnLhppEntryForm = 37
mnUserRpt = 38
mnBeliRpt = 39
mnTransferRpt = 40
mnKeluarLainRpt = 41
mnJualRpt = 42
mnJualFeeCustomerRpt = 43
mnJualMonitoringFakturRpt = 44
mnJualRptByProduk = 45
mnKwitansiRpt = 46
mnTransaksiByProdukRpt = 47
mnGLTransGLRpt = 48
mnGLTransGLRkpRpt = 49
mnJualRptByKomisi = 50
mnLabaRugiByCustomerByProductRpt = 51
mnAggingDetailRpt = 52
mnMonitoringBayarFakturRpt = 53
mnMonitoringLHPPEntryRpt = 54
mnLHPPTFReport = 55
mnLHPPReport = 56
mnKartuStockRpt = 57
mnMutasiStockRpt = 58
mnProsesBulanan = 59
mnBrandRpt = 60
mnCategoryRpt = 61
mnFungsiRpt = 62
mnDiskonRpt = 63
mnHargaRpt = 64
mnSatuanProdukRpt = 65
mnTipeGudangRpt = 66
mnGudang_Rpt = 67
mnProdukRpt = 68
mnSalesmanRpt = 69
mnKolektorRpt = 70
mnMaster_Supplier_Rpt = 71
mnMaster_Customer_Rpt = 72
mnGLMasterAkunRpt = 73
mnGLMasterGroupAkunRpt = 74
mnLoginFrm = 75
mnMasukLainRpt = 76
mnAreaFrm = 77
mnAreaRpt = 78
mnBeliRptByProduk = 79
mnProsesExportFiles = 80
mnProsesImportFiles = 81
mnSettingLockSalesFrm = 82
mnJualFeeSalesPurchaseCustomerRpt = 83
mnMonitoringPelunasanFaktur = 84
mnMonitoringPelunasanFakturRpt = 85

End Enum

Public Enum TampilData
    SesuaiID = 0
    Awal = 1
    Sebelum = 2
    Setelah = 3
    Akhir = 4
End Enum
Public Enum sTatusCell
    Insert = 1
    Update = 2
    Delete = 3
End Enum
Public Enum StatusForm
    Normal = 0
    DataBaru = 1
    MainMenu = 2
    SaveMati = 3
    ActvClose = 4
    RefrshRpt = 5
    MultiItem = 6
    StatusClose = 7
    NormalPlusExec = 8
    SettingForm = 9
    NormalClosePlusExec = 10

End Enum
Public Enum StatusMessage
    Normalmsg = 0
    dataBarumsg = 1
    MainMenumsg = 2
    WaitMsg = 3
    DataBaruPOSMsg = 4
End Enum
Public Enum ShowLabel
    TranParkir = 0
End Enum

Public Enum BrowseTables
    BrowsUser = 1
    BrowsUserAkses = 2
    BrowsLogin = 3
    BrowsAgama = 4
    BrowsPekerjaan = 5
    BrowsCompanyLogin = 6
    BrowsPelajaranGroup = 7
    BrowsPelajaranLevel = 8
    BrowsSettingDocument = 9
    BrowsPendaftaran = 10
    BrowsKwitansi = 11
    BrowsKelas = 12
    BrowsMasterPreferencesSpecial = 13
    BrowsKartuKelas = 14
    BrowsCuti = 15
    BrowsBrand = 16
    BrowsCategory = 17
    BrowsFunction = 18
    BrowsHarga = 19
    BrowsDiskon = 20
    BrowsTipeGudang = 21
    BrowsMasterProduk = 22
    BrowsSatuanProduk = 23
    BrowsGudang = 24
    BrowsSupplier = 25
    Browscustomer = 26
    BrowsPembelian = 27
    BrowsMasukLain = 28
    BrowsPenjualan = 29
    BrowsKeluarLain = 30
    BrowsUserGroup = 31
    BrowsModule = 32
    BrowsSiswaKeluar = 33
    BrowsJenisMateri = 34
    BrowsTransfer = 35
    BrowsMasterDefaultPelajaran = 36
    
    BrowsAkunGroupSumberData = 38
    BrowsAkunSumberData = 39
    BrowsAkunMasterCOA = 40
    BrowsAkunRAB = 41
    BrowsAkunTransRAB = 42
    BrowsAkunTransGL = 43
    BrowsFee = 44
    BrowsSalesman = 45
    
    BrowsBarSettLabelBarcode = 46
    BrowsKolektor = 47
    Browslhpp = 48
    Browslhppdetail1 = 49
    
    Browslhppentry = 50
    Browslhppentrydetail1 = 51
    
    Browslhpptf = 52
    Browslhpptfdetail1 = 53
    BrowsCompany = 54
    BrowsArea = 55
    
    
End Enum
Public Enum BolehEdit
    ya = 0
    tidak = 1
End Enum
Public Enum ModeIncDec
    increase = 0
    decrease = 1
End Enum
Public Enum DBaseConection
    login = 0
    Modul = 1
End Enum

Public Enum FlagFind
    First = 1
    Last = 2
End Enum

Public Enum FlagUp
    byUp = 1
    byDown = 2
    byAuto = 3
End Enum


Public Enum sBottom
sNewData = 1
sUndoData = 2
sSaveData = 3
sDeleteData = 4
sMoveFirst = 5
sMovePrevious = 6
sMoveNext = 7
sMoveLast = 8
sExportToExecl = 9
sImportFromExecl = 10
sCloseform = 11
sExecution = 12
End Enum

Public Enum urutBy
ubDescending = 1
ubAscending = 2
End Enum
Public Sub Main()
MenuFrm.Show
LoginFrm.Show 1
End Sub
Public Property Get Conectionku(sDBaseConection As DBaseConection) As Variant
If sDBaseConection = login Then
    Conectionku = "DRIVER={" & MenuFrm.Driverku & "};SERVER=" & MenuFrm.Serverku & ";DATABASE=" & MenuFrm.Databaseku & ";UID=" & "sa" & ";PWD=" & "spvsql" & ";PORT=" & MenuFrm.Portku & ";OPTION=3"
Else
    Conectionku = "DRIVER={" & MenuFrm.Driverku & "};SERVER=" & MenuFrm.Serverku & ";DATABASE=" & MenuFrm.sCompnyDatabase & ";UID=" & "sa" & ";PWD=" & "spvsql" & ";PORT=" & MenuFrm.Portku & ";OPTION=3"
End If
End Property
Public Property Get ConStringAlternatif(sDatabaseku As String) As Variant
Select Case sDatabaseku
Case "SAIPST"
ConStringAlternatif = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=ODPS;Data Source=MTGMIS22"
Case "DATACENTER"
ConStringAlternatif = "Provider=SQLOLEDB.1;Password=spvsql;Persist Security Info=True;User ID=sa;Initial Catalog=DataCenter;Data Source=10.17.0.12"
                      'Provider=SQLOLEDB.1;Password=spvsql;Persist Security Info=True;User ID=sa;Initial Catalog=DataCenter;Data Source=10.17.0.12
Case "ODPS"
ConStringAlternatif = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=ODPS;Data Source=MTGMIS22"
End Select
'ConString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=ODPS;Data Source=MTGMIS22"
'"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Malaka.mdb;Persist Security Info=False"
End Property
Public Sub ShowMessage(sDesc As String, sProcName As String)
MsgBox sDesc, , sProcName

End Sub

Public Sub highlighttext(oText As TextBox)
oText.SelStart = 0
oText.SelLength = Len(oText.text)
End Sub

Public Sub DoKeyDown(KeyCode As Integer, istatus As StatusForm)
Select Case KeyCode
Case 13
     'SendKeys "{Tab}"
Case vbKeyF2
    MenuFrm.NewData
Case vbKeyF3
    MenuFrm.SaveData
Case vbKeyF4
    If istatus = Normal Then
        MenuFrm.DeleteData
    Else
        MenuFrm.UndoData
    End If
    
Case vbKeyF9
    '------------------ Bayar ---------------
    MenuFrm.Bayar
Case vbKeyF10
    '--------------- Simpan/Cetak-------------
    MenuFrm.SavePrint
Case vbKeyEscape
     MenuFrm.Closeform
       
End Select
End Sub
Public Sub ShowShortMenu(ShowLabel)
Select Case ShowLabel
Case 0
MenuFrm.LblPesanku = "F2 : New | F3 Save | F4 : Delete | F5 : CLose"
End Select
End Sub
Public Sub ShowFormMessage(istatus As StatusMessage)
Select Case istatus
Case StatusMessage.WaitMsg
    MenuFrm.LblPesanku.Caption = " Tunggu , Sedang Dalam Proses Pemanggilan Data !!"
Case 0
    MenuFrm.LblPesanku.Caption = " Tekan : F2 -> Data Baru | F3 -> Simpan | F4 -> Hapus Data | Esc -> Tutup"
Case 1
    MenuFrm.LblPesanku.Caption = " Tekan : F3 -> Simpan | F4 -> Batal | Esc -> Tutup"
Case 2
    MenuFrm.LblPesanku.Caption = " Harap Pilih Menu Paling Atas"
Case 3
    MenuFrm.LblPesanku.Caption = " Tekan : Insert = Sisip Data | Delete : Hapus Data | F2 = Simpan "
Case StatusMessage.DataBaruPOSMsg
                                '" Tekan : F2 -> Data Baru | F3 -> Simpan Data | F4 -> Hapus Data | Esc -> Tutup"
    MenuFrm.LblPesanku.Caption = " Tekan : F3 -> Simpan Data | F4 -> Batal | F9 -> Bayar | F10 -> Simpan/Cetak | ESC -> Tutup"
End Select
End Sub
Public Function EncryptPassword(KataKunci As String) As String
Dim X As Integer
Dim MyPassword As String, enc As String
For X = 1 To Len(KataKunci)
    enc = Mid(KataKunci, X, 1)
    MyPassword = MyPassword & Chr(Asc(enc) + 13)
Next
EncryptPassword = MyPassword
End Function
Public Function DecryptPassword(KataKunci As String) As String
Dim X As Integer
Dim MyPassword As String, enc As String
For X = 1 To Len(KataKunci)
    enc = Mid(KataKunci, X, 1)
    MyPassword = MyPassword & Chr(Asc(enc) - 13)
Next
DecryptPassword = MyPassword
End Function

Public Function ToText(svalue As Variant) As String
If IsNull(svalue) = True Then
    ToText = ""
Else
    ToText = svalue
End If
ToText = Replace(ToText, "\", "\\")
ToText = Replace(ToText, "'", "\'")
End Function

Public Function ToNumber(svalue As Variant) As Double
If IsNull(svalue) = True Then
        ToNumber = 0
Else
    If IsNumeric(svalue) = False Then
        ToNumber = 0
    Else
        ToNumber = CDbl(svalue)
    End If
End If
End Function
Public Function ToDate(svalue As Variant) As Date
If IsNull(svalue) = True Then
        ToDate = Null
Else
    If IsDate(svalue) = False Then
        ToDate = Null
    Else
        ToDate = CDate(svalue)
    End If
End If
End Function

Public Sub TypeNumber(oText As TextBox)
Dim svalue As Double
svalue = ToNumber(oText.text)
oText.text = Format(svalue, "#,##0")
oText.SelStart = Len(oText.text)
End Sub

Public Function CheckKey(ByVal KeyAscii As Integer) As Integer
    If KeyAscii = 13 Then
        CheckKey = 13
        Exit Function
    End If
    If KeyAscii = vbKeyBack Then
        CheckKey = vbKeyBack
        Exit Function
    End If
    If KeyAscii = 46 Then
        CheckKey = 46
        Exit Function
    End If
    If KeyAscii = 45 Then
        CheckKey = 45
        Exit Function
    End If
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
        CheckKey = KeyAscii
    Else
        CheckKey = 0
    End If
End Function

Public Function formatRupiah(sNumber As Double)
formatRupiah = Format(sNumber, "###,###,##0")
End Function


Public Property Get oConstring(sServer As String, sDatabase As String) As Variant
'oConstring = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=D:\dataKu\Hotelku\hotelku.mdb;Persist Security Info=False"
'Constring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sServer & sDatabase & ";Persist Security Info=False"
'ConString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=ODPS;Data Source=MTGMIS22"
'"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Malaka.mdb;Persist Security Info=False"
End Property


'Public Property Get Conectionku(cnDataBase As DBaseConection)
''Conectionku = "Provider=SQLOLEDB.1;Password=spvsql;Persist Security Info=True;User ID=sa;Initial Catalog=DataCenter;Data Source=10.17.0.12"
''Conectionku = "Provider=SQLOLEDB.1;Password=spvsql;Persist Security Info=True;User ID=sa;Initial Catalog=" & sDatabaseku & ";Data Source=" & sServerku
''Constring DBaseConection.Modul
'End Property

Public Function ReadExcel(ByVal sfile As String, ByVal sViewData As String) As ADODB.Recordset
    On Error GoTo errhandler
    Const MODULENAME_ = "Ordering System"
    Const PROCNAME_ = "RunSPReturnRS"
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sconn As String
 
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockBatchOptimistic
 
    sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & sfile
    rs.Open sViewData, sconn
    Set ReadExcel = rs
    Set rs = Nothing
    Exit Function
errhandler:
    Set rs = Nothing
    Debug.Print Err.Description
End Function
Public Sub oFormatWarnaLabel(wLabel As iWarna, wText As iWarna, wBackground As iWarna, uForm As Form)
Dim i As Integer
For i = 0 To uForm.Label1.Count - 1
    uForm.Label1(i).ForeColor = warna(sWarnaLabel)  '&HFFC0C0
    uForm.Label1(i).BackColor = warna(sWarnaBackcolour) '&HFF&
    uForm.Label1(i).FontBold = True
Next

For i = 0 To uForm.Text1.Count - 1
    uForm.Text1(i).ForeColor = warna(sWarnaText)  '&HFFC0C0
    uForm.Text1(i).BorderStyle = 0
    uForm.Text1(i).FontBold = True
Next
For i = 0 To uForm.Frame1.Count - 1
    uForm.Frame1(i).BackColor = warna(sWarnaBackcolour)  '&HFFC0C0
    uForm.Frame1(i).ForeColor = &HFFFF&
    uForm.Frame1(i).FontBold = True
Next

End Sub

Public Property Get warna(iWarna As iWarna)
    Select Case iWarna
    Case 1
        warna = &H0&
    Case 2
        warna = &H404040
    Case 3
        warna = &HC0&
    Case 4
        warna = &H80&
    Case 5
        warna = &HFFFF&
    Case 6
        warna = &HC0C0&
    Case 7
        warna = &H8000&
    Case 8
        warna = &HFF00&
    Case 9
        warna = &HFFC0C0
    Case 10
            warna = &HC0C000
    Case 11
            warna = &HFFFF00
    Case 12
            warna = &H808000
End Select
        '
        ' hitam = 1 &H00000000& abuabu = 2 &H00404040& merahmuda = 3 &H000000C0&
        ' merahtua = 4 &H00000080& kuningnyala = 5 &H0000FFFF&
        ' kuningbiasa = 6 &H0000C0C0& hijaubiasa = 7 &H00008000&
        ' hijaumenyala = 8 &H0000FF00& background = 9 birumuda=10 &H00C0C000&
        ' birumenyala=11 &H00FFFF00& birutua=12 &H00808000&
        
End Property

Public Sub oFormatWarnaText(wText As iWarna, wBackground As iWarna, uForm As Form)
Dim i As Integer
For i = 0 To uForm.Label1.Count - 1
    uForm.Label1(i).ForeColor = warna(wText)  '&HFFC0C0
    'uForm.Label1(i).BackColor = warna(wBackground) '&HFF&
Next
End Sub
Public Sub oFormatFrameBackground(sFrame As Frame)
sFrame.BackColor = warna(background)
End Sub

Public Sub oFormatOption(sSampaiDengan As Integer, sForm As Form)
Dim irow As Integer
For irow = 0 To sForm.Option1.Count - 1
    sForm.Option1(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Option1(irow).ForeColor = warna(sWarnaTextOption)
Next
If sSampaiDengan = 1 Then Exit Sub

For irow = 0 To sForm.Option2.Count - 1
    sForm.Option2(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Option2(irow).ForeColor = warna(sWarnaTextOption)
Next
If sSampaiDengan = 2 Then Exit Sub

For irow = 0 To sForm.Option3.Count - 1
    sForm.Option3(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Option3(irow).ForeColor = warna(sWarnaTextOption)
Next
If sSampaiDengan = 3 Then Exit Sub

For irow = 0 To sForm.Option4.Count - 1
    sForm.Option4(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Option4(irow).ForeColor = warna(sWarnaTextOption)
Next
If sSampaiDengan = 4 Then Exit Sub

For irow = 0 To sForm.Option5.Count - 1
    sForm.Option5(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Option5(irow).ForeColor = warna(sWarnaTextOption)
Next

End Sub

Public Sub oFormatCheckList(sSampaiDengan As Integer, sForm As Form)
Dim irow As Integer
For irow = 0 To sForm.Check1.Count - 1
    sForm.Check1(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Check1(irow).ForeColor = warna(sWarnaTextCheck)
Next
If sSampaiDengan = 1 Then Exit Sub

For irow = 0 To sForm.Check2.Count - 1
    sForm.Check2(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Check2(irow).ForeColor = warna(sWarnaTextCheck)
Next
If sSampaiDengan = 2 Then Exit Sub

For irow = 0 To sForm.Check3.Count - 1
    sForm.Check3(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Check3(irow).ForeColor = warna(sWarnaTextCheck)
Next
If sSampaiDengan = 3 Then Exit Sub

For irow = 0 To sForm.Check4.Count - 1
    sForm.Check4(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Check4(irow).ForeColor = warna(sWarnaTextCheck)
Next
If sSampaiDengan = 4 Then Exit Sub

For irow = 0 To sForm.Check5.Count - 1
    sForm.Check5(irow).BackColor = warna(sWarnaBackcolour)
    sForm.Check5(irow).ForeColor = warna(sWarnaTextCheck)
Next

End Sub

Public Function oRubahKutip(sText As String)
oRubahKutip = Replace(sText, "'", "\'")
End Function

Public Function toNumberIndonesia(sText As String) As String
Dim stextQ As String

If MenuFrm.sisIndonesianFormat = "Y" Then
    stextQ = sText
    stextQ = Replace(sText, ".", "x")
    stextQ = Replace(stextQ, ",", ".")
    stextQ = Replace(stextQ, "x", "", DBaseConection.Modul)
  
    toNumberIndonesia = stextQ
Else
    toNumberIndonesia = sText
    stextQ = sText
    stextQ = Replace(sText, ".", "x")
    stextQ = Replace(stextQ, ",", "", DBaseConection.Modul)
    stextQ = Replace(stextQ, "x", ".")
  
    toNumberIndonesia = stextQ
End If
End Function
Public Sub Sendkeys(text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys text, wait
   Set WshShell = Nothing
End Sub
