Attribute VB_Name = "FindData"
Public Function GetTable(sTable As String, sField As String, sKondisi As String, sOrderBy As String) As ADODB.Recordset
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.Modul)
sQuery = "Select " & IIf(sField = "", "*", sField) & " From sTable Where " & IIf(sKondisi = "", "''=''", sKondisi) & " Order By " & sOrderBy
Set GetTable = oConku.Execute(sQuery)
'GetTable = oRstku
End Function


Public Function GetProductID(sKeyFind As String, sPriority As Integer, sCustmrID As String) As String
On Error GoTo errhandler
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.Modul)
sQuery = "Select Dbo.FgetProductID('" & Trim(sKeyFind) & "'," & sPriority & ",'" & sCustmrID & "')"
Set oRstku = oConku.Execute(sQuery)
GetProductID = oRstku(0)
oConku.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "GetProductID"
End Function

Public Sub oGetAccesMenuByModulID(sGroupUser As String, sModulID As Integer)
On Error GoTo errhandler
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
Dim sAccess As String
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.login)
'sQuery = "SELECT kodegroup,a.modulid,CONCAT(baca,tulis,edit,hapus,cetak) AS accesmodule,b.`transid` FROM master_moduleaccess a"
'sQuery = sQuery & " INNER JOIN `master_module` b ON b.`Modulid`=a.`modulid` "
'sQuery = sQuery & "where kodegroup='" & sGroupUser & "' and a.modulid=" & sModulID & ""

sQuery = "SELECT a.modulid,CONCAT(MAX(baca),MAX(tulis),MAX(edit),MAX(hapus),MAX(cetak)) AS accesmodule,d.`transid` FROM master_moduleaccess a "
sQuery = sQuery & " INNER JOIN `master_group_user` b ON b.`kodegroup`=a.kodegroup "
sQuery = sQuery & " INNER JOIN `groupaccess` c ON c.`groupid`=b.`groupid`  "
sQuery = sQuery & " INNER JOIN `master_module` d ON d.`Modulid`=a.`Modulid`  "
sQuery = sQuery & " WHERE c.`emplcode`='" & sGroupUser & "' AND c.`accessctrl`='Y' and a.modulid=" & sModulID & ""
sQuery = sQuery & " GROUP BY a.modulid "

Set oRstku = oConku.Execute(sQuery)

''MenuFrm.toolsEnableTrue

With oRstku
     If .EOF() Then Exit Sub
     sAccess = .Fields("accesmodule")
     
     '----- menu bottom access Data Baru -----
     If Mid(sAccess, 1, 1) = "Y" Or (MenuFrm.isAdmin = "Y") Then
     
        If .Fields("transid") = "3" Then
            MenuFrm.Toolbar1.Buttons(btm_First).Enabled = False
            MenuFrm.Toolbar1.Buttons(btm_prev).Enabled = False
            MenuFrm.Toolbar1.Buttons(btm_next).Enabled = False
            MenuFrm.Toolbar1.Buttons(btm_Last).Enabled = False
        Else
            If sModulID = 8 And MenuFrm.isAdmin = "N" Then
                MenuFrm.Toolbar1.Buttons(btm_First).Enabled = False
                MenuFrm.Toolbar1.Buttons(btm_prev).Enabled = False
                MenuFrm.Toolbar1.Buttons(btm_next).Enabled = False
                MenuFrm.Toolbar1.Buttons(btm_Last).Enabled = False
            Else
            
                MenuFrm.Toolbar1.Buttons(btm_First).Enabled = True
                MenuFrm.Toolbar1.Buttons(btm_prev).Enabled = True
                MenuFrm.Toolbar1.Buttons(btm_next).Enabled = True
                MenuFrm.Toolbar1.Buttons(btm_Last).Enabled = True
            End If
        End If
        
     Else
     
        MenuFrm.Toolbar1.Buttons(btm_First).Enabled = False
        MenuFrm.Toolbar1.Buttons(btm_prev).Enabled = False
        MenuFrm.Toolbar1.Buttons(btm_next).Enabled = False
        MenuFrm.Toolbar1.Buttons(btm_Last).Enabled = False
     End If
     
     '----- menu bottom access Data Baru -----
     If Mid(sAccess, 2, 1) = "Y" Or (MenuFrm.isAdmin = "Y") Then
        If .Fields("transid") = "2" Then
            MenuFrm.Toolbar1.Buttons(1).Enabled = True
        Else
            MenuFrm.Toolbar1.Buttons(1).Enabled = False
        End If
     Else
        MenuFrm.Toolbar1.Buttons(1).Enabled = False
     End If
     
     '----- menu bottom access Edit Baru -----
     If Mid(sAccess, 3, 1) = "Y" Or (MenuFrm.isAdmin = "Y") Then
        If .Fields("transid") = "3" Then
            MenuFrm.Toolbar1.Buttons(3).Enabled = False
        Else
            MenuFrm.Toolbar1.Buttons(3).Enabled = True
        End If
     Else
        MenuFrm.Toolbar1.Buttons(3).Enabled = False
     End If

     '----- menu bottom access Edit Baru -----
     If Mid(sAccess, 4, 1) = "Y" Or (MenuFrm.isAdmin = "Y") Then
        If .Fields("transid") = "2" Then
            MenuFrm.Toolbar1.Buttons(4).Enabled = True
        Else
            MenuFrm.Toolbar1.Buttons(4).Enabled = False
        End If
     Else
        MenuFrm.Toolbar1.Buttons(4).Enabled = False
     End If
     
     If Mid(sAccess, 5, 1) = "Y" Then
        MenuFrm.Toolbar1.Buttons(12).Enabled = True
     Else
        MenuFrm.Toolbar1.Buttons(12).Enabled = False
     End If
     

     
     
End With
oConku.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "GetProductID"
End Sub

Public Sub oGetCekModul(sUserID As String)
On Error GoTo errhandler
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
Dim sAccess As String
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.login)
'sQuery = "select kodegroup,modulid,CONCAT(baca,tulis,edit,hapus,cetak) as accesmodule from master_moduleaccess "
'sQuery = sQuery & "where kodegroup='" & sGroupUser & "' "

sQuery = "SELECT a.modulid,CONCAT(MAX(baca),MAX(tulis),MAX(edit),MAX(hapus),MAX(cetak)) AS accesmodule FROM master_moduleaccess a "
sQuery = sQuery & "INNER JOIN `master_group_user` b ON b.`kodegroup`=a.kodegroup "
sQuery = sQuery & "INNER JOIN `groupaccess` c ON c.`groupid`=b.`groupid`  "
sQuery = sQuery & "WHERE c.`emplcode`='" & MenuFrm.sUserID & "' AND c.`accessctrl`='Y' "
sQuery = sQuery & "GROUP BY a.modulid "

Set oRstku = oConku.Execute(sQuery)

     Do While Not oRstku.EOF()
     sAccess = oRstku.Fields("accesmodule")
     If oRstku.Fields("modulid") = mnPreferences Then
        If MenuFrm.sPassSuperUser = "210309" Then
            MenuFrm.mnPreferenceFrm.Visible = True
        Else
            MenuFrm.mnPreferenceFrm.Visible = False
        End If
     End If
     If sAccess = "NNNNN" Then
        oMenuEnable oRstku.Fields("modulid"), False
     Else
        oMenuEnable oRstku.Fields("modulid"), True
     End If
     

     oRstku.MoveNext
     Loop
oConku.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "GetProductID"
End Sub


Public Sub oMenuEnable(sModule As Integer, sBoolean As Boolean)

With MenuFrm
    Select Case sModule
    
    Case mnbar_sett_label
        .mnbar_sett_label.Enabled = sBoolean
    Case mnSetting_DocumentFrm
        .mnSetting_DocumentFrm.Enabled = sBoolean
    Case mnPreferenceFrm
        .mnPreferenceFrm.Enabled = sBoolean
    Case mnPreference_special_Frm
        .mnPreference_special_Frm.Enabled = sBoolean
    Case mnModulFrm
        
        .mnModulFrm.Enabled = sBoolean
    Case mnmriAccessModuleFrm
        .mnmriAccessModuleFrm.Enabled = sBoolean
    Case mnBackupDatabaseFrm
        .mnBackupDatabaseFrm.Enabled = sBoolean
    Case mnUserFrm
        .mnUserFrm.Enabled = sBoolean
    Case mnUserGroup
        .mnUserGroup.Enabled = sBoolean
    Case mnmst_company_frm
        .mnmst_company_frm.Enabled = sBoolean
    Case mnBrandFrm
        .mnBrandFrm.Enabled = sBoolean
    Case mnCategoryFrm
        .mnCategoryFrm.Enabled = sBoolean
    Case mnFunctionFrm
        .mnFunctionFrm.Enabled = sBoolean
    Case mnHargaFrm
        .mnHargaFrm.Enabled = sBoolean
    Case mnDiskonFrm
        .mnDiskonFrm.Enabled = sBoolean
    Case mnFeeFrm
        .mnFeeFrm.Enabled = sBoolean
    Case mnSatuanProduk
        .mnSatuanProduk.Enabled = sBoolean
    Case mnTipeGudangFrm
        .mnTipeGudangFrm.Enabled = sBoolean
    Case mnGudangFrm
        .mnGudangFrm.Enabled = sBoolean
    Case mnProdukFrm
        .mnProdukFrm.Enabled = sBoolean
    Case mnSalesmanFrm
        .mnSalesmanFrm.Enabled = sBoolean
    Case mnKolektorFrm
        .mnKolektorFrm.Enabled = sBoolean
    Case mnSupplierFrm
        .mnSupplierFrm.Enabled = sBoolean
    Case mnCustomerFrm
        .mnCustomerFrm.Enabled = sBoolean
    Case mnGLMasterAkun
        .mnGLMasterAkun.Enabled = sBoolean
    Case mnGroupEntryDataFrm
        .mnGroupEntryDataFrm.Enabled = sBoolean
    Case mnBeliFrm
        .mnBeliFrm.Enabled = sBoolean
    Case mnMasukLainFrm
        .mnMasukLainFrm.Enabled = sBoolean
    Case mnTransferFrm
        .mnTransferFrm.Enabled = sBoolean
    Case mnKeluarLainFrm
        .mnKeluarLainFrm.Enabled = sBoolean
    Case mnJualFrm
        .mnJualFrm.Enabled = sBoolean
    Case mnMonitoringTukarFaktur
        .mnMonitoringTukarFaktur.Enabled = sBoolean
    Case mnKwitansiFrm
        .mnKwitansiFrm.Enabled = sBoolean
    Case mnGLTransGL
        .mnGLTransGL.Enabled = sBoolean
    Case mnLhppFromTF
        .mnLhppFromTF.Enabled = sBoolean
    Case mnLhppForm
        .mnLhppForm.Enabled = sBoolean
    Case mnLhppEntryForm
        .mnLhppEntryForm.Enabled = sBoolean
    Case mnUserRpt
        .mnUserRpt.Enabled = sBoolean
    Case mnBeliRpt
        .mnBeliRpt.Enabled = sBoolean
    Case mnTransferRpt
        .mnTransferRpt.Enabled = sBoolean
    Case mnKeluarLainRpt
        .mnKeluarLainRpt.Enabled = sBoolean
    Case mnJualRpt
        .mnJualRpt.Enabled = sBoolean
    Case mnJualFeeCustomerRpt
        .mnJualFeeCustomerRpt.Enabled = sBoolean
    Case mnJualMonitoringFakturRpt
        .mnJualMonitoringFakturRpt.Enabled = sBoolean
    Case mnJualRptByProduk
        .mnJualRptByProduk.Enabled = sBoolean
    Case mnKwitansiRpt
        .mnKwitansiRpt.Enabled = sBoolean
    Case mnTransaksiByProdukRpt
        .mnTransaksiByProdukRpt.Enabled = sBoolean
    Case mnGLTransGLRpt
        .mnGLTransGLRpt.Enabled = sBoolean
    Case mnGLTransGLRkpRpt
        .mnGLTransGLRkpRpt.Enabled = sBoolean
    Case mnJualRptByKomisi
        .mnJualRptByKomisi.Enabled = sBoolean
    Case mnLabaRugiByCustomerByProductRpt
        .mnLabaRugiByCustomerByProductRpt.Enabled = sBoolean
    Case mnAggingDetailRpt
        .mnAggingDetailRpt.Enabled = sBoolean
    Case mnMonitoringBayarFakturRpt
        .mnMonitoringBayarFakturRpt.Enabled = sBoolean
    Case mnMonitoringLHPPEntryRpt
        .mnMonitoringLHPPEntryRpt.Enabled = sBoolean
    Case mnLHPPTFReport
        .mnLHPPTFReport.Enabled = sBoolean
    Case mnLHPPReport
        .mnLHPPReport.Enabled = sBoolean
    Case mnKartuStockRpt
        .mnKartuStockRpt.Enabled = sBoolean
    Case mnMutasiStockRpt
        .mnMutasiStockRpt.Enabled = sBoolean
    Case mnProsesBulanan
        .mnProsesBulanan.Enabled = sBoolean
    Case mnBrandRpt
        .mnBrandRpt.Enabled = sBoolean
    Case mnCategoryRpt
        .mnCategoryRpt.Enabled = sBoolean
    Case mnFungsiRpt
        .mnFungsiRpt.Enabled = sBoolean
    Case mnDiskonRpt
        .mnDiskonRpt.Enabled = sBoolean
    Case mnHargaRpt
        .mnHargaRpt.Enabled = sBoolean
    Case mnSatuanProdukRpt
        .mnSatuanProdukRpt.Enabled = sBoolean
    Case mnTipeGudangRpt
        .mnTipeGudangRpt.Enabled = sBoolean
    Case mnGudang_Rpt
        .mnGudang_Rpt.Enabled = sBoolean
    Case mnProdukRpt
        .mnProdukRpt.Enabled = sBoolean
    Case mnSalesmanRpt
        .mnSalesmanRpt.Enabled = sBoolean
    Case mnKolektorRpt
        .mnKolektorRpt.Enabled = sBoolean
    Case mnMaster_Supplier_Rpt
        .mnMaster_Supplier_Rpt.Enabled = sBoolean
    Case mnMaster_Customer_Rpt
        .mnMaster_Customer_Rpt.Enabled = sBoolean
    Case mnGLMasterAkunRpt
        .mnGLMasterAkunRpt.Enabled = sBoolean
    Case mnGLMasterGroupAkunRpt
        .mnGLMasterGroupAkunRpt.Enabled = sBoolean
    Case mnLoginFrm
        .mnLoginFrm.Enabled = sBoolean
    Case mnMasukLainRpt
        .mnMasukLainRpt.Enabled = sBoolean
    Case mnAreaFrm
        .mnAreaFrm.Enabled = sBoolean
    Case mnAreaRpt
        .mnAreaRpt.Enabled = sBoolean
    Case mnBeliRptByProduk
        .mnBeliRptByProduk.Enabled = sBoolean
    Case mnProsesExportFiles
        .mnProsesExportFiles.Enabled = sBoolean
    Case mnProsesImportFiles
        .mnProsesImportFiles.Enabled = sBoolean
    Case mnSettingLockSalesFrm
        .mnSettingLockSalesFrm.Enabled = sBoolean
    Case mnJualFeeSalesPurchaseCustomerRpt
        .mnJualFeeSalesPurchaseCustomerRpt.Enabled = sBoolean
    Case mnMonitoringPelunasanFaktur
        .mnMonitoringPelunasanFaktur.Enabled = sBoolean
    Case mnMonitoringPelunasanFakturRpt
        .mnMonitoringPelunasanFakturRpt.Enabled = sBoolean

    End Select
    End With
End Sub
Public Sub oGetComboBoxTahun(fcb As FlatComboBox)
Dim i As Integer
Dim smulai As Integer
smulai = 2011
For i = IIf(smulai > Year(Now), Year(Now), smulai) To Year(Now)
    fcb.AddItem i
Next
fcb.text = Year(Now())
End Sub


Public Sub oGetComboBulanan(fcb As FlatComboBox)
On Error GoTo errhandler
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
Dim i As Integer

        If oConku.State = 1 Then oConku.Close
        oConku.Open MainModule.Conectionku(DBaseConection.Modul)
        sQuery = "select namabulan from master_bulan"
        Set oRstku = oConku.Execute(sQuery) 'master_moduleaccess
        With oRstku
        Do While Not .EOF
            fcb.AddItem .Fields(0)
        .MoveNext
        Loop
        fcb.ListIndex = Month(Date) - 1
        End With
        oConku.Close
        
        istatus = Normal
        Exit Sub

errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Sub
Public Sub oUpdateKodeGudang(skodegudang As String, ogrid As VSFlexGrid)
Dim irow As Integer
With ogrid
    For irow = 1 To .Rows - 1
        .TextMatrix(irow, 8) = skodegudang
    Next
End With
End Sub

Public Sub oUpdateKodeGudangKolom(skodegudang As String, sKolom As Integer, ogrid As VSFlexGrid)
Dim irow As Integer
With ogrid
    For irow = 1 To .Rows - 1
        .TextMatrix(irow, sKolom) = skodegudang
    Next
End With
End Sub

Public Sub WriteData(sFileBackupSQL As String)
Dim a As Double
Dim sQuery As String
Dim sFileBat As String
Dim s As New FileSystemObject
Dim s1 As TextStream
'If Dir("C:\BackupDatabase\backupdatabase.bat", vbReadOnly) = "" Then
'           MkDir "C:\BackupDatabase\"
'           FileCopy App.Path & "\BackupDatabase\backupdatabase.bat", "C:\BackupDatabase\backupdatabase.bat"
'End If

If Dir("C:\BackupDatabase\backupdatabase.bat", vbReadOnly) = "" Then
           MkDir "C:\BackupDatabase\"
           FileCopy App.Path & "\BackupDatabase\backupdatabase.bat", "C:\BackupDatabase\backupdatabase.bat"
End If
sFileBat = "C:\BackupDatabase\backupdatabase.bat"
sQuery = "C:\xampp\mysql\bin\mysqldump -P " & MenuFrm.Portku & " -h " & MenuFrm.Serverku & " -usa -pspvsql --routines " & MenuFrm.sCompnyDatabase & "> "   ' mysqldump -P 3306 -h 10.11.3.141 -usa -pspvsql test > db_backup.sql
' sQuery = "C:\xampp\mysql\bin\mysqldump -uroot --routines " & MenuFrm.sCompnyDatabase & "> "
'sTxtTujuan = App.Path & "\" & InputBox("Ketik Nama File Export !!", "Cuman Ngingetin !!") & ".SQL"
s.CreateTextFile sFileBat
s.OpenTextFile sFileBat, ForWriting, True
Set s1 = s.OpenTextFile(sFileBat, ForWriting, True)
        s1.WriteLine sQuery & sFileBackupSQL
s1.Close
a = Shell("C:\BackupDatabase\backupdatabase.bat", vbHide)

End Sub
Public Sub oGetbackupDatabase()
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
Dim sbackup_login_aplikasi As String
Dim sbackup_exit_aplikasi As String
Dim sawal As String
Dim sQuery As String
Dim sfile_backup As String
Dim sTujuanBackup As String
Dim sTempFileBat As String
Dim i As Integer
Dim a As Double
Dim s As New FileSystemObject
Dim s1 As TextStream
'sTxtTujuan = App.Path & "\" & InputBox("Ketik Nama File Export !!", "Cuman Ngingetin !!") & ".SQL"


        If oConku.State = 1 Then oConku.Close
        oConku.Open MainModule.Conectionku(DBaseConection.Modul)
        sQuery = "select backup_login_aplikasi,backup_exit_aplikasi,awal from setting_backup_database"
        Set oRstku = oConku.Execute(sQuery)
        sbackup_login_aplikasi = oRstku(0)
        sbackup_exit_aplikasi = oRstku(1)
        sawal = oRstku(2)

        sTempFileBat = "C:\TempFileBat\"
        If Dir(sTempFileBat, vbDirectory) = "" Then
            MkDir sTempFileBat
        End If
        
        sTujuanBackup = App.Path & "\BackupDatabase_" & MenuFrm.sCompnyDatabase & "_Temp\"
        If Dir(sTujuanBackup, vbDirectory) = "" Then
           MkDir sTujuanBackup
           ' FileCopy App.Path & "\BackupDatabase\backupdatabase.bat", "C:\BackupDatabase\backupdatabase.bat"
        End If
        
        sQuery = MenuFrm.SQLBinLocation & "mysqldump -P " & MenuFrm.Portku & " -h " & MenuFrm.Serverku ' & " -usa -pspvsql --routines " & MenuFrm.sCompnyDatabase & "> "   ' mysqldump -P 3306 -h 10.11.3.141 -usa -pspvsql test > db_backup.sql
        sQuery = sQuery & " -usa -pspvsql --routines " & MenuFrm.sCompnyDatabase & " > " & Chr(34) & sTujuanBackup & Chr(34) & MenuFrm.sCompnyDatabase & ".SQL"
        
        sFileBat = sTempFileBat & "backupdatabase.bat"
        s.CreateTextFile sFileBat
        s.OpenTextFile sFileBat, ForWriting, True
        Set s1 = s.OpenTextFile(sFileBat, ForWriting, True)
                 s1.WriteLine sQuery & sFileBackupSQL
        s1.Close
        a = Shell(sTempFileBat & "backupdatabase.bat", vbHide)
        oConku.Close
End Sub
Public Sub oGetUpdateDatabase()
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
Dim sbackup_login_aplikasi As String
Dim sbackup_exit_aplikasi As String
Dim sawal As String
Dim sQuery As String
Dim sfile_backup As String
Dim sTujuanBackup As String
Dim sTempFileBat As String
Dim i As Integer
Dim a As Double
Dim s As New FileSystemObject
Dim s1 As TextStream

sTujuanBackup = App.Path & "\updateDatabasefile\fileupdate\"
 
        sTempFileBat = "C:\TempFileBat\"
        If Dir(sTempFileBat, vbDirectory) = "" Then
            MkDir sTempFileBat
        End If
        
        If Dir(sTujuanBackup, vbDirectory) = "" Then
           MkDir sTujuanBackup
        End If
        
        If Dir(sTujuanBackup & MenuFrm.FileUpdate, vbReadOnly) = "" Then
           FileCopy App.Path & "\updateDatabasefile\" & MenuFrm.FileUpdate, sTujuanBackup & MenuFrm.FileUpdate
        
        
     
            sQuery = MenuFrm.SQLBinLocation & "mysql -P " & MenuFrm.Portku & " -h " & MenuFrm.Serverku ' & " -usa -pspvsql --routines " & MenuFrm.sCompnyDatabase & "> "   ' mysqldump -P 3306 -h 10.11.3.141 -usa -pspvsql test > db_backup.sql
            sQuery = sQuery & " -usa -pspvsql " & MenuFrm.sCompnyDatabase & " < " & Chr(34) & sTujuanBackup & Chr(34) & MenuFrm.FileUpdate
            
            sFileBat = sTempFileBat & "backupdatabase.bat"
            s.CreateTextFile sFileBat
            s.OpenTextFile sFileBat, ForWriting, True
            Set s1 = s.OpenTextFile(sFileBat, ForWriting, True)
                     s1.WriteLine sQuery & sFileBackupSQL
            s1.Close
            a = Shell(sTempFileBat & "backupdatabase.bat", vbHide)
        
        End If
'        oConku.Close
End Sub
Public Function oCekJumlahTrx(sTable As String, sMaxIsiTable As Integer) As Boolean
On Error GoTo errhandler
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.Modul)
sQuery = "Select count(*) from " & sTable
Set oRstku = oConku.Execute(sQuery)
If sMaxIsiTable <= oRstku(0) Then
    MsgBox "Aplikasi Demo ini dibatasi Maksimal " & sMaxIsiTable & " Record ," & MenuFrm.sPicProgram & " jika berminat dengan Aplikasi ini ", vbInformation
    oCekJumlahTrx = True
Else
    oCekJumlahTrx = False
End If

oConku.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "CekJumlahTrx"
End Function

Public Function oExecute(sQueryCommand As String) As Boolean
On Error GoTo errhandler
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.Modul)
oConku.Execute (sQueryCommand)
oConku.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "oExecute"
End Function
Public Sub oInsertModulMenu(sModuleMenu As String, sDscription As String, stipemenu As typemenu, sInsert As Boolean)
On Error GoTo errhandler
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
If Not sInsert Then Exit Sub
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.login)
'sp_master_module_insert`(IN sModuleMenu  VARCHAR(50),Dscription  VARCHAR(100),stransid CHAR(1),saudituser CHAR(10))
sQuery = "call sp_master_module_insert_auto('" & sModuleMenu & "','" & sDscription & "','" & stipemenu & "','" & MenuFrm.sUserID & "')"
oConku.Execute sQuery
oConku.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "GetProductID"
End Sub
