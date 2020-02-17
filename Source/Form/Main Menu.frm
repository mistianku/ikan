VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MenuFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   Caption         =   "IkanMD Release 19.11.12"
   ClientHeight    =   7845
   ClientLeft      =   -120
   ClientTop       =   390
   ClientWidth     =   14220
   Icon            =   "Main Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   948
      TabIndex        =   11
      Top             =   1755
      Visible         =   0   'False
      Width           =   14220
      Begin VB.Label txtModul 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   9135
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   0
      ScaleHeight     =   4860
      ScaleWidth      =   135
      TabIndex        =   8
      Top             =   2250
      Width           =   135
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   14220
      TabIndex        =   6
      Top             =   420
      Width           =   14220
      Begin VB.Label txtHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   8295
      End
      Begin VB.Label txtHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   8295
      End
      Begin VB.Label txtHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   8295
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      DrawMode        =   6  'Mask Pen Not
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   14220
      TabIndex        =   3
      Top             =   7110
      Width           =   14220
      Begin VB.ComboBox cmbMenuku 
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label LblPesanku 
         BackColor       =   &H00FF8080&
         Caption         =   "LblPesanku"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   11415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7470
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Text            =   "Tanggal"
            TextSave        =   "CAPS"
            Object.Tag             =   "Tanggal"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Text            =   "Jam"
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Text            =   "Tanggal : "
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "12/11/2019"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "09:33"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16776960
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":65A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":68BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":6BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":6D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":6E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":6FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":7140
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":729A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":73F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":754E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":76A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":7802
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":795C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":7AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":7C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":7D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Menu.frx":7EC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Entry Data Baru"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Batal "
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Simpan Data"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Hapus Data"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ke Record Awal"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ke Record Sebelumnya"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ke Record Berikutnya"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ke Record Terakhir"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Export Ke File Excel"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Import dari File Excel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tutup Layar"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eksekusi Perintah"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   4080
         TabIndex        =   1
         Top             =   2760
         Width           =   2415
      End
   End
   Begin VB.Menu Sistem 
      Caption         =   "Sistem"
      Begin VB.Menu mnLoginFrm 
         Caption         =   "Login"
      End
      Begin VB.Menu grsstm1 
         Caption         =   "-"
      End
      Begin VB.Menu SKS 
         Caption         =   "Exit From System"
      End
   End
   Begin VB.Menu mnSetting 
      Caption         =   "Setting"
      Begin VB.Menu mnbar_sett_label 
         Caption         =   "Master Setting Faktur"
      End
      Begin VB.Menu mnSetting_DocumentFrm2 
         Caption         =   "Master Setting Form Kwitansi"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSetting_DocumentFrm 
         Caption         =   "Setting Penomoran Dokumen"
      End
      Begin VB.Menu mnPreferenceFrm 
         Caption         =   "Preferences"
      End
      Begin VB.Menu mnPreference_special_Frm 
         Caption         =   "Preferences Special"
      End
      Begin VB.Menu mnSettingGrs1 
         Caption         =   "-"
      End
      Begin VB.Menu mnModulFrm 
         Caption         =   "Modul Form"
      End
      Begin VB.Menu mnmriAccessModuleFrm 
         Caption         =   "User Akses Modul"
      End
      Begin VB.Menu mnSettingGrs2 
         Caption         =   "-"
      End
      Begin VB.Menu mnBackupDatabaseFrm 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu mnSettingGrs3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSettingLockSalesFrm 
         Caption         =   "Setting Lock Sales"
      End
   End
   Begin VB.Menu mnMasterData 
      Caption         =   "Master Data"
      Begin VB.Menu mnMaster 
         Caption         =   "Master"
         Begin VB.Menu mnUserFrm 
            Caption         =   "User List"
         End
         Begin VB.Menu mnUserGroup 
            Caption         =   "Group"
         End
         Begin VB.Menu mnmst_company_frm 
            Caption         =   "Company"
         End
      End
      Begin VB.Menu grsMd1 
         Caption         =   "-"
      End
      Begin VB.Menu mnProduk 
         Caption         =   "Produk"
         Begin VB.Menu mnBrandFrm 
            Caption         =   "Brand"
         End
         Begin VB.Menu mnCategoryFrm 
            Caption         =   "Kategori"
         End
         Begin VB.Menu mnFunctionFrm 
            Caption         =   "Fungsi"
         End
         Begin VB.Menu mnHargaFrm 
            Caption         =   "Harga"
         End
         Begin VB.Menu mnDiskonFrm 
            Caption         =   "Diskon"
         End
         Begin VB.Menu mnFeeFrm 
            Caption         =   "Fee Produk"
         End
         Begin VB.Menu mnSatuanProduk 
            Caption         =   "Satuan Produk"
         End
         Begin VB.Menu grsproduk1 
            Caption         =   "-"
         End
         Begin VB.Menu mnTipeGudangFrm 
            Caption         =   "Tipe Gudang"
         End
         Begin VB.Menu mnGudangFrm 
            Caption         =   "Gudang"
         End
         Begin VB.Menu grsproduk2 
            Caption         =   "-"
         End
         Begin VB.Menu mnProdukFrm 
            Caption         =   "Master Produk"
         End
      End
      Begin VB.Menu grsproduk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalesmanFrm 
         Caption         =   "Salesman"
      End
      Begin VB.Menu mnKolektorFrm 
         Caption         =   "Kolektor"
      End
      Begin VB.Menu mnSupplierFrm 
         Caption         =   "Supplier"
      End
      Begin VB.Menu grsCustmr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnCustomerFrm 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnAreaFrm 
         Caption         =   "Area"
      End
      Begin VB.Menu grsproduk4 
         Caption         =   "-"
      End
      Begin VB.Menu mnGLMasterAkun 
         Caption         =   "Master Akun"
      End
      Begin VB.Menu mnGroupEntryDataFrm 
         Caption         =   "Group Entri"
      End
   End
   Begin VB.Menu mnTransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnBeliFrm 
         Caption         =   "Pembelian"
      End
      Begin VB.Menu grsmonitoring3 
         Caption         =   "-"
      End
      Begin VB.Menu mnMasukLainFrm 
         Caption         =   "Masuk Lain-lain"
      End
      Begin VB.Menu mnTransferFrm 
         Caption         =   "Pindah Antar Gudang"
      End
      Begin VB.Menu mnKeluarLainFrm 
         Caption         =   "Keluar Lain-lain"
      End
      Begin VB.Menu grsmonitoring4 
         Caption         =   "-"
      End
      Begin VB.Menu mnJualFrm 
         Caption         =   "Penjualan"
      End
      Begin VB.Menu mnMonitoringTukarFaktur 
         Caption         =   "Monitoring Tukar Faktur"
      End
      Begin VB.Menu mnKwitansiFrm 
         Caption         =   "Tanda Terima Pembayaran (Kwitansi)"
      End
      Begin VB.Menu mnMonitoringPelunasanFaktur 
         Caption         =   "Monitoring Pelunasan Faktur"
      End
      Begin VB.Menu grsmonitoring5 
         Caption         =   "-"
      End
      Begin VB.Menu mnGLTransGL 
         Caption         =   "Entri Jurnal"
      End
      Begin VB.Menu grsmonitoring6 
         Caption         =   "-"
      End
      Begin VB.Menu mnLhppFromTF 
         Caption         =   "Penerbitan Lembar Tukar Faktur"
         Visible         =   0   'False
      End
      Begin VB.Menu mnLhppForm 
         Caption         =   "Penerbitan LHPP"
      End
      Begin VB.Menu mnLhppEntryForm 
         Caption         =   "Entri LHPP"
      End
   End
   Begin VB.Menu mnBackOffice 
      Caption         =   "Back Office"
      Visible         =   0   'False
      Begin VB.Menu mnTandaBuktiPembayaran 
         Caption         =   "Tanda Bukti Pembayaran"
      End
   End
   Begin VB.Menu mnReport 
      Caption         =   "Report"
      Begin VB.Menu mnMasterRpt2 
         Caption         =   "Master"
         Begin VB.Menu mnUserRpt 
            Caption         =   "User"
         End
         Begin VB.Menu mnSiswaRpt 
            Caption         =   "Siswa"
            Visible         =   0   'False
         End
         Begin VB.Menu grsMasterRpt1 
            Caption         =   "-"
         End
         Begin VB.Menu mnProdukRpt 
            Caption         =   "Produk"
            Begin VB.Menu mnBrandRpt 
               Caption         =   "Brand"
            End
            Begin VB.Menu mnCategoryRpt 
               Caption         =   "Kategori"
            End
            Begin VB.Menu mnFungsiRpt 
               Caption         =   "Fungsi"
            End
            Begin VB.Menu mnDiskonRpt 
               Caption         =   "Diskon"
            End
            Begin VB.Menu mnHargaRpt 
               Caption         =   "Harga"
            End
            Begin VB.Menu mnSatuanProdukRpt 
               Caption         =   "Satuan Produk"
            End
            Begin VB.Menu grsMasterRpt2 
               Caption         =   "-"
            End
            Begin VB.Menu mnTipeGudangRpt 
               Caption         =   "Tipe Gudang"
            End
            Begin VB.Menu mnGudang_Rpt 
               Caption         =   "Gudang"
            End
            Begin VB.Menu grsMasterRpt3 
               Caption         =   "-"
            End
            Begin VB.Menu mnMasterProdukRpt 
               Caption         =   "Master Produk"
            End
         End
         Begin VB.Menu grsMasterRpt4 
            Caption         =   "-"
         End
         Begin VB.Menu mnSalesmanRpt 
            Caption         =   "Salesman"
         End
         Begin VB.Menu mnKolektorRpt 
            Caption         =   "Kolektor"
         End
         Begin VB.Menu mnMaster_Supplier_Rpt 
            Caption         =   "Supplier"
         End
         Begin VB.Menu brscustomer 
            Caption         =   "-"
         End
         Begin VB.Menu mnMaster_Customer_Rpt 
            Caption         =   "Customer"
         End
         Begin VB.Menu mnAreaRpt 
            Caption         =   "Area"
         End
         Begin VB.Menu grsMasterRpt5 
            Caption         =   "-"
         End
         Begin VB.Menu mnGLMasterAkunRpt 
            Caption         =   "Akun"
         End
         Begin VB.Menu mnGLMasterGroupAkunRpt 
            Caption         =   "Group Entri"
         End
      End
      Begin VB.Menu mnTransaksiRpt2 
         Caption         =   "Transaksi"
         Begin VB.Menu mnBeliRpt 
            Caption         =   "Pembelian"
         End
         Begin VB.Menu mnBeliRptByProduk 
            Caption         =   "Pembelian Per Produk/Customer"
         End
         Begin VB.Menu grsfronoffice2 
            Caption         =   "-"
         End
         Begin VB.Menu mnMasukLainRpt 
            Caption         =   "Masuk Lain-lain"
         End
         Begin VB.Menu mnTransferRpt 
            Caption         =   "Pindah Antar Gudang"
         End
         Begin VB.Menu mnKeluarLainRpt 
            Caption         =   "Keluar Lain-lain"
         End
         Begin VB.Menu grsfronoffice3 
            Caption         =   "-"
         End
         Begin VB.Menu mnJualRpt 
            Caption         =   "Penjualan"
         End
         Begin VB.Menu mnJualFeeCustomerRpt 
            Caption         =   "Fee Customer"
         End
         Begin VB.Menu mnJualFeeSalesPurchaseCustomerRpt 
            Caption         =   "Fee Sales Purchase Customer"
         End
         Begin VB.Menu mnJualMonitoringFakturRpt 
            Caption         =   "Monitoring Tukar Faktur"
         End
         Begin VB.Menu mnJualRptByProduk 
            Caption         =   "Penjualan Per Produk/Customer"
         End
         Begin VB.Menu mnKwitansiRpt 
            Caption         =   "Tanda Terima Pembayaran"
         End
         Begin VB.Menu mnMonitoringPelunasanFakturRpt 
            Caption         =   "Monitoring Pelunasan Faktur"
         End
         Begin VB.Menu grsfronoffice4 
            Caption         =   "-"
         End
         Begin VB.Menu mnTransaksiByProdukRpt 
            Caption         =   "Transaksi Per Produk"
         End
         Begin VB.Menu grsfronoffice5 
            Caption         =   "-"
         End
         Begin VB.Menu mnEntri 
            Caption         =   "Entri Jurnal"
            Begin VB.Menu mnGLTransGLRpt 
               Caption         =   "Form"
            End
            Begin VB.Menu mnGLTransGLRkpRpt 
               Caption         =   "Report"
            End
         End
         Begin VB.Menu grsfronoffice6 
            Caption         =   "-"
         End
         Begin VB.Menu mnJualRptByKomisi 
            Caption         =   "Komisi Per Produk/Customer"
         End
         Begin VB.Menu mnLabaRugiByCustomerByProductRpt 
            Caption         =   "Laba Rugi Per Customer Per Product"
         End
         Begin VB.Menu grsfronoffice7 
            Caption         =   "-"
         End
         Begin VB.Menu mnAggingDetailRpt 
            Caption         =   "Umur Faktur"
         End
         Begin VB.Menu mnMonitoringBayarFakturRpt 
            Caption         =   "Monitoring Bayar Faktur"
         End
         Begin VB.Menu mnMonitoringLHPPEntryRpt 
            Caption         =   "Monitoring LHPP"
         End
         Begin VB.Menu grsfronoffice8 
            Caption         =   "-"
         End
         Begin VB.Menu mnLHPPTFReport 
            Caption         =   "Tukar Faktur Form"
         End
         Begin VB.Menu mnLHPPReport 
            Caption         =   "LHPP Form"
         End
      End
      Begin VB.Menu mnStok1 
         Caption         =   "Stock"
         Begin VB.Menu mnKartuStockRpt 
            Caption         =   "Kartu Stock"
         End
         Begin VB.Menu mnMutasiStockRpt 
            Caption         =   "Mutasi Stock"
         End
      End
   End
   Begin VB.Menu mnProses 
      Caption         =   "Proses"
      Begin VB.Menu mnProsesExportFiles 
         Caption         =   "Proses Export File"
      End
      Begin VB.Menu mnProsesImportFiles 
         Caption         =   "Proses Import File"
      End
      Begin VB.Menu grsProses1 
         Caption         =   "-"
      End
      Begin VB.Menu mnProsesBulanan 
         Caption         =   "Proses Bulanan"
      End
   End
   Begin VB.Menu mnViewWindows 
      Caption         =   "View Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MenuFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public itools As Integer
Public bukaForm As Integer
Dim namaFrom(10)      As String
Public strConection As String
Public sUserID As String
Public Serverku, Databaseku, Driverku, Portku, SQLBinLocation, FileUpdate As String
Public RefreshForm As String
Public isKodeGroup As String
Public isAdmin As String
Public sGroupUserID As String
Public sisIndonesianFormat As String
Public sCompnyDatabase As String
Public sCmpnyID As String
Public sCmnyName As String
Public sAddress1 As String
Public sAddress2 As String
Public sCity As String
Public sZipCode As String
Public sState As String
Public sPhone1 As String
Public sPhone2 As String
Public sFaximale As String
Public sPassSuperUser As String
Public sKodeHargaCost As String
Public sjtempo As Integer
Public sinsertmodul As Boolean
Public sis_image As Integer
Public simage_name As String
Public skodeareaDefault As String
Public sissj_sama_inv As String
Public stextprefix As String
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Public sMaxIsiTable As Integer
Public sPicProgram As String
Public sCariotomatis As Integer
Public sAplikasiDemo As Boolean

Public smodelkwitansi As String
Public slebar As Integer
Public stinggi As Integer
Public skiri As Integer
Public skanan As Integer
Public stxtpesan As String

Private Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private Sub MDIForm_Load()
sCariotomatis = 1
sAplikasiDemo = False
sMaxIsiTable = 39
sPicProgram = "Hubungi Bp. Amrih di 08128127771"
sWarnaLabel = merahmuda
sWarnaBackcolour = background
sWarnaText = hijaubiasa
sWarnaTextFrame = kuningnyala
sWarnaTextOption = birutua
sWarnaTextCheck = birumuda
bukaForm = 0
strConection = "driver={sql server};server=mtgmis22;initial catalog=Payroll;uid=sa;pwd=admin"
Me.WindowState = 2

Enable 0

OpenFileIni
OpenFileIniUpdate

''oGetbackupDatabase
If sAplikasiDemo Then
Me.Caption = Me.Caption & " ( Demo ) !!! "
End If
RemoveCancelMenuItem Me
'sjtempo = oFindByQuery("SELECT jtempo FROM master_preferences where id=1", DBaseConection.Modul)
'sKodeHargaCost = oFindByQuery("SELECT CostDefault FROM master_preferences where id=1", DBaseConection.Modul)
'smodelkwitansi = oFindByQuery("select modellabel from bar_sett_label_barcode", DBaseConection.Modul)
'slebar = oFindByQuery("select lebarmm from bar_sett_label_barcode", DBaseConection.Modul)
'stinggi = oFindByQuery("select tinggimm from bar_sett_label_barcode", DBaseConection.Modul)
'skiri = oFindByQuery("select kiri from bar_sett_label_barcode", DBaseConection.Modul)
'skanan = oFindByQuery("select kanan from bar_sett_label_barcode", DBaseConection.Modul)
'stxtpesan = oFindByQuery("select txtpesan tinggi from master_setting_kwitansi_form", DBaseConection.Modul)


End Sub
Private Sub MDPA_Click()
UserFrm.Show
End Sub

Private Sub mnAggingRpt_Click()
AggingDetailRpt.Show
End Sub

Private Sub mnAggingDetailRpt_Click()
AggingDetailRpt.Show
End Sub

Private Sub mnAkunRpt_Click()
GLMasterAkunRpt.Show
End Sub

Private Sub mnAreaFrm_Click()
AreaFrm.Show
End Sub

Private Sub mnAreaRpt_Click()
AreaRpt.Show
End Sub

Private Sub mnBackupDatabaseFrm_Click()
BackupDatabaseFrm.Show
End Sub

Private Sub mnBrand_Click()
BrandFrm.Show
End Sub

Private Sub mnbar_sett_label_Click()
bar_sett_label.Show
End Sub

Private Sub mnBeliFrm_Click()
BeliFrm.Show
End Sub

Private Sub mnBeliRpt_Click()
BeliRpt.Show
End Sub

Private Sub mnBeliRptByProduk_Click()
BeliRptByProduk.Show
End Sub

Private Sub mnBrandFrm_Click()
BrandFrm.Show
End Sub

Private Sub mnBrandRpt_Click()
BrandRpt.Show
End Sub

Private Sub mnCustomer_Click()
CustomerFrm.Show
End Sub

Private Sub mnCategoryFrm_Click()
CategoryFrm.Show
End Sub

Private Sub mnCustomerFrm_Click()
CustomerFrm.Show
End Sub

Private Sub mnCustomerRpt_Click()
Master_Customer_Rpt.Show
End Sub



Private Sub mnDiskon_Click()
DiskonFrm.Show
End Sub

Private Sub mnDiskonFrm_Click()
DiskonFrm.Show
End Sub

Private Sub mnDiskonRpt_Click()
DiskonRpt.Show
End Sub



Private Sub mnEntriJurnal_Click()
GLTransGL.Show
End Sub

Private Sub mnEntriJurnalRptForm_Click()
GLTransGLRpt.Show
End Sub

Private Sub mnEntriJurnalRptReport_Click()
GLTransGLRkpRpt.Show
End Sub

Private Sub mnFeeCustomerRpt_Click()
JualFeeCustomerRpt.Show
End Sub

Private Sub mnFeeProduk_Click()
FeeFrm.Show
End Sub

Private Sub mnFungsi_Click()
FunctionFrm.Show
End Sub

Private Sub mnFeeFrm_Click()
FeeFrm.Show
End Sub

Private Sub mnFeeSalesPurchaseCustomerRpt_Click()
JualFeeSalesPurchaseCustomerRpt.Show
End Sub

Private Sub mnFunctionFrm_Click()
FunctionFrm.Show
End Sub

Private Sub mnFungsiRpt_Click()
FungsiRpt.Show
End Sub


Private Sub mnGroupEntri_Click()
GroupEntryDataFrm.Show
End Sub

Private Sub mnGLMasterAkun_Click()
GLMasterAkun.Show
End Sub

Private Sub mnGLMasterAkunRpt_Click()
GLMasterAkunRpt.Show
End Sub

Private Sub mnGLTransGL_Click()
GLTransGL.Show
End Sub

Private Sub mnGLTransGLRkpRpt_Click()
GLTransGLRkpRpt.Show
End Sub

Private Sub mnGLTransGLRpt_Click()
GLTransGLRpt.Show
End Sub

Private Sub mnGroupEntriRpt_Click()
GLMasterGroupAkunRpt.Show
End Sub

Private Sub mnGroupUser_Click()
UserGroup.Show
End Sub

Private Sub mnGudang_Click()
GudangFrm.Show
End Sub

Private Sub mnGroupEntryDataFrm_Click()
GroupEntryDataFrm.Show
End Sub

Private Sub mnGudang_Rpt_Click()
Gudang_Rpt.Show
End Sub

Private Sub mnGudangFrm_Click()
GudangFrm.Show
End Sub

Private Sub mnGudangRpt_Click()
Gudang_Rpt.Show
End Sub

Private Sub mnHarga_Click()
HargaFrm.Show
End Sub

Private Sub mnHargaFrm_Click()
HargaFrm.Show
End Sub

Private Sub mnHargaRpt_Click()
HargaRpt.Show
End Sub

Private Sub mnJenisMateri_Click()
JenisFrm.Show
End Sub



Private Sub mnKartuStock_Click()
KartuStockRpt.Show
End Sub

Private Sub mnKategori_Click()
CategoryFrm.Show
End Sub

Private Sub mnJualFeeCustomerRpt_Click()
JualFeeCustomerRpt.Show
End Sub

Private Sub mnJualFeeSalesPurchaseCustomerRpt_Click()
JualFeeSalesPurchaseCustomerRpt.Show
End Sub

Private Sub mnJualFrm_Click()
JualFrm.Show
End Sub

Private Sub mnJualMonitoringFakturRpt_Click()
JualMonitoringFakturRpt.Show
End Sub

Private Sub mnJualRpt_Click()
JualRpt.Show
End Sub

Private Sub mnJualRptByKomisi_Click()
JualRptByKomisi.Show
End Sub

Private Sub mnJualRptByProduk_Click()
JualRptByProduk.Show
End Sub

Private Sub mnKartuStockRpt_Click()
KartuStockRpt.Show
End Sub

Private Sub mnKategoriRpt_Click()
CategoryRpt.Show
End Sub

Private Sub mnKeluarLainlain_Click()
KeluarLainFrm.Show
End Sub

Private Sub mnKeluarLainlainRpt_Click()
KeluarLainRpt.Show
End Sub

Private Sub mnKolektor_Click()
KolektorFrm.Show
End Sub

Private Sub mnKeluarLainFrm_Click()
KeluarLainFrm.Show
End Sub

Private Sub mnKeluarLainRpt_Click()
KeluarLainRpt.Show
End Sub

Private Sub mnKolektorFrm_Click()
KolektorFrm.Show
End Sub

Private Sub mnKolektorRpt_Click()
KolektorRpt.Show
End Sub

Private Sub mnKomisiPerProduk_Click()
JualRptByKomisi.Show
End Sub

Private Sub mnKwitansiFrm_Click()
KwitansiFrm.Show
End Sub

Private Sub mnKwitansiRpt_Click()
KwitansiRpt.Show
End Sub

Private Sub mnLabaRugiByCustomerByProductRpt_Click()
LabaRugiByCustomerByProductRpt.Show
End Sub

Private Sub mnLHPP_Click()
LhppForm.Show
End Sub

Private Sub mnLHPPEntry_Click()
LhppEntryForm.Show
End Sub

Private Sub mnLhppEntryForm_Click()
LhppEntryForm.Show
End Sub

Private Sub mnLhppForm_Click()
LhppForm.Show
End Sub

Private Sub mnLhppFromTF_Click()
LhppFromTF.Show
End Sub

Private Sub mnLHPPRpt_Click()
LHPPReport.Show
End Sub

Private Sub mnMasterAkun_Click()
GLMasterAkun.Show
End Sub

Private Sub mnMasterProduk_Click()
ProdukFrm.Show
End Sub

Private Sub mnLHPPReport_Click()
LHPPReport.Show
End Sub

Private Sub mnLHPPTFReport_Click()
LHPPTFReport.Show
End Sub

Private Sub mnLoginFrm_Click()
LoginFrm.Show
End Sub

Private Sub mnMaster_Customer_Rpt_Click()
Master_Customer_Rpt.Show
End Sub

Private Sub mnMaster_Supplier_Rpt_Click()
Master_Supplier_Rpt.Show
End Sub

Private Sub mnMasterProdukRpt_Click()
ProdukRpt.Show
End Sub

Private Sub mnMasterSettingFaktur_Click()
bar_sett_label.Show
End Sub

Private Sub mnMasterSettingFormKwitansi_Click()
master_setting_kwitansi_form.Show
End Sub


Private Sub mnMasukLainlain_Click()
MasukLainFrm.Show
End Sub

Private Sub mnMasukLainFrm_Click()
MasukLainFrm.Show
End Sub

Private Sub mnMasukLainlainRpt_Click()
MasukLainRpt.Show
End Sub



Private Sub mnMasukLainRpt_Click()
MasukLainRpt.Show
End Sub

Private Sub mnModulForm_Click()
ModulFrm.Show
End Sub

Private Sub mnModulFrm_Click()
ModulFrm.Show
End Sub

Private Sub mnMonitoringBayarFakturRpt_Click()
MonitoringBayarFakturRpt.Show
End Sub

Private Sub mnMonitoringLHPPRpt_Click()
MonitoringLHPPEntryRpt.Show
End Sub

Private Sub mnMonitoringLHPPEntryRpt_Click()
MonitoringLHPPEntryRpt.Show
End Sub

Private Sub mnMonitoringPelunasanFakturFrm_Click()
MonitoringPelunasanFaktur.Show
End Sub

Private Sub mnMonitoringPelunasanFaktur_Click()
MonitoringPelunasanFaktur.Show
End Sub

Private Sub mnMonitoringPelunasanFakturRpt_Click()
MonitoringPelunasanFakturRpt.Show
End Sub

Private Sub mnMonitoringTukarFaktur_Click()
MonitoringTukarFaktur.Show
End Sub

Private Sub mnMonitoringTukarFakturRpt_Click()
JualMonitoringFakturRpt.Show
End Sub

Private Sub mnMutasiStock_Click()
MutasiStockRpt.Show
End Sub

Private Sub mnPembelian_Click()
BeliFrm.Show
End Sub

Private Sub mnPembelianRpt_Click()
BeliRpt.Show
End Sub

Private Sub mnPenjualan_Click()
JualFrm.Show
End Sub

Private Sub mnPenjualanPerProduk_Click()
JualRptByProduk.Show
End Sub

Private Sub mnPenjualanRpt_Click()
JualRpt.Show
End Sub


Private Sub mnPindahAntarGudang_Click()
TransferFrm.Show
End Sub

Private Sub mnPindahAntarGudangRpt_Click()
TransferRpt.Show
End Sub


Private Sub mnPreferences_Click()
PreferenceFrm.Show
End Sub

Private Sub mnPreferencesSpecial_Click()
Preference_special_Frm.Show
End Sub

Private Sub mnmriAccessModuleFrm_Click()
mriAccessModuleFrm.Show
End Sub

Private Sub mnmst_company_frm_Click()
mst_company_frm.Show
End Sub

Private Sub mnMutasiStockRpt_Click()
MutasiStockRpt.Show
End Sub

Private Sub mnPreference_special_Frm_Click()
Preference_special_Frm.Show
End Sub

Private Sub mnPreferenceFrm_Click()
PreferenceFrm.Show
End Sub

Private Sub mnProdukFrm_Click()
ProdukFrm.Show
End Sub

Private Sub mnProseExportFile_Click()
prosesExportFiles.Show
End Sub

Private Sub mnProsesBulanan_Click()
ProsesBulanan.Show
End Sub


Private Sub mnSalesman_Click()
SalesmanFrm.Show
End Sub

Private Sub mnProsesExportFiles_Click()
prosesExportFiles.Show
End Sub

Private Sub mnProsesImportFiles_Click()
ProsesImportFiles.Show
End Sub

Private Sub mnSalesmanFrm_Click()
SalesmanFrm.Show
End Sub

Private Sub mnSalesmanRpt_Click()
SalesmanRpt.Show
End Sub

Private Sub mnSatuanProduk_Click()
SatuanProduk.Show
End Sub

Private Sub mnSatuanProdukRpt_Click()
SatuanProdukRpt.Show
End Sub

Private Sub mnSettingPenomoranDokumen_Click()
Setting_DocumentFrm.Show
End Sub



Private Sub mnSupplier_Click()
SupplierFrm.Show
End Sub

Private Sub mnSetting_DocumentFrm_Click()
Setting_DocumentFrm.Show
End Sub

Private Sub mnSettingLockSalesFrm_Click()
SettingLockSalesFrm.Show
End Sub

Private Sub mnSupplierFrm_Click()
SupplierFrm.Show
End Sub

Private Sub mnSupplierRpt_Click()
Master_Supplier_Rpt.Show
End Sub

Private Sub mnTandaTerimaPembayaran_Click()
KwitansiFrm.Show
End Sub

Private Sub mnTandaTerimaPembayaranRpt_Click()
KwitansiRpt.Show
End Sub

Private Sub mnTipeGudang_Click()
TipeGudangFrm.Show
End Sub

Private Sub mnTipeGudangFrm_Click()
TipeGudangFrm.Show
End Sub

Private Sub mnTipeGudangRpt_Click()
TipeGudangRpt.Show
End Sub

Private Sub mnTransaksiPerProduk_Click()
TransaksiByProdukRpt.Show
End Sub

Private Sub mnTukarFakturFormRpt_Click()
LHPPTFReport.Show
End Sub

Private Sub mnUserAksesModul_Click()
mriAccessModuleFrm.Show
End Sub

Private Sub mnUserList_Click()
UserFrm.Show
End Sub

Private Sub mnTransaksiByProdukRpt_Click()
TransaksiByProdukRpt.Show
End Sub

Private Sub mnTransferFrm_Click()
TransferFrm.Show
End Sub

Private Sub mnTransferRpt_Click()
TransferRpt.Show
End Sub

Private Sub mnUserFrm_Click()
UserFrm.Show
End Sub

Private Sub mnUserGroup_Click()
UserGroup.Show
End Sub

Private Sub mnUserRpt_Click()
UserRpt.Show
End Sub

Private Sub mst_company_frmmn_Click()
mst_company_frm.Show
End Sub

Private Sub ProsesImportFileFrm_Click()
ProsesImportFiles.Show
End Sub

Private Sub SKS_Click()
    If oFindByQuery("SELECT backup_exit_aplikasi FROM setting_backup_database where id=1", DBaseConection.Modul) = "Y" Then
        oGetbackupDatabase
    End If
End
End Sub



Public Sub toolsEnableTrue()
Dim iBottom As Integer
For iBottom = 1 To Toolbar1.Buttons.Count
    Toolbar1.Buttons(iBottom).Enabled = True
Next
End Sub

Public Sub toolsEnableFalse()
Dim iBottom As Integer
For iBottom = 1 To Toolbar1.Buttons.Count
    Toolbar1.Buttons(iBottom).Enabled = False
Next
End Sub

Public Sub SetToolbar(istatus As StatusForm)
Dim i As Integer
Select Case istatus
Case StatusForm.NormalPlusExec
    
    toolsEnableFalse
    For i = 1 To 8
        Toolbar1.Buttons(i).Enabled = True
    Next
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(11).Enabled = True
    Toolbar1.Buttons(12).Enabled = True

    ShowFormMessage Normalmsg
    
Case StatusForm.Normal
    
    toolsEnableFalse
    For i = 1 To 8
        Toolbar1.Buttons(i).Enabled = True
    Next
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(11).Enabled = True

    ShowFormMessage Normalmsg
Case StatusForm.DataBaru
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(11).Enabled = True

    ShowFormMessage dataBarumsg
Case StatusForm.SaveMati
    Toolbar1.Buttons(3).Enabled = False
Case StatusForm.MainMenu
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
    ShowFormMessage MainMenumsg
Case StatusForm.ActvClose
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
    Toolbar1.Buttons(11).Enabled = True
    ShowFormMessage MainMenumsg
Case StatusForm.RefrshRpt
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
    Toolbar1.Buttons(11).Enabled = True
    Toolbar1.Buttons(12).Enabled = True
    ShowFormMessage MainMenumsg
Case StatusForm.MultiItem
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(8).Enabled = True
        Toolbar1.Buttons(11).Enabled = True
Case StatusForm.StatusClose
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
        Select Case i
        Case 1, 5, 6, 7, 8, 11
            Toolbar1.Buttons(i).Enabled = True
        End Select
        
        
    Next
Case StatusForm.SettingForm
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(11).Enabled = True
        

End Select

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
NewData
Case 2
UndoData
Case 3
SaveData
Case 4
DeleteData
Case 5
MoveFirst
Case 6
MovePrevious
Case 7
MoveNext
Case 8
MoveLast
Case 9
ExportToExecl
Case 10
ImportFromExecl
Case 11
Closeform
Case 12
Execution
End Select
End Sub

Public Sub MoveFirst()
Me.ActiveForm.MoveFirst
End Sub
Public Sub MovePrevious()
Me.ActiveForm.MovePrevious
End Sub
Public Sub MoveNext()
Me.ActiveForm.MoveNext
End Sub
Public Sub MoveLast()
Me.ActiveForm.MoveLast
End Sub
Public Sub Closeform()
Me.ActiveForm.Closeform
MenuFrm.Picture3.Visible = True
MenuFrm.Picture4.Visible = False
MenuFrm.Toolbar1.Visible = False
End Sub
Public Sub DeleteData()
Me.ActiveForm.DeleteData
End Sub
Public Sub SaveData()
Me.ActiveForm.SaveData
End Sub
Public Sub NewData()
Me.ActiveForm.NewData
End Sub
Public Sub UndoData()
Me.ActiveForm.Undo
End Sub
Public Sub Bayar()
Me.ActiveForm.Bayar
End Sub
Public Sub SavePrint()
Me.ActiveForm.SavePrint
End Sub

Public Sub Execution()
Me.ActiveForm.Execution
End Sub
Public Sub OpenFileIni()
Dim a, b, c, d, e, f As String
Dim jRec As Integer
Dim sPos1 As Integer
Dim sPos2 As Integer
Dim sPos3 As Integer
Dim sPos4 As Integer
Dim sPos5 As Integer
'sPos1 = InStr(1, "database=", "=")
Open App.Path & "\ServerKu.Ini" For Input As #1
Input #1, a, b, c, d, e, f
sPos1 = InStr(1, b, "=")
sPos2 = InStr(1, c, "=")
sPos3 = InStr(1, d, "=")
sPos4 = InStr(1, e, "=")
sPos5 = InStr(1, f, "=")

Databaseku = Mid(b, sPos1 + 1, Len(b) - sPos1)
Serverku = Mid(c, sPos2 + 1, Len(c) - sPos2)
Driverku = Mid(d, sPos3 + 1, Len(d) - sPos3)
Portku = Mid(e, sPos4 + 1, Len(e) - sPos4)
SQLBinLocation = Mid(f, sPos5 + 1, Len(f) - sPos5)
Close #1
End Sub
Public Sub OpenFileIniUpdate()
Dim a, b, c, d, e As String
Dim jRec As Integer
Dim sPos1 As Integer
Dim sPos2 As Integer
Dim sPos3 As Integer
Dim sPos4 As Integer
Dim sPos5 As Integer
'sPos1 = InStr(1, "database=", "=")
Open App.Path & "\updateDatabasefile\updatedatabase.Ini" For Input As #1
Input #1, a
sPos1 = InStr(1, a, "=")
FileUpdate = Mid(a, sPos1 + 1, Len(a) - sPos1)
Close #1
End Sub
Public Sub ExportToExecl()
Me.ActiveForm.ExportToExecl
End Sub

Public Sub ImportFromExecl()
Me.ActiveForm.ImportFromExecl
End Sub

Public Sub MoveNextActive()
Dim iBottom As Integer
For iBottom = 5 To 8
    Toolbar1.Buttons(iBottom).Enabled = True
Next

End Sub


Public Sub SetToolbarku(sForm As Form, istatus As StatusForm, sGroupUser As String, sModulID As Modul)
Dim i As Integer
Dim iModulid As Integer
MenuFrm.Picture3.Visible = False
MenuFrm.Picture4.Visible = True
MenuFrm.Toolbar1.Visible = True
txtModul = sForm.Caption
sForm.BackColor = &HFFC0C0
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, sForm

iModulid = sModulID
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(6).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Toolbar1.Buttons(8).Enabled = False
Select Case istatus
Case StatusForm.NormalPlusExec
    
    toolsEnableFalse
    For i = 1 To 8
        Toolbar1.Buttons(i).Enabled = True
    Next
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(11).Enabled = True
    Toolbar1.Buttons(12).Enabled = True
    FindData.oGetAccesMenuByModulID sGroupUser, iModulid
    ShowFormMessage Normalmsg
    
Case StatusForm.Normal
    
    toolsEnableFalse
    For i = 1 To 8
        
        Toolbar1.Buttons(i).Enabled = True
    Next
    If isAdmin = "N" Then
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Toolbar1.Buttons(8).Enabled = False
    End If
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(11).Enabled = True
    FindData.oGetAccesMenuByModulID MenuFrm.sUserID, iModulid
    ShowFormMessage Normalmsg
Case StatusForm.DataBaru
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(11).Enabled = True

    ShowFormMessage dataBarumsg
Case StatusForm.SaveMati
    Toolbar1.Buttons(3).Enabled = False
    FindData.oGetAccesMenuByModulID sGroupUser, iModulid
    
Case StatusForm.MainMenu
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
    ShowFormMessage MainMenumsg
Case StatusForm.ActvClose
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
    Toolbar1.Buttons(11).Enabled = True
    FindData.oGetAccesMenuByModulID sGroupUser, iModulid
    ShowFormMessage MainMenumsg
Case StatusForm.RefrshRpt
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
    Toolbar1.Buttons(11).Enabled = True
    Toolbar1.Buttons(12).Enabled = True
    ShowFormMessage MainMenumsg
    FindData.oGetAccesMenuByModulID sGroupUser, iModulid
    Toolbar1.Buttons(12).Enabled = True
Case StatusForm.MultiItem
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(8).Enabled = True
        Toolbar1.Buttons(11).Enabled = True
        FindData.oGetAccesMenuByModulID sGroupUser, iModulid
Case StatusForm.StatusClose
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
        Select Case i
        Case 1, 5, 6, 7, 8, 11
            Toolbar1.Buttons(i).Enabled = True
        End Select
        FindData.oGetAccesMenuByModulID sGroupUser, iModulid
        
    Next
Case StatusForm.SettingForm
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = False
    Next
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(11).Enabled = True
        FindData.oGetAccesMenuByModulID sGroupUser, iModulid
        
Case StatusForm.NormalClosePlusExec
    
    toolsEnableFalse
    For i = 1 To 8
        Toolbar1.Buttons(i).Enabled = True
    Next
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(11).Enabled = True
    Toolbar1.Buttons(12).Enabled = True
    FindData.oGetAccesMenuByModulID sGroupUser, iModulid
    ShowFormMessage Normalmsg
    

End Select


End Sub


Public Sub oGetPreference()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select  *  from master_preferences limit 1"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
    
        MenuFrm.sCmpnyID = (oRs("CmpnyID"))
        MenuFrm.sCmnyName = (oRs("CmnyName"))
        MenuFrm.sAddress1 = (oRs("Address1"))
        MenuFrm.sAddress2 = (oRs("Address2"))
        MenuFrm.sCity = (oRs("City"))
        MenuFrm.sZipCode = (oRs("ZipCode"))
        MenuFrm.sState = (oRs("State"))
        MenuFrm.sPhone1 = (oRs("Phone1"))
        MenuFrm.sPhone2 = (oRs("Phone2"))
        MenuFrm.sFaximale = (oRs("Faximale"))
        MenuFrm.sisIndonesianFormat = (oRs("isIndonesianFormat"))
        MenuFrm.txtHeader(0) = sCmnyName
        MenuFrm.txtHeader(1) = sAddress1 & IIf(sAddress2 = "", "", "," & sAddress2)
        MenuFrm.txtHeader(2) = sCity & IIf(sPhone1 = "", "", " Telp : " & sPhone1) & IIf(sPhone2 = "", "", "," & sPhone2)
        MenuFrm.sinsertmodul = IIf(oRs("insertmodul") = "Y", True, False)
        MenuFrm.sis_image = (oRs("is_image"))
        MenuFrm.simage_name = (oRs("image_name"))
        MenuFrm.skodeareaDefault = (oRs("kodeareaDefault"))
        MenuFrm.sissj_sama_inv = (oRs("issj_sama_inv"))
        
    End If
    oCon.Close
    
    MenuFrm.stextprefix = oFindByQuery("select textprefix from setting_document where docid=8", DBaseConection.Modul)
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveFirst"
End Sub



Public Sub RemoveCancelMenuItem(frm As Form)
Dim hSysMenu As Long
'Ambil menu system untuk form ini
hSysMenu = GetSystemMenu(frm.hWnd, 0)
'Hilangkan tombol Close (X)
Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
'Hilangkan pemisah yang melalui tombol Close tsb
Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

