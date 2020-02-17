/*
SQLyog Ultimate v10.00 Beta1
MySQL - 5.1.41 : Database - common_ikan
*********************************************************************
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;
/*Table structure for table `company` */

DROP TABLE IF EXISTS `company`;

CREATE TABLE `company` (
  `dbid` int(11) NOT NULL AUTO_INCREMENT,
  `cmpnyid` varchar(20) NOT NULL DEFAULT '',
  `cmpnyname` varchar(50) DEFAULT NULL,
  `databasename` varchar(10) DEFAULT NULL,
  `auditdate` datetime DEFAULT NULL,
  `audituser` varchar(8) DEFAULT NULL,
  UNIQUE KEY `NewIndex1` (`cmpnyid`),
  UNIQUE KEY `dbid` (`dbid`),
  UNIQUE KEY `NewIndex2` (`databasename`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

/*Data for the table `company` */

insert  into `company`(`dbid`,`cmpnyid`,`cmpnyname`,`databasename`,`auditdate`,`audituser`) values (2,'JKT','Cab Jakarta','sayur','2016-08-14 01:53:02','admin'),(1,'SBY','Cab Surabaya','ikan','2016-08-13 15:41:21','admin');

/*Table structure for table `companyaccess` */

DROP TABLE IF EXISTS `companyaccess`;

CREATE TABLE `companyaccess` (
  `emplcode` char(8) NOT NULL,
  `dbid` int(11) NOT NULL,
  `accessctrl` char(1) DEFAULT NULL,
  `auditdate` datetime DEFAULT NULL,
  `audituser` char(8) DEFAULT NULL,
  PRIMARY KEY (`emplcode`,`dbid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `companyaccess` */

insert  into `companyaccess`(`emplcode`,`dbid`,`accessctrl`,`auditdate`,`audituser`) values ('007',1,'Y','2012-10-18 15:16:41',''),('008',1,'N','2013-01-05 18:13:19','admin'),('abay',1,'N','2012-10-13 23:30:31','Admin'),('acep',1,'N','2012-10-13 23:30:31','Admin'),('ade',1,'N','2012-10-13 23:30:31','Admin'),('admin',1,'Y','2012-10-13 23:30:31','Admin'),('admin',2,'Y','2016-08-13 15:41:21','admin'),('Masri',1,'N','2012-10-18 15:18:06',''),('User1',1,'Y','2013-01-05 21:07:22','admin'),('User1',2,'N','2016-08-13 15:41:21','admin'),('User2',1,'Y','2013-01-05 21:07:49','admin'),('User2',2,'N','2016-08-13 15:41:21','admin');

/*Table structure for table `master_group_user` */

DROP TABLE IF EXISTS `master_group_user`;

CREATE TABLE `master_group_user` (
  `kodegroup` varchar(10) NOT NULL,
  `groupuser` varchar(30) DEFAULT NULL,
  `objtype` smallint(6) DEFAULT NULL,
  `audituser` varchar(10) DEFAULT NULL,
  `auditdate` datetime DEFAULT NULL,
  PRIMARY KEY (`kodegroup`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `master_group_user` */

insert  into `master_group_user`(`kodegroup`,`groupuser`,`objtype`,`audituser`,`auditdate`) values ('01','Office',NULL,'admin','0000-00-00 00:00:00'),('88','Admin',NULL,'admin','0000-00-00 00:00:00'),('02','Purchasing',NULL,'admin','0000-00-00 00:00:00'),('04','Sales',NULL,'admin','0000-00-00 00:00:00'),('03','Inventorys',NULL,'admin','0000-00-00 00:00:00');

/*Table structure for table `master_module` */

DROP TABLE IF EXISTS `master_module`;

CREATE TABLE `master_module` (
  `Modulid` int(11) NOT NULL AUTO_INCREMENT,
  `ModuleMenu` varchar(50) DEFAULT NULL,
  `Dscription` varchar(100) DEFAULT NULL,
  `visorder` int(10) DEFAULT NULL,
  `transid` char(1) DEFAULT NULL COMMENT '1 Setting 2 Transaksi 3 report',
  `MenuIndex` int(11) DEFAULT '0' COMMENT 'Index Modul Menu Bar',
  `Auditdate` timestamp NULL DEFAULT NULL,
  `AuditUser` varchar(10) DEFAULT NULL,
  PRIMARY KEY (`Modulid`)
) ENGINE=MyISAM AUTO_INCREMENT=79 DEFAULT CHARSET=utf8;

/*Data for the table `master_module` */

insert  into `master_module`(`Modulid`,`ModuleMenu`,`Dscription`,`visorder`,`transid`,`MenuIndex`,`Auditdate`,`AuditUser`) values (1,'bar_sett_label','Master Setting Label Form',10,'1',0,'2016-08-14 05:58:46','admin'),(2,'Setting_DocumentFrm','Master Setting Dokumen Form',20,'1',0,'2016-08-14 06:00:58','admin'),(3,'PreferenceFrm','Master Prefernce Form',30,'1',0,'2016-08-14 06:01:01','admin'),(4,'Preference_special_Frm','Master Prefernce Spesial Form',40,'2',0,'2016-08-14 06:01:04','admin'),(5,'ModulFrm','Master Modul Aplikasi Form',50,'2',0,'2016-08-14 06:01:08','admin'),(6,'mriAccessModuleFrm','Master Group Akses Modul Form',60,'1',0,'2016-08-14 06:01:12','admin'),(7,'BackupDatabaseFrm','Backup Database Form',70,'1',0,'2016-08-14 06:01:15','admin'),(8,'UserFrm','Master Pengguna Form',80,'2',0,'2016-08-14 06:01:24','admin'),(9,'UserGroup','Master Group User Form',90,'2',0,'2016-08-14 06:01:28','admin'),(10,'mst_company_frm','Master Company Form',100,'2',0,'2016-08-14 06:01:31','admin'),(11,'BrandFrm','Master Brand Form',110,'2',0,'2016-08-14 06:01:34','admin'),(12,'CategoryFrm','Master Kategori Produk Form',120,'2',0,'2016-08-14 06:01:54','admin'),(13,'FunctionFrm','Master Fungsi Produk Form',130,'2',0,'2016-08-14 06:01:58','admin'),(14,'HargaFrm','Master Harga Produk Form',140,'2',0,'2016-08-14 06:02:01','admin'),(15,'DiskonFrm','Master Diskon Produk Form',150,'2',0,'2016-08-14 06:02:07','admin'),(16,'FeeFrm','Master Fee Produk Form',160,'2',0,'2016-08-14 06:02:11','admin'),(17,'SatuanProduk','Master Satuan Produk Form',170,'2',0,'2016-08-14 06:02:15','admin'),(18,'TipeGudangFrm','Master Tipe Gudang Form',180,'2',0,'2016-08-14 06:02:20','admin'),(19,'GudangFrm','Master Gudang Form',190,'2',0,'2016-08-14 06:02:24','admin'),(20,'ProdukFrm','Master Produk Form',200,'2',0,'2016-08-14 06:02:29','admin'),(21,'SalesmanFrm','Master Salesman Form',210,'2',0,'2016-08-14 06:28:59','admin'),(22,'KolektorFrm','Master Kolektor Form',220,'2',0,'2016-08-14 06:29:04','admin'),(23,'SupplierFrm','Master Supplier Form',230,'2',0,'2016-08-14 06:29:09','admin'),(24,'CustomerFrm','Master Customer Form',240,'2',0,'2016-08-14 06:29:13','admin'),(25,'GLMasterAkun','Master Kode Akun Form',250,'2',0,'2016-08-14 06:29:17','admin'),(26,'GroupEntryDataFrm','Master Group Entry Data Form',260,'2',0,'2016-08-14 06:29:20','admin'),(27,'BeliFrm','Transaksi Pembelian Form',270,'2',0,'2016-08-14 06:29:47','admin'),(28,'MasukLainFrm','Transaksi Masuk Lain-Lain Form',280,'2',0,'2016-08-14 06:29:51','admin'),(29,'TransferFrm','Transaksi Pindah Antar Gudang Form',290,'2',0,'2016-08-14 06:29:54','admin'),(30,'KeluarLainFrm','Transaksi Keluar Lain Lain Form',300,'2',0,'2016-08-14 06:29:57','admin'),(31,'JualFrm','Transaksi Penjualan Form',310,'2',0,'2016-08-14 06:30:01','admin'),(32,'MonitoringTukarFaktur','Monitoring Entry Tukar Faktur Form',320,'2',0,'2016-08-14 06:30:04','admin'),(33,'KwitansiFrm','Bukti Penerimaan Pembayaran Form',330,'2',0,'2016-08-14 06:30:12','admin'),(34,'GLTransGL','Transaksi GL Entri Jurnal Form',340,'2',0,'2016-08-14 06:30:15','admin'),(35,'LhppFromTF','Penerbitan Lembar Tukar Faktur Form',350,'2',0,'2016-08-14 06:32:55','admin'),(36,'LhppForm','Penerbitan LHPP Form',360,'2',0,'2016-08-14 06:33:08','admin'),(37,'LhppEntryForm','LHPP Entry Form',370,'2',0,'2016-08-14 06:33:18','admin'),(38,'UserRpt','Master Pengguna Report',380,'3',0,'2016-08-14 06:34:02','admin'),(39,'BeliRpt','Pembelian Report',390,'3',0,'2016-08-14 06:34:13','admin'),(40,'TransferRpt','Transaksi Pindah Antar Gudang Report',400,'3',0,'2016-08-14 06:34:20','admin'),(41,'KeluarLainRpt','Keluar Lain-lain Report',410,'3',0,'2016-08-14 06:34:24','admin'),(42,'JualRpt','Penjualan Report',420,'3',0,'2016-08-14 06:34:28','admin'),(43,'JualFeeCustomerRpt','Fee Customer Report',430,'3',0,'2016-08-14 06:34:31','admin'),(44,'JualMonitoringFakturRpt','Monitoring Tukar Faktur Form',440,'3',0,'2016-08-14 06:34:37','admin'),(45,'JualRptByProduk','Penjualan By Produk Report',450,'3',0,'2016-08-14 06:34:42','admin'),(46,'KwitansiRpt','Kwitansi Report',460,'3',0,'2016-08-14 06:34:47','admin'),(47,'TransaksiByProdukRpt','Sales By Product Report',470,'3',0,'2016-08-14 06:34:53','admin'),(48,'GLTransGLRpt','Jurnal Entri Report',480,'3',0,'2016-08-14 06:35:02','admin'),(49,'GLTransGLRkpRpt','Transaksi GL Rekap Report ',490,'3',0,'2016-08-14 06:35:08','admin'),(50,'JualRptByKomisi','Fee Penjualan Report',500,'3',0,'2016-08-14 06:36:48','admin'),(51,'LabaRugiByCustomerByProductRpt','Laba Rugi Customer By Product Report',510,'3',0,'2016-08-14 06:37:02','admin'),(52,'AggingDetailRpt','Umur Faktur Report',520,'3',0,'2016-08-14 06:37:15','admin'),(53,'MonitoringBayarFakturRpt','Monitoring Bayar Faktur Report',530,'3',0,'2016-08-14 06:37:19','admin'),(54,'MonitoringLHPPEntryRpt','Monitoring Penerimaan Pembayaran (LHPP) Report',540,'3',0,'2016-08-14 06:37:29','admin'),(55,'LHPPTFReport','Tukar Faktur LHPP Report',550,'3',0,'2016-08-14 06:37:35','admin'),(56,'LHPPReport','LHPP Report',560,'3',0,'2016-08-14 06:37:40','admin'),(57,'KartuStockRpt','Kartu Stock Report',570,'3',0,'2016-08-14 06:38:04','admin'),(58,'MutasiStockRpt','Mutasi Stok Report',580,'3',0,'2016-08-14 06:38:07','admin'),(59,'ProsesBulanan','Proses Bulanan Form',590,'3',0,'2016-08-14 06:38:14','admin'),(60,'BrandRpt','Master Brand Produk Report',600,'3',0,'2016-08-14 07:57:15','admin'),(61,'CategoryRpt','Master Kategori Produk Report',610,'3',0,'2016-08-14 07:58:16','admin'),(62,'FungsiRpt','Master Fungsi Produk Report',620,'3',0,'2016-08-14 07:59:30','admin'),(63,'DiskonRpt','Master Diskon Produk Report',630,'3',0,'2016-08-14 07:59:55','admin'),(64,'HargaRpt','Master Harga Produk Report',640,'3',0,'2016-08-14 08:00:31','admin'),(65,'SatuanProdukRpt','Master Satuan Produk Report',650,'3',0,'2016-08-14 08:01:41','admin'),(66,'TipeGudangRpt','Master Tipe Gudang Report',660,'3',0,'2016-08-14 08:02:10','admin'),(67,'Gudang_Rpt','Master Gudang Report',670,'3',0,'2016-08-14 08:02:24','admin'),(68,'ProdukRpt','Master Produk Report',680,'3',0,'2016-08-14 08:02:36','admin'),(69,'SalesmanRpt','Master Salesman Report',690,'3',0,'2016-08-14 08:02:48','admin'),(70,'KolektorRpt','Master Kolektor Report',700,'3',0,'2016-08-14 08:03:59','admin'),(71,'Master_Supplier_Rpt','Master Supplier Report',710,'3',0,'2016-08-14 08:04:05','admin'),(72,'Master_Customer_Rpt','Master Customer Form',720,'3',0,'2016-08-14 08:04:15','admin'),(73,'GLMasterAkunRpt','Master Daftar Akun Report ',730,'3',0,'2016-08-14 08:04:36','admin'),(74,'GLMasterGroupAkunRpt','Master Group Entri Report',740,'3',0,'2016-08-14 08:04:52','admin'),(75,'LoginFrm','Otorisasi Masuk Ke Sistem',750,'3',0,'2016-08-14 10:39:39','User1'),(76,'MasukLainRpt','Masuk Lain-lain Report',760,'3',0,'2016-08-15 06:01:27','admin'),(77,'AreaFrm','Master Area Customer Form',770,'2',0,'2016-08-29 11:02:45','admin'),(78,'AreaRpt','Master Area Customer Report',780,'3',0,'2016-08-31 17:11:49','admin');

/*Table structure for table `master_moduleaccess` */

DROP TABLE IF EXISTS `master_moduleaccess`;

CREATE TABLE `master_moduleaccess` (
  `kodegroup` varchar(2) NOT NULL,
  `modulid` int(10) NOT NULL,
  `Baca` varchar(1) DEFAULT NULL,
  `Tulis` varchar(1) DEFAULT NULL,
  `Edit` varchar(1) DEFAULT NULL,
  `Hapus` varchar(1) DEFAULT NULL,
  `Cetak` varchar(1) DEFAULT NULL,
  PRIMARY KEY (`kodegroup`,`modulid`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

/*Data for the table `master_moduleaccess` */

insert  into `master_moduleaccess`(`kodegroup`,`modulid`,`Baca`,`Tulis`,`Edit`,`Hapus`,`Cetak`) values ('01',1,'Y','N','N','N','N'),('02',1,'N','N','N','N','N'),('03',76,'N','N','N','N','Y'),('88',1,'Y','N','Y','N','N'),('04',76,'N','N','N','N','N'),('01',2,'Y','N','N','N','N'),('02',2,'N','N','N','N','N'),('03',75,'N','N','N','N','N'),('88',2,'Y','N','Y','N','N'),('04',75,'N','N','N','N','N'),('01',3,'Y','N','N','N','N'),('02',3,'N','N','N','N','N'),('03',74,'N','N','N','N','N'),('88',3,'Y','N','Y','N','N'),('04',74,'N','N','N','N','N'),('01',4,'Y','N','N','N','N'),('02',4,'N','N','N','N','N'),('03',73,'N','N','N','N','N'),('88',4,'Y','Y','Y','Y','Y'),('04',73,'N','N','N','N','N'),('01',5,'Y','N','N','N','N'),('02',5,'N','N','N','N','N'),('03',72,'N','N','N','N','N'),('88',5,'Y','Y','Y','Y','Y'),('04',72,'N','N','N','N','Y'),('01',6,'Y','N','N','N','N'),('02',6,'N','N','N','N','N'),('03',71,'N','N','N','N','N'),('88',6,'Y','Y','N','Y','Y'),('04',71,'N','N','N','N','N'),('01',7,'Y','N','N','N','N'),('02',7,'N','N','N','N','N'),('03',70,'N','N','N','N','N'),('88',7,'Y','N','Y','N','N'),('04',70,'N','N','N','N','Y'),('01',8,'Y','N','Y','N','N'),('02',8,'Y','N','Y','N','N'),('03',69,'N','N','N','N','N'),('88',8,'Y','Y','Y','Y','N'),('04',69,'N','N','N','N','Y'),('01',9,'Y','N','N','N','N'),('02',9,'N','N','N','N','N'),('03',68,'N','N','N','N','Y'),('88',9,'Y','Y','Y','Y','N'),('04',68,'N','N','N','N','Y'),('01',10,'Y','N','N','N','N'),('02',10,'N','N','N','N','N'),('03',67,'N','N','N','N','Y'),('88',10,'Y','Y','Y','Y','Y'),('04',67,'N','N','N','N','Y'),('01',11,'Y','Y','Y','Y','Y'),('02',11,'Y','Y','Y','Y','Y'),('03',66,'N','N','N','N','Y'),('88',11,'Y','Y','Y','Y','Y'),('04',66,'N','N','N','N','Y'),('01',12,'Y','N','N','N','N'),('02',12,'Y','Y','Y','Y','Y'),('03',65,'N','N','N','N','Y'),('88',12,'Y','Y','Y','Y','Y'),('04',65,'N','N','N','N','Y'),('01',13,'Y','N','N','N','N'),('02',13,'Y','Y','Y','Y','Y'),('03',64,'N','N','N','N','Y'),('88',13,'Y','Y','Y','Y','Y'),('04',64,'N','N','N','N','Y'),('01',14,'Y','N','N','N','N'),('02',14,'Y','Y','Y','Y','Y'),('03',63,'N','N','N','N','Y'),('88',14,'Y','Y','Y','Y','Y'),('04',63,'N','N','N','N','Y'),('01',15,'Y','N','N','N','N'),('02',15,'Y','Y','Y','Y','Y'),('03',62,'N','N','N','N','Y'),('88',15,'Y','Y','Y','Y','Y'),('04',62,'N','N','N','N','Y'),('01',16,'Y','N','N','N','N'),('02',16,'Y','Y','Y','Y','Y'),('03',61,'N','N','N','N','Y'),('88',16,'Y','Y','Y','Y','Y'),('04',61,'N','N','N','N','Y'),('01',17,'Y','N','N','N','N'),('02',17,'Y','Y','Y','Y','Y'),('03',60,'N','N','N','N','Y'),('88',17,'Y','Y','Y','Y','Y'),('04',60,'N','N','N','N','Y'),('01',18,'Y','N','N','N','N'),('02',18,'N','N','N','N','N'),('03',59,'N','N','N','N','Y'),('88',18,'Y','Y','Y','Y','Y'),('04',59,'N','N','N','N','Y'),('01',19,'Y','N','N','N','N'),('02',19,'Y','N','N','N','Y'),('03',58,'N','N','N','N','Y'),('88',19,'Y','Y','Y','Y','Y'),('04',58,'N','N','N','N','Y'),('01',20,'Y','N','N','N','N'),('02',20,'Y','Y','Y','Y','Y'),('03',57,'N','N','N','N','Y'),('88',20,'Y','Y','Y','Y','Y'),('04',57,'N','N','N','N','Y'),('01',21,'Y','N','N','N','N'),('02',21,'N','N','N','N','N'),('03',56,'N','N','N','N','N'),('88',21,'Y','Y','Y','Y','Y'),('04',56,'N','N','N','N','Y'),('01',22,'Y','N','N','N','N'),('02',22,'N','N','N','N','N'),('03',55,'N','N','N','N','N'),('88',22,'Y','Y','Y','Y','Y'),('04',55,'N','N','N','N','Y'),('01',23,'Y','N','N','N','N'),('02',23,'Y','Y','Y','Y','Y'),('03',54,'N','N','N','N','N'),('88',23,'Y','Y','Y','Y','Y'),('04',54,'N','N','N','N','Y'),('01',24,'Y','N','N','N','N'),('02',24,'N','N','N','N','N'),('03',53,'N','N','N','N','N'),('88',24,'Y','Y','Y','Y','Y'),('04',53,'N','N','N','N','Y'),('01',25,'Y','N','N','N','N'),('02',25,'N','N','N','N','N'),('03',52,'N','N','N','N','N'),('88',25,'Y','Y','Y','Y','Y'),('04',52,'N','N','N','N','Y'),('01',26,'Y','N','N','N','N'),('02',26,'N','N','N','N','N'),('03',51,'N','N','N','N','N'),('88',26,'Y','Y','Y','Y','Y'),('04',51,'N','N','N','N','Y'),('01',27,'Y','N','N','N','N'),('02',27,'Y','Y','Y','Y','Y'),('03',50,'N','N','N','N','N'),('88',27,'Y','Y','Y','Y','Y'),('04',50,'N','N','N','N','Y'),('01',28,'Y','N','N','N','N'),('02',28,'N','N','N','N','N'),('03',49,'N','N','N','N','N'),('88',28,'Y','Y','Y','Y','Y'),('04',49,'N','N','N','N','N'),('01',29,'Y','N','N','N','N'),('02',29,'N','N','N','N','N'),('03',48,'N','N','N','N','N'),('88',29,'Y','Y','Y','Y','Y'),('04',48,'N','N','N','N','N'),('01',30,'Y','N','N','N','N'),('02',30,'N','N','N','N','N'),('03',47,'N','N','N','N','N'),('88',30,'Y','Y','Y','Y','Y'),('04',47,'N','N','N','N','Y'),('01',31,'Y','N','N','N','N'),('02',31,'N','N','N','N','N'),('03',46,'N','N','N','N','N'),('88',31,'Y','Y','Y','Y','Y'),('04',46,'N','N','N','N','Y'),('01',32,'Y','N','N','N','N'),('02',32,'N','N','N','N','N'),('03',45,'N','N','N','N','N'),('88',32,'Y','Y','Y','Y','Y'),('04',45,'N','N','N','N','Y'),('01',33,'Y','N','N','N','N'),('02',33,'N','N','N','N','N'),('03',44,'N','N','N','N','N'),('88',33,'Y','Y','Y','Y','Y'),('04',44,'N','N','N','N','Y'),('01',34,'Y','N','N','N','N'),('02',34,'N','N','N','N','N'),('03',43,'N','N','N','N','N'),('88',34,'Y','Y','Y','Y','Y'),('04',43,'N','N','N','N','Y'),('01',35,'Y','N','N','N','N'),('02',35,'N','N','N','N','N'),('03',42,'N','N','N','N','N'),('88',35,'Y','Y','Y','Y','Y'),('04',42,'N','N','N','N','Y'),('01',36,'Y','N','N','N','N'),('02',36,'N','N','N','N','N'),('03',41,'N','N','N','N','Y'),('88',36,'Y','Y','Y','Y','Y'),('04',41,'N','N','N','N','N'),('01',37,'Y','N','N','N','N'),('02',37,'N','N','N','N','N'),('03',40,'N','N','N','N','Y'),('88',37,'Y','Y','Y','Y','Y'),('04',40,'N','N','N','N','N'),('01',38,'N','N','N','N','Y'),('02',38,'N','N','N','N','N'),('03',39,'N','N','N','N','N'),('88',38,'N','N','N','N','Y'),('04',39,'N','N','N','N','N'),('01',39,'N','N','N','N','N'),('02',39,'N','N','N','N','Y'),('03',38,'N','N','N','N','N'),('88',39,'N','N','N','N','Y'),('04',38,'N','N','N','N','N'),('01',40,'N','N','N','N','N'),('02',40,'N','N','N','N','N'),('03',37,'N','N','N','N','N'),('88',40,'N','N','N','N','Y'),('04',37,'Y','Y','Y','Y','Y'),('01',41,'N','N','N','N','N'),('02',41,'N','N','N','N','N'),('03',36,'N','N','N','N','N'),('88',41,'N','N','N','N','Y'),('04',36,'Y','Y','Y','Y','Y'),('01',42,'N','N','N','N','N'),('02',42,'N','N','N','N','N'),('03',35,'N','N','N','N','N'),('88',42,'N','N','N','N','Y'),('04',35,'Y','Y','Y','Y','Y'),('01',43,'N','N','N','N','N'),('02',43,'N','N','N','N','N'),('03',34,'N','N','N','N','N'),('88',43,'N','N','N','N','Y'),('04',34,'Y','Y','Y','Y','Y'),('01',44,'N','N','N','N','N'),('02',44,'N','N','N','N','N'),('03',33,'N','N','N','N','N'),('88',44,'N','N','N','N','Y'),('04',33,'Y','Y','Y','Y','Y'),('01',45,'N','N','N','N','N'),('02',45,'N','N','N','N','N'),('03',32,'N','N','N','N','N'),('88',45,'N','N','N','N','Y'),('04',32,'Y','Y','Y','Y','Y'),('01',46,'N','N','N','N','N'),('02',46,'N','N','N','N','N'),('03',31,'N','N','N','N','N'),('88',46,'N','N','N','N','Y'),('04',31,'Y','Y','Y','Y','Y'),('01',47,'N','N','N','N','N'),('02',47,'N','N','N','N','N'),('03',30,'Y','Y','Y','Y','Y'),('88',47,'N','N','N','N','Y'),('04',30,'N','N','N','N','N'),('01',48,'N','N','N','N','N'),('02',48,'N','N','N','N','N'),('03',29,'Y','Y','Y','Y','Y'),('88',48,'N','N','N','N','Y'),('04',29,'N','N','N','N','N'),('01',49,'N','N','N','N','N'),('02',49,'N','N','N','N','N'),('03',28,'Y','Y','Y','Y','Y'),('88',49,'N','N','N','N','Y'),('04',28,'N','N','N','N','N'),('01',50,'N','N','N','N','N'),('02',50,'N','N','N','N','N'),('03',27,'N','N','N','N','N'),('88',50,'N','N','N','N','Y'),('04',27,'N','N','N','N','N'),('01',51,'N','N','N','N','N'),('02',51,'N','N','N','N','N'),('03',26,'N','N','N','N','N'),('88',51,'N','N','N','N','Y'),('04',26,'N','N','N','N','N'),('01',52,'N','N','N','N','N'),('02',52,'N','N','N','N','N'),('03',25,'N','N','N','N','N'),('88',52,'N','N','N','N','Y'),('04',25,'N','N','N','N','N'),('01',53,'N','N','N','N','N'),('02',53,'N','N','N','N','N'),('03',24,'N','N','N','N','N'),('88',53,'N','N','N','N','Y'),('04',24,'Y','Y','Y','Y','Y'),('01',54,'N','N','N','N','N'),('02',54,'N','N','N','N','N'),('03',23,'N','N','N','N','N'),('88',54,'N','N','N','N','Y'),('04',23,'Y','Y','Y','Y','Y'),('01',55,'N','N','N','N','N'),('02',55,'N','N','N','N','N'),('03',22,'N','N','N','N','N'),('88',55,'N','N','N','N','Y'),('04',22,'Y','Y','Y','Y','Y'),('01',56,'N','N','N','N','N'),('02',56,'N','N','N','N','N'),('03',21,'N','N','N','N','N'),('88',56,'N','N','N','N','Y'),('04',21,'Y','Y','Y','Y','Y'),('01',57,'N','N','N','N','N'),('02',57,'N','N','N','N','Y'),('03',20,'Y','Y','Y','Y','Y'),('88',57,'N','N','N','N','Y'),('04',20,'Y','Y','Y','Y','Y'),('01',58,'N','N','N','N','N'),('02',58,'N','N','N','N','Y'),('03',19,'Y','Y','Y','Y','Y'),('88',58,'N','N','N','N','Y'),('04',19,'Y','Y','Y','Y','Y'),('01',59,'N','N','N','N','N'),('02',59,'N','N','N','N','N'),('03',18,'Y','Y','Y','Y','Y'),('88',59,'N','N','N','N','Y'),('04',18,'Y','Y','Y','Y','Y'),('01',60,'N','N','N','N','N'),('02',60,'N','N','N','N','Y'),('03',17,'Y','Y','Y','Y','Y'),('88',60,'N','N','N','N','Y'),('04',17,'Y','Y','Y','Y','Y'),('01',61,'N','N','N','N','N'),('02',61,'N','N','N','N','Y'),('03',16,'Y','Y','Y','Y','Y'),('88',61,'N','N','N','N','Y'),('04',16,'Y','Y','Y','Y','Y'),('01',62,'N','N','N','N','N'),('02',62,'N','N','N','N','Y'),('03',15,'Y','Y','Y','Y','Y'),('88',62,'N','N','N','N','Y'),('04',15,'Y','Y','Y','Y','Y'),('01',63,'N','N','N','N','N'),('02',63,'N','N','N','N','Y'),('03',14,'Y','Y','Y','Y','Y'),('88',63,'N','N','N','N','Y'),('04',14,'Y','Y','Y','Y','Y'),('01',64,'N','N','N','N','N'),('02',64,'N','N','N','N','Y'),('03',13,'Y','Y','Y','Y','Y'),('88',64,'N','N','N','N','Y'),('04',13,'Y','Y','Y','Y','Y'),('01',65,'N','N','N','N','N'),('02',65,'N','N','N','N','Y'),('03',12,'Y','Y','Y','Y','Y'),('88',65,'N','N','N','N','Y'),('04',12,'Y','Y','Y','Y','Y'),('01',66,'N','N','N','N','N'),('02',66,'N','N','N','N','Y'),('03',11,'Y','Y','Y','Y','Y'),('88',66,'N','N','N','N','Y'),('04',11,'Y','Y','Y','Y','Y'),('01',67,'N','N','N','N','N'),('02',67,'N','N','N','N','Y'),('03',10,'N','N','N','N','N'),('88',67,'N','N','N','N','Y'),('04',10,'N','N','N','N','N'),('01',68,'N','N','N','N','N'),('02',68,'N','N','N','N','Y'),('03',9,'N','N','N','N','N'),('88',68,'N','N','N','N','Y'),('04',9,'Y','N','Y','N','N'),('01',69,'N','N','N','N','N'),('02',69,'N','N','N','N','Y'),('03',8,'Y','N','Y','N','N'),('88',69,'N','N','N','N','Y'),('04',8,'N','N','N','N','N'),('01',70,'N','N','N','N','N'),('02',70,'N','N','N','N','Y'),('03',7,'Y','Y','Y','Y','N'),('88',70,'N','N','N','N','Y'),('04',7,'Y','Y','Y','Y','Y'),('01',71,'N','N','N','N','N'),('02',71,'N','N','N','N','Y'),('03',6,'N','N','N','N','N'),('88',71,'N','N','N','N','Y'),('04',6,'N','N','N','N','N'),('01',72,'N','N','N','N','N'),('02',72,'N','N','N','N','Y'),('03',5,'N','N','N','N','N'),('88',72,'N','N','N','N','Y'),('04',5,'N','N','N','N','N'),('01',73,'N','N','N','N','N'),('02',73,'N','N','N','N','Y'),('03',4,'Y','N','N','N','Y'),('88',73,'N','N','N','N','Y'),('04',4,'Y','Y','Y','Y','Y'),('01',74,'N','N','N','N','N'),('02',74,'N','N','N','N','Y'),('03',3,'Y','N','N','N','Y'),('88',74,'N','N','N','N','Y'),('04',3,'Y','N','N','N','Y'),('01',75,'N','N','N','N','N'),('02',75,'N','N','N','N','Y'),('03',2,'N','N','N','N','N'),('88',75,'N','N','N','N','Y'),('04',2,'Y','Y','Y','Y','Y'),('01',76,'N','N','N','N','N'),('02',76,'N','N','N','N','Y'),('03',1,'N','N','N','N','N'),('88',76,'N','N','N','N','Y'),('04',1,'Y','Y','Y','Y','Y'),('01',77,'N','N','N','N','N'),('02',77,'N','N','N','N','N'),('03',77,'N','N','N','N','N'),('04',77,'N','N','N','N','N'),('88',77,'Y','Y','N','Y','Y'),('01',78,'N','N','N','N','N'),('02',78,'N','N','N','N','N'),('03',78,'N','N','N','N','N'),('04',78,'N','N','N','N','N'),('88',78,'Y','Y','N','Y','Y');

/*Table structure for table `master_user` */

DROP TABLE IF EXISTS `master_user`;

CREATE TABLE `master_user` (
  `UserID` char(15) NOT NULL,
  `NamaUser` varchar(100) DEFAULT NULL,
  `Password` varchar(50) DEFAULT NULL,
  `kodegroup` varchar(10) DEFAULT NULL COMMENT 'Group Outlet',
  `locked` char(1) DEFAULT NULL,
  `objtype` int(11) DEFAULT NULL,
  `auditdate` datetime DEFAULT NULL,
  `audituser` char(8) DEFAULT NULL,
  `admin` char(1) DEFAULT NULL,
  PRIMARY KEY (`UserID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `master_user` */

insert  into `master_user`(`UserID`,`NamaUser`,`Password`,`kodegroup`,`locked`,`objtype`,`auditdate`,`audituser`,`admin`) values ('admin','administrator','nqzv{','88','N',0,'2010-09-01 03:46:53','admin','Y'),('User1','User 1','‚€r>','02','N',NULL,'2013-01-05 21:07:22','admin','N'),('User2','User 2','‚€r?','04','N',NULL,'2013-01-05 21:07:49','admin','N');

/*Table structure for table `sqlcommandlogs` */

DROP TABLE IF EXISTS `sqlcommandlogs`;

CREATE TABLE `sqlcommandlogs` (
  `cmdid` int(11) NOT NULL,
  `cmdsql` varchar(8000) NOT NULL,
  `cmdsyncsts` char(1) NOT NULL,
  `auditdate` timestamp NULL DEFAULT NULL,
  `audituser` varchar(15) DEFAULT NULL,
  PRIMARY KEY (`cmdid`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `sqlcommandlogs` */

/*Table structure for table `sr_menu_copy` */

DROP TABLE IF EXISTS `sr_menu_copy`;

CREATE TABLE `sr_menu_copy` (
  `moduleid` int(11) DEFAULT NULL,
  `modulename` varchar(100) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `sr_menu_copy` */

insert  into `sr_menu_copy`(`moduleid`,`modulename`) values (51,'mnu_MasterSetting'),(52,'mnu_sr_counter'),(53,'mnu_sr_area'),(54,'mnu_ser_category'),(55,'mnu_sr_groupcategory'),(56,'mnu_sr_racktype'),(57,'mnu_sr_locate'),(58,'mnu_sr_product'),(59,'mnu_IntransactionModules'),(60,'mnu_sr_transin'),(61,'mnu_sr_autoputawayprocess'),(62,'mnu_sr_transin_confirmation'),(63,'mnu_OutTransactionModules'),(64,'mnu_sr_transout'),(65,'mnu_sr_autopickingprocess'),(66,'mnu_sr_transout_confirmation'),(67,'mnu_TransferTransactionModules'),(68,'mnu_sr_transtransfer'),(69,'mnu_sr_replanishmentprocess'),(70,'mnu_sr_transtransfer_confirmation'),(71,'mnu_Utility'),(72,'mnu_sr_soentry'),(73,'mnu_sr_dashboard'),(74,'mnu_sr_updaterackbalance'),(75,'mnu_SR_ReportForm'),(76,'mnuSR_StockCardForm'),(77,'mnuTransferForm'),(78,'mnuPackingListForm'),(79,'mnuPickingListForm'),(80,'mnuPutAwayForm'),(81,'mnuSR_itembylocation');

/* Procedure structure for procedure `sp_companyaccess_get` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_companyaccess_get` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_companyaccess_get`(in semplcode char(8))
BEGIN
	SELECT 	emplcode, 
	a.dbid, b.cmpnyid,b.cmpnyname,
	accessctrl
	FROM 
	companyaccess a INNER JOIN company b ON a.dbid=b.dbid
	where emplcode=semplcode;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_companyaccess_get_access` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_companyaccess_get_access` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_companyaccess_get_access`(in sEmplCode char(8),sCmpnyId varchar(20))
BEGIN
	declare sDBId char(1);
	select DBId into sDBId from company where CmpnyId=sCmpnyId;
	select AccessCtrl from companyaccess where EmplCode=sEmplCode and DBId=sDBId;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_companyaccess_insert` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_companyaccess_insert` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_companyaccess_insert`(in
	semplcode char(8),
	sdbid int(11),
	saccessctrl char(1),
	saudituser char(8))
BEGIN
	insert into companyaccess
	(
	emplcode,
	dbid,
	accessctrl,
	auditdate,
	audituser
	)
	values
	(
	semplcode,
	sdbid,
	saccessctrl,
	now(),
	saudituser
	);
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_companyaccess_update` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_companyaccess_update` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_companyaccess_update`(in
	semplcode char(8),
	sdbid int(11),
	saccessctrl char(1),
	saudituser char(8))
BEGIN
	update companyaccess
	set 
	
	accessctrl=saccessctrl,
	auditdate=now(),
	audituser=saudituser where emplcode=semplcode and dbid=sdbid;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_companyaccess_view` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_companyaccess_view` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_companyaccess_view`(IN sEmplCode CHAR(8))
BEGIN
	SELECT 	emplcode, a.dbid,b.cmpnyid,b.cmpnyname, accessctrl  
	FROM 
	companyaccess a INNER JOIN 
	company b ON a.dbid=b.dbid AND emplcode=sEmplCode;
	
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_company_delete` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_company_delete` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_company_delete`(in
	sdbid int(11),
	scmpnyid varchar(20),
	scmpnyname varchar(50),
	sdatabasename VARCHAR(10),
	saudituser varchar(8)
	)
BEGIN
	
	delete from company
	where dbid=sdbid;
	
	delete from companyaccess where dbid=sdbid;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_company_get` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_company_get` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_company_get`(in
	scmpnyid varchar(20),
	sget int(1)
	)
BEGIN
	if sget=0 then
	
	select dbid,cmpnyid,cmpnyname,databasename from company 
	where cmpnyid=scmpnyid order by dbid asc limit 1;
	
	end if;
	
	IF sget=1 THEN
	
	SELECT dbid,cmpnyid,cmpnyname,databasename FROM company 
	ORDER BY dbid ASC LIMIT 1;
	
	END IF;
	
	IF sget=2 THEN
	
	SELECT dbid,cmpnyid,cmpnyname,databasename FROM company 
	WHERE cmpnyid<scmpnyid ORDER BY dbid desc LIMIT 1;
	
	END IF;
	
	IF sget=3 THEN
	
	SELECT dbid,cmpnyid,cmpnyname,databasename FROM company 
	WHERE cmpnyid>scmpnyid ORDER BY dbid ASC LIMIT 1;
	
	END IF;
	
	IF sget=4 THEN
	
	SELECT dbid,cmpnyid,cmpnyname,databasename FROM company 
	ORDER BY dbid desc LIMIT 1;
	
	END IF;
	
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_company_insert` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_company_insert` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_company_insert`(in
	sdbid int(11),
	scmpnyid varchar(20),
	scmpnyname varchar(50),
	sdatabasename VARCHAR(10),
	saudituser varchar(8)
	)
BEGIN
	
	insert into company
	(
	cmpnyid,
	cmpnyname,databasename,
	auditdate,
	audituser
	)
	values
	(
	scmpnyid,
	scmpnyname,sdatabasename,
	now(),
	saudituser
	);
	
	select max(dbid) into sdbid from company;
	insert into companyaccess
	SELECT 	`UserID`,sdbid dbid, 'N' accessctrl, NOW() auditdate, saudituser audituser	 
	FROM `master_user`;
	
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_company_update` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_company_update` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_company_update`(in
	sdbid int(11),
	scmpnyid varchar(20),
	scmpnyname varchar(50),
	sdatabasename VARCHAR(10),
	saudituser varchar(8)
	)
BEGIN
	
	update company
	set
	cmpnyid=scmpnyid,
	cmpnyname=scmpnyname,databasename=sdatabasename,
	auditdate=now(),
	audituser=saudituser where dbid=sdbid;
	
		SELECT MAX(dbid) INTO sdbid FROM company;
	if 	(select count(*) from companyaccess where sdbid= dbid)=0 then
		INSERT INTO companyaccess
		SELECT 	`UserID`,sdbid dbid, 'N' accessctrl, NOW() auditdate, saudituser audituser	 
		FROM `master_user`;
	end if;
	
	
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_employee_delete` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_employee_delete` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_employee_delete`(in
	semplcode char(8),
	semplname varchar(100),
	spass varchar(50),
	skodegroup varchar(10),
	slocked char(1),
	saudituser char(8),
	sadmin char(1))
BEGIN
	delete from `master_user` where `UserID`=semplcode;
	delete from companyaccess where  emplcode=semplcode;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_employee_get` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_employee_get` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_employee_get`(in
	semplcode char(8),sget int(1))
BEGIN
if sget=0 then 
	SELECT 	`UserID` emplcode,`NamaUser` emplname,`Password` pass,a.kodegroup,b.groupuser, locked,admin
	FROM 
	`master_user` a INNER JOIN master_group_user b ON a.kodegroup=b.kodegroup AND `UserID`=semplcode
	ORDER BY `UserID` ASC LIMIT 1;
end if;
IF sget=1 THEN 
	SELECT 	`UserID` emplcode,`NamaUser` emplname,`Password`  pass,a.kodegroup,b.groupuser, locked,admin
	FROM 
	`master_user` a INNER JOIN master_group_user b ON a.kodegroup=b.kodegroup 
	ORDER BY `UserID` ASC LIMIT 1;
END IF;
IF sget=2 THEN 
	SELECT 	`UserID` emplcode,`NamaUser` emplname,`Password`  pass,a.kodegroup,b.groupuser, locked,admin
	FROM 
	`master_user` a INNER JOIN master_group_user b ON a.kodegroup=b.kodegroup AND `UserID`<semplcode
	ORDER BY `UserID` desc LIMIT 1;
END IF;
IF sget=3 THEN 
	SELECT 	`UserID` emplcode,`NamaUser`  emplname,`Password` pass,a.kodegroup,b.groupuser, locked,admin
	FROM 
	`master_user` a INNER JOIN master_group_user b ON a.kodegroup=b.kodegroup AND `UserID`>semplcode
	ORDER BY `UserID` ASC LIMIT 1;
END IF;
IF sget=4 THEN 
	SELECT 	`UserID` emplcode,`NamaUser`  emplname,`Password`  pass,a.kodegroup,b.groupuser, locked,admin
	FROM 
	`master_user` a INNER JOIN master_group_user b ON a.kodegroup=b.kodegroup 
	ORDER BY `UserID` desc LIMIT 1;
END IF;
END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_employee_insert` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_employee_insert` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_employee_insert`(in
	semplcode char(8),
	semplname varchar(100),
	spass varchar(50),
	skodegroup varchar(10),
	slocked char(1),
	saudituser char(8),
	sadmin char(1))
BEGIN
	insert into employee
	(
	emplcode,
	emplname,
	pass,
	kodegroup,
	locked,auditdate,
	audituser,
	admin
	)
	values
	(
	semplcode,
	semplname,
	spass,
	skodegroup,
	slocked,now(),
	saudituser,
	sadmin
	);
	
	INSERT INTO common_ic.companyaccess 
	(emplcode, 
	dbid, 
	accessctrl, 
	auditdate, 
	audituser
	)
	
	select semplcode, dbid, 'N', now(), saudituser
	from company ;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_employee_update` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_employee_update` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_employee_update`(in
	semplcode char(8),
	semplname varchar(100),
	spass varchar(50),
	skodegroup varchar(10),
	slocked char(1),
	saudituser char(8),
	sadmin char(1))
BEGIN
	update `master_user` set 
	`NamaUser`=semplname,
	`Password`=spass,
	kodegroup=skodegroup,
	locked=slocked,
	audituser=saudituser,
	admin=sadmin where `UserID`=semplcode;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_master_group_user_delete` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_master_group_user_delete` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_group_user_delete`(in
	skodegroup varchar(10),
	sgroupuser varchar(30),
	saudituser varchar(10))
BEGIN
		delete from master_group_user
		where kodegroup=skodegroup;
		
		delete from master_moduleaccess where kodegroup=skodegroup;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_master_group_user_insert` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_master_group_user_insert` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_group_user_insert`(in
	skodegroup varchar(10),
	sgroupuser varchar(30),
	saudituser varchar(10))
BEGIN
		insert into master_group_user
		(
		kodegroup,
		groupuser,auditdate,
		audituser
		)
		values
		(
		skodegroup,
		sgroupuser,now(),
		saudituser
		);
		
		insert into master_moduleaccess
		SELECT 	skodegroup, modulid, 'N' Baca,'N' Tulis,'N' Edit,'N' Hapus,'N' Cetak
		FROM master_module;
	
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_master_group_user_update` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_master_group_user_update` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_group_user_update`(in
	skodegroup varchar(10),
	sgroupuser varchar(30),
	saudituser varchar(10))
BEGIN
		update master_group_user
		set
		
		groupuser=sgroupuser,
		audituser=saudituser,
		auditdate=now() where kodegroup=skodegroup;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_master_module_delete` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_master_module_delete` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_module_delete`(in
	sModulid int(11),
	sModuleMenu varchar(50),
	sDscription varchar(100),
	svisorder int(10),
	stransid char(1),
	sMenuIndex int(11),saudituser CHAR(10))
BEGIN
	delete from  master_module
	where Modulid=sModulid;
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_master_module_insert` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_master_module_insert` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_module_insert`(in
	sModulid int(11),
	sModuleMenu varchar(50),
	sDscription varchar(100),
	svisorder int(10),
	stransid char(1),
	sMenuIndex int(11),saudituser CHAR(10))
BEGIN
	insert into master_module
	(
	Modulid,
	ModuleMenu,
	Dscription,
	visorder,
	transid,
	MenuIndex,audituser
	)
	values
	(
	sModulid,
	sModuleMenu,
	sDscription,
	svisorder,
	stransid,
	sMenuIndex,saudituser
	);
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_master_module_insert_auto` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_master_module_insert_auto` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_module_insert_auto`(in sModuleMenu  VARCHAR(50),sDscription  VARCHAR(100),stransid char(1),saudituser char(10))
BEGIN
	declare svisorder int default 0;
	declare sModulid int;
	select ifnull(max(visorder)+10,10) into svisorder from master_module;
	
	if (select count(*) from master_module where ModuleMenu=sModuleMenu)=0 then
		INSERT INTO master_module 
		(ModuleMenu, Dscription, visorder, transid, 
		Auditdate, 
		AuditUser
		)
		VALUES
		(sModuleMenu,sDscription, svisorder, stransid, 
		now(), 
		saudituser
		);
		
		select Modulid into sModulid from master_module where ModuleMenu=sModuleMenu;
		insert into master_moduleaccess
		SELECT 	kodegroup, sModulid as modulid, 
		IF(kodegroup='88',if(stransid='2','Y','N'),'N') AS Baca, 
		IF(kodegroup='88',IF(stransid='2','Y','N'),'N') AS Tulis, 
		IF(kodegroup='88',IF(stransid='3','Y','N'),'N') AS Edit, 
		IF(kodegroup='88',IF(stransid='2','Y','N'),'N') AS Hapus, 
		IF(kodegroup='88','Y','N') AS Cetak	 
		FROM 
		master_group_user ;
		
	end if;
	
    END */$$
DELIMITER ;

/* Procedure structure for procedure `sp_master_module_update` */

/*!50003 DROP PROCEDURE IF EXISTS  `sp_master_module_update` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_module_update`(in
	sModulid int(11),
	sModuleMenu varchar(50),
	sDscription varchar(100),
	svisorder int(10),
	stransid char(1),
	sMenuIndex int(11),saudituser char(10))
BEGIN
	update master_module
	set
	
	ModuleMenu=sModuleMenu,
	Dscription=sDscription,
	visorder=svisorder,
	transid=stransid,
	MenuIndex=sMenuIndex,
	audituser=saudituser where Modulid=sModulid;
    END */$$
DELIMITER ;

/*Table structure for table `v_company_login_browse` */

DROP TABLE IF EXISTS `v_company_login_browse`;

/*!50001 DROP VIEW IF EXISTS `v_company_login_browse` */;
/*!50001 DROP TABLE IF EXISTS `v_company_login_browse` */;

/*!50001 CREATE TABLE  `v_company_login_browse`(
 `emplcode` char(8) ,
 `cmpnyid` varchar(20) ,
 `cmpnyname` varchar(50) ,
 `accessctrl` char(1) 
)*/;

/*View structure for view v_company_login_browse */

/*!50001 DROP TABLE IF EXISTS `v_company_login_browse` */;
/*!50001 DROP VIEW IF EXISTS `v_company_login_browse` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `v_company_login_browse` AS (select `b`.`emplcode` AS `emplcode`,`a`.`cmpnyid` AS `cmpnyid`,`a`.`cmpnyname` AS `cmpnyname`,`b`.`accessctrl` AS `accessctrl` from (`company` `a` join `companyaccess` `b` on((`a`.`dbid` = `b`.`dbid`)))) */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
