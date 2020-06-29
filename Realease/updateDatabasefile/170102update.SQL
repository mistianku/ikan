/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;

USE `ikan_sby`;

/* Alter table in target */
ALTER TABLE `master_customer` 
	ADD COLUMN `is_print_company` char(1)  COLLATE latin1_swedish_ci NULL DEFAULT '1' COMMENT '0 tidak 1 ya' after `is_invoice_spc`;


/* Alter table in target */
ALTER TABLE `master_produk` 
	CHANGE `kodebarcode` `kodebarcode` varchar(20)  COLLATE latin1_swedish_ci NULL after `kodefungsi`;

/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_master_customer_delete`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_customer_delete`(in
    skodecustomer CHAR(15),
	snamacustomer VARCHAR(50),skodesalesman CHAR(10),
	skodeharga VARCHAR(6),
	skodediskon VARCHAR(6),
	skodegudang VARCHAR(6),sfee VARCHAR(6),
	sppn INT(1),
	sjtempo INT(1),
	sjbayar CHAR(1),
	salamat1 VARCHAR(100),
	salamat2 VARCHAR(100),
	skota VARCHAR(20),
	stelp VARCHAR(30),
	sfaximale VARCHAR(30),
	saktif CHAR(1),
	spic VARCHAR(20),
	spichp VARCHAR(20),skodearea VARCHAR(5),sis_invoice_spc CHAR(1),
	sis_print_company CHAR(1),
	saudituser VARCHAR(10))
BEGIN
	delete from master_customer
	where kodecustomer=skodecustomer;
	delete from master_produk_harga_customer where kodecustomer=skodecustomer;
	
    END$$
DELIMITER ;


/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_master_customer_insert`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_customer_insert`(in
        skodecustomer char(15),
	snamacustomer varchar(50),skodesalesman CHAR(10),
	skodeharga varchar(6),
	skodediskon varchar(6),
	skodegudang varchar(6),sfee VARCHAR(6),
	sppn int(1),
	sjtempo int(1),
	sjbayar char(1),
	salamat1 varchar(100),
	salamat2 varchar(100),
	skota varchar(20),
	stelp varchar(30),
	sfaximale varchar(30),
	saktif char(1),
	spic varchar(20),
	spichp varchar(20),skodearea varchar(5),sis_invoice_spc char(1),
	sis_print_company CHAR(1),
	saudituser varchar(10))
BEGIN
	insert into master_customer
	(
	
	kodecustomer,
	namacustomer,kodesalesman,
	kodeharga,
	kodediskon,
	kodegudang,fee,
	ppn,
	jtempo,
	jbayar,
	alamat1,
	alamat2,
	kota,
	telp,
	faximale,
	aktif,
	pic,
	pichp,
	audituser,
	kodearea,is_invoice_spc,is_print_company,
	auditdate
	)
	values
	(
	skodecustomer,
	snamacustomer,skodesalesman,
	skodeharga,
	skodediskon,
	skodegudang,sfee,
	sppn,
	sjtempo,
	sjbayar,
	salamat1,
	salamat2,
	skota,
	stelp,
	sfaximale,
	saktif,
	spic,
	spichp,
	saudituser,
	skodearea,sis_invoice_spc,sis_print_company,
	now()
	);
	INSERT INTO master_produk_harga_customer
	(kodecustomer, kodeproduk, kodeharga, harga, fee, audituser, auditdate,namaproduk,kodeproduk_custmr	)
	SELECT  skodecustomer,a.kodeproduk,a.kodeharga,a.harga,a.harga*b.diskon/100 AS fee ,'' saudituser,NOW(),
	mp.namaproduk,a.kodeproduk
	FROM master_produk_harga a 
	INNER JOIN master_produk_fee b ON a.kodeproduk=b.kodeproduk 
	INNER JOIN master_produk mp ON mp.`kodeproduk`=a.`kodeproduk`
	AND a.kodeharga=skodeharga AND b.kodediskon=sfee ;
	
	call sp_master_customer_company_insert_from_customer_all(saudituser);
	
    END$$
DELIMITER ;


/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_master_customer_mov`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_customer_mov`(in skodecustomer varchar(15),sget int(1))
BEGIN
	if sget=0 then
	SELECT 	kodecustomer, namacustomer, a.kodesalesman,ms.namasalesman, 
	a.kodeharga,IFNULL(mh.namaharga,'') AS namaharga, a.kodediskon,IFNULL(md.keterangan,'') AS namadiskon,
	a.kodegudang,IFNULL(mg.namagudang,'') AS namagudang,a.fee,ifnull(mf.keterangan,'') namafee,
	ppn, 
	jtempo, jbayar, a.alamat1, a.alamat2, a.kota, a.telp, a.faximale, a.aktif, a.pic, a.pichp,ifnull(a.kodearea,'') kodearea,a.is_invoice_spc
	,is_print_company 
	FROM 
	master_customer a
	LEFT JOIN master_harga mh ON a.kodeharga=mh.kodeharga
	LEFT JOIN master_gudang mg ON mg.kodegudang=a.kodegudang
	LEFT JOIN master_diskon md ON md.kodediskon=a.kodediskon
	LEFT JOIN master_fee mf ON mf.kodediskon=a.kodediskon
	LEFT JOIN master_salesman ms ON ms.kodesalesman=a.kodesalesman
	where kodecustomer=skodecustomer 
	order by kodecustomer asc limit 1;
	end if;
	IF sget=1 THEN
	SELECT 	kodecustomer, namacustomer, a.kodesalesman,ms.namasalesman,  
	a.kodeharga,IFNULL(mh.namaharga,'') AS namaharga, a.kodediskon,IFNULL(md.keterangan,'') AS namadiskon,
	a.kodegudang,IFNULL(mg.namagudang,'') AS namagudang,a.fee,IFNULL(mf.keterangan,'') namafee,
	ppn, 
	jtempo, jbayar, a.alamat1, a.alamat2, a.kota, a.telp, a.faximale, a.aktif, a.pic, a.pichp,IFNULL(a.kodearea,'') kodearea,a.is_invoice_spc
	,is_print_company 
	FROM 
	master_customer a
	LEFT JOIN master_harga mh ON a.kodeharga=mh.kodeharga
	LEFT JOIN master_gudang mg ON mg.kodegudang=a.kodegudang
	LEFT JOIN master_diskon md ON md.kodediskon=a.kodediskon
	LEFT JOIN master_fee mf ON mf.kodediskon=a.kodediskon
	LEFT JOIN master_salesman ms ON ms.kodesalesman=a.kodesalesman
	ORDER BY kodecustomer ASC limit 1;
	END IF;
	IF sget=2 THEN
	SELECT 	kodecustomer, namacustomer, a.kodesalesman,ms.namasalesman,  
	a.kodeharga,IFNULL(mh.namaharga,'') AS namaharga, a.kodediskon,IFNULL(md.keterangan,'') AS namadiskon,
	a.kodegudang,IFNULL(mg.namagudang,'') AS namagudang,a.fee,IFNULL(mf.keterangan,'') namafee,
	ppn, 
	jtempo, jbayar, a.alamat1, a.alamat2, a.kota, a.telp, a.faximale, a.aktif, a.pic, a.pichp,IFNULL(a.kodearea,'') kodearea,a.is_invoice_spc
	,is_print_company 
	FROM 
	master_customer a
	LEFT JOIN master_harga mh ON a.kodeharga=mh.kodeharga
	LEFT JOIN master_gudang mg ON mg.kodegudang=a.kodegudang
	LEFT JOIN master_diskon md ON md.kodediskon=a.kodediskon
	LEFT JOIN master_fee mf ON mf.kodediskon=a.kodediskon
	LEFT JOIN master_salesman ms ON ms.kodesalesman=a.kodesalesman
	WHERE kodecustomer<skodecustomer 
	ORDER BY kodecustomer desc limit 1;
	END IF;
	IF sget=3 THEN
	SELECT 	kodecustomer, namacustomer, a.kodesalesman,ms.namasalesman, 
	a.kodeharga,IFNULL(mh.namaharga,'') AS namaharga, a.kodediskon,IFNULL(md.keterangan,'') AS namadiskon,
	a.kodegudang,IFNULL(mg.namagudang,'') AS namagudang,a.fee,IFNULL(mf.keterangan,'') namafee,
	ppn, 
	jtempo, jbayar, a.alamat1, a.alamat2, a.kota, a.telp, a.faximale, a.aktif, a.pic, a.pichp,IFNULL(a.kodearea,'') kodearea,a.is_invoice_spc
	,is_print_company 
	FROM 
	master_customer a
	LEFT JOIN master_harga mh ON a.kodeharga=mh.kodeharga
	LEFT JOIN master_gudang mg ON mg.kodegudang=a.kodegudang
	LEFT JOIN master_diskon md ON md.kodediskon=a.kodediskon
	LEFT JOIN master_fee mf ON mf.kodediskon=a.kodediskon
	LEFT JOIN master_salesman ms ON ms.kodesalesman=a.kodesalesman
	WHERE kodecustomer>skodecustomer 
	ORDER BY kodecustomer ASC limit 1;
	END IF;
	IF sget=4 THEN
	SELECT 	kodecustomer, namacustomer, a.kodesalesman,ms.namasalesman, 
	a.kodeharga,IFNULL(mh.namaharga,'') AS namaharga, a.kodediskon,IFNULL(md.keterangan,'') AS namadiskon,
	a.kodegudang,IFNULL(mg.namagudang,'') AS namagudang,a.fee,IFNULL(mf.keterangan,'') namafee,
	ppn, 
	jtempo, jbayar, a.alamat1, a.alamat2, a.kota, a.telp, a.faximale, a.aktif, a.pic, a.pichp,IFNULL(a.kodearea,'') kodearea,a.is_invoice_spc
	,is_print_company 
	FROM 
	master_customer a
	LEFT JOIN master_harga mh ON a.kodeharga=mh.kodeharga
	LEFT JOIN master_gudang mg ON mg.kodegudang=a.kodegudang
	LEFT JOIN master_diskon md ON md.kodediskon=a.kodediskon
	LEFT JOIN master_fee mf ON mf.kodediskon=a.kodediskon
	LEFT JOIN master_salesman ms ON ms.kodesalesman=a.kodesalesman
	ORDER BY kodecustomer desc limit 1;
	END IF;
    END$$
DELIMITER ;


/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_master_customer_update`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_customer_update`(in
    skodecustomer varchar(15),
	snamacustomer varchar(50),skodesalesman char(10),
	skodeharga varchar(6),
	skodediskon varchar(6),
	skodegudang varchar(6),sfee VARCHAR(6),
	sppn int(1),
	sjtempo int(1),
	sjbayar char(1),
	salamat1 varchar(100),
	salamat2 varchar(100),
	skota varchar(20),
	stelp varchar(30),
	sfaximale varchar(30),
	saktif char(1),
	spic varchar(20),
	spichp VARCHAR(20),skodearea VARCHAR(5),sis_invoice_spc CHAR(1),
	sis_print_company CHAR(1),
	saudituser varchar(10))
BEGIN
	update master_customer
	set
	namacustomer=snamacustomer,kodesalesman=skodesalesman,
	kodeharga=skodeharga,
	kodediskon=skodediskon,
	kodegudang=skodegudang,fee=sfee,
	ppn=sppn,
	jtempo=sjtempo,
	jbayar=sjbayar,
	alamat1=salamat1,
	alamat2=salamat2,
	kota=skota,
	telp=stelp,
	faximale=sfaximale,
	aktif=saktif,
	pic=spic,
	pichp=spichp,
	audituser=saudituser,kodearea=skodearea,is_invoice_spc=sis_invoice_spc,
	is_print_company=sis_print_company,
	auditdate=now()
	where
	kodecustomer=skodecustomer;
	
	INSERT INTO master_produk_harga_customer
	(kodecustomer, kodeproduk, kodeharga, harga, fee, audituser, auditdate	)
	SELECT skodecustomer,a.kodeproduk,a.kodeharga,a.harga,if(ifnull(b.diskon,0)=0,0, a.harga*b.diskon/100) AS fee ,saudituser,NOW()
	FROM master_produk_harga a 
	left JOIN master_produk_fee b ON a.kodeproduk=b.kodeproduk 
	AND a.kodeharga=skodeharga AND b.kodediskon=sfee WHERE NOT EXISTS
	(SELECT * FROM master_produk_harga_customer WHERE kodecustomer=skodecustomer AND kodeproduk=a.kodeproduk AND 
	kodeharga=a.kodeharga);
	
	
    END$$
DELIMITER ;


/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_master_produk_delete`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_produk_delete`(in 
	skodeproduk char(15),snamaproduk varchar(50),saktif varchar(1),skodebrand char(6),skodekategori char(6),
	skodefungsi char(6),skodebarcode varchar(20),suom1 int(11),suom2 int(11),sumo3 int(11),suom1sat varchar(5),
	suom2sat varchar(5),suom3sat varchar(5),sregisterdate datetime,saudituser varchar(10))
BEGIN
	delete from master_produk
	where kodeproduk=skodeproduk;
	
	DELETE FROM master_produk_harga WHERE kodeproduk=skodeproduk;
	DELETE FROM master_produk_diskon WHERE kodeproduk=skodeproduk;
	delete from master_inventori WHERE kodeproduk=skodeproduk;
	delete from master_produk_harga_customer where kodeproduk=skodeproduk;
	
    END$$
DELIMITER ;


/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_master_produk_insert`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_produk_insert`(in 
	skodeproduk char(15),snamaproduk varchar(50),saktif varchar(1),skodebrand char(6),skodekategori char(6),
	skodefungsi char(6),skodebarcode varchar(20),suom1 int(11),suom2 int(11),sumo3 int(11),suom1sat varchar(5),
	suom2sat varchar(5),suom3sat varchar(5),sregisterdate datetime,saudituser varchar(10))
BEGIN
	insert into master_produk
	(
	kodeproduk,
	namaproduk,
	aktif,
	kodebrand,
	kodekategori,
	kodefungsi,
	kodebarcode,
	uom1,
	uom2,
	umo3,
	uom1sat,
	uom2sat,
	uom3sat,
	registerdate,
	audituser,
	auditdate
	)
	values
	(
	skodeproduk,
	snamaproduk,
	saktif,
	skodebrand,
	skodekategori,
	skodefungsi,
	skodebarcode,
	suom1,
	suom2,
	sumo3,
	suom1sat,
	suom2sat,
	suom3sat,
	sregisterdate,
	saudituser,
	now()
	);
	
	INSERT INTO master_produk_diskon ( kodeproduk, kodediskon, diskon ) 
	SELECT mp.kodeproduk,md.kodediskon,md.diskon  FROM master_produk mp , master_diskon md  
	WHERE NOT EXISTS (SELECT * FROM master_produk_diskon mpd 
	WHERE mpd.kodeproduk=mp.kodeproduk AND mpd.kodediskon=md.kodediskon);
	
	INSERT INTO master_produk_harga (kodeproduk, kodeharga, harga) 
	SELECT a.kodeproduk,a.kodeharga, a.harga FROM 
	( SELECT a.kodeproduk,a.kodeharga, a.kelipatan*IFNULL(mph.harga,0) AS harga FROM 
	( SELECT mp.kodeproduk,mh.kodeharga, mh.kelipatan, mh.baseharga FROM master_produk mp , master_harga mh  ) AS a 
	LEFT JOIN master_produk_harga mph ON mph.kodeproduk=a.kodeproduk AND mph.kodeharga=a.baseharga) AS a 
	WHERE NOT EXISTS (SELECT * FROM master_produk_harga c WHERE c.kodeproduk=a.kodeproduk AND c.kodeharga=a.kodeharga);
	
	INSERT INTO master_inventori (kodegudang, kodeproduk, stock) SELECT mg.kodegudang, mp.kodeproduk,0 AS stock 
	FROM master_produk mp, master_gudang mg 
	WHERE NOT EXISTS 
	(SELECT * FROM master_inventori mi WHERE mi.kodegudang=mg.kodegudang AND mi.kodeproduk=mp.kodeproduk);
	
	INSERT INTO master_produk_fee (kodeproduk,kodediskon, diskon,audituser,auditdate)
	SELECT 	a.kodeproduk,kodediskon, diskon, saudituser,NOW() 
	FROM 
	master_produk a , master_fee b
	WHERE NOT EXISTS 
	( SELECT * FROM master_produk_fee WHERE a.kodeproduk =kodeproduk AND b.kodediskon=kodediskon);
	
	
	Call sp_master_produk_harga_customer_insert(skodeproduk,saudituser);
	
    END$$
DELIMITER ;


/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_master_produk_update`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_master_produk_update`(in 
	skodeproduk char(15),snamaproduk varchar(50),saktif varchar(1),skodebrand char(6),skodekategori char(6),
	skodefungsi char(6),skodebarcode varchar(20),suom1 int(11),suom2 int(11),sumo3 int(11),suom1sat varchar(5),
	suom2sat varchar(5),suom3sat varchar(5),sregisterdate datetime,saudituser varchar(10))
BEGIN
	update master_produk
	set
	namaproduk=snamaproduk,
	aktif=saktif,
	kodebrand=skodebrand,
	kodekategori=skodekategori,
	kodefungsi=skodefungsi,
	kodebarcode=skodebarcode,
	uom1=suom1,
	uom2=suom2,
	umo3=sumo3,
	uom1sat=suom1sat,
	uom2sat=suom2sat,
	uom3sat=suom3sat,
	registerdate=sregisterdate,
	audituser=saudituser,
	auditdate=now()
	where kodeproduk=skodeproduk;
	
    END$$
DELIMITER ;


/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_transaksi_keluar_form`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_transaksi_keluar_form`(in snodokumenfr char(15),snodokumento CHAR(15),
sisbarcode int(1),sget int(1))
BEGIN
if sget = 0 then   
	SELECT 	tk.docentry, 
	tk.nodokumen, 
	tk.tgldokumen, 
	tk.jtempo, 
	tk.jbayar, IF(tk.jbayar=1,'Cash',IF(tk.jbayar=2,'Transfer','Credit')) AS jbayardesc , 
	tk.totalsebpotongan, 
	tk.totalpotongan, 
	tk.totalsetpotongan, 
	tk.totalppn, 
	tk.totalsetppn,
	if(sisbarcode=0,tk1.kodeproduk,mp.kodebarcode) as kodeproduk,mp.namaproduk,
	tk1.jumlah,mp.uom1sat,
	tk1.harga,
	tk1.totalsebdiskon,tk.kodecustomer,mc.namacustomer,mc.alamat1, ifnull(p.CmpnyId,'') CmpnyId,
	ifnull(p.CmnyName,'') CmnyName,
	IFNULL(p.Address1,'') Address1,IFNULL(p.City,'') City,IFNULL(p.Faximale,'') Faximale,
	IFNULL(p.Phone1,'') Phone1,
	IFNULL(p.`NPWP`,'') NPWP,
	IFNULL(p.`PKPName`,'') PKPName,IFNULL(p.`PKPAddress1`,'') PKPAddress1,IFNULL(p.PKPAddress2,'') PKPAddress2,
	IFNULL(p.is_image,0) is_image,
	IFNULL(p.image_name,'') image_name,tk.referensi,mphc.namaproduk namaproduk_custmr,is_print_company,
	tk.nodokumen_sj
	FROM 
	transaksi_keluar tk 
	left join master_customer_company mcc ON mcc.kodecustomer=tk.kodecustomer AND mcc.aktive='Y'
	INNER JOIN transaksi_keluar_detail1 tk1 ON tk.docentry=tk1.docentry
	INNER JOIN master_customer mc ON tk.kodecustomer=mc.kodecustomer
	INNER JOIN master_produk mp ON mp.kodeproduk=tk1.kodeproduk 
	inner join `master_produk_harga_customer` mphc on mphc.`kodecustomer`=tk.kodecustomer 
	and mphc.`kodeproduk`=tk1.kodeproduk and mphc.kodeharga=tk.kodeharga
	left JOIN master_preferences_special p ON p.CmpnyID=mcc.CmpnyID 
	where nodokumen between snodokumenfr and snodokumento
	 ;
end if;
if sget=1 then
select count(*)
	FROM 
	transaksi_keluar tk 
	left JOIN master_customer_company mcc ON mcc.kodecustomer=tk.kodecustomer AND mcc.aktive='Y'
	INNER JOIN transaksi_keluar_detail1 tk1 ON tk.docentry=tk1.docentry
	INNER JOIN master_customer mc ON tk.kodecustomer=mc.kodecustomer
	INNER JOIN master_produk mp ON mp.kodeproduk=tk1.kodeproduk 
	left JOIN master_preferences_special p ON p.CmpnyID=mcc.CmpnyID 
	WHERE nodokumen BETWEEN snodokumenfr AND snodokumento;
end if;
	
	
    END$$
DELIMITER ;


/*  Alter Procedure in target  */

DELIMITER $$
DROP PROCEDURE IF EXISTS `sp_transaksi_kwitansi_form_tbp`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `sp_transaksi_kwitansi_form_tbp`(in snodokumenfr char(15),snodokumento CHAR(15),sget int(1))
BEGIN
if sget = 0 then   
	SELECT 	tk.docentry, tk1.linenum,
	tk.nodokumen, 
	tk.tgldokumen,
	tk.totalsebpotongan, 
	tk.totalpotongan, 
	tk.totalsetpotongan, 
	tk.totalppn, 
	tk.totalsetppn,tk.jml_faktur as jmlfaktur,txtnodokumen AS nofaktur,
	tk.kodecustomer,mc.namacustomer,mc.alamat1, p.CmpnyID ,
	p.CmnyName,p.Address1,p.City,p.Faximale,p.Phone1,p.PKPAddress2,is_print_company,
	terbilang1,terbilang2,terbilang3,IFNULL(p.is_image,0) is_image,IFNULL(p.image_name,'') image_name
	FROM transaksi_kwitansi tk 
	inner join master_customer_company mcc on mcc.kodecustomer=tk.kodecustomer and mcc.aktive='Y'
	left JOIN transaksi_kwitansidetail2 tk1 ON tk.docentry=tk1.docentry
	INNER JOIN master_customer mc ON tk.kodecustomer=mc.kodecustomer
	inner join master_preferences_special p on p.CmpnyID=mcc.CmpnyID
	where tk.nodokumen between snodokumenfr and snodokumento
	order by 	tk.docentry, tk1.linenum
	;
end if;
if sget=1 then
select count(*)
	FROM transaksi_kwitansi tk 
	INNER JOIN master_customer_company mcc ON mcc.kodecustomer=tk.kodecustomer AND mcc.aktive='Y'
	LEFT JOIN transaksi_kwitansidetail2 tk1 ON tk.docentry=tk1.docentry
	INNER JOIN master_customer mc ON tk.kodecustomer=mc.kodecustomer
	INNER JOIN master_preferences_special p ON p.CmpnyID=mcc.CmpnyID
	WHERE tk.nodokumen BETWEEN snodokumenfr AND snodokumento;
end if;
	
	
    END$$
DELIMITER ;


/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;