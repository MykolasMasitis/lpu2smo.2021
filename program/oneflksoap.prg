FUNCTION OneFlkSoap(para1, para2, para3)
 m.mcod      = para1
 m.lpuid     = para2
 m.lIsSilent = para3
 
 m.loForm = mailsoap
 m.IsVed = .F.
 
 m.Mmy = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 m.blpuid = 'b' + STR(m.lpuid,4) + '.' + m.mmy
 m.bmcod  = 'b' + m.mcod + '.' + m.mmy

 m.bfile = ''
* IF fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.blpuid)
*  m.bfile = m.blpuid
* ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.bmcod)
  m.bfile = m.bmcod
 ENDIF 

 IF EMPTY(m.bfile)
  RETURN .F.
 ENDIF 
 ffile = fso.GetFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.bfile)
 IF ffile.size >= 2
  fhandl = ffile.OpenAsTextStream
  lcHead = fhandl.Read(2)
  fhandl.Close
 ELSE 
  RETURN .F.
 ENDIF 
 ZipDir  = m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\'
 InDir = pBase + '\' + m.gcPeriod + '\' + m.mcod
 ZipName = ZipDir + m.bfile

 IF lcHead == 'PK' && Это zip-файл!
  IF !UnzipOpen(ZipName)
   RETURN .F.
  ENDIF 
  UnzipClose()
 ELSE 
  RETURN .F.
 ENDIF 
 && Проверяем комплектность посылки - наличие 5 файлов!
 UnzipOpen(ZipName)
 hItem    = 'H'  + STR(m.lpuid,4) + '.' + m.mmy
 dItem    = 'D'  + STR(m.lpuid,4) + '.' + m.mmy
 nvItem   = 'NV' + STR(m.lpuid,4) + '.' + m.mmy
 nsItem   = 'NS' + STR(m.lpuid,4) + '.' + m.mmy
 rItem    = 'R' + m.qcod + '.' + m.mmy
 sItem    = 'S' + m.qcod + '.' + m.mmy
 hoItem   = 'HO' + m.qcod + '.' + m.mmy
 dsItem   = 'D79S' + m.qcod + '.' + m.mmy
 sprItem  = 'SPR' + STR(m.lpuid,4) + '.' + m.mmy

 * Файлы онкологии
 onkItem  = 'ONK_SL' + m.qcod + '.' + m.mmy && "старый" файл

 SLItem    = 'ONK_SL' + m.qcod + '.' + m.mmy
 USLItem   = 'ONK_USL' + m.qcod + '.' + m.mmy
 CONSItem  = 'ONK_CONS' + m.qcod + '.' + m.mmy
 LSItem    = 'ONK_LS' + m.qcod + '.' + m.mmy
 NAPRItem  = 'ONK_NAPR_V_OUT' + m.qcod + '.' + m.mmy
 DIAGItem  = 'ONK_DIAG' + m.qcod + '.' + m.mmy
 PROTItem  = 'ONK_PROT' + m.qcod + '.' + m.mmy
 * Файлы онкологии
 
 * Файл ЛС ковид
 CVItem    = 'CV_LS' + m.qcod + '.' + m.mmy
 * Файл ЛС ковид

 IF !UnzipGotoFileByName(dItem)
  UnzipClose()
  RETURN .F. 
 ENDIF 
 IF  !UnzipGotoFileByName(nvItem)
  UnzipClose()
  RETURN .F. 
 ENDIF 
 IF !UnzipGotoFileByName(rItem)
  UnzipClose()
  RETURN .F. 
 ENDIF 
 IF !UnzipGotoFileByName(sItem)
  UnzipClose()
  RETURN .F. 
 ENDIF 
 IF !UnzipGotoFileByName(sprItem)
  UnzipClose()
  RETURN .F. 
 ENDIF 
 UnzipClose()
 && Проверяем комплектность посылки - наличие 5 файлов!

 =ClDir()
 IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\b_flk_'+m.mcod)
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\b_flk_'+m.mcod)
 ENDIF 
 IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\b_mek_'+m.mcod)
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\b_mek_'+m.mcod)
 ENDIF 

 UnzipOpen(ZipName)
 UnZipSetFolder(m.pbase+'\'+m.gcperiod+'\'+m.mcod)
 UnZip()

 *UnzipGotoFileByName(rItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(sItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(dItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(nvItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(sprItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(HOItem)
 *UnzipFile(ZipDir)
 *IF UnzipGotoFileByName(OnkItem)
 * UnzipFile(ZipDir)
 *ENDIF 

 UnzipClose()

 SET DEFAULT TO (ZipDir)

 =OpenTemplates()

 =OpenFile("&dItem",  "dfile",  "SHARED")
 =OpenFile("&nvItem", "nvfile", "SHARED")
 =OpenFile("&rItem",  "rfile",  "SHARED")
 =OpenFile("&sItem",  "sfile",  "SHARED")
 =OpenFile("&sprItem", "sprfile", "SHARED")
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+HOitem)
  =OpenFile("&HOitem", "hofile", "SHARED")
 ENDIF 
 *IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\onk_sl'+m.qcod)
 * =OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\onk_sl'+m.qcod, "onkfile", "SHARED")
 *ENDIF 
 
 *IF !CheckFilesStucture()
 * =CloseTemplates() 
 * RETURN .F.  
 *ENDIF 
 =CloseItems()

 lcDir  = m.pBase + '\' + m.gcPeriod + '\' + m.mcod
 People = lcDir + '\people'
 Talon  = lcDir + '\talon'
 Otdel  = lcDir + '\otdel'
 Doctor = lcDir + '\doctor'
 Error  = lcDir + '\e' + m.mcod
 mError = lcDir + '\m' + m.mcod

 =CreateFilesStructure()
 =OpenLocalFiles()

 m.t_0 = SECONDS()

 m.paz     = 0 
 m.s_pred  = 0
 m.nsch    = 0
 m.krank   = 0
 m.paz_dst = 0
 m.paz_st  = 0
 m.paz_vmp = 0
 m.s_lek   = 0

 =MakePeople()
 =MakeTalon()
 =MakeOtdel() 
 =MakeDoctor()
 =MakeHO()
 
 =OpenFile(lcDir + '\talon', 'talon', 'shar', 'recid_lpu')

 =MakeOnkFile(SLItem)
 IF fso.FileExists(STRTRAN(SLItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(SLItem, m.mmy, 'dbf'), 'onk_sl', 'excl')=0
   SELECT onk_sl
   SET SAFETY OFF
   ALTER table onk_sl ADD COLUMN rid i
   ALTER table onk_sl ADD COLUMN sqlid i
   ALTER table onk_sl ADD COLUMN sqldt t
   IF USED('talon')
    SET RELATION TO recid_s INTO talon 
    REPLACE ALL rid WITH talon.recid 
    SET RELATION OFF INTO talon 
   ENDIF 
   INDEX on rid TAG rid 
   INDEX on recid_s TAG recid_s
   INDEX on recid TAG recid
   SET ORDER TO recid 
   SET SAFETY ON 
   *USE IN onk_sl
  ELSE 
   IF USED('onk_sl')
    USE IN onk_sl
   ENDIF 
  ENDIF 
 ENDIF 
 =MakeOnkFile(USLItem)
 IF fso.FileExists(STRTRAN(USLItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(USLItem, m.mmy, 'dbf'), 'onk_usl', 'excl')=0
   SELECT onk_usl
   SET SAFETY OFF
   ALTER table onk_usl ADD COLUMN rid i
   ALTER table onk_usl ADD COLUMN sqlid i
   ALTER table onk_usl ADD COLUMN sqldt t
   IF USED('onk_sl')
    SET RELATION TO recid_sl INTO onk_sl
    REPLACE ALL rid WITH onk_sl.rid 
    SET RELATION OFF INTO onk_sl
   ENDIF 
   INDEX on recid TAG recid
   INDEX on recid_sl TAG recid_s
   SET ORDER TO recid 
   SET SAFETY ON 
   *USE IN onk_usl
  ELSE 
   IF USED('onk_usl')
    USE IN onk_usl
   ENDIF 
  ENDIF 
 ENDIF 

 =MakeOnkFile(LSItem)
 IF fso.FileExists(STRTRAN(LSItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(LSItem, m.mmy, 'dbf'), 'onk_ls', 'excl')=0
   SELECT onk_ls
   SET SAFETY OFF
   ALTER table onk_ls ADD COLUMN rid i
   ALTER table onk_ls ADD COLUMN sqlid i
   ALTER table onk_ls ADD COLUMN sqldt t
   IF USED('onk_usl')
    SET RELATION TO recid_usl INTO onk_usl
    REPLACE ALL rid WITH onk_usl.rid 
    SET RELATION OFF INTO onk_usl
   ENDIF 
   ALTER TABLE onk_ls ADD COLUMN s_all n(11,2)
   ALTER TABLE onk_ls ADD COLUMN oms l
   INDEX on recid_usl TAG recid_s

   CREATE CURSOR sss (recid i, s_all n(11,2))
   SELECT sss
   INDEX on recid TAG recid
   SET ORDER TO recid
   
   SELECT onk_ls
   SCAN 
    m.tip_opl = IIF(FIELD('tip_opl')=UPPER('tip_opl'), tip_opl, 1)
    IF m.tip_opl!=1
     LOOP 
    ENDIF 
    m.cod = cod
    IF !INLIST(m.cod, 97158, 81094)
     LOOP 
    ENDIF 
    m.date_inj = date_inj
    IF !BETWEEN(m.date_inj, {01.03.2019}, m.tdat2)
     LOOP 
    ENDIF 
    
    m.recid = rid   
    m.ds_c = IIF(USED('talon') AND SEEK(m.recid, 'talon', 'recid'), talon.ds, '')
    IF EMPTY(m.ds_c)
     LOOP 
    ENDIF 
    IF !SEEK(m.ds_c, 'mkb_c')
     LOOP 
    ENDIF 

    m.r_up   = ALLTRIM(r_up) && розничая упаковка
    IF EMPTY(m.r_up)
     LOOP 
    ENDIF 

    m.dd_sid = sid
    m.dt_d   = dt_d && курсовая (дневная) доза в единицах назначения!

    m.s_all = FLS(m.dd_sid, m.dt_d, m.r_up)
    IF m.s_all<=0
     LOOP 
    ENDIF 
    
    IF m.s_all>0
     IF !SEEK(m.recid, 'sss')
      INSERT INTO sss FROM MEMVAR 
     ELSE 
      m.o_s_all = sss.s_all
      m.n_s_all = m.o_s_all + m.s_all
      UPDATE sss SET s_all = m.n_s_all WHERE recid=m.recid
     ENDIF 
    ENDIF 

    REPLACE s_all WITH m.s_all

   ENDSCAN 
   
   SELECT Talon
   IF USED('sss')
    SET RELATION TO recid INTO sss
    REPLACE ALL s_lek WITH sss.s_all
    SET RELATION OFF INTO sss
    USE IN sss 
   ENDIF 
   SUM s_lek TO m.s_lek

   SET SAFETY ON 
   USE IN onk_ls
  ELSE 
   IF USED('onk_ls')
    USE IN onk_ls
   ENDIF 
  ENDIF 
 ENDIF 

 =MakeOnkFile(CONSItem)
 IF fso.FileExists(STRTRAN(CONSItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(CONSItem, m.mmy, 'dbf'), 'onk_cons', 'excl')=0
   SELECT onk_cons
   SET SAFETY OFF
   ALTER table onk_cons ADD COLUMN rid i
   ALTER table onk_cons ADD COLUMN sqlid i
   ALTER table onk_cons ADD COLUMN sqldt t
   IF USED('talon')
    SET RELATION TO recid_s INTO talon
    REPLACE ALL rid WITH talon.recid
    SET RELATION OFF INTO talon
   ENDIF 
   INDEX on recid_s TAG recid
   SET SAFETY ON 
   USE IN onk_cons
  ELSE 
   IF USED('onk_cons')
    USE IN onk_cons
   ENDIF 
  ENDIF 
 ENDIF 

 =MakeOnkFile(NAPRItem)
 IF fso.FileExists(STRTRAN(NAPRItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(NAPRItem, m.mmy, 'dbf'), 'onk_napr', 'excl')=0
   SELECT onk_napr
   SET SAFETY OFF
   ALTER table onk_napr ADD COLUMN rid i
   ALTER table onk_napr ADD COLUMN sqlid i
   ALTER table onk_napr ADD COLUMN sqldt t
   IF USED('talon')
    SET RELATION TO recid_s INTO talon
    REPLACE ALL rid WITH talon.recid
    SET RELATION OFF INTO talon
   ENDIF 
   INDEX on recid_s TAG recid
   SET SAFETY ON
   USE IN onk_napr
  ELSE 
   IF USED('onk_napr')
    USE IN onk_napr
   ENDIF 
  ENDIF 
 ENDIF 
 
 =MakeOnkFile(DIAGItem)
 IF fso.FileExists(STRTRAN(DIAGItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(DIAGItem, m.mmy, 'dbf'), 'onk_diag', 'excl')=0
   SELECT onk_diag
   SET SAFETY OFF
   ALTER table onk_diag ADD COLUMN rid i
   ALTER table onk_diag ADD COLUMN sqlid i
   ALTER table onk_diag ADD COLUMN sqldt t
   IF USED('onk_sl')
    SET RELATION TO recid_sl INTO onk_sl
    REPLACE ALL rid WITH onk_sl.rid
    SET RELATION OFF INTO onk_sl
   ENDIF 
   INDEX on recid_sl TAG recid
   SET SAFETY ON 
   USE IN onk_diag
  ELSE 
   IF USED('onk_diag')
    USE IN onk_diag
   ENDIF 
  ENDIF 
 ENDIF 

 =MakeOnkFile(PROTItem)

 =MakeCVLS(CVItem)
 IF fso.FileExists(STRTRAN(CVItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(CVItem, m.mmy, 'dbf'), 'cv_ls', 'excl')=0
   SELECT cv_ls
   SET SAFETY OFF
   ALTER TABLE cv_ls ADD COLUMN rid i
   IF USED('talon')
    SET RELATION TO recid_s INTO talon 
    REPLACE ALL rid WITH talon.recid 
    SET RELATION OFF INTO talon 
   ENDIF 
   ALTER TABLE cv_ls ADD COLUMN s_all n(11,2)
   ALTER TABLE cv_ls ADD COLUMN oms l
   INDEX on rid TAG rid 
   INDEX on recid_s TAG recid_s
   *INDEX on recid TAG recid
   USE IN cv_ls
  ELSE 
   IF USED('cv_ls')
    USE IN cv_ls
   ENDIF 
  ENDIF 
 ENDIF 

 USE IN talon 
 IF USED('onk_sl')
  USE IN onk_sl
 ENDIF 
 IF USED('onk_usl')
  USE IN onk_usl
 ENDIF 
 
 IF fso.FileExists(STRTRAN(SLItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(SLItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(USLItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(USLItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(LSItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(LSItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(DIAGItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(DIAGItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(CONSItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(CONSItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(NAPRItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(NAPRItem, m.mmy, 'bak'))
 ENDIF 

 m.t_6 = SECONDS()
 
 loForm.get_recs.value = loForm.get_recs.value + 1
 loForm.get_paz.value  = loForm.get_paz.value + m.paz
 loForm.get_nsch.value = loForm.get_nsch.value + m.nsch
 loForm.get_sum.value  = loForm.get_sum.value + m.s_pred
 
 m.processed = DATETIME()

* UPDATE aisoms SET processed=m.processed, paz=m.paz, nsch=m.nsch, ;
 	s_pred=m.s_pred, s_lek=m.s_lek, krank=m.krank, paz_dst=m.paz_dst, paz_st=m.paz_st, paz_vmp=m.paz_vmp,; 
 	erz_id='', erz_status=0, sum_flk=0, ispr=.f., t_1=m.t_6-m.t_0, ; 
 	polltag='', polltagdt={}, soapstatus='' WHERE mcod=m.mcod

 ** Удаляем файлы-признаки отправки ФЛК и МЭК
 m.b_flk = pbase+'\'+m.gcperiod+'\'+m.mcod+'\b_flk_'+mcod
 IF fso.FileExists(m.b_flk)
  fso.DeleteFile(m.b_flk)
 ENDIF 
 m.b_mek = pbase+'\'+m.gcperiod+'\'+m.mcod+'\b_mek_'+mcod
 IF fso.FileExists(m.b_mek)
  fso.DeleteFile(m.b_mek)
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\expselected.dbf')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\expselected.dbf')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\expselected.cdx')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\expselected.cdx')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.zip')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.zip')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.xml')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.xml')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.http')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.http')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\polltag.xml')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\polltag.xml')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\polltag.http')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\polltag.http')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\soapans.dbf')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\soapans.dbf')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\request.http')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\request.http')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\request.xml')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\request.xml')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\data.xml')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\data.xml')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ctrl'+m.qcod+'.dbf')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ctrl'+m.qcod+'.dbf')
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\t_y_'+m.mcod)
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\t_y_'+m.mcod)
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+STR(m.lpuid,4)+m.qcod+'.'+m.mmy)
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+STR(m.lpuid,4)+m.qcod+'.'+m.mmy)
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ot'+STR(m.lpuid,4)+m.qcod+'.'+m.mmy)
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ot'+STR(m.lpuid,4)+m.qcod+'.'+m.mmy)
 ENDIF 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\d'+m.qcod+STR(m.lpuid,4)+'.'+m.mmy)
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\d'+m.qcod+STR(m.lpuid,4)+'.'+m.mmy)
 ENDIF 
 ** Удаляем файлы-признаки отправки ФЛК и МЭК
   
 ** Удаляем все файлы: протокол, акт, реестр актов, табличную форму актов
 m.l_path = pbase+'\'+m.gcperiod+'\'+m.mcod
 m.mmy    = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
   
 DIMENSION dim_files(5)
 dim_files(1) = "Pr"+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 dim_files(2) = "Mk" + STR(m.lpuid,4) + m.qcod + m.mmy
 dim_files(3) = "Mt" + STR(m.lpuid,4) + m.qcod + m.mmy
 dim_files(4) = "Mc" + STR(m.lpuid,4) + m.qcod + m.mmy
 dim_files(5) = 'pdf'+m.qcod+m.mmy
   
 FOR i=1 TO ALEN(dim_files,1)
  IF fso.FileExists(m.l_path+'\'+ALLTRIM(dim_files(i))+'.xls')
   fso.DeleteFile(m.l_path+'\'+ALLTRIM(dim_files(i))+'.xls')
  ENDIF 
  IF fso.FileExists(m.l_path+'\'+ALLTRIM(dim_files(i))+'.pdf')
   fso.DeleteFile(m.l_path+'\'+ALLTRIM(dim_files(i))+'.pdf')
  ENDIF 
 ENDFOR 
   
 RELEASE dim_files, l_path
 ** Удаляем все файлы: протокол, акт, реестр актов, табличную форму актов

 SELECT AisOms

 REPLACE processed WITH m.processed, paz WITH m.paz, nsch WITH m.nsch, ;
 	s_pred WITH m.s_pred, s_lek WITH m.s_lek, krank WITH m.krank, paz_dst WITH m.paz_dst, ; 
 	paz_st WITH m.paz_st, paz_vmp WITH m.paz_vmp, erz_id WITH '', erz_status WITH 0, ; 
 	sum_flk WITH 0, ls_flk WITH 0, ispr WITH .f., t_1 WITH m.t_6-m.t_0, polltag WITH '', polltagdt WITH {}, ; 
 	soapstatus WITH '' 
 	

 *UPDATE aisoms SET processed=m.processed, paz=m.paz, nsch=m.nsch, ;
 	s_pred=m.s_pred, s_lek=m.s_lek, krank=m.krank, paz_dst=m.paz_dst, paz_st=m.paz_st, paz_vmp=m.paz_vmp,; 
 	erz_id='', erz_status=0, sum_flk=0, ispr=.f., t_1=m.t_6-m.t_0, ; 
 	polltag='', polltagdt={}, soapstatus='' WHERE mcod=m.mcod

 loForm.Refresh 

 =CloseTemplates() 

 loForm.Refresh
 loForm.LockScreen=.f.

RETURN .T.

FUNCTION ClDir
 =CloseItems()
 IF fso.FileExists(dItem)
  DELETE FILE &dItem
 ENDIF 
 IF fso.FileExists(hItem)
  DELETE FILE &hItem
 ENDIF 
 IF fso.FileExists(nvItem)
  DELETE FILE &nvItem
 ENDIF 
 IF fso.FileExists(nsItem)
  DELETE FILE &nsItem
 ENDIF 
 IF fso.FileExists(rItem)
  DELETE FILE &ritem
 ENDIF 
 IF fso.FileExists(sItem)
  DELETE FILE &sItem
 ENDIF 
 IF fso.FileExists(dsItem)
  DELETE FILE &dsItem
 ENDIF 
 IF fso.FileExists(sprItem)
  DELETE FILE &sprItem
 ENDIF 
 IF fso.FileExists(hoItem)
  DELETE FILE &hoItem
 ENDIF 
 IF fso.FileExists(SLItem)
  DELETE FILE &SLItem
 ENDIF 
 IF fso.FileExists(USLItem)
  DELETE FILE &USLItem
 ENDIF 
 IF fso.FileExists(CONSItem)
  DELETE FILE &CONSItem
 ENDIF 
 IF fso.FileExists(LSItem)
  DELETE FILE &LSItem
 ENDIF 
 IF fso.FileExists(NAPRItem)
  DELETE FILE &NAPRItem
 ENDIF 
 IF fso.FileExists(DIAGItem)
  DELETE FILE &DIAGItem
 ENDIF 
 IF fso.FileExists(PROTItem)
  DELETE FILE &PROTItem
 ENDIF 
RETURN 

FUNCTION OpenTemplates
 tn_result = 0
 tn_result = tn_result + OpenFile("&ptempl\dxxxx.mmy", "d_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\hxxxx.mmy", "h_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\nvxxxx.mmy", "nv_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\nsxxxx.mmy", "ns_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\rqq.mmy", "r_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\sqq.mmy", "s_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\sqqv01.mmy", "sv01_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\d79sqq.mmy", "d79s_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\sprxxxx.mmy", "spr_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\HOqq.mmy", "spr_ho", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\ONK_SLqq.mmy", "spr_onk", "SHARED")
 
RETURN tn_result

FUNCTION CloseTemplates
 IF USED('people_sv')
  USE IN people_sv
 ENDIF 
 IF USED('talon_sv')
  USE IN talon_sv
 ENDIF 
 IF USED('otdel_sv')
  USE IN otdel_sv
 ENDIF 
 IF USED('doctor_sv')
  USE IN doctor_sv
 ENDIF 
 IF USED('d_et')
  USE IN d_et
 ENDIF 
 IF USED('h_et')
  USE IN h_et
 ENDIF 
 IF USED('nv_et')
  USE IN nv_et
 ENDIF 
 IF USED('ns_et')
  USE IN ns_et
 ENDIF 
 IF USED('r_et')
  USE IN r_et
 ENDIF 
 IF USED('s_et')
  USE IN s_et
 ENDIF 
 IF USED('sv01_et')
  USE IN sv01_et
 ENDIF 
 IF USED('d79s_et')
  USE IN d79s_et
 ENDIF 
 IF USED('spr_et')
  USE IN spr_et
 ENDIF 
 IF USED('spr_ho')
  USE IN spr_ho
 ENDIF 
 IF USED('spr_onk')
  USE IN spr_onk
 ENDIF 
RETURN 

FUNCTION CloseItems
 IF USED('dfile')
  USE IN dfile
 ENDIF 
 IF USED('hfile')
  USE IN hfile
 ENDIF 
 IF USED('nvfile')
  USE IN nvfile
 ENDIF 
 IF USED('nsfile')
  USE IN nsfile
 ENDIF 
 IF USED('rfile')
  USE IN rfile
 ENDIF 
 IF USED('sfile')
  USE IN sfile
 ENDIF 
 IF USED('dsfile')
  USE IN dsfile
 ENDIF 
 IF USED('sprfile')
  USE IN sprfile
 ENDIF 
 IF USED('hofile')
  USE IN hofile
 ENDIF 
 IF USED('onkfile')
  USE IN onkfile
 ENDIF 
RETURN 

FUNCTION IsAisDir()
 IF !fso.FolderExists(pAisOms)
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms, 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\INPUT')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser\INPUT', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\OUTPUT')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser\OUTPUT', 0+16, '')
  RETURN .F. 
 ENDIF

RETURN .T. 

FUNCTION WriteInBFile(BFullName, TextToWrite)
 CFG = FOPEN(BFullName,12)
 IsMyCommentExists = .F.
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  IF UPPER(READCFG) = 'MYCOMMENT'
   IsMyCommentExists = .T.
   LOOP 
  ENDIF 
 ENDDO
 IF !IsMyCommentExists
  nFileSize = FSEEK(CFG,0,2)
  =FWRITE(CFG, TextToWrite)
 ENDIF 
 = FCLOSE (CFG)
RETURN 

FUNCTION MakeDoctor
 tnvFile = lcDir+'\'+nvItem
 oSettings.CodePage('&tnvFile', 866, .t.)
* tnsFile = lcDir+'\'+nsItem
* oSettings.CodePage('&tnsFile', 866, .t.)
* USE (tnsFile) IN 0 ALIAS lcDoctor  EXCLUSIVE 
 USE (tnvFile) IN 0 ALIAS lcDoctor2 EXCLUSIVE 
 SELECT lcDoctor2
* INDEX on pcod TAG pcod 
* SET ORDER TO pcod 
* SELECT lcDoctor
* SET RELATION TO pcod INTO lcDoctor2
 SCAN 
  SCATTER MEMVAR
*  m.prvs_1 = lcDoctor2.prvs_1
*  m.prvs_2 = lcDoctor2.prvs_2
*  m.prvs_3 = lcDoctor2.prvs_3
*  m.prvs_4 = lcDoctor2.prvs_4
*  m.prvs_5 = lcDoctor2.prvs_5
*  m.prvs_6 = lcDoctor2.prvs_6
*  m.d_ser_1 = lcDoctor2.d_ser_1
*  m.d_ser_2 = lcDoctor2.d_ser_2
*  m.d_ser_3 = lcDoctor2.d_ser_3
*  m.d_ser_4 = lcDoctor2.d_ser_4
*  m.d_ser_5 = lcDoctor2.d_ser_5
*  m.d_ser_6 = lcDoctor2.d_ser_6
*  m.ps_1 = lcDoctor2.ps_1
*  m.ps_2 = lcDoctor2.ps_2
*  m.ps_3 = lcDoctor2.ps_3
*  m.ps_4 = lcDoctor2.ps_4
*  m.ps_5 = lcDoctor2.ps_5
*  m.ps_6 = lcDoctor2.ps_6

  m.dr = CTOD(SUBSTR(m.dr,7,2)+'.'+SUBSTR(m.dr,5,2)+'.'+SUBSTR(m.dr,1,4))

  INSERT INTO Doctor FROM MEMVAR 
*  m.un_key = m.mcod + ' ' + m.pcod
*  IF !SEEK(m.un_key, 'doctor_sv', 'unkey')
*   INSERT INTO Doctor_sv FROM MEMVAR 
*  ENDIF 
 ENDSCAN 
* SET RELATION OFF INTO lcDoctor2
 USE 
* SELECT lcDoctor2
* SET ORDER TO 
* DELETE TAG ALL 
* USE 
 USE IN Doctor
* fso.DeleteFile(lcDir+'\'+nvItem)
RETURN 

FUNCTION MakeOtdel
 tFile = lcDir+'\'+dItem
 oSettings.CodePage('&tFile', 866, .t.)
 USE (tFile) IN 0 ALIAS lcOtdel  EXCLUSIVE 
 SELECT lcOtdel
 SCAN 
  SCATTER FIELDS EXCEPT mcod MEMVAR
  INSERT INTO Otdel FROM MEMVAR 
*  m.un_key = m.mcod+' '+m.iotd
*  IF !SEEK(m.un_key, 'otdel_sv', 'unkey')
*   INSERT INTO Otdel_sv FROM MEMVAR 
*  ENDIF 
 ENDSCAN 
 USE 
 USE IN Otdel

 m.err = .f. 
 TRY 
  fso.DeleteFile(m.lcDir+'\'+m.dItem)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 IF m.err = .t. 
  MESSAGEBOX('ОШИБКА ПРИ УДАЛЕНИИ ФАЙЛА!'+CHR(13)+CHR(10)+oEx.Message, 0+64, m.mcod+' oneFlkSoap')
 ENDIF 

RETURN 

FUNCTION MakePeople
 tFile = lcDir+'\'+rItem
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&rItem IN 0 ALIAS lcRFile  EXCLUSIVE 
 SELECT lcRFile
 m.paz = 0 
 SCAN 
  SCATTER FIELDS EXCEPT tip_p MEMVAR 
  m.recid_lpu = m.recid
  m.period    = m.gcPeriod
  m.tipp      = tip_p
  
  m.prmcod  = IIF(SEEK(m.prik, 'sprlpu'), sprlpu.mcod, '')
  m.prmcods = IIF(SEEK(m.priks, 'sprlpu'), sprlpu.mcod, '')
  
  IF m.SaveInitPr = 1 && Сверка с номерником включена, режим по умолчанию

   m.qq = ''
   m.sv = ''
   DO CASE 
    CASE m.tipp='В'
     m.polis = ALLTRIM(sn_pol)
     IF LEN(m.polis)=9
      m.lpuid   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_tera, 0)
      m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
      m.lpuids  = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_stom, 0)
      m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')
     ENDIF 

    CASE INLIST(m.tipp,'П','Э','К')
     *m.polis   = enp
     m.polis   = LEFT(sn_pol,16)
     m.lpuid   = IIF(SEEK(m.polis, 'enp'), enp.lpu_tera, 0)
     m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
     m.lpuids  = IIF(SEEK(m.polis, 'enp'), enp.lpu_stom, 0)
     m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')

    CASE m.tipp='С'
     m.polis = ALLTRIM(sn_pol)
     m.lpuid   = IIF(SEEK(m.polis, 'kms'), kms.lpu_tera, 0)
     m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
     m.lpuids = IIF(SEEK(m.polis, 'kms'), kms.lpu_stom, 0)
     m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')

    OTHERWISE 
     && оставляем так, как подало МО
   ENDCASE 
  
  ENDIF 
  
  m.prmcod  = IIF(m.d_type='9', '', m.prmcod)
  m.prmcods = IIF(m.d_type='9', '', m.prmcods)

  RELEASE m.recid, m.d_beg, m.d_end, m.tip_p, m.s_all
  INSERT INTO People FROM MEMVAR
  m.paz = m.paz + 1
 ENDSCAN 
 USE 

 m.err = .f. 
 TRY 
  fso.DeleteFile(m.lcDir+'\'+m.rItem)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 IF m.err = .t. 
  MESSAGEBOX('ОШИБКА ПРИ УДАЛЕНИИ ФАЙЛА!'+CHR(13)+CHR(10)+oEx.Message, 0+64, m.mcod+' oneFlkSoap')
 ENDIF 

RETURN 

FUNCTION MakeHO
 tFile = lcDir+'\'+hoItem
 IF !fso.FileExists(tFile)
  RETURN 
 ENDIF 
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&hoItem IN 0 ALIAS lcHOFile  EXCLUSIVE 
 SELECT lcHOFile
 IF fso.FileExists(lcDir+'\ho'+qcod+'.dbf')
  fso.DeleteFile(lcDir+'\ho'+qcod+'.dbf')
 ENDIF 
 COPY TO &lcDir\ho&qcod
 USE 
 USE &lcDir\ho&qcod IN 0 ALIAS lcHOFile  EXCLUSIVE 
 ALTER TABLE lcHOFile ALTER COLUMN c_i c(30)
 SELECT lcHOFile
 INDEX on sn_pol+c_i+PADL(cod,6,'0') TAG unik
 INDEX on sn_pol+c_i TAG snp_ci
 USE 
 IF fso.FileExists(lcDir+'\'+hoItem)
  fso.DeleteFile(lcDir+'\'+hoItem)
 ENDIF 
RETURN 

FUNCTION MakeOnkFile(para1)
 PRIVATE tFile
 m.tFile  = ALLTRIM(para1)

 m.ntFile = STRTRAN(m.tFile, m.mmy, 'dbf')
 IF fso.FileExists(m.ntFile)
  fso.DeleteFile(m.ntFile)
 ENDIF 
 IF !fso.FileExists(m.tFile)
  RETURN 
 ENDIF 

 fso.CopyFile(m.tFile, m.ntFile)

 IF !fso.FileExists(m.ntFile)
  RETURN 
 ENDIF 
 
 m.err = .f. 
 TRY 
  fso.DeleteFile(m.tFile)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 IF m.err = .t. 
  MESSAGEBOX('ОШИБКА ПРИ УДАЛЕНИИ ФАЙЛА!'+CHR(13)+CHR(10)+oEx.Message, 0+64, m.mcod+' oneFlkSoap')
 ENDIF 

 oSettings.CodePage('&ntFile', 866, .t.)
RETURN 

FUNCTION MakeCVLS(para1)
 PRIVATE tFile
 m.tFile  = ALLTRIM(para1)

 m.ntFile = STRTRAN(m.tFile, m.mmy, 'dbf')
 IF fso.FileExists(m.ntFile)
  fso.DeleteFile(m.ntFile)
 ENDIF 
 IF !fso.FileExists(m.tFile)
  RETURN 
 ENDIF 

 fso.CopyFile(m.tFile, m.ntFile)

 IF !fso.FileExists(m.ntFile)
  RETURN 
 ENDIF 
 
 m.err = .f. 
 TRY 
  fso.DeleteFile(m.tFile)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 IF m.err = .t. 
  MESSAGEBOX('ОШИБКА ПРИ УДАЛЕНИИ ФАЙЛА!'+CHR(13)+CHR(10)+oEx.Message, 0+64, m.mcod+' oneFlkSoap')
 ENDIF 

 oSettings.CodePage('&ntFile', 866, .t.)
RETURN 

FUNCTION MakeTalon
 tFile = lcDir+'\'+sItem
 oSettings.CodePage('&tFile', 866, .t.)
* tFile = lcDir+'\'+dsItem
* oSettings.CodePage('&tFile', 866, .t.)
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN' IN 0 ALIAS Tarif SHARED ORDER cod 
 USE &lcDir\&sItem  IN 0 ALIAS lcSFile EXCLUSIVE 
* USE &lcDir\&dsItem IN 0 ALIAS lcDSFile EXCLUSIVE 
* SELECT lcDSFile
* INDEX ON recid TAG recid 
* SET ORDER TO recid
* SELECT lcSFile
* SET RELATION TO recid INTO lcDSFile

 SELECT lcSFile
 m.nsch = RECCOUNT('lcSFile')
 SCAN
  SCATTER MEMVAR 
*  SCATTER FIELDS lcDSFile.novor, lcDSFile.ds_s, lcDSFile.ds_p, lcDSFile.profil, lcDSFile.rslt,;
   lcDSFile.prvs, lcDSFile.ord, lcDSFile.ishod, lcDSFile.fil_id MEMVAR 
  m.recid_lpu = m.recid
  RELEASE m.recid

  IF IsUsl(m.cod) OR (IsKdP(m.cod) AND !IsEko(m.cod))
   IF !SEEK(m.sn_pol, 'pazamb')
    INSERT INTO pazamb (sn_pol) VALUES (m.sn_pol)
    m.krank = m.krank + 1
   ENDIF 
  ENDIF 
   
  IF IsMes(m.cod)
   IF !SEEK(m.c_i, 'pazst')
    INSERT INTO pazst (c_i) VALUES (m.c_i)
    m.paz_st = m.paz_st + 1
   ENDIF 
  ENDIF 

  IF IsKdS(m.cod) OR IsEko(m.cod)
   IF !SEEK(m.sn_pol, 'pazdst')
    INSERT INTO pazdst (sn_pol) VALUES (m.sn_pol)
    m.paz_dst = m.paz_dst + 1
   ENDIF 
  ENDIF 

  IF IsVMP(m.cod)
   IF !SEEK(m.sn_pol, 'pazvmp')
    INSERT INTO pazvmp (sn_pol) VALUES (m.sn_pol)
    m.paz_vmp = m.paz_vmp + 1
   ENDIF 
  ENDIF 
  
  m.k_u = IIF(m.k_u=0 AND FIELD('KD_FACT')='KD_FACT', m.kd_fact, m.k_u)

  m.otd   = m.iotd
  *m.s_all = fsumm(m.cod, m.tip, m.k_u, m.IsVed)
  *m.s_all = fsumm(m.cod, m.tip, IIF(BETWEEN(m.cod,97107,97158), m.kd_fact, m.k_u), m.IsVed)
  m.s_all = fsumm(m.cod, m.tip, IIF(INLIST(INT(m.cod/1000),97,197) AND !INLIST(97041,97013,197013), m.kd_fact, m.k_u), m.IsVed)
  m.period = m.gcPeriod
  *m.profil = IIF(SEEK(m.cod, 'profus'), ALLTRIM(profus.profil), '')
  m.profil = SUBSTR(m.iotd,4,3)
  m.n_kd = IIF(SEEK(m.cod,'tarif'), tarif.n_kd, 0)

  m.s_pred = m.s_pred + s_all
  
  IF OCCURS(' ',ALLTRIM(m.pcod)) > 0 && Составной код врача
   m.pcod  = SUBSTR(ALLTRIM(m.pcod),1,AT(' ',ALLTRIM(m.pcod))-1)
  ELSE 
   m.pcod  = ALLTRIM(LEFT(ALLTRIM(m.pcod),10))
  ENDIF 
  
  IF VARTYPE(m.lpu_ord)='C'
   m.lpu_ord = INT(VAL(m.lpu_ord))
  ENDIF 
  
  INSERT INTO Talon FROM MEMVAR 
 ENDSCAN 
 USE 

 IF IsStac(m.mcod)
  SELECT c_i DISTINCT FROM Talon ;
 	WHERE INLIST(cod,61400,161400,161401,70150,70160,70170,170150,170151,170160,170161,170170,170171) AND Tip='7' ;
  	INTO CURSOR t_tst READWRITE 
  SELECT t_tst
  INDEX on c_i TAG c_i 
  SET ORDER TO c_i 
  
  SELECT * FROM Talon WHERE c_i IN (SELECT c_i FROM t_tst);
  	ORDER BY c_i, d_u DESC INTO CURSOR Gosp READWRITE 
 	
  USE IN t_tst

  SELECT Gosp
  INDEX on c_i TAG c_i
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO c_i

  m.s_pred = 0
  SELECT Talon 
  SCAN 
   m.cod     = cod 
   m.tip     = tip 
   m.d_u     = d_u
   m.kd_fact = kd_fact
   m.k_u     = k_u
   m.k_u     = IIF(m.k_u=0 AND FIELD('KD_FACT')='KD_FACT', m.kd_fact, m.k_u)
   m.c_i     = c_i
  
   *m.summa = fsumm(m.cod, m.tip, IIF(BETWEEN(m.cod,97107,97158), m.kd_fact, m.k_u), m.IsVed)
   m.summa = fsumm(m.cod, m.tip, IIF(INLIST(INT(m.cod/1000),97,197) AND !INLIST(97041,97013,197013), m.kd_fact, m.k_u), m.IsVed)
   * Сделано для расчета covid
   IF INLIST(m.cod,61400,161400,161401,70150,70160,70170,170150,170151,170160,170161,170170,170171) AND m.tip='7'
    IF SEEK(m.c_i, 'Gosp')
     DO WHILE Gosp.c_i = m.c_i
      IF m.cod<>Gosp.cod AND (INLIST(INT(Gosp.cod/1000),83,183) OR INLIST(Gosp.cod,56029,156003)) AND Gosp.d_u>=m.d_u AND IIF(INLIST(INT(Gosp.cod/1000),83,183), Gosp.Tip='5', Gosp.d_type='5')
       *m.summa = fsumm(m.cod, '0', IIF(BETWEEN(m.cod,97107,97158), m.kd_fact, m.k_u), m.IsVed)
       m.summa = fsumm(m.cod, m.tip, IIF(INLIST(INT(m.cod/1000),97,197) AND !INLIST(97041,97013,197013), m.kd_fact, m.k_u), m.IsVed)
       *MESSAGEBOX(TRANSFORM(m.summa,'999999.99'),0+64,m.c_i)
      ENDIF        
      SKIP IN Gosp
     ENDDO 
    ENDIF 
   ENDIF 
   * Сделано для расчета covid
    m.s_pred = m.s_pred + m.summa
   REPLACE s_all WITH m.summa IN talon 
  ENDSCAN 
 ENDIF 

 m.err = .f. 
 TRY 
  fso.DeleteFile(m.lcDir+'\'+m.sItem)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 IF m.err = .t. 
  MESSAGEBOX('ОШИБКА ПРИ УДАЛЕНИИ ФАЙЛА!'+CHR(13)+CHR(10)+oEx.Message, 0+64, m.mcod+' oneFlkSoap')
 ENDIF 
 
* fso.DeleteFile(lcDir+'\'+dsItem)
 USE IN Tarif

 SELECT sn_pol, 1 AS tip_p, MIN(d_u) as min_p, MAX(d_u) as max_p, SUM(s_all) as s_all FROM talon WHERE EMPTY(tip) ;
   GROUP BY sn_pol INTO CURSOR intp
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 SELECT sn_pol, 2 AS tip_p, MIN(d_u-k_u) as min_s, MAX(d_u) as max_s, SUM(s_all) as s_all FROM talon GROUP BY sn_pol ;
  WHERE !EMPTY(tip) INTO CURSOR ints
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 * сюда вставить создание файла hosp
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\hosp.dbf')
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\hosp.dbf')
  ENDIF 

  SELECT c_i, SPACE(25) as sn_pol, MAX(d_u)-SUM(k_u)+1 as d_pos, MAX(d_u) as d_vip, coun(*) as cnt, ;
  	SUM(k_u) as k_u FROM talon WHERE IsMes(cod) OR IsVMP(cod) ;
  	GROUP BY c_i ORDER BY c_i ASC INTO CURSOR cur_h READWRITE 
  SELECT cur_h
  INDEX on c_i TAG c_i 
  INDEX on sn_pol TAG sn_pol
  INDEX on d_pos TAG d_pos
  SET ORDER TO c_i
  SET ORDER TO c_i IN talon 
  SET RELATION TO c_i INTO talon 
  REPLACE ALL sn_pol WITH talon.sn_pol
  SET RELATION OFF INTO talon 

  IF tMonth=1
  ELSE
   m.p_period = STR(tYear,4)+PADL(tMonth-1,2,'0') 
   IF fso.FileExists(pBase+'\'+m.p_period+'\'+m.mcod+'\hosp.dbf')
    APPEND FROM &pBase\&p_period\&mcod\hosp
   ENDIF 
  ENDIF 
  
  IF RECCOUNT('cur_h')>0
   SET ORDER TO d_pos
   COPY TO &pBase\&gcPeriod\&mcod\hosp CDX 
  ENDIF 
  USE 
 * сюда вставить создание файла hosp

 USE IN talon
 SELECT people
 SET RELATION TO sn_pol INTO intp
 SET RELATION TO sn_pol INTO ints ADDITIVE 
* SET ORDER TO unkey IN people_sv
* SET RELATION TO mcod+sn_pol INTO people_sv ADDITIVE 
* SET ORDER TO sn_pol IN people_sv
* SET RELATION TO sn_pol INTO people_sv ADDITIVE 
 SCAN 
  m.d_beg = MIN(IIF(!EMPTY(intp.min_p), intp.min_p, m.tdat2), IIF(!EMPTY(ints.min_s), ints.min_s, m.tdat2))
  m.d_end = MAX(intp.max_p, ints.max_s)
  DO CASE 
   CASE intp.tip_p == 1 AND ints.tip_p != 2
    m.tip_p = 1
   CASE intp.tip_p != 1 AND ints.tip_p == 2
    m.tip_p = 2
   CASE intp.tip_p == 1 AND ints.tip_p == 2
    m.tip_p = 3
   OTHERWISE 
    m.tip_p = 0
  ENDCASE 
  REPLACE people.d_beg WITH m.d_beg, people.d_end WITH m.d_end, tip_p WITH m.tip_p
  m.s_all = IIF(!EMPTY(intp.s_all), intp.s_all, 0) + IIF(!EMPTY(ints.s_all), ints.s_all, 0)
  REPLACE people.s_all WITH m.s_all

*  REPLACE people_sv.d_beg WITH m.d_beg, people_sv.d_end WITH m.d_end, people_sv.tip_p WITH m.tip_p
  m.s_all = IIF(!EMPTY(intp.s_all), intp.s_all, 0) + IIF(!EMPTY(ints.s_all), ints.s_all, 0)
*  REPLACE people_sv.s_all WITH m.s_all

 ENDSCAN 
* SET RELATION OFF INTO people_sv
 SET RELATION OFF INTO ints
 SET RELATION OFF INTO intp
 USE IN people
 USE IN ints
 USE IN intp 
RETURN 

FUNCTION CreateFilesStructure
 IF fso.FileExists(People+'.dbf')
  fso.DeleteFile(People+'.dbf')
 ENDIF 
 CREATE TABLE (People) ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1,;
   mcod c(7), prmcod c(7), prmcods c(7), period c(6), d_beg d, d_end d, s_all n(11,2), ;
   tip_p n(1), sn_pol c(25), tipp c(1), enp c(16), qq c(2), ;
   fam c(25), im c(20), ot c(20), w n(1), dr d, d_type c(1), ;
   sv c(3), recid_lpu c(7), IsPr L, fil_id n(6))
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON sn_pol TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX on dr TAG dr
 INDEX on s_all TAG s_all
 USE 

 IF fso.FileExists(Talon+'.dbf')
  fso.DeleteFile(Talon+'.dbf')
 ENDIF 
 *CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6),  ;
	 pcod c(10), otd c(8), cod n(6), tip c(1), d_u d, ;
	 k_u n(3), d_type c(1), s_all n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3),;
	 codnom c(14), kur n(5,3), ds_2 c(6), ds_3 c(6), det n(1), k2 n(5,3), tipgr c(1), ;
	 vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17),;
	 ord n(1), date_ord d, lpu_ord n(6), recid_lpu c(7), fil_id n(6), ;
	 ds_onk n(1), p_cel c(3), dn n(1), reab n(1), tal_d d, napr_v_out n(1), napr_v_in n(1), ;
	 IsPr L, vz l, mp c(1), n_kd n(3), f_type c(2), mm c(1)) && Новая структура с 01082018
* CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6),  ;
	 pcod c(10), otd c(8), cod n(6), tip c(1), d_u d, ;
	 k_u n(4), kd_fact n(3), d_type c(1), s_all n(11,2), s_lek n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3),;
	 codnom c(14), kur n(5,3), ds_2 c(6), ds_3 c(6), det n(1), k2 n(5,3), tipgr c(1), ;
	 vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17),;
	 ord n(1), date_ord d, lpu_ord n(6), recid_lpu c(7), fil_id n(6), ;
	 ds_onk n(1), p_cel c(3), dn n(1), reab n(1), tal_d d, napr_v_in n(1), ;
	 c_zab n(1), napr_usl c(15), vid_vme c(15),IsPr L, vz l, mp c(1), n_kd n(3), f_type c(2), ;
	 mm c(1), typ c(1)) && Новая структура с 01022019
 CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6), pcod c(10), otd c(8), ;
	 cod n(6), tip c(1), d_u d, k_u n(4), kd_fact n(4), n_kd n(3), d_type c(1), s_all n(11,2), ;
	 s_lek n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3), kur n(5,3), ds_2 c(6), ds_3 c(6), ;
	 det n(1), k2 n(5,3), vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17), ord n(1), date_ord d, ;
	 lpu_ord n(6), recid_lpu c(7), fil_id n(6), ds_onk n(1), p_cel c(3), dn n(1), reab n(1), tal_d d, napr_v_in n(1),;
	 c_zab n(1), mp c(1), typ c(1), dop_r n(2), vz n(1), IsPr L,;
	 sqlid i, sqldt t, prcell c(3), nsif n(1)) && Убрал поля codnom, napr_usl, vid_vme, tipgr, mm, vz, ispr, f_type 01.06.2019

 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol TAG sn_pol
 INDEX ON otd TAG otd
 INDEX on pcod TAG pcod
 INDEX ON ds TAG ds
 INDEX ON d_u TAG d_u
 INDEX ON cod TAG cod
 INDEX ON profil TAG profil
 INDEX ON sn_pol+STR(cod,6)+DTOS(d_u) TAG ExpTag
 INDEX ON sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u) TAG unik 
 INDEX ON tip TAG tip
 INDEX ON s_all TAG s_all
 USE 
 
 IF fso.FileExists(Otdel+'.dbf')
  fso.DeleteFile(Otdel+'.dbf')
 ENDIF 
 CREATE TABLE (Otdel) ;
	(recid c(6), mcod c(7), iotd c(8), name c(100), pr_name c(100), cnt_bed n(5), fil_id n(6))
 INDEX ON iotd TAG iotd
 USE 

 IF fso.FileExists(Doctor+'.dbf')
  fso.DeleteFile(Doctor+'.dbf')
 ENDIF 
 CREATE TABLE (Doctor) ;
   (pcod c(10),sn_pol c(25),fam c(25),im c(20),ot c(20),dr d, w n(1),;
    prvs n(4), d_ser d, d_ser2 d, d_prik d, iotd c(8),;
	lgot_r c(1),c_ogrn c(15),lpu_id n(6), fil_id n(6))
 INDEX ON pcod TAG pcod
 USE 

 IF fso.FileExists(Error+'.dbf')
  fso.DeleteFile(Error+'.dbf')
 ENDIF 
 *CREATE TABLE (Error) (f c(1), c_err c(3), detail c(1), rid i, "comment" c(250))
 CREATE TABLE (Error) (f c(1), c_err c(3), et n(1), detail c(1), rid i, tip n(1), dt t, usr c(6), ;
 "comment" c(250), sqlid i, sqldt t)
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 USE 

 IF fso.FileExists(mError+'.dbf')
  fso.DeleteFile(mError+'.dbf')
 ENDIF 
 CREATE TABLE (mError) ;
  (rid i autoinc, RecId i, cod n(6), k_u n(4), tip c(1), et c(1), ee c(1), usr c(6), d_exp d,;
   e_cod n(6), e_ku n(3), e_tip c(1), err_mee c(3), osn230 c(5), e_period c(6),  ;
   koeff n(4,2), straf n(4,2), docexp c(7), s_all n(11,2), s_1 n(11,2), s_2 n(11,2), impdata d,;
   subet n(1), reason c(1), n_akt c(15), t_akt c(2), d_edit d, d_akt d)
 INDEX ON rid TAG rid 
 INDEX ON RecId TAG recid
 *INDEX ON PADL(recid,6,'0')+et TAG id_et
 *INDEX ON PADL(recid,6,'0')+et+LEFT(err_mee,2) TAG unik
 INDEX ON PADL(recid,6,'0')+et+docexp+reason TAG id_et
 INDEX ON PADL(recid,6,'0')+et+docexp+reason+LEFT(err_mee,2) TAG unik
  INDEX ON PADL(recid,6,'0')+et TAG uniket
 USE 

 CREATE CURSOR pazamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR pazdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazst (c_i c(30))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR pazvmp (sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol
RETURN 

FUNCTION OpenLocalFiles
 USE (People) IN 0 ALIAS People SHARED
 USE (Talon)  IN 0 ALIAS Talon  SHARED   
 USE (Otdel)  IN 0 ALIAS Otdel  SHARED 
 USE (Doctor) IN 0 ALIAS Doctor SHARED 
RETURN 

FUNCTION CheckFilesStucture
 fld_1 = AFIELDS(tabl_1, 'dfile') && Проверка d-файла 
 fld_2 = AFIELDS(tabl_2, 'd_et')  && 1 столбец - название, 2 - тип,  3 - размерность, 4 - нулей после запятой
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(dItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(dItem)
*  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

* fld_1 = AFIELDS(tabl_1, 'hfile') && Проверка h-файла 
* fld_2 = AFIELDS(tabl_2, 'h_et')  && 1 столбец - название, 2 - тип,  3 - размерность, 4 - нулей после запятой
* IF fld_1 == fld_2 && Кол-во полей совпадает!
*  FieldsIdent = CompFields(hItem) && 0 - есть отличия, 1 - полное совпадение
*  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
*   =ClDir()
*   RETURN .F.
*  ENDIF 
* ELSE 
*  =DiffFields(hItem)
*  =CopyToTrash(m.InDir,2)
*  =ClDir()
*  RETURN .F.
* ENDIF 

 fld_1 = AFIELDS(tabl_1, 'nvfile') && проверка nv-файла
 fld_2 = AFIELDS(tabl_2, 'nv_et')
 IF 3=2
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(nvItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(nvItem)
*  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
 ENDIF 

* fld_1 = AFIELDS(tabl_1, 'nsfile') && проверка ns-файла
* fld_2 = AFIELDS(tabl_2, 'ns_et')
* IF fld_1 == fld_2 && Кол-во полей совпадает!
*  FieldsIdent = CompFields(nsItem) && 0 - есть отличия, 1 - полное совпадение
*  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
*   =ClDir()
*   RETURN .F.
*  ENDIF 
* ELSE 
*  =DiffFields(nsItem)
*  =CopyToTrash(m.InDir,2)
*  =ClDir()
*  RETURN .F.
* ENDIF 

 fld_1 = AFIELDS(tabl_1, 'rfile') && проверка r-файла
 fld_2 = AFIELDS(tabl_2, 'r_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(rItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(rItem)
*  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

 m.SCompResult = .T.
 fld_1 = AFIELDS(tabl_1, 'sfile') &&& проверка s-файла
 fld_2 = AFIELDS(tabl_2, 's_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(sItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   m.SCompResult = .F.
*   =CopyToTrash(m.InDir,2)
*   =ClDir()
*   RETURN .F.
  ENDIF 
 ELSE 
  m.SCompResult = .F.
*  =DiffFields(sItem)
*  =CopyToTrash(m.InDir,2)
*  =ClDir()
*  RETURN .F.
 ENDIF 

 IF m.SCompResult = .F.
 fld_1 = AFIELDS(tabl_1, 'sfile') &&& проверка s-файла
 fld_2 = AFIELDS(tabl_2, 'sv01_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(sItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(sItem)
*  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
 ENDIF 

 fld_1 = AFIELDS(tabl_1, 'sprfile') && проверка spr-файла
 fld_2 = AFIELDS(tabl_2, 'spr_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(sprItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(sprItem)
*  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

* IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+HOitem)
 IF USED('hofile') AND 1=2
 fld_1 = AFIELDS(tabl_1, 'hofile') && проверка ho-файла
 fld_2 = AFIELDS(tabl_2, 'spr_ho')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(hoItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(hoItem)
*  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
 ENDIF 

 IF USED('onkfile') AND 1=2
 fld_1 = AFIELDS(tabl_1, 'onkfile') && проверка onk_sl-файла
 fld_2 = AFIELDS(tabl_2, 'spr_onk')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(onkItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(onkItem)
  =ClDir()
  RETURN .F.
 ENDIF 
 ENDIF 

RETURN .T. 

FUNCTION CompFields(NameOfFile, para2)
 FOR nFld = 1 TO fld_1
  IF (tabl_1(nFld,1) == tabl_2(nFld,1)) AND ;
     (tabl_1(nFld,2) == tabl_2(nFld,2)) AND ;
     (tabl_1(nFld,3) == tabl_2(nFld,3))
  ELSE 
   RETURN 0 
  ENDIF 
 ENDFOR 
RETURN 1

FUNCTION DiffFields(NameOfFile)
RETURN 
