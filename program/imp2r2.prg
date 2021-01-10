PROCEDURE Imp2R2
 IF MESSAGEBOX('ИМПОРТИРОВАТЬ ДАННЫЕ В ФОРМАТ ВТБ МС?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod)
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+CHR(13)+CHR(10)+;
  	UPPER(m.pBase+'\'+m.gcPeriod),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ '+CHR(13)+CHR(10)+;
  	UPPER(m.pBase+'\'+m.gcPeriod+'\aisom.dbf'),0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  USE IN aisoms
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  USE IN aisoms
  USE IN tarif
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 
 pImpDir = fso.GetParentFolderName(pbin)+'\IMP2R2\'+m.gcPeriod
 IF !fso.FolderExists(pImpDir)
  fso.CreateFolder(pImpDir)
 ENDIF 
 
 IF !fso.FileExists(pTempl+'\pgxxxxmm.dbf')
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ-ШАБЛОН РЕГИСТРА '+CHR(13)+CHR(10)+;
  	UPPER(pTempl+'\pgxxxxmm.dbf'),0+64,'')
  RETURN 
 ENDIF 
 
 m.mmy = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 m.period = m.gcPeriod

 SELECT aisoms
 SCAN 
  m.mcod  = mcod 
  m.IsVed = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
  m.lpuid = lpuid
  m.cokr  = cokr
  
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  
  IF m.mcod = '0371001' && скорая помощь

  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\People.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\Talon.dbf')
   LOOP 
  ENDIF 
  
  ** Модуль заполнения pg - регистра
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'share')>0
   IF USED('people')
    USE IN peoople
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  m.pg_file = 'pg'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\pgxxxxmm.dbf', pImpDir+'\'+m.pg_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.pg_file, 'pg', 'share')>0
   IF USED('pg')
    USE IN pg
   ENDIF 
   USE IN people
   SELECT aisoms
   LOOP 
  ENDIF 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'ctrl', 'shar', 'rrid')>0
    IF USED('ctrl')
     USE IN Ctrl
    ENDIF 
   ENDIF 
  ENDIF 
  
  WAIT m.mcod + '...' WINDOW NOWAIT 
  
  SELECT people
  IF USED('ctrl')
   SET RELATION TO recid INTO ctrl
  ENDIF 
  SCAN 
   IF FIELD('prmcods')='PRMCODS'
    SCATTER FIELDS sn_pol, qq, enp, fam, im, ot, w, dr, d_type, sv, prmcod, prmcods MEMVAR
    m.priks = IIF(SEEK(m.prmcods, 'sprlpu'), sprlpu.lpu_id, 0)
    m.LPU_ID_ERS = IIF(SEEK(m.prmcods, 'sprlpu'), sprlpu.lpu_id, 0)
   ELSE 
    SCATTER FIELDS sn_pol, qq, enp, fam, im, ot, w, dr, d_type, sv, prmcod MEMVAR
   ENDIF 
   
   m.tip_p = tipp
   m.prik  = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   
   m.recid = recid_lpu
   m.LPU_ID_ERZ = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   
   IF USED('ctrl')
    m.koder = ctrl.c_err
   ENDIF 
   INSERT INTO pg FROM MEMVAR
  ENDSCAN 
  IF USED('ctrl')
   SET RELATION OFF INTO ctrl
   SELECT ctrl 
   USE IN ctrl
  ENDIF 
  USE IN people
  USE IN pg 
  
  ** Модуль заполнения pg - регистра

  ** Модуль заполнения rg - талона

  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\Talon', 'talon', 'share')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  m.rg_file = 'rg'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\rgxxxxmm.dbf', pImpDir+'\'+m.rg_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.rg_file, 'rg', 'share')>0
   IF USED('rg')
    USE IN rg
   ENDIF 
   USE IN talon
   SELECT aisoms
   LOOP 
  ENDIF 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'ctrl', 'excl', 'rid')>0
    IF USED('ctrl')
     USE IN Ctrl
    ENDIF 
   ENDIF 
  ENDIF 
  
  SELECT talon
  IF USED('ctrl')
   SET RELATION TO recid INTO ctrl
  ENDIF 
  SCAN 
   SCATTER  FIELDS EXCEPT recid MEMVAR 
   m.lpu_id = 4708
   m.recid = recid_lpu
   *m.s_all  = fsumm(m.cod, m.tip, m.k_u, m.IsVed)
   *m.c_i    = ALLTRIM(pst)+'#'+ALLTRIM(c_br)+'#'+ALLTRIM(n_u)
   m.pst  = SUBSTR(m.c_i, 1, AT('#',m.c_i,1)-1)
   m.c_br = SUBSTR(m.c_i, AT('#',m.c_i,1)+1, AT('#',m.c_i,2)-AT('#',m.c_i,1)-1)
   m.n_u  = SUBSTR(m.c_i, AT('#',m.c_i,2)+1)
   
   m.c_i = m.n_u
   
   m.tar = m.s_all

   IF VARTYPE(m.lpu_ord)='C'
    m.lpu_ord = INT(VAL(m.lpu_ord))
   ENDIF 
   
   IF USED('ctrl')
    m.koder = ctrl.c_err
   ENDIF 
   INSERT INTO rg FROM MEMVAR
  ENDSCAN 
  IF USED('ctrl')
   SET RELATION OFF INTO ctrl
   SELECT ctrl 
   USE 
  ENDIF 
  USE IN talon
  USE IN rg 
  
  RELEASE C_BR, PROFBR, T_UZ, T_UP, TAR, PST
  
  ** Модуль заполнения rg - талона

  ELSE && IF m.mcod != '0371001' && скорая помощь

  *m.bfile = 'b'+m.mcod+'.'+m.mmy
  m.bfile = 'd'+m.qcod+STR(m.lpuid,4)+'.'+m.mmy
  
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.bfile)
   LOOP 
  ENDIF 
  ffile = fso.GetFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.bfile)
  IF ffile.size >= 2
   fhandl = ffile.OpenAsTextStream
   lcHead = fhandl.Read(2)
  ELSE 
   lcHead = ''
  ENDIF 
  fhandl.Close

  ** Модуль заполнения pg - регистра

  IF lcHead == 'PK' && Это zip-файл!
   UnzipOpen(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.bfile)
   *rItem  = 'R' + m.qcod + '.' + m.mmy
   *sItem  = 'S' + m.qcod + '.' + m.mmy
   rItem  = 'R' + m.qcod + 'Y.' + m.mmy
   sItem  = 'S' + m.qcod + 'Y.' + m.mmy
   nvItem  = 'NV' + STR(m.lpuid,4) + '.' + m.mmy
   hoItem  = 'HO' + m.qcod + '.' + m.mmy

   slItem   = 'ONK_SL' + m.qcod + '.' + m.mmy
   uslItem  = 'ONK_USL' + m.qcod + '.' + m.mmy
   lsItem   = 'ONK_LS' + m.qcod + '.' + m.mmy
   dsItem   = 'ONK_DIAG' + m.qcod + '.' + m.mmy
   consItem = 'ONK_CONS' + m.qcod + '.' + m.mmy
   ptItem   = 'ONK_PROT' + m.qcod + '.' + m.mmy
   npItem   = 'ONK_NAPR_V_OUT' + m.qcod + '.' + m.mmy

   cvItem   = 'CV_LS' + m.qcod + '.' + m.mmy

   IF UnzipGotoFileByName(rItem)
    llIsOneZip = .t.
   ELSE 
    llIsOneZip = .f.
   ENDIF 
   UnzipClose()
  ENDIF 
 
  IF llIsOneZip = .f.
   LOOP 
  ENDIF 

  ZipName = m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.bfile
  ZipDir  = m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'

  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rItem)
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rItem)
  ENDIF  

  UnzipOpen(ZipName)

  UnzipGotoFileByName(rItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rItem)
   SELECT aisoms
   LOOP 
  ENDIF  

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rItem, 'rfile', 'share')>0
   IF USED('rfile')
    USE IN rfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rItem)
   SELECT aisoms
   LOOP 
  ENDIF 
  
  m.pg_file = 'pg'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\pgxxxxmm.dbf', pImpDir+'\'+m.pg_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.pg_file, 'pg', 'share')>0
   IF USED('pg')
    USE IN pg
   ENDIF 
   USE IN rfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rItem)
   SELECT aisoms
   LOOP 
  ENDIF 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\Ctrl'+m.qcod+'.dbf')
   IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\Ctrl'+m.qcod, 'ctrl', 'excl')>0
    IF USED('ctrl')
     USE IN Ctrl
    ENDIF 
   ELSE 
    SELECT Ctrl
    INDEX ON recid FOR UPPER(LEFT(file,1))='R' TAG recid 
    SET ORDER TO recid
   ENDIF 
  ENDIF 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people.dbf')
   IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'shar', 'recid_lpu')>0
    IF USED('people')
     USE IN people
    ENDIF 
   ELSE 
   ENDIF 
  ENDIF 
  
  WAIT m.mcod + '...' WINDOW NOWAIT 
  
  SELECT rfile
  IF USED('ctrl')
   SET RELATION TO recid INTO ctrl
  ENDIF 
  IF USED('people')
   SET RELATION TO recid INTO people
  ENDIF 
  SCAN 
   SCATTER MEMVAR 
   IF USED('people')
    m.prmcod  = people.prmcod
    IF FIELD('prmcods','people')='PRMCODS'
     m.prmcods = people.prmcods
     m.LPU_ID_ERS = IIF(SEEK(m.prmcods, 'sprlpu'), sprlpu.lpu_id, 0)
    ENDIF 
    m.LPU_ID_ERZ = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   ENDIF 
   *IF USED('ctrl')
   * m.koder = ctrl.errors
   *ENDIF 
   m.koder = er_c
   INSERT INTO pg FROM MEMVAR
  ENDSCAN 
  IF USED('ctrl')
   SET RELATION OFF INTO ctrl
   SELECT ctrl 
   SET ORDER to 
   DELETE TAG ALL 
   USE IN ctrl
  ENDIF 
  IF USED('people')
   SELECT rfile
   SET RELATION OFF INTO people
   USE IN people
  ENDIF 
  USE IN rfile
  USE IN pg 
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rItem)
  
  ** Модуль заполнения pg - регистра

  ** Модуль заполнения rg - талона
  UnzipOpen(ZipName)

  UnzipGotoFileByName(sItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.sItem)
   SELECT aisoms
   LOOP 
  ENDIF  

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.sItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.sItem, 'sfile', 'share')>0
   IF USED('sfile')
    USE IN sfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.sItem)
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\Talon', 'talon', 'share', 'recid_lpu')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   USE IN sfile
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'share', 'sn_pol')>0
   USE IN talon
   IF USED('people')
    USE IN people
   ENDIF 
   USE IN sfile
   SELECT aisoms
   LOOP 
  ENDIF 
  SELECT talon 
  SET RELATION TO sn_pol INTO people 
  
  m.rg_file = 'rg'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\rgxxxxmm.dbf', pImpDir+'\'+m.rg_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.rg_file, 'rg', 'share')>0
   IF USED('rg')
    USE IN rg
   ENDIF 
   USE IN sfile
   USE IN talon
   USE IN people
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.sItem)
   SELECT aisoms
   LOOP 
  ENDIF 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\Ctrl'+m.qcod+'.dbf')
   IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\Ctrl'+m.qcod, 'ctrl', 'excl')>0
    IF USED('ctrl')
     USE IN Ctrl
    ENDIF 
   ELSE 
    SELECT Ctrl
    INDEX ON recid FOR UPPER(LEFT(file,1))='S' TAG recid 
    SET ORDER TO recid
   ENDIF 
  ENDIF 
  
  SELECT sfile
  SET RELATION TO recid INTO talon ADDITIVE 
  IF USED('ctrl')
   SET RELATION TO recid INTO ctrl ADDITIVE 
  ENDIF
  SCAN 
   SCATTER MEMVAR 
   IF USED('people')
    m.prmcod  = people.prmcod
    IF FIELD('prmcods','people')='PRMCODS'
     m.prmcods = people.prmcods
     m.LPU_ID_ERS = IIF(SEEK(m.prmcods, 'sprlpu'), sprlpu.lpu_id, 0)
    ENDIF 
    m.LPU_ID_ERZ = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   ENDIF 
   
   
   IF INT(VAL(m.gcperiod)) > 201712
    *m.s_all  = FSumm(m.cod, m.tip, IIF(BETWEEN(m.cod,97107,97158), m.kd_fact, m.k_u), m.IsVed)
    m.tarif = IIF(FIELD('s_lek', 'talon')='S_LEK', talon.s_lek, 0)
    m.s_all = talon.s_all + IIF(FIELD('s_lek', 'talon')='S_LEK', talon.s_lek, 0)
   ELSE 
    m.s_all  = FSummVeryOld(m.cod, m.tip, m.k_u, m.IsVed)
   ENDIF 
  
   *IF USED('ctrl')
   * m.koder = ctrl.errors
   *ENDIF 
   m.koder = er_c
   IF VARTYPE(m.lpu_ord)='C'
    m.lpu_ord = INT(VAL(m.lpu_ord))
   ENDIF 

   INSERT INTO rg FROM MEMVAR
  ENDSCAN 
  SET RELATION OFF INTO talon 
  IF USED('ctrl')
   SET RELATION OFF INTO ctrl
   SELECT ctrl 
   SET ORDER to 
   DELETE TAG ALL 
   USE 
  ENDIF 
  USE IN sfile
  USE IN talon 
  USE IN people
  USE IN rg 
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.sItem)

  ** Модуль заполнения rg - талона

  ** Модуль заполнения nv
  UnzipOpen(ZipName)

  UnzipGotoFileByName(nvItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.nvItem)
   SELECT aisoms
   LOOP 
  ENDIF  

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.nvItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.nvItem, 'nvfile', 'share')>0
   IF USED('nvfile')
    USE IN nvfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.nvItem)
   SELECT aisoms
   LOOP 
  ENDIF 
  
  m.nv_file = 'nv'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\nvxxxxmm.dbf', pImpDir+'\'+m.nv_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.nv_file, 'nv', 'share')>0
   IF USED('nv')
    USE IN nv
   ENDIF 
   USE IN nvfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.nvItem)
   SELECT aisoms
   LOOP 
  ENDIF 
  
  SELECT nvfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO nv FROM MEMVAR
  ENDSCAN 
  USE IN nvfile
  USE IN nv 
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.nvItem)

  ** Модуль заполнения nv

  ** Модуль заполнения ho
  UnzipOpen(ZipName)

  UnzipGotoFileByName(hoItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.hoItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.hoItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.hoItem, 'hofile', 'share')>0
   IF USED('hofile')
    USE IN hofile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.hoItem)
   EXIT 
  ENDIF 
  
  m.ho_file = 'op'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\opxxxxmm.dbf', pImpDir+'\'+m.ho_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.ho_file, 'ho', 'share')>0
   IF USED('ho')
    USE IN ho
   ENDIF 
   USE IN hofile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.hoItem)
   EXIT 
  ENDIF 
  
  SELECT hofile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO ho FROM MEMVAR
  ENDSCAN 
  USE IN hofile
  USE IN ho 
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.hoItem)
  
  ENDIF 

  ** Модуль заполнения ho

  ** Модуль заполнения sl
  UnzipOpen(ZipName)

  UnzipGotoFileByName(slItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.slItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.slItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.slItem, 'slfile', 'share')>0
   IF USED('slfile')
    USE IN slfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.slItem)
   EXIT 
  ENDIF 
  
  m.sl_file = 'sl'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\slxxxxmm.dbf', pImpDir+'\'+m.sl_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.sl_file, 'sl', 'share')>0
   IF USED('sl')
    USE IN sl
   ENDIF 
   USE IN slfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.slItem)
   EXIT 
  ENDIF 
  
  SELECT slfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO sl FROM MEMVAR
  ENDSCAN 
  USE IN slfile
  USE IN sl 
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.slItem)
  
  ENDIF 

  ** Модуль заполнения sl

  ** Модуль заполнения usl
  UnzipOpen(ZipName)

  UnzipGotoFileByName(uslItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.uslItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.uslItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.uslItem, 'uslfile', 'share')>0
   IF USED('uslfile')
    USE IN uslfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.uslItem)
   EXIT 
  ENDIF 
  
  m.usl_file = 'us'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\usxxxxmm.dbf', pImpDir+'\'+m.usl_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.usl_file, 'usl', 'share')>0
   IF USED('usl')
    USE IN usl
   ENDIF 
   USE IN uslfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.uslItem)
   EXIT 
  ENDIF 
  
  SELECT uslfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO usl FROM MEMVAR
  ENDSCAN 
  USE IN uslfile
  USE IN usl 
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.uslItem)
  
  ENDIF 

  ** Модуль заполнения usl

  ** Модуль заполнения ls
  UnzipOpen(ZipName)

  UnzipGotoFileByName(lsItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.lsItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.lsItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.lsItem, 'lsfile', 'share')>0
   IF USED('lsfile')
    USE IN lsfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.lsItem)
   EXIT 
  ENDIF 
  
  m.ls_file = 'ls'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\lsxxxxmm.dbf', pImpDir+'\'+m.ls_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.ls_file, 'ls', 'share')>0
   IF USED('ls')
    USE IN ls
   ENDIF 
   USE IN lsfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.lsItem)
   EXIT 
  ENDIF 
  
  SELECT lsfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO ls FROM MEMVAR
  ENDSCAN 
  USE IN lsfile
  USE IN ls 
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.lsItem)
  
  ENDIF 

  ** Модуль заполнения ls

  ** Модуль заполнения cons
  UnzipOpen(ZipName)

  UnzipGotoFileByName(consItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.consItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.consItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.consItem, 'consfile', 'share')>0
   IF USED('consfile')
    USE IN consfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.consItem)
   EXIT 
  ENDIF 
  
  m.cons_file = 'cn'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\cnxxxxmm.dbf', pImpDir+'\'+m.cons_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.cons_file, 'cn', 'share')>0
   IF USED('cn')
    USE IN cn
   ENDIF 
   USE IN consfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.consItem)
   EXIT 
  ENDIF 
  
  SELECT consfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO cn FROM MEMVAR
  ENDSCAN 
  USE IN consfile
  USE IN cn
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.consItem)
  
  ENDIF 

  ** Модуль заполнения cons

  ** Модуль заполнения diag
  UnzipOpen(ZipName)

  UnzipGotoFileByName(dsItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dsItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dsItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dsItem, 'dsfile', 'share')>0
   IF USED('dsfile')
    USE IN dsfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dsItem)
   EXIT 
  ENDIF 
  
  m.ds_file = 'dg'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\dgxxxxmm.dbf', pImpDir+'\'+m.ds_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.ds_file, 'dg', 'share')>0
   IF USED('dg')
    USE IN dg
   ENDIF 
   USE IN dsfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dsItem)
   EXIT 
  ENDIF 
  
  SELECT dsfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO dg FROM MEMVAR
  ENDSCAN 
  USE IN dsfile
  USE IN dg
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dsItem)
  
  ENDIF 

  ** Модуль заполнения diag

  ** Модуль заполнения pt
  UnzipOpen(ZipName)

  UnzipGotoFileByName(ptItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ptItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ptItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ptItem, 'ptfile', 'share')>0
   IF USED('ptfile')
    USE IN ptfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ptItem)
   EXIT 
  ENDIF 
  
  m.pt_file = 'pt'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\ptxxxxmm.dbf', pImpDir+'\'+m.pt_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.pt_file, 'pt', 'share')>0
   IF USED('pt')
    USE IN pt
   ENDIF 
   USE IN ptfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ptItem)
   EXIT 
  ENDIF 
  
  SELECT ptfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO pt FROM MEMVAR
  ENDSCAN 
  USE IN ptfile
  USE IN pt
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ptItem)
  
  ENDIF 
  ** Модуль заполнения pt
  
  ** Модуль заполнения np && napr_v_out
  UnzipOpen(ZipName)

  UnzipGotoFileByName(npItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.npItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.npItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.npItem, 'npfile', 'share')>0
   IF USED('npfile')
    USE IN npfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.npItem)
   EXIT 
  ENDIF 
  
  m.np_file = 'np'+STR(m.lpuid,4)+PADL(m.tMonth,2,'0')
  fso.CopyFile(pTempl+'\npxxxxmm.dbf', pImpDir+'\'+m.np_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.np_file, 'np', 'share')>0
   IF USED('np')
    USE IN np
   ENDIF 
   USE IN npfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.npItem)
   EXIT 
  ENDIF 
  
  SELECT npfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO np FROM MEMVAR
  ENDSCAN 
  USE IN npfile
  USE IN np
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.npItem)
  
  ENDIF 

  ** Модуль заполнения np && napr_v_out
  
  ** Модуль заполнения cv_ls
  UnzipOpen(ZipName)

  UnzipGotoFileByName(cvItem)
  UnzipFile(ZipDir)
  UnzipClose()
 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.cvItem)

  oSettings.CodePage(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.cvItem, 866, .t.)
 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.cvItem, 'cvfile', 'share')>0
   IF USED('cvfile')
    USE IN cvfile
   ENDIF 
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.cvItem)
   EXIT 
  ENDIF 
  
  m.cv_file = 'cvls'+STR(m.lpuid,4)
  fso.CopyFile(pTempl+'\cvlsxxxx.dbf', pImpDir+'\'+m.cv_file+'.dbf')
  
  IF OpenFile(pImpDir+'\'+m.cv_file, 'cv', 'share')>0
   IF USED('cv')
    USE IN cv
   ENDIF 
   USE IN cvfile
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.cvItem)
   EXIT 
  ENDIF 
  
  SELECT cvfile
  SCAN 
   SCATTER MEMVAR 
   INSERT INTO cv FROM MEMVAR
  ENDSCAN 
  USE IN cvfile
  USE IN cv 
  
  fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.cvItem)
  
  ENDIF 

  ** Модуль заполнения ls

  ENDIF && if m.mcod = '0371001'

  WAIT CLEAR 
  
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms 
 USE IN tarif
 USE IN sprlpu
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 

FUNCTION MagePG(para1) && Модуль создания регистра
 LOCAL m.mcod
 m.mcod = para1
RETURN .T. 