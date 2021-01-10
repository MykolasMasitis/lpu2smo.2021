PROCEDURE ConsMEE
 IF MESSAGEBOX('СОБРАТЬ БАЗЫ ДЛЯ ЭКСПЕРТОВ?',4+32,'ИНГОССТРАХ-М')=7
  RETURN 
 ENDIF 
 
 *PUBLIC pMeeAll
 pMeeAll = fso.GetParentFolderName(pMee)+'\MEE.ALL'
 CreateAllDir()
 
 CREATE CURSOR dirlist (d_name c(11))
 ScanSubDirs()
 SELECT dirlist
 BROWSE 
 IF RECCOUNT('dirlist')<=0
  USE 
  RETURN 
 ENDIF 
 
 IF OpenAllFiles()>0
  CloseAllFiles()
  RETURN 
 ENDIF 
 
 SELECT dirlist
 SCAN 
  m.d_dir = ALLTRIM(d_name)
  IF !fso.FolderExists(fso.GetParentFolderName(pMee)+'\'+m.d_dir)
   LOOP 
  ENDIF 
  
  WAIT fso.GetParentFolderName(pMee)+'\'+m.d_dir+'...' WINDOW NOWAIT 
  
  m.w_dir = fso.GetParentFolderName(pMee)+'\'+m.d_dir
  m.t_svacts = m.w_dir + '\SVACTS'
  m.t_ssacts = m.w_dir + '\SSACTS'
  m.t_reqs   = m.w_dir + '\REQUESTS'
  m.t_rss    = m.w_dir + '\RSS'
  
  IF fso.FolderExists(m.t_svacts)
   IF fso.FileExists(m.t_svacts+'\svacts.dbf')
    IF OpenFile(m.t_svacts+'\svacts', 't_sv', 'shar')>0
     IF USED('t_sv')
      USE IN t_sv
     ENDIF 
    ELSE 
    
    SELECT t_sv 
    SCAN 
     SCATTER MEMVAR 
     IF SEEK(m.recid, 'svacts')
      LOOP 
     ENDIF 

     INSERT INTO svacts FROM MEMVAR 

    ENDSCAN 
    USE IN t_sv 
     
    ENDIF 
   ENDIF 

   IF fso.FileExists(m.t_svacts+'\moves.dbf')
    IF OpenFile(m.t_svacts+'\moves', 't_svmoves', 'shar')>0
     IF USED('t_svmoves')
      USE IN t_svmoves
     ENDIF 
    ELSE 
    
    SELECT t_svmoves
    RELEASE recid 
    SCAN 
     SCATTER FIELDS EXCEPT recid MEMVAR 
     INSERT INTO svmoves FROM MEMVAR 
    ENDSCAN 
    USE IN t_svmoves
     
    ENDIF 
   ENDIF 

  ENDIF 

  IF fso.FolderExists(m.t_ssacts)
   IF fso.FileExists(m.t_ssacts+'\ssacts.dbf')
    IF OpenFile(m.t_ssacts+'\ssacts', 't_ss', 'shar')>0
     IF USED('t_ss')
      USE IN t_ss
     ENDIF 
    ELSE 
    
    SELECT t_ss
    SCAN 
     SCATTER MEMVAR 
     IF SEEK(m.recid, 'ssacts')
      LOOP 
     ENDIF 
     
     INSERT INTO ssacts FROM MEMVAR 
     
    ENDSCAN 
    USE IN t_ss
     
    ENDIF 
   ENDIF 

   IF fso.FileExists(m.t_ssacts+'\moves.dbf')
    IF OpenFile(m.t_ssacts+'\moves', 't_ssmoves', 'shar')>0
     IF USED('t_ssmoves')
      USE IN t_ssmoves
     ENDIF 
    ELSE 
    
    SELECT t_ssmoves
    RELEASE recid 
    SCAN 
     SCATTER FIELDS EXCEPT recid MEMVAR 
     INSERT INTO ssmoves FROM MEMVAR
    ENDSCAN 
    USE IN t_ssmoves
     
    ENDIF 
   ENDIF 

  ENDIF 

  IF fso.FolderExists(m.t_reqs)
   IF fso.FileExists(m.t_reqs+'\catalog.dbf')
    IF OpenFile(m.t_reqs+'\catalog', 't_cat', 'shar')>0
     IF USED('t_cat')
      USE IN t_cat
     ENDIF 
    ELSE 
    
    SELECT t_cat
    SCAN 
     SCATTER MEMVAR 
     IF SEEK(m.recid, 'catalog')
      LOOP 
     ENDIF 
     
     INSERT INTO catalog FROM MEMVAR 
    ENDSCAN 
    USE IN t_cat
     
    ENDIF 
   ENDIF 
  ENDIF 

  IF fso.FolderExists(m.t_rss)
   IF fso.FileExists(m.t_rss+'\rss.dbf')
    IF OpenFile(m.t_rss+'\rss', 't_rss', 'shar')>0
     IF USED('t_rss')
      USE IN t_rss
     ENDIF 
    ELSE 
    
    SELECT t_rss
    SCAN 
     SCATTER MEMVAR 
     IF SEEK(m.recid, 'rss')
      LOOP 
     ENDIF 

     INSERT INTO rss FROM MEMVAR 

    ENDSCAN 
    USE IN t_rss
     
    ENDIF 
   ENDIF 
  ENDIF 
  
  WAIT CLEAR 

 ENDSCAN 
 USE IN dirlist 
 
 SELECT svacts 
 CALCULATE MAX(recid) TO m.max_id
 ALTER table svacts alter COLUMN recid i AUTOINC NEXTVALUE m.max_id STEP 1
 
 SELECT ssacts 
 CALCULATE MAX(recid) TO m.max_id
 ALTER table ssacts alter COLUMN recid i AUTOINC NEXTVALUE m.max_id STEP 1

 SELECT catalog
 CALCULATE MAX(recid) TO m.max_id
 ALTER table catalog alter COLUMN recid i AUTOINC NEXTVALUE m.max_id STEP 1

 SELECT rss
 CALCULATE MAX(recid) TO m.max_id
 ALTER table rss alter COLUMN recid i AUTOINC NEXTVALUE m.max_id STEP 1

 CloseAllFiles()
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 

FUNCTION CreateAllDir
 IF !fso.FolderExists(pMeeAll)
  fso.CreateFolder(pMeeAll)
 ENDIF 
 IF !fso.FolderExists(pMeeAll)
  MESSAGEBOX('НЕ УДАЛОСЬ СОЗДАТЬ ДИРЕКТОРИЮ'+CHR(13)+CHR(10)+pMeeAll,;
  	0+63,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pMeeAll+'\svacts\svacts.dbf')
  IF !fso.FolderExists(pMeeAll+'\SVACTS')
   fso.CreateFolder(pMeeAll+'\SVACTS')
  ENDIF 
  * Автоинкремент выключен!
  CREATE TABLE &pMeeAll\svacts\svacts ;
  (RecId i, n_rss i, n_akt c(15), period c(6), e_period c(6), flcod c(12), lpu_id n(6), mcod c(7), ;
   codexp n(1), reason c(1), smoexp c(6), docexp c(7), actname c(250), actdate t, n_ss n(3), n_st n(3), ;
   n_dst n(3), n_plk n(3), s_all n(11,2), s_exp n(11,2), s_me n(11,2), et c(1), resume m, conclusion m, ;
   recommend m, qr l, status c(1), n_02 n(3))

  INDEX ON recid TAG recid 
  INDEX FOR qr ON recid TAG qrrecid 
  INDEX ON period TAG period
  INDEX ON e_period TAG e_period
  INDEX ON mcod TAG mcod 
  INDEX ON actdate TAG actdate
  INDEX ON period+e_period+mcod+STR(codexp,1)+docexp TAG unik
  INDEX on status TAG status 
  USE 
 ENDIF 

 IF !fso.FileExists(pMeeAll+'\svacts\moves.dbf')
  * Автоинкремент выключен!
  CREATE TABLE &pMeeAll\svacts\moves ;
  (recid i AUTOINC , actid int, et c(1), usr c(6), dat datetime)
  INDEX on recid TAG recid 
  INDEX on actid TAG actid 
  USE IN moves 
 ENDIF 
 
 IF !fso.FileExists(pMeeAll+'\ssacts\ssacts.dbf')
  IF !fso.FolderExists(pMeeAll+'\SSACTS')
   fso.CreateFolder(pMeeAll+'\SSACTS')
  ENDIF 
  * Автоинкремент выключен!
  CREATE TABLE &pMeeAll\ssacts\ssacts ;
  (RecId i, n_rss i, n_akt c(15), doctyp c(3), period c(6), e_period c(6), flcod c(12), lpu_id n(6), mcod c(7), ;
   codexp n(1), tipacc n(1), isok l, sn_pol c(25), fam c(25), im c(20), ot c(20), actname c(250), actdate t, docexp c(7),;
   s_all n(11,2), s_def n(11,2), s_fee n(11,2), reason c(1), resume m, conclusion m, recommend m, smoexp c(6), qr l, status c(1),;
   n_st n(3), n_dst n(3), n_plk n(3), n_02 n(3))
 
  INDEX ON recid TAG recid 
  INDEX FOR qr ON recid TAG qrrecid 
  INDEX ON period TAG period
  INDEX ON mcod TAG mcod 
  INDEX ON sn_pol TAG sn_pol
  INDEX ON actdate TAG actdate
  INDEX ON PADR(ALLTRIM(fam)+' '+LEFT(im,1)+LEFT(ot,1),28) TAG fio 
  USE 
 ENDIF 

 IF !fso.FileExists(pMeeAll+'\ssacts\moves.dbf')
  * Автоинкремент выключен!
  CREATE TABLE &pMeeAll\ssacts\moves ;
   (recid i AUTOINC , actid int, et c(1), usr c(6), dat datetime)
  INDEX on recid TAG recid 
  INDEX on actid TAG actid 
  USE IN moves 
 ENDIF 

 IF !fso.FolderExists(pMeeAll+'\RSS')
  fso.CreateFolder(pMeeAll+'\RSS')
 ENDIF 

 IF !fso.FileExists(pMeeAll+'\rss\rss.dbf')
  * Автоинкремент выключен!
  CREATE TABLE &pMeeAll\rss\rss ;
   (RecId i, lpu_id n(6), mcod c(7), d_u d, e_period c(6), smoexp c(6), k_acts n(2), s_all n(11,2), s_fee n(11,2))

  INDEX ON mcod+DTOS(d_u) TAG unik
  INDEX ON recid TAG recid 
  INDEX ON e_period TAG e_period
  INDEX ON lpu_id TAG lpu_id
  INDEX ON mcod TAG mcod 
  USE 
 ENDIF 

 IF !fso.FolderExists(pMeeAll+'\REQUESTS')
  fso.CreateFolder(pMeeAll+'\REQUESTS')
 ENDIF 

 IF !fso.FileExists(pMeeAll+'\requests\catalog.dbf')
  CREATE TABLE &pMeeAll\requests\catalog ;
   (RecId i, lpu_id n(6), mcod c(7), d_u d, period c(6), e_period c(6), smoexp c(6), ;
  	supexp c(7), et c(1), rs c(1), n_recs n(3), n_chkd n(3))
  INDEX on recid TAG recid 
  INDEX on mcod+period+et+rs TAG unik 
  USE 
 ENDIF 
RETURN 

FUNCTION ScanSubDirs
 m.oSubDirs = fso.GetFolder(fso.GetParentFolderName(pMee)).SubFolders
 FOR EACH oSubDir IN oSubDirs
  m.d_name = UPPER(ALLTRIM(oSubDir.name))
  IF LEFT(m.d_name,3) != 'MEE'
   LOOP 
  ENDIF 
  IF m.d_name = 'MEE.ALL'
   LOOP 
  ENDIF 
  IF 'MEE' = m.d_name
   LOOP 
  ENDIF 
  
  INSERT INTO dirlist FROM MEMVAR 

 ENDFOR 
 RELEASE oSubDirs
RETURN 

FUNCTION OpenAllFiles
 m.t_r = 0 
 m.t_r = m.t_r + OpenFile(pMeeAll+'\SVACTS\svacts', 'svacts', 'excl', 'recid')
 m.t_r = m.t_r + OpenFile(pMeeAll+'\SVACTS\moves', 'svmoves', 'excl')

 m.t_r = m.t_r + OpenFile(pMeeAll+'\SSACTS\ssacts', 'ssacts', 'excl', 'recid')
 m.t_r = m.t_r + OpenFile(pMeeAll+'\SSACTS\moves', 'ssmoves', 'excl')
 
 m.t_r = m.t_r + OpenFile(pMeeAll+'\REQUESTS\catalog', 'catalog', 'excl', 'recid')
 m.t_r = m.t_r + OpenFile(pMeeAll+'\RSS\rss', 'rss', 'excl', 'recid')
RETURN m.t_r

FUNCTION CloseAllFiles
 IF USED('svacts')
  USE IN svacts
 ENDIF 
 IF USED('svmoves')
  USE IN svmoves
 ENDIF 
 IF USED('ssacts')
  USE IN ssacts
 ENDIF 
 IF USED('ssmoves')
  USE IN ssmoves
 ENDIF 
 IF USED('catalog')
  USE IN catalog
 ENDIF 
 IF USED('rss')
  USE IN rss
 ENDIF 
RETURN 