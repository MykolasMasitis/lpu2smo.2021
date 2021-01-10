PROCEDURE ConsBase
 IF MESSAGEBOX('—ќЅ–ј“№ Ѕј«џ ƒЋя Ё —ѕ≈–“ќ¬?',4+32,'BASE')=7
  RETURN 
 ENDIF 
 
 CREATE CURSOR BaseDirList (d_name c(11))
 ScanSubDirs()
 SELECT BaseDirList

 BROWSE 
 
 IF RECCOUNT('BaseDirList')<=0
  USE 
  RETURN 
 ENDIF 
 
 SELECT BaseDirList
 SCAN 
  m.BaseDir = ALLTRIM(d_name)
  IF !fso.FolderExists(fso.GetParentFolderName(pBase)+'\'+m.BaseDir)
   LOOP 
  ENDIF 
  
  WAIT fso.GetParentFolderName(pBase)+'\'+m.BaseDir+'...' WINDOW NOWAIT 
  
  m.PeriodDirs = fso.GetFolder(fso.GetParentFolderName(pBase)+'\'+m.BaseDir).SubFolders
  FOR EACH PeriodDir IN PeriodDirs
   m.p_name = UPPER(ALLTRIM(PeriodDir.name))
   IF LEN(m.p_name)!=6
    LOOP 
   ENDIF 
   IF SUBSTR(m.p_name,1,2)!='20'
    LOOP 
   ENDIF 
   IF !INLIST(SUBSTR(m.p_name,3,2),'19','18','17','16','15','14','13','12')
    LOOP 
   ENDIF 
   IF !INLIST(SUBSTR(m.p_name,5,2),'01','02','03','04','05','06','07','08','09','10','11','12')
    LOOP 
   ENDIF 
   
   m.lcPeriod = fso.GetParentFolderName(pBase)+'\'+m.BaseDir+'\'+m.p_name 
   IF !fso.FolderExists(m.lcPeriod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(m.lcPeriod+'\aisoms.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(m.lcPeriod+'\aisoms', 'aisoms', 'shar')>0
    IF USED('aisoms')
     USE IN aisoms
    ENDIF 
    LOOP 
   ENDIF 
   
   SELECT aisoms 
   SCAN 
    m.mcod = mcod 
    IF !fso.FolderExists(m.lcPeriod+'\'+m.mcod)
     LOOP 
    ENDIF 
    IF !fso.FileExists(m.lcPeriod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
     LOOP 
    ENDIF 
    IF OpenFile(m.lcPeriod+'\'+m.mcod+'\m'+m.mcod, 'merr', 'shar')>0
     IF USED('merr')
      USE IN merr
     ENDIF 
     SELECT aisoms
     LOOP 
    ENDIF 
    IF RECCOUNT('merr')<=0 
     USE IN merr
     SELECT aisoms 
     LOOP 
    ENDIF 
    
    * ќткрыть аналогичную директорию в главной базе!
    IF !fso.FolderExists(m.pBase+'\'+m.p_name+'\'+m.mcod)
     USE IN merr
     SELECT aisoms
     LOOP 
    ENDIF 
    IF !fso.FileExists(m.pBase+'\'+m.p_name+'\'+m.mcod+'\m'+m.mcod+'.dbf')
     USE IN merr
     SELECT aisoms
     LOOP 
    ENDIF 
    IF OpenFile(m.pBase+'\'+m.p_name+'\'+m.mcod+'\m'+m.mcod, 'gc_merr', 'shar', 'unik')>0
     IF USED('gc_merr')
      USE IN gc_merr
     ENDIF 
     USE IN merr
     SELECT aisoms
     LOOP 
    ENDIF 
    * ќткрыть аналогичную директорию в главной базе!
    
    SELECT merr
    SCAN 
     SCATTER FIELDS EXCEPT rid MEMVAR 
     m.k_key = PADL(M.RECID,6,"0")+M.ET+M.DOCEXP+M.REASON+LEFT(M.ERR_MEE,2)
     IF !SEEK(m.k_key, 'gc_merr')
      INSERT INTO gc_merr FROM MEMVAR 
     ENDIF 
    ENDSCAN 
    USE IN merr 
    USE IN gc_merr
    *MESSAGEBOX(m.mcod, 0+64, m.lcPeriod)
    
   ENDSCAN 
   USE IN aisoms 
   
   *MESSAGEBOX(m.lcPeriod, 0+64, m.BaseDir)
  
  ENDFOR 
  RELEASE PeriodDirs
  
  WAIT CLEAR 

 ENDSCAN 
 USE IN BaseDirList 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 

FUNCTION ScanSubDirs
 m.oSubDirs = fso.GetFolder(fso.GetParentFolderName(pBase)).SubFolders
 FOR EACH oSubDir IN oSubDirs
  m.d_name = UPPER(ALLTRIM(oSubDir.name))
  IF LEFT(m.d_name,4) != 'BASE'
   LOOP 
  ENDIF 
  IF 'BASE' = m.d_name
   LOOP 
  ENDIF 
  
  INSERT INTO BaseDirList FROM MEMVAR 

 ENDFOR 
 RELEASE oSubDirs
RETURN 

