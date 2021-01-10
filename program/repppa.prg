PROCEDURE RepPPA
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ ÏÎ PPA?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\ppa.xls')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ ÎÒ×ÅÒÀ PPA.XLS',0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN aisoms 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curss (mcod c(7), moname c(40), ved c(3), s_pred n(13,2), s_def n(13,2), ;
 	s_ok n(13,2), n_pp n(6))
 
 SELECT aisoms 
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('err')
    USE IN err
   ENDIF 
   USE IN talon
   SELECT aisoms
   LOOP 
  ENDIF 
  
  SELECT aisoms
  
  m.moname = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.name, '')
  m.ved    = IIF(SEEK(m.mcod, 'sprlpu'), STR(sprlpu.prn_kodved,3), '')
  m.s_pred = s_pred + s_lek
  m.s_def  = sum_flk + ls_flk
  m.s_ok   = m.s_pred - m.s_def
  
  SELECT err 
  COUNT FOR c_err='PPA' TO m.n_pp
  USE
  
  USE IN talon 
  
  INSERT INTO curss FROM MEMVAR 
  
  SELECT aisoms 
  
 ENDSCAN 
 USE 
 USE IN sprlpu
 
 m.llResult = X_Report(pTempl+'\ppa.xls', pBase+'\'+m.gcPeriod+'\ppa.xls', .T.)
 USE IN curss
 
RETURN 