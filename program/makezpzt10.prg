PROCEDURE MakeZPZT10
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ ÏÎ ÔÎÐÌÅ ÇÏÇ?',4+32,'ÒÀÁËÈÖÀ 10')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\zpz_t10.xls')
  MESSAGEBOX('ÎÒÑÓÒÑÂÓÅÒ ÔÀÉË '+UPPER(pTempl+'\zpz_t10.xls'),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\sookodxx', 'sookod', 'shar', 'er_c')>0
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 
 DIMENSION dimdata(20,10)
 dimdata = 0

 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  m.IsPuchok = IIF(m.mcod = '0371001', .T., .F.)
  
  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'er', 'shar', 'rid')>0
   IF USED('er')
    USE IN er
   ENDIF 
   USE IN talon 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  SELECT er
  SET RELATION TO LEFT(c_err,2) INTO sookod
  SELECT talon 
  SET RELATION TO recid INTO er ADDITIVE 
  SCAN 
   m.otd  = SUBSTR(otd,2,2)
   m.cod  = cod
   m.ds   = ds
   m.ds_2 = ds_2
   m.s_all = s_all
   
   m.IsOnk = IIF(INLIST(SUBSTR(otd,4,3),'018','060'), .T., .F.)
   	
   m.IsErr = IIF(!EMPTY(er.c_err), .T., .F.)
   m.osn230 = sookod.osn230
   
   dimdata(1,5) = dimdata(1,5) + m.s_all
   dimdata(3,5) = dimdata(3,5) + IIF(m.IsErr, m.s_all, 0)
   dimdata(4,5) = dimdata(4,5) + IIF(m.IsOnk AND m.IsErr, m.s_all, 0)
   
  ENDSCAN 
  SELECT talon 
  SET RELATION OFF INTO er
  SELECT er
  SET RELATION OFF INTO sookod
  USE IN talon 
  USE IN er
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 USE IN sookod
 
 CREATE CURSOR curdata (recid i)
 INSERT INTO curdata (recid) VALUES (0)

 m.llResult = X_Report(pTempl+'\zpz_t10.xls', pBase+'\'+m.gcperiod+'\zpz_t10.xls', .T.)
 
RETURN 