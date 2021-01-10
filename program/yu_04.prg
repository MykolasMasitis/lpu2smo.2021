PROCEDURE yu_04
 IF MESSAGEBOX('ÑÔÎÐÌÈÎÂÀÒÜ ÎÒ×ÅÒ Þ-04?',4+16,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\yu_04.xls')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ'*pTempl+'\yu_04.xls')
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
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\profot', 'profot', 'share', 'otd')>0
  USE IN aisoms 
  IF USED('profot')
   USE IN profot
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'share', 'cod')>0
  USE IN aisoms 
  USE IN profot
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 

 
 DIMENSION dimdata(10,7)
 dimdata = 0
 
 SELECT aisoms
 SCAN
  m.mcod = mcod
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
   SELECT aiosms 
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT talon 
  SCAN  
   m.cod   = cod  
   m.otd   = otd 
   m.k_u   = k_u 
   m.s_all = s_all+s_lek
   m.ds    = ds 

   m.usl_ok  = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), INT(VAL(profot.usl_ok)), 0)
   m.isonk = IIF(INLIST(SUBSTR(m.otd,4,3),'018','060'),.T.,.F.)
   m.iscovid = IIF(INLIST(m.ds, 'B34.2','J02','J04','J06','J20','U07.1','U07.2') OR ;
  	 	BETWEEN(LEFT(m.ds,3),'J09','J18') ,.T., .F.)
   m.isdsp = IIF(SEEK(m.cod,'dspcodes') AND INLIST(dspcodes.tip,1,3), .T., .F.)
   m.ispmo = IIF(SEEK(m.cod,'dspcodes') AND INLIST(dspcodes.tip,2,4), .T., .F.)
   
   m.isvmp = IIF(INLIST(FLOOR(m.Cod/1000), 200, 297, 300, 397), .T., .F.)
   
   dimdata(1,1) = dimdata(1,1) + m.s_all
   dimdata(2,1) = dimdata(2,1) + IIF(m.isonk, m.s_all, 0)
   dimdata(3,1) = dimdata(3,1) + IIF(m.isdsp, m.s_all, 0)
   dimdata(4,1) = dimdata(4,1) + IIF(m.ispmo, m.s_all, 0)
   dimdata(5,1) = dimdata(5,1) + IIF(m.iscovid, m.s_all, 0)

   dimdata(6,1)  = dimdata(6,1) + m.k_u
   dimdata(7,1)  = dimdata(7,1) + IIF(m.isonk, m.k_u, 0)
   dimdata(8,1)  = dimdata(8,1) + IIF(m.isdsp, m.k_u, 0)
   dimdata(9,1)  = dimdata(9,1) + IIF(m.ispmo, m.k_u, 0)
   dimdata(10,1) = dimdata(10,1) + IIF(m.iscovid, m.k_u, 0)

   IF m.mcod='0371001'
    dimdata(1,7) = dimdata(1,7) + m.s_all
    dimdata(2,7) = dimdata(2,7) + IIF(m.isonk, m.s_all, 0)
    dimdata(3,7) = dimdata(3,7) + IIF(m.isdsp, m.s_all, 0)
    dimdata(4,7) = dimdata(4,7) + IIF(m.ispmo, m.s_all, 0)
    dimdata(5,7) = dimdata(5,7) + IIF(m.iscovid, m.s_all, 0)

    dimdata(6,7) = dimdata(6,7) + m.k_u
    dimdata(7,7) = dimdata(7,7) + IIF(m.isonk, m.k_u, 0)
    dimdata(8,7) = dimdata(8,7) + IIF(m.isdsp, m.k_u, 0)
    dimdata(9,7) = dimdata(9,7) + IIF(m.ispmo, m.k_u, 0)
    dimdata(10,7) = dimdata(10,7) + IIF(m.iscovid, m.k_u, 0)
   ELSE 
    dimdata(1,2) = dimdata(1,2) + IIF(m.usl_ok=3, m.s_all, 0)
    dimdata(2,2) = dimdata(2,2) + IIF(m.usl_ok=3 and m.isonk, m.s_all, 0)
    dimdata(3,2) = dimdata(3,2) + IIF(m.usl_ok=3 and m.isdsp, m.s_all, 0)
    dimdata(4,2) = dimdata(4,2) + IIF(m.usl_ok=3 and m.ispmo, m.s_all, 0)
    dimdata(5,2) = dimdata(5,2) + IIF(m.usl_ok=3 and m.iscovid, m.s_all, 0)

    dimdata(1,3) = dimdata(1,3) + IIF(m.usl_ok=2, m.s_all, 0)
    dimdata(2,3) = dimdata(2,3) + IIF(m.usl_ok=2 and m.isonk, m.s_all, 0)
    dimdata(3,3) = dimdata(3,3) + IIF(m.usl_ok=2 and m.isdsp, m.s_all, 0)
    dimdata(4,3) = dimdata(4,3) + IIF(m.usl_ok=2 and m.ispmo, m.s_all, 0)
    dimdata(5,3) = dimdata(5,3) + IIF(m.usl_ok=2 and m.iscovid, m.s_all, 0)

    dimdata(1,4) = dimdata(1,4) + IIF(m.usl_ok=2 and m.isvmp, m.s_all, 0)
    dimdata(2,4) = dimdata(2,4) + IIF(m.usl_ok=2 and m.isvmp and m.isonk, m.s_all, 0)
    dimdata(3,4) = dimdata(3,4) + IIF(m.usl_ok=2 and m.isvmp and m.isdsp, m.s_all, 0)
    dimdata(4,4) = dimdata(4,4) + IIF(m.usl_ok=2 and m.isvmp and m.ispmo, m.s_all, 0)
    dimdata(5,4) = dimdata(5,4) + IIF(m.usl_ok=2 and m.isvmp and m.iscovid, m.s_all, 0)

    dimdata(1,5) = dimdata(1,5) + IIF(m.usl_ok=1, m.s_all, 0)
    dimdata(2,5) = dimdata(2,5) + IIF(m.usl_ok=1 and m.isonk, m.s_all, 0)
    dimdata(3,5) = dimdata(3,5) + IIF(m.usl_ok=1 and m.isdsp, m.s_all, 0)
    dimdata(4,5) = dimdata(4,5) + IIF(m.usl_ok=1 and m.ispmo, m.s_all, 0)
    dimdata(5,5) = dimdata(5,5) + IIF(m.usl_ok=1 and m.iscovid, m.s_all, 0)

    dimdata(1,6) = dimdata(1,6) + IIF(m.usl_ok=1 and m.isvmp, m.s_all, 0)
    dimdata(2,6) = dimdata(2,6) + IIF(m.usl_ok=1 and m.isvmp and m.isonk, m.s_all, 0)
    dimdata(3,6) = dimdata(3,6) + IIF(m.usl_ok=1 and m.isvmp and m.isdsp, m.s_all, 0)
    dimdata(4,6) = dimdata(4,6) + IIF(m.usl_ok=1 and m.isvmp and m.ispmo, m.s_all, 0)
    dimdata(5,6) = dimdata(5,6) + IIF(m.usl_ok=1 and m.isvmp  and m.iscovid, m.s_all, 0)

    dimdata(6,2) = dimdata(6,2) + IIF(m.usl_ok=3, m.k_u, 0)
    dimdata(7,2) = dimdata(7,2) + IIF(m.usl_ok=3 and m.isonk, m.k_u, 0)
    dimdata(8,2) = dimdata(8,2) + IIF(m.usl_ok=3 and m.isdsp, m.k_u, 0)
    dimdata(9,2) = dimdata(9,2) + IIF(m.usl_ok=3 and m.ispmo, m.k_u, 0)
    dimdata(10,2) = dimdata(10,2) + IIF(m.usl_ok=3 and m.iscovid, m.k_u, 0)

    dimdata(6,3) = dimdata(6,3) + IIF(m.usl_ok=2, m.k_u, 0)
    dimdata(7,3) = dimdata(7,3) + IIF(m.usl_ok=2 and m.isonk, m.k_u, 0)
    dimdata(8,3) = dimdata(8,3) + IIF(m.usl_ok=2 and m.isdsp, m.k_u, 0)
    dimdata(9,3) = dimdata(9,3) + IIF(m.usl_ok=2 and m.ispmo, m.k_u, 0)
    dimdata(10,3) = dimdata(10,3) + IIF(m.usl_ok=2 and m.iscovid, m.k_u, 0)

    dimdata(6,4) = dimdata(6,4) + IIF(m.usl_ok=2 and m.isvmp, m.k_u, 0)
    dimdata(7,4) = dimdata(7,4) + IIF(m.usl_ok=2 and m.isvmp and m.isonk, m.k_u, 0)
    dimdata(8,4) = dimdata(8,4) + IIF(m.usl_ok=2 and m.isvmp and m.isdsp, m.k_u, 0)
    dimdata(9,4) = dimdata(9,4) + IIF(m.usl_ok=2 and m.isvmp and m.ispmo, m.k_u, 0)
    dimdata(10,4) = dimdata(10,4) + IIF(m.usl_ok=2 and m.isvmp and m.iscovid, m.k_u, 0)

    dimdata(6,5) = dimdata(6,5) + IIF(m.usl_ok=1, m.k_u, 0)
    dimdata(7,5) = dimdata(7,5) + IIF(m.usl_ok=1 and m.isonk, m.k_u, 0)
    dimdata(8,5) = dimdata(8,5) + IIF(m.usl_ok=1 and m.isdsp, m.k_u, 0)
    dimdata(9,5) = dimdata(9,5) + IIF(m.usl_ok=1 and m.ispmo, m.k_u, 0)
    dimdata(10,5) = dimdata(10,5) + IIF(m.usl_ok=1 and m.iscovid, m.k_u, 0)

    dimdata(6,6) = dimdata(6,6) + IIF(m.usl_ok=1 and m.isvmp, m.k_u, 0)
    dimdata(7,6) = dimdata(7,6) + IIF(m.usl_ok=1 and m.isvmp and m.isonk, m.k_u, 0)
    dimdata(8,6) = dimdata(8,6) + IIF(m.usl_ok=1 and m.isvmp and m.isdsp, m.k_u, 0)
    dimdata(9,6) = dimdata(9,6) + IIF(m.usl_ok=1 and m.isvmp and m.ispmo, m.k_u, 0)
    dimdata(10,6) = dimdata(10,6) + IIF(m.usl_ok=1 and m.isvmp and m.iscovid, m.k_u, 0)
   ENDIF 

  ENDSCAN 
  USE IN talon 
  SELECT aisoms 
  
  WAIT CLEAR 
 
 ENDSCAN 
 USE IN aisoms 
 USE IN profot
 USE IN dspcodes

 CREATE CURSOR curdata (recid i)
 INSERT INTO curdata (recid) VALUES (0)
 m.llResult = X_Report(pTempl+'\yu_04.xls', pBase+'\'+m.gcperiod+'\yu_04.xls', .T.)
 USE IN curdata 

RETURN 