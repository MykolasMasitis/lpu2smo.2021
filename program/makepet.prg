# DEFINE CURMONTH .T.
# DEFINE ALLPERIOD .F.

PROCEDURE MakePet

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ'+CHR(13)+CHR(10)+;
 	'ÏÎ ÏÝÒ?'+CHR(13)+CHR(10), 4+32, '')=7
  RETURN 
 ENDIF 
 
 CREATE CURSOR pet2018 ;
	(RecId i, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6), pcod c(10), otd c(8), ;
	 cod n(6), tip c(1), d_u d, k_u n(4), kd_fact n(3), n_kd n(3), d_type c(1), s_all n(11,2), ;
	 s_lek n(11,2), profil c(3))

 DIMENSION dimdata(1,10)
 dimdata = 0 

 FOR i=1 TO 12
  m.lc_period = '2018'+PADL(i,2,'0')
  IF !fso.FolderExists(pBase+'\'+m.lc_period)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.lc_period+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.lc_period+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms 
   ENDIF 
   LOOP 
  ENDIF 
  WAIT m.lc_period+'...' WINDOW NOWAIT 
  SELECT aisoms
  SCAN 
   m.mcod = mcod 
   IF !fso.FolderExists(pBase+'\'+m.lc_period+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+m.lc_period+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pBase+'\'+m.lc_period+'\'+m.mcod+'\talon', 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon 
    ENDIF 
    SELECT aisoms 
    LOOP 
   ENDIF 
   
   SELECT talon 
   SCAN 
    m.cod = cod 
    IF !INLIST(m.cod,37047,137047)
     LOOP 
    ENDIF 
    SCATTER MEMVAR 
    INSERT INTO pet2018 FROM MEMVAR 
   ENDSCAN 
   USE IN talon 
   
   SELECT aisoms 
   
  ENDSCAN 
  USE IN aisoms 
  WAIT CLEAR    

 ENDFOR  
 SELECT pet2018
 dimdata(1,4) = RECCOUNT('pet2018')
 SUM s_all TO dimdata(1,5)
 COPY TO &pBase\pet2018
 USE 
 
 CREATE CURSOR pet2019 ;
	(RecId i, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6), pcod c(10), otd c(8), ;
	 cod n(6), tip c(1), d_u d, k_u n(4), kd_fact n(3), n_kd n(3), d_type c(1), s_all n(11,2), ;
	 s_lek n(11,2), profil c(3))

 FOR i=1 TO 12
  m.lc_period = '2019'+PADL(i,2,'0')
  IF !fso.FolderExists(pBase+'\'+m.lc_period)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.lc_period+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.lc_period+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms 
   ENDIF 
   LOOP 
  ENDIF 
  WAIT m.lc_period+'...' WINDOW NOWAIT 
  SELECT aisoms
  SCAN 
   m.mcod = mcod 
   IF !fso.FolderExists(pBase+'\'+m.lc_period+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+m.lc_period+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pBase+'\'+m.lc_period+'\'+m.mcod+'\talon', 'talon', 'shar', 'cod')>0
    IF USED('talon')
     USE IN talon 
    ENDIF 
    SELECT aisoms 
    LOOP 
   ENDIF 
   
   SELECT talon 
   SCAN 
    m.cod = cod 
    IF !INLIST(m.cod,37047,137047)
     LOOP 
    ENDIF 
    SCATTER MEMVAR 
    INSERT INTO pet2019 FROM MEMVAR 
   ENDSCAN 
   USE IN talon 
   
   SELECT aisoms 
   
  ENDSCAN 
  USE IN aisoms 
  WAIT CLEAR 

 ENDFOR  
 SELECT pet2019
 dimdata(1,6) = RECCOUNT('pet2019')
 SUM s_all TO dimdata(1,7)
 COPY TO &pBase\pet2019
 USE 
 
 CREATE CURSOR curdata (recid i)
 INSERT INTO curdata (recid) VALUES (0)

 m.llResult = X_Report(pTempl+'\pet.xls', pBase+'\'+m.gcperiod+'\pet.xls', .T.)
 
 USE IN curdata 
 
 
RETURN 
