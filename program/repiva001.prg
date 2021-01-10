PROCEDURE RepIVA001
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ ÄËß ÈÂÀÍÀ ÔÈËÈÍÀ?',4+32,'')=7
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
 
 CREATE CURSOR curss (mcod c(7), k_u n(6), s_all n(13,2))
 
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
  
  m.k_u   = 0
  m.s_all = 0
  SELECT talon 
  SCAN 
   IF !BETWEEN(ds,'I10','I15')
    LOOP 
   ENDIF 
   IF SUBSTR(otd,4,3)<>'034'
    LOOP 
   ENDIF 
   m.k_u   = m.k_u + talon.k_u
   m.s_all = m.s_all + talon.s_all
   
  ENDSCAN 
  USE IN talon 
  USE IN err 
  
  INSERT INTO curss FROM MEMVAR 
  
  SELECT aisoms 
  
 ENDSCAN 
 USE 
 
 SELECT curss
 COPY TO &pBase\&gcPeriod\iva001
 
 MESSAGEBOX('OK!',0+64,m.gcperiod)
 
 
RETURN 