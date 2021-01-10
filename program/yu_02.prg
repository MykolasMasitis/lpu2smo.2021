PROCEDURE yu_02
 IF MESSAGEBOX('ÑÔÎÎÌÈÐÎÂÀÒÜ ÔÀÉËÛ ÄËß ÌÎÍÈÒÎÐÈÍÃÀ ÝÊÎ?',4+32,'97041')=7
  RETURN 
 ENDIF 

 CREATE CURSOR eco (period c(6), mcod c(7), lpu_id n(4), sn_pol c(25), cod n(6), d_u d, k_u n(3), s_all n(11,2))
 
 FOR m.nmonth=1 TO m.tmonth
  m.lcperiod = LEFT(m.gcperiod,4)+PADL(m.nmonth,2,'0')
  m.lcmonth  = PADL(m.nmonth,2,'0')

  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  
  WAIT m.lcperiod + '...' WINDOW NOWAIT 
  =yu_one(m.lcperiod)
  WAIT CLEAR 

 ENDFOR 
 
 SELECT eco
 COPY TO &pbase\&gcperiod\eco
 USE 
 
 MESSAGEBOX('OK!',0+64,'')

RETURN 

FUNCTION yu_one(para1)
 PRIVATE m.lcperiod
 m.lcperiod = para1
 
 IF OpenFile(m.pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR eco&lcperiod (period c(6), mcod c(7), lpu_id n(4), sn_pol c(25), cod n(6), d_u d, k_u n(3), s_all n(11,2))
 SELECT aisoms
 SCAN  
  m.mcod   = mcod 
  m.lpu_id = lpuid
  IF !fso.FolderExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  SELECT talon 
  SCAN 
   m.cod = cod
   IF m.cod!=97041
    LOOP 
   ENDIF 
   
   m.period = m.lcperiod
   m.sn_pol = sn_pol
   m.cod    = cod 
   m.d_u    = d_u
   m.k_u    = k_u
   m.s_all  = s_all
   
   INSERT INTO eco FROM MEMVAR 
   INSERT INTO eco&lcperiod FROM MEMVAR 
   
  ENDSCAN 
  USE IN talon 
  
  SELECT aisoms

 ENDSCAN 
 USE IN aisoms
 SELECT eco&lcperiod
 COPY TO &pbase\&lcperiod\eco&lcperiod
 USE 
 
RETURN 