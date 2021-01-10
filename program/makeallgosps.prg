PROCEDURE MakeAllGosps
 IF MESSAGEBOX('ÑÎÁÐÀÒÜ HOSP-ÔÀÉËÛ?',4+32,'')=7
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
 
 SELECT aisoms 
 SCAN 
  m.mcod = mcod 
  IF !IsStac(m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar', 'c_i')>0 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\hosp.dbf')
   fso.DeleteFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\hosp.dbf')
  ENDIF 

  SELECT c_i, SPACE(25) as sn_pol, MAX(d_u)-SUM(k_u)+1 as d_pos, MAX(d_u) as d_vip, coun(*) as cnt, ;
  	SUM(k_u) as k_u FROM talon WHERE IsMes(cod) OR IsVMP(cod) ;
  	GROUP BY c_i ORDER BY c_i ASC INTO CURSOR cur_h READWRITE 
  SELECT cur_h
  INDEX on c_i TAG c_i 
  INDEX on sn_pol TAG sn_pol
  INDEX on d_pos TAG d_pos
  SET ORDER TO c_i
  SET RELATION TO c_i INTO talon 
  REPLACE ALL sn_pol WITH talon.sn_pol

  IF tMonth=1
  ELSE
   m.p_period = STR(tYear,4)+PADL(tMonth-1,2,'0') 
   IF fso.FileExists(pBase+'\'+m.p_period+'\'+m.mcod+'\hosp.dbf')
    APPEND FROM &pBase\&p_period\&mcod\hosp
   ENDIF 
  ENDIF 
  
  SET ORDER TO d_pos
  COPY TO &pBase\&gcPeriod\&mcod\hosp CDX 
  USE 
  
  USE IN talon 
  
  WAIT CLEAR 
  
  SELECT aisoms 
  
 ENDSCAN 
 USE IN aisoms 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 