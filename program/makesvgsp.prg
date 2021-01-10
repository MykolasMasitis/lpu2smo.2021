PROCEDURE MakeSvGsp
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÑÂÎÄÍÛÉ ÔÀÉË ÃÎÑÏÈÒÀËÈÇÀÖÈÉ?', 4+32, '')=7
  RETURN
 ENDIF 

 CREATE CURSOR Gosp (recid i, mcod c(7), sn_pol c(25), c_i c(30), cod n(6), d_u d, k_u n(3))
 INDEX ON sn_pol TAG sn_pol && FOR IsMes(cod) OR IsVMP(cod)

 IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
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
  IF !fso.FolderExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 

  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod + '...' WINDOW NOWAIT 
  
  SELECT talon 
  SCAN 
   SCATTER MEMVAR 
   IF !IsMes(m.cod) AND !IsVmp(m.cod)
    LOOP 
   ENDIF 
   INSERT INTO Gosp FROM MEMVAR 
  ENDSCAN 
  USE 
  
  WAIT CLEAR 
 
  SELECT aisoms 
  
 ENDSCAN 
 USE IN aisoms 
 
 SELECT Gosp

 SELECT mcod,sn_pol,c_i,MAX(d_u) as d_u,MAX(cod) as cod,SUM(k_u) as k_u FROM gosp ;
	GROUP BY sn_pol,c_i, mcod ORDER BY mcod,sn_pol,c_i INTO CURSOR Gsp

 USE IN Gosp
 SELECT Gsp
 COPY TO &pBase\&gcPeriod\Gosp WITH cdx 
 USE 
 
 MESSAGEBOX('OK!', 0+64, '')

RETURN 