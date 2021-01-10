PROCEDURE StatFillFiles
 IF MESSAGEBOX('ÐÀÑ×ÈÒÀÒÜ ÑÒÀÒÈÑÒÈÊÓ ÇÀÏÎËÍÅÍÈß ÔÀÉËÎÂ?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('asioms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curss (f_name c(10), total n(11), f_full n(11))
 SELECT curss 
 INDEX on f_name TAG f_name
 SET ORDER TO f_name
 m.IsFilled = .F.
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod 
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
  
  SELECT talon 
  IF !m.IsFilled
   getfcount = AFIELDS(dTalon)
   FOR nCount = 1 TO getfcount
    INSERT INTO curss (f_name) VALUES (dTalon(nCount,1))
   ENDFOR   
   m.IsFilled = .T.
   *RELEASE dTalon
  ENDIF
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
   
  SCAN 
   FOR nCount = 1 TO getfcount
    tField = PADR(dTalon(nCount,1),10)
    IF FIELD(dTalon(nCount,1))=UPPER(dTalon(nCount,1))
     IF !EMPTY(&tField)
      m.o_full = IIF(SEEK(tField, 'curss'), curss.f_full, -1)
      IF m.o_full = -1
       INSERT INTO curss (f_name) VALUES (tField)
      ENDIF 
      m.n_full = m.o_full + 1
      UPDATE curss SET f_full=m.n_full WHERE f_name=tField
     ENDIF 
    ENDIF 
   ENDFOR   
  ENDSCAN 

  USE IN talon 
  
  WAIT CLEAR 
  
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 
 SELECT curss
 COPY TO &pBase\&gcPeriod\stattalon 
 BROWSE 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 