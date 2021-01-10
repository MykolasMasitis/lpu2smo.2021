PROCEDURE Make15001
 IF MESSAGEBOX('ÑÎÁÐÀÒÜ ÔÀÉË "ÖÅÍÒÐ ÇÄÎÐÎÂÜß"',4+32,'')=7
  RETURN 
 ENDIF 
 
 *CREATE CURSOR cur_cz (sn_pol c(25), mcod c(7), d_u d, cod n(6))
 *INDEX on sn_pol TAG sn_pol
 
 m.prperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')
 IF !fso.FileExists(pbase+'\'+m.prperiod+'\nsi\polic_h.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ POLIC_H ÇÀ ÏÐÅÄÛÄÓÙÈÉ ÏÅÐÈÎÄ!'+CHR(13)+CHR(10)+;
  	'ÑÔÎÐÌÈÐÎÂÀÒÜ ÏÓÑÒÎÉ ÔÀÉË Â ÒÅÊÓÙÅÌ ÏÅÐÈÎÄÅ?',0+16,'') = 7
   RETURN
  ELSE
   CREATE TABLE &pBase\&gcPeriod\nsi\polic_h (sn_pol c(25), mcod c(7), d_u d, cod n(6))
   INDEX on sn_pol TAG sn_pol
   USE 
  ENDIF
 ELSE 
  fso.CopyFile(pbase+'\'+m.prperiod+'\nsi\polic_h.dbf', pbase+'\'+m.gcperiod+'\nsi\polic_h.dbf')
  fso.CopyFile(pbase+'\'+m.prperiod+'\nsi\polic_h.cdx', pbase+'\'+m.gcperiod+'\nsi\polic_h.cdx')
 ENDIF 
 
 IF OpenFile(pCommon+'\lpu_cz', 'lpu_cz', 'shar', 'mcod')>0
  IF USED('lpu_cz')
   USE IN lpu_cz
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\polic_h', 'polic_h', 'shar', 'sn_pol')>0
  IF USED('polic_h')
   USE IN polic_h
  ENDIF 
  USE IN lpu_cz
  RETURN 
 ENDIF 

 *FOR m.i=1 TO m.tmonth
  m.lc_period = STR(tYear,4)+PADL(m.i,2,'0')
  m.lc_period = m.gcPeriod
  
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
   IF !SEEK(m.mcod, 'lpu_cz')
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
    IF !INLIST(m.cod,15001,115001)
     LOOP 
    ENDIF 
    m.sn_pol = sn_pol
    m.d_u    = d_u
    *INSERT INTO cur_cz FROM MEMVAR 
    INSERT INTO polic_h FROM MEMVAR 
   ENDSCAN 
   USE IN talon 
   SELECT aisoms

  ENDSCAN 
  USE IN aisoms 
  WAIT CLEAR 
  
  USE IN lpu_cz

  SELECT polic_h
  USE 
 
 *ENDFOR 
 
 *SELECT cur_cz
 *IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\nsi\polic_h.dbf')
 * fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\nsi\polic_h.dbf')
 *ENDIF 
 *IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\nsi\polic_h.cdx')
 * fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\nsi\polic_h.cdx')
 *ENDIF 

 *COPY TO &pBase\&gcPeriod\nsi\polic_h WITH cdx 
 *USE 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 