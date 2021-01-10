PROCEDURE CleanMFiles
 IF MESSAGEBOX('ÏÐÎÂÅÐÈÒÜ M-ÔÀÉËÛ ÍÀ ÓÍÈÊÀËÜÍÎÑÒÜ?',4+32,'')=7
  RETURN 
 ENDIF 

 FOR m.nmonth = m.tmonth TO m.tmonth
  m.lcmonth  = PADL(m.nmonth,2,'0')
  m.lcperiod = LEFT(m.gcperiod,4) + m.lcmonth

  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 

  =CleanOne(m.lcperiod)
  
 ENDFOR 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 

FUNCTION CleanOne(para01)
 PRIVATE m.lcperiod
 m.lcperiod = para01
 IF OpenFile(pBase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  WAIT m.mcod WINDOW NOWAIT 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'err', 'excl')>0
   IF USED('err')
    USE IN err
   ENDIF 
   LOOP 
  ENDIF 

  SELECT err
  PACK 
  INDEX ON PADL(recid,6,'0')+et+docexp+reason+LEFT(err_mee,2) TAG unikk UNIQUE 
  SET ORDER TO UNIK
  IF KEY(TAGCOUNT()) = UPPER('PADL(recid,6,"0")+et+docexp+reason+LEFT(err_mee,2)')
   SET ORDER TO 
   CALCULATE coun() TO m.nallrecs
   SET ORDER TO unikk
   CALCULATE coun() TO m.nunikrecs
   IF m.nunikrecs = m.nallrecs
    SET ORDER TO 
    DELETE TAG unikk
    USE IN err
    SELECT aisoms 
    LOOP 
   ENDIF 
  ELSE 
   DELETE TAG unik 
   INDEX ON PADL(recid,6,'0')+et+docexp+reason+LEFT(err_mee,2) TAG unik
   SET ORDER TO 
   CALCULATE coun() TO m.nallrecs
   SET ORDER TO unikk
   CALCULATE coun() TO m.nunikrecs
   IF m.nunikrecs = m.nallrecs
    SET ORDER TO 
    DELETE TAG unikk
    USE IN err
    SELECT aisoms 
    LOOP 
   ENDIF 
  ENDIF 

  MESSAGEBOX(m.lcperiod+': '+m.mcod,0+64,'')

  DELETE TAG ALL 
*  IF !m.IsServer
   ALTER table err drop COLUMN rid
*  ENDIF 
  PACK 
  INDEX ON PADL(recid,6,'0')+et+docexp+reason+LEFT(err_mee,2) TAG unik UNIQUE 
  SET ORDER TO unik 
  COPY TO &pbase\&lcperiod\&mcod\e_tmp
  ZAP 
*  IF !m.IsServer
   ALTER table err ADD COLUMN rid i AUTOINC 
*  ENDIF 
  APPEND FROM &pbase\&lcperiod\&mcod\e_tmp
  fso.DeleteFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\e_tmp.dbf')
  DELETE TAG unik
  INDEX ON PADL(recid,6,'0')+et+docexp+reason+LEFT(err_mee,2) TAG unik
  INDEX ON rid TAG rid 
  INDEX ON RecId TAG recid
  INDEX ON PADL(recid,6,'0')+et+docexp+reason TAG id_et
  USE IN err 

  SELECT aisoms
 ENDSCAN 
 USE IN aisoms
 WAIT CLEAR 
RETURN 