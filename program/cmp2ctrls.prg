PROCEDURE Cmp2Ctrls
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\Ctrl'+m.qcod+'.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\Ctrl'+PADL(m.tMonth,2,'0')+RIGHT(STR(tYear,4),2)+'.dbf')
  RETURN 
 ENDIF 

 * Èץ פאיכ! 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\Ctrl'+PADL(m.tMonth,2,'0')+RIGHT(STR(tYear,4),2), 'ctrlr2', 'shar')>0
  IF USED('ctrlr2')
   USE IN ctrlr2
  ENDIF 
  RETURN 
 ENDIF 

 SELECT ctrlr2
 INDEX ON STR(lpu_id,6) + file + recid + errors TAG unik 
 SET ORDER TO unik
 
 * Íאר פאיכ
 IF OpenFile(pBase+'\'+m.gcPeriod+'\Ctrl'+m.qcod, 'myCtrl', 'shar')>0
  IF USED('myCtrl')
   USE IN myCtrl
  ENDIF 
  USE IN ctrlr2
  RETURN 
 ENDIF 
 
 SELECT myCtrl
 INDEX ON STR(lpu_id,6) + file + recid + errors TAG unik 
 SET ORDER TO unik
 
 CREATE CURSOR rslt (lpu_id n(6), mcod c(7), file c(12), recid c(7), errors c(5), c_err_r2 c(5))
 INDEX ON STR(lpu_id,6) + file + recid + errors TAG unik 
 SET ORDER TO unik
 
 SELECT rslt 
 APPEND FROM pBase+'\'+m.gcPeriod+'\Ctrl'+m.qcod
 
 SELECT ctrlr2
 SCAN 
  SCATTER MEMVAR 
  m.c_err_r2 = m.errors
  RELEASE m.errors
  m.key = STR(m.lpu_id,6) + m.file + m.recid+' ' + m.c_err_r2

  IF SEEK(m.key, 'rslt')
   REPLACE c_err_r2 WITH m.c_err_r2 IN rslt 
  ELSE 
   INSERT INTO rslt FROM MEMVAR 
  ENDIF 
 ENDSCAN 
 USE 
 
 SELECT rslt
 COPY TO &pBase\&gcPeriod\cmp_ctrls WITH cdx 
 
 USE IN rslt
 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 