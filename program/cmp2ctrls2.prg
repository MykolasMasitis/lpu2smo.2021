PROCEDURE Cmp2Ctrls2
 IF MESSAGEBOX('ÑÐÀÂÍÈÒÜ ÄÂÅ ÝÊÑÏÅÐÒÈÇÛ?',4+32,'')=7
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
 
 CREATE CURSOR c_c (mcod c(7), e_rid i, e_f c(1), ec_err c(3), x_rid i, x_f c(1), xc_err c(3))
 
 SELECT aisoms 
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   *MESSAGEBOX('ÔÀÉË '+pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf'+' ÍÅ ÍÀÉÄÅÍ!',0+64,'')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod+'.dbf')
   MESSAGEBOX('ÔÀÉË '+pBase+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod+'.dbf'+' ÍÅ ÍÀÉÄÅÍ!',0+64,'')
   LOOP 
  ENDIF 
  
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'ef', 'shar')>0
   IF USED('ef')
    USE IN ef
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod, 'xf', 'shar')>0
   USE IN ef
   IF USED('xf')
    USE IN xf
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT m.mcod as mcod, IIF(!ISNULL(ef.rid), ef.rid, 0) as e_rid, IIF(!ISNULL(ef.f), ef.f, '') as e_f, IIF(!ISNULL(ef.c_err), ef.c_err, '') as ec_err,;
  	IIF(!ISNULL(xf.rid), xf.rid, 0) as x_rid, IIF(!ISNULL(xf.f), xf.f, '') as x_f, IIF(!ISNULL(xf.c_err), xf.c_err, '') as xc_err;
  	FROM ef FULL JOIN xf ON ef.f=xf.f AND ef.rid=xf.rid AND ;
  	ef.c_err=xf.c_err INTO TABLE  &pBase\&gcPeriod\&mcod\cmp_ctrls

  USE IN cmp_ctrls
  
  SELECT c_c
  APPEND FROM &pBase\&gcPeriod\&mcod\cmp_ctrls
  
  USE IN ef
  USE IN xf
  
  WAIT CLEAR 
  
  SELECT aisoms 
   
 ENDSCAN 
 USE IN aisoms
 
 SELECT c_c
 COPY TO &pBase\&gcPeriod\cmp_ctrls

  SELECT xc_err, coun(*) as cnt  FROM c_c WHERE EMPTY(ec_err) ;
  	GROUP BY xc_err	ORDER BY cnt DESC INTO TABLE &pBase\&gcPeriod\x_stat
  USE IN x_stat

  SELECT ec_err, coun(*) as cnt  FROM c_c WHERE EMPTY(xc_err) ;
  	GROUP BY ec_err	ORDER BY cnt DESC INTO TABLE &pBase\&gcPeriod\e_stat
  USE IN e_stat

 USE 
 
 MESSAGEBOX('OK!',0+64,'')
RETURN 