PROCEDURE  FormOnk01
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÎÐÌÓ ÎÍÊ-01,',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\ONK01.xls')
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdata (osn230 c(6), c_err c(3), k_u n(9), s_all n(11,2))
 SELECT curdata
 INDEX ON osn230 TAG osn230
 INDEX ON c_err TAG c_err
 SET ORDER TO c_err
 
 FOR m.nmonth = m.tmonth TO m.tmonth
  m.lcmonth  = PADL(m.nmonth,2,'0')
  m.lcperiod = LEFT(m.gcperiod,4) + m.lcmonth

  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 

  =ONK01(m.lcperiod)
  
 ENDFOR 
 
 SELECT curdata
 
 IF OpenFile(pBase+'\'+m.lcperiod+'\nsi\sookodxx', 'sookod', 'shar', 'er_c')>0
  USE IN curdata
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT curdata
 SET RELATION TO LEFT(c_err,2) INTO sookod
 REPLACE ALL osn230 WITH sookod.osn230
 SET RELATION OFF INTO sookod
 USE IN sookod
 
 IF RECCOUNT('curdata')=0
  USE IN curdata
  MESSAGEBOX('ÇÀÏÈÑÅÉ ÍÅ ÎÁÍÀÐÓÆÅÍÎ!',0+64,'')
  RETURN 
 ENDIF 
 
 m.llResult = X_Report(pTempl+'\ONK01.xls', pBase+'\'+m.gcperiod+'\ONK01.xls', .T.)
 
 USE 
 
RETURN 

FUNCTION ONK01(para01)
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
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 

  IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('err')
    USE IN err
   ENDIF 
   LOOP 
  ENDIF 
  SELECT err
  IF RECCOUNT()<=0
   IF USED('err')
    USE IN err
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   IF USED('err')
    USE IN err
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT err 
  SET RELATION TO rid INTO talon 

  SCAN 
  
   m.ds = talon.ds
   IF UPPER(LEFT(m.ds,1)) != 'C'
*    LOOP 
   ENDIF 
   m.c_err = c_err
   m.k_u   = talon.k_u
   m.s_all = talon.s_all

   IF !SEEK(m.c_err, 'curdata')
    INSERT INTO curdata (osn230, c_err, k_u, s_all) VALUES ('', m.c_err, m.k_u, m.s_all)
   ELSE 
    m.ok_u   = curdata.k_u
    m.os_all = curdata.s_all
    UPDATE curdata SET s_all = m.os_all + m.s_all, k_u = m.ok_u+m.k_u WHERE c_err = m.c_err
   ENDIF 
  ENDSCAN 

  SET RELATION OFF INTO talon
  USE 
  USE IN talon 
  SELECT aisoms
 ENDSCAN 
 USE IN aisoms

RETURN 