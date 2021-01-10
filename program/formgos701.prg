PROCEDURE FormGOS701
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÎÐÌÓ ÃÎ-01?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\GOS701.xls')
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdata (recid i AUTOINC , lpu_id n(4), mcod c(7), cod n(6), ds c(6), profil c(3), c_err c(3), osn230 c(6), "name" c(100), s_all n(11,2))
 INDEX on recid tag recid
 SET ORDER TO recid 
 SELECT curdata
 
 FOR m.nmonth = m.tmonth TO m.tmonth
  m.lcmonth  = PADL(m.nmonth,2,'0')
  m.lcperiod = LEFT(m.gcperiod,4) + m.lcmonth

  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 

  =GO01one(m.lcperiod)
  
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
 REPLACE ALL osn230 WITH sookod.osn230, name WITH sookod.comment
 SET RELATION OFF INTO sookod
 USE IN sookod
 
 IF RECCOUNT('curdata')=0
  USE IN curdata
  MESSAGEBOX('ÇÀÏÈÑÅÉ ÍÅ ÎÁÍÀÐÓÆÅÍÎ!',0+64,'')
  RETURN 
 ENDIF 
 
 m.llResult = X_Report(pTempl+'\GOS701.xls', pBase+'\'+m.gcperiod+'\GOS701.xls', .T.)
 
 USE 
 
RETURN 

FUNCTION GO01one(para01)
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
  m.lpu_id = lpuid
  m.mcod   = mcod 
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
  
   m.c_err  = c_err
   m.profil = talon.profil
   m.cod    = talon.cod
   m.ds     = talon.ds
   m.s_all  = talon.s_all

   INSERT INTO curdata FROM MEMVAR 
  ENDSCAN 

  SET RELATION OFF INTO talon
  USE 
  USE IN talon 
  SELECT aisoms
 ENDSCAN 
 USE IN aisoms

RETURN 