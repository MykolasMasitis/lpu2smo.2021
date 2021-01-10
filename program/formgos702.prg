PROCEDURE FormGOS702
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÎÐÌÓ ÎÍÊ-02?',4+32,'')=7
  RETURN 
 ENDIF 
* IF !fso.FileExists(ptempl+'\GOS701.xls')
*  RETURN 
* ENDIF 
 
* CREATE CURSOR curdata (recid i AUTOINC , lpu_id n(4), mcod c(7), cod n(6), ds c(6), c_err c(3), osn230 c(6), "name" c(100), ;
 	isonk_sl l, ds1_t n(1), stad n(3), onk_t n(3), onk_n n(3), onk_m n(3), s_all n(11,2))
 CREATE CURSOR curdata (recid i AUTOINC , lpu_id n(4), mcod c(7), cod n(6), ds c(6), c_err c(3), ;
 	isonk_sl l, ds1_t n(1), stad n(3), onk_t n(3), onk_n n(3), onk_m n(3), s_all n(11,2))
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
* SET RELATION TO LEFT(c_err,2) INTO sookod
* REPLACE ALL osn230 WITH sookod.osn230, name WITH sookod.comment
* SET RELATION OFF INTO sookod
* USE IN sookod
 
 IF RECCOUNT('curdata')=0
  USE IN curdata
  MESSAGEBOX('ÇÀÏÈÑÅÉ ÍÅ ÎÁÍÀÐÓÆÅÍÎ!',0+64,'')
  RETURN 
 ENDIF 
 
 SELECT curdata
 COPY TO &pBase\&gcPeriod\ONKS701
 
 *m.llResult = X_Report(pTempl+'\ONKS701.xls', pBase+'\'+m.gcperiod+'\ONKS701.xls', .T.)
 
 USE 
 
 MESSAGEBOX('OK!',0+64,'')
 
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
  IF fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\onk_sl'+m.qcod+'.dbf')
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\onk_sl'+m.qcod, 'onk_sl', 'shar', 'sn_pol')>0
    IF USED('onk_sl')
     USE IN onk_sl
    ENDIF 
    IF USED('talon')
     USE IN talon 
    ENDIF 
    IF USED('err')
     USE IN err
    ENDIF 
    LOOP 
   ENDIF 
  ENDIF 
  
  IF USED('onk_sl')
   SELECT talon 
   SET RELATION TO sn_pol INTO onk_sl
  ENDIF 
  SELECT err 
  SET RELATION TO rid INTO talon 

  SCAN 
  
   m.c_err  = c_err
   IF !INLIST(m.c_err, 'X1B','X2B','X3B','X4B','X5B','X6B','X7B','X8B','X9B')
    LOOP 
   ENDIF 
   
   m.profil = talon.profil
   m.cod    = talon.cod
   m.ds     = talon.ds
   m.s_all  = talon.s_all
   
   IF !USED('onk_sl')
    m.isonk_sl = .F.
   ELSE 
    m.isonk_sl = .T.
    ds1_t = onk_sl.ds1_t
    stad  = onk_sl.stad
    onk_t = onk_sl.onk_t
    onk_n = onk_sl.onk_n
    onk_m = onk_sl.onk_m
   ENDIF 

   INSERT INTO curdata FROM MEMVAR 
  ENDSCAN 

  SET RELATION OFF INTO talon
  USE 
  USE IN talon 
  IF USED('onk_sl')
   USE IN onk_sl
  ENDIF 
  SELECT aisoms
 ENDSCAN 
 USE IN aisoms

RETURN 