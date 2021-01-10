PROCEDURE  FormOnk03
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÎÐÌÓ ÎÍÊ-03,',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\ONK03.xls')
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdata (osn230 c(6), c_err c(3), k_u n(9), s_all n(11,2))
 SELECT curdata
 INDEX ON osn230 TAG osn230
 INDEX ON c_err TAG c_err
 SET ORDER TO c_err
 
 PUBLIC dimdata(6,2)
 dimdata = 0
 
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
 
 m.llResult = X_Report(pTempl+'\ONK03.xls', pBase+'\'+m.gcperiod+'\ONK03.xls', .T.)
 
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
  IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   IF USED('err')
    USE IN err
   ENDIF 
   LOOP 
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_SL'+m.qcod+'.dbf')
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_SL'+m.qcod, 'onk_sl', 'share', 'recid_s')>0
    IF USED('onk_sl')
     USE IN onk_sl
    ENDIF 
   ENDIF 
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_USL'+m.qcod+'.dbf')
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_USL'+m.qcod, 'onk_usl', 'share', 'recid_s')>0
    IF USED('onk_usl')
     USE IN onk_usl
    ENDIF 
   ENDIF 
  ENDIF 
  
  SELECT talon
  SET RELATION TO recid INTO err

  SCAN 
  
   m.ds   = ds
   m.ds_2 = ds_2
   *m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR ;
    	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
   m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
   m.s_all = s_all
   
   m.recid_s = recid_lpu
   m.recid_sl = ''
   m.usl_tip = 0
   
   IF USED('onk_sl')
    IF SEEK(m.recid_s, 'onk_sl')
     m.recid_sl = onk_sl.recid
    ENDIF 
   ENDIF 
   IF USED('onk_usl')
    IF !EMPTY(m.recid_sl)
     IF SEEK(m.recid_sl, 'onk_usl')
      m.usl_tip = onk_usl.usl_tip
     ENDIF 
    ENDIF 
   ENDIF 

   m.IsErr = IIF(!EMPTY(err.c_err), .T., .F.)
   
   dimdata(1,1) = dimdata(1,1) + 1
   dimdata(1,2) = dimdata(1,2) + m.s_all

   dimdata(2,1) = dimdata(2,1) + IIF(m.IsOnkDs, 1, 0)
   dimdata(2,2) = dimdata(2,2) + IIF(m.IsOnkDs, m.s_all, 0)
   
   dimdata(5,1) = dimdata(5,1) + IIF(m.IsOnkDs and m.usl_tip=2, 1, 0)
   dimdata(5,2) = dimdata(5,2) + IIF(m.IsOnkDs and m.usl_tip=2, m.s_all, 0)
   
   dimdata(3,1) = dimdata(3,1) + IIF(m.IsErr, 1, 0)
   dimdata(3,2) = dimdata(3,2) + IIF(m.IsErr, m.s_all, 0)
   
   dimdata(4,1) = dimdata(4,1) + IIF(m.IsOnkDs and m.IsErr, 1, 0)
   dimdata(4,2) = dimdata(4,2) + IIF(m.IsOnkDs and m.IsErr, m.s_all, 0)

   dimdata(6,1) = dimdata(6,1) + IIF(m.IsOnkDs and m.IsErr and m.usl_tip=2, 1, 0)
   dimdata(6,2) = dimdata(6,2) + IIF(m.IsOnkDs and m.IsErr and m.usl_tip=2, m.s_all, 0)

  ENDSCAN 
  SET RELATION OFF INTO err 
  USE 
  USE IN err
  IF USED('onk_sl')
   USE IN onk_sl
  ENDIF 
  IF USED('onk_usl')
   USE IN onk_usl
  ENDIF 
  
  SELECT aisoms
 ENDSCAN 
 USE IN aisoms

RETURN 