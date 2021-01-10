PROCEDURE rep_kv
 IF MESSAGEBOX('СФОРМИРОВАТЬ КВ-ОТЧЕТЫ',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pTempl+'\kv.xls')
  MESSAGEBOX('ШАБЛОН '+m.pTempl+'\kv.xls НЕ НАЙДЕН!',0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\profot', 'profot', 'shar', 'otd')>0
  USE IN aisoms 
  IF USED('profot')
   USE IN profot
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR  rep_kv (mcod c(7), lpu_id n(4),;
 	sum_all n(13,2), st_all n(13,2), st_onk n(13,2), st_ext n(13,2), st_pln n(13,2),;
 	dst_all n(13,2), dst_onk n(13,2), eco n(13,2), gem n(13,2), n_02 n(6), ;
 	s_02 n(13,2), s_def_02 n(13,2), n_ekmo n(6), s_ekmo n(13,2), def_ekmo n(13,2),;
 	n_gosp n(6), s_gosp n(13,2), def_gosp n(13,2), n_amb n(6), s_amb n(13,2), def_amb n(13,2))

 CREATE CURSOR  rep_covid (mcod c(7), lpu_id n(4),;
 	sum_all n(13,2), st_all n(13,2), st_onk n(13,2), st_ext n(13,2), st_pln n(13,2),;
 	dst_all n(13,2), dst_onk n(13,2), eco n(13,2), gem n(13,2), n_02 n(6), ;
 	s_02 n(13,2), s_def_02 n(13,2), n_ekmo n(6), s_ekmo n(13,2), def_ekmo n(13,2),;
 	n_gosp n(6), s_gosp n(13,2), def_gosp n(13,2), n_amb n(6), s_amb n(13,2), def_amb n(13,2))
 
 SELECT aisoms 
 SCAN 
  m.mcod = mcod 
  m.lpu_id = lpuid
  m.s_pred = s_pred
  IF m.s_pred<=0
   LOOP 
  ENDIF 
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   USE IN talon 
   IF USED('err')
    USE IN err 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  CREATE CURSOR gosp (c_i c(25))
  SELECT gosp 
  INDEX on c_i TAG c_i
  SET ORDER TO c_i
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT talon 
  SET RELATION TO recid INTO err
  SET RELATION TO SUBSTR(otd,2,2) INTO profot ADDITIVE 
  m.sum_all = 0
  m.st_all  = 0
  m.st_onk  = 0 
  m.st_ext  = 0
  m.st_pln  = 0
  
  m.dst_all = 0
  m.dst_onk = 0 
  
  m.eco     = 0
  m.gem     = 0
  
  m.n_02    = 0
  m.s_02    = 0
  m.s_def_02 = 0
  
  m.n_ekmo = 0
  m.s_ekmo = 0
  m.def_ekmo = 0
  
  m.n_gosp = 0
  m.s_gosp = 0
  m.def_gosp = 0
  
  m.n_amb = 0
  m.s_amb = 0
  m.def_amb = 0

  SCAN 
   m.c_i     = c_i
   m.cod     = cod
   m.s_all   = s_all && + s_lek
   m.usl_ok  = INT(VAL(profot.usl_ok))
   m.ds      = ds
   m.p_cel   = p_cel

   IF !(INLIST(m.ds, 'U07.1', 'U07.2','Z03.8','Z22.8','Z20.8','Z11.5','B34.2','B33.8') OR BETWEEN(LEFT(m.ds,3),'J12','J18')) && COVID-19

    IF !EMPTY(err.c_err)
     LOOP 
    ENDIF
    
    m.ds_2    = ds_2

    m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) OR ;
  		BETWEEN(LEFT(m.ds,3),'D00','D09') , .T., .F.)

    m.ord = ord
    m.IsExt = IIF(INLIST(m.ord,2,3), .T., .F.)

    m.sum_all = m.sum_all + m.s_all && Строка 4. Всего за медицинскую помощь.
    m.st_all  = m.st_all + IIF(m.usl_ok=1, m.s_all, 0) && Строка 5. Медицинская помощь в условиях круглосуточного стационара
    m.st_onk  = m.st_onk + IIF(m.usl_ok=1 AND m.IsOnkDs, m.s_all, 0) && Строка 5.1. В т.ч. по профилю "онкология"
   
    m.st_ext  = m.st_ext + IIF(m.IsExt AND !m.IsOnkDs, m.s_all, 0) && Строка 5.2. госпитализации в экстренной форме
    m.st_pln  = m.st_pln + IIF(!m.IsExt AND !m.IsOnkDs, m.s_all, 0) && Строка 5.3. госпитализации в плановой форме
   
    m.dst_all  = m.dst_all + IIF(m.usl_ok=2, m.s_all, 0) && Строка 6. Медицинская помощь в условиях дневного стационара
    m.dst_onk  = m.dst_onk + IIF(m.usl_ok=2 AND m.IsOnkDs, m.s_all, 0) && Строка 5.1. В т.ч. по профилю "онкология"
   
    IF m.usl_ok=2
     DO CASE 
      CASE m.cod = 97041
       m.eco = m.eco + m.s_all
      CASE INLIST(m.cod,97010,197010)
       m.gem = m.gem + m.s_all
      OTHERWISE 
     
     ENDCASE 
    ENDIF 
   
   ELSE 
   
    *IF m.usl_ok=3
    * m.n_amb   = m.n_amb + IIF(m.p_cel='1.1', 1, 0)
    * m.s_amb   = m.s_amb + IIF(m.p_cel='1.1', m.s_all, 0)
    * m.def_amb = m.def_amb + IIF(m.p_cel='1.1' AND !EMPTY(err.c_err), m.s_all, 0)
    *ENDIF 
   
    IF m.mcod='0371001'
     IF INLIST(m.cod, 56031,156002)
      m.n_amb   = m.n_amb + 1
      m.s_amb   = m.s_amb + m.s_all
      m.def_amb = m.def_amb + IIF(!EMPTY(err.c_err), m.s_all, 0)
     ELSE 
      m.n_02 = m.n_02 + 1
      m.s_02 = m.s_02 + m.s_all
      m.s_def_02 = m.s_def_02 + IIF(!EMPTY(err.c_err), m.s_all, 0)
     ENDIF
    ELSE 
     IF m.usl_ok=1
      IF !SEEK(m.c_i, 'gosp')
       INSERT INTO gosp FROM MEMVAR 
      ENDIF 
      m.s_gosp   = m.s_gosp + m.s_all
      m.def_gosp = m.def_gosp + IIF(!EMPTY(err.c_err), m.s_all, 0)
     ENDIF 
     IF INLIST(m.cod,49011,49012,49013,149011,149012,149013)
      m.n_ekmo = m.n_ekmo + 1
      m.s_ekmo = m.s_ekmo + m.s_all
      m.def_ekmo = m.def_ekmo + IIF(!EMPTY(err.c_err), m.s_all, 0)
     ENDIF 
    ENDIF 
   
   ENDIF 

  ENDSCAN 
  
  m.n_gosp = RECCOUNT('gosp')
  USE IN gosp 
  INSERT INTO rep_kv FROM MEMVAR 
  
  SET RELATION OFF INTO err
  SET RELATION OFF INTO profot 
  USE IN talon 
  USE IN err 
  
  WAIT CLEAR 
  
  SELECT aisoms

 ENDSCAN 
 USE IN aisoms 
 
 USE IN profot
 
 SELECT rep_kv
 COPY TO &pBase\&gcPeriod\rep_kv
 
 SCAN 
  SCATTER MEMVAR 
  IF m.n_amb>0 OR m.n_gosp>0 OR m.n_ekmo>0 OR m.n_02>0
   INSERT INTO rep_covid FROM MEMVAR 
  ENDIF 
 ENDSCAN 
 SELECT rep_covid
 COPY TO &pBase\&gcPeriod\rep_covid
 
 m.llResult = X_Report(m.pTempl+'\kv.xls', m.pBase+'\'+m.gcPeriod+'\rep_kv.xls', .T.)
 
 MESSAGEBOX('OK!', 0+64,'')
 
RETURN 