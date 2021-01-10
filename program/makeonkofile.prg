PROCEDURE MakeOnkoFile
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÑÂÎÄÍÛÉ ÎÍÊÎ-ÔÀÉË?', 4+32, '')=7
  RETURN
 ENDIF 
 
 m.period = m.gcperiod
 
 CREATE CURSOR sl (mcod c(7), period c(6), recid i, sn_pol c(25), c_i c(30), cod n(6), ds1_t n(1), stad n(3), ;
 	onk_t n(3), onk_n n(3), onk_m n(3), mtstz n(1), sod n(6,2), k_fr n(2), wei n(5,1), hei n(3), bsa n(4,2), err c(3), rid i)

 CREATE CURSOR usl (mcod c(7), period c(6), recid i, sn_pol c(25), c_i c(30), cod n(6), usl_tip n(1), hir_tip n(1),;
 	lek_tip_l n(1), lek_tip_v n(1), luch_tip n(1), pptr n(1), err c(3), rid i)
 SELECT usl 
 *INDEX ON PADL(rid,6,'0')+' '+STR(usl_tip,1) TAG unik
 *SET ORDER TO unik
 
 CREATE CURSOR diag (mcod c(7), period c(6), recid i, sn_pol c(25), c_i c(30), cod n(6), diag_tip n(1), diag_code n(1),;
 	diag_rslt n(1), diag_date d, rec_rslt n(1), met_issl n(2), err c(3), rid i)
 CREATE CURSOR ls (mcod c(7), period c(6), recid i, sn_pol c(25), c_i c(30), cod n(6), regnum c(40), date_inj d, code_sh c(10),;
 	err c(3), rid i, sid c(10))
 CREATE CURSOR cons (mcod c(7), period c(6), recid i, sn_pol c(25), c_i c(30), cod n(6), pr_cons n(1), dt_cons d,;
 	err c(3), rid i)
 CREATE CURSOR napr (mcod c(7), period c(6), recid i, sn_pol c(25), c_i c(30), cod n(6), ;
 	napr_date d, nap_number c(16), napr_v_out n(1), napr_mo n(4),;
 	err c(3), rid i)

 CREATE CURSOR cvls (mcod c(7), period c(6), recid i, sn_pol c(25), c_i c(30), cod n(6), regnum c(40), date_inj d, code_sh c(10), sid c(10),;
 	err c(3), rid i)

 FOR n_mon = m.tMonth TO m.tMonth
  m.lcperiod = STR(m.tYear,4)+PADL(n_mon,2,'0')
  IF !fso.FolderExists(m.pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  
  =OnePeriod(m.lcperiod)
  
 ENDFOR 
 
 SELECT sl 
 COPY TO &pBase\&gcPeriod\onk_sl
 USE 
 SELECT usl 
 *SET ORDER TO 
 COPY TO &pBase\&gcPeriod\onk_usl
 USE 
 SELECT diag
 COPY TO &pBase\&gcPeriod\onk_diag
 USE 
 SELECT ls
 COPY TO &pBase\&gcPeriod\onk_ls	
 USE 
 SELECT cons
 COPY TO &pBase\&gcPeriod\onk_cons
 USE 
 SELECT napr
 COPY TO &pBase\&gcPeriod\onk_napr_v_out
 USE 
 SELECT cvls
 COPY TO &pBase\&gcPeriod\cv_ls	
 USE 
 
 MESSAGEBOX('OK!', 0+64, '')

RETURN 

FUNCTION OnePeriod
 PARAMETERS para1
 PRIVATE m.lcperiod
 m.lcperiod = para1

 IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 

  =OneMO(m.mcod)
  
 ENDSCAN 
 USE IN aisoms 
 
RETURN 

FUNCTION OneMO
 PARAMETERS para1
 PRIVATE m.mcod
 m.mcod = para1
 
 IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
  RETURN 
 ENDIF 

 IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid_lpu')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod, 'eerr', 'shar', 'rid')>0
  USE IN talon
  IF USED('eerr')
   USE IN eerr
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merr', 'shar', 'rid')>0
  USE IN talon
  USE IN eerr
  IF USED('merr')
   USE IN merr
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_SL'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_SL'+m.qcod, 'onk_sl', 'shar', 'recid')>0
   IF USED('onk_sl')
    USE IN onk_sl
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_USL'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_USL'+m.qcod, 'onk_usl', 'shar', 'recid')>0
   IF USED('onk_usl')
    USE IN onk_usl
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_DIAG'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_DIAG'+m.qcod, 'onk_diag', 'shar', 'recid')>0
   IF USED('onk_diag')
    USE IN onk_diag
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_CONS'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_CONS'+m.qcod, 'onk_cons', 'shar', 'recid')>0
   IF USED('onk_cons')
    USE IN onk_cons
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_LS'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_LS'+m.qcod, 'onk_ls', 'shar', 'recid_s')>0
   IF USED('onk_ls')
    USE IN onk_ls
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_NAPR_V_OUT'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_NAPR_V_OUT'+m.qcod, 'onk_napr', 'shar', 'recid')>0
   IF USED('onk_napr')
    USE IN onk_napr
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\CV_LS'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\CV_LS'+m.qcod, 'cv_ls', 'shar', 'recid_s')>0
   IF USED('cv_ls')
    USE IN cv_ls
   ENDIF 
  ENDIF 
 ENDIF 
 
 SELECT * FROM eerr WHERE f='S' AND LEFT(c_err,1)='O' INTO CURSOR c_err
 SELECT c_err
 INDEX on rid TAG rid 
 SET ORDER TO rid 
 
 USE IN eerr
 
 IF USED('onk_sl')
  SELECT onk_sl
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   m.recid = IIF(SEEK(m.recid_s, 'talon'), talon.recid, 0)

   m.err = IIF(SEEK(m.recid, 'c_err'), c_err.c_err, '')
   
   INSERT INTO sl FROM MEMVAR 
  
  ENDSCAN 
 ENDIF 
 
 IF USED('onk_usl')
  SELECT onk_usl
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   m.recid_s = IIF(SEEK(m.recid_sl, 'onk_sl'), onk_sl.recid_s, 'qwert')
   m.recid   = IIF(SEEK(m.recid_s, 'talon'), talon.recid, 0)
   
   m.err = IIF(SEEK(m.recid, 'c_err'), c_err.c_err, '')
   
   *m.vir = PADL(rid,6,'0')+' '+STR(usl_tip,1)
   *IF SEEK(m.vir, 'usl')
   * LOOP 
   *ENDIF 

   INSERT INTO usl FROM MEMVAR 
  
  ENDSCAN 
 ENDIF 
 
 IF USED('onk_diag')
  SELECT onk_diag
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   m.recid_s = IIF(SEEK(m.recid_sl, 'onk_sl'), onk_sl.recid_s, 'qwert')
   m.recid   = IIF(SEEK(m.recid_s, 'talon'), talon.recid, 0)
   
   m.err = IIF(SEEK(m.recid, 'c_err'), c_err.c_err, '')

   INSERT INTO diag FROM MEMVAR 
  
  ENDSCAN 
 ENDIF 

 IF USED('onk_cons')
  SELECT onk_cons
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   IF m.pr_cons=0
    LOOP 
   ENDIF 
   m.recid   = IIF(SEEK(m.recid_s, 'talon'), talon.recid, 0)
   
   m.err = IIF(SEEK(m.recid, 'c_err'), c_err.c_err, '')

   INSERT INTO cons FROM MEMVAR 
  
  ENDSCAN 
 ENDIF 

 IF USED('onk_napr')
  SELECT onk_napr
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   m.recid   = IIF(SEEK(m.recid_s, 'talon'), talon.recid, 0)
   
   m.err = IIF(SEEK(m.recid, 'c_err'), c_err.c_err, '')

   INSERT INTO napr FROM MEMVAR 
  
  ENDSCAN 
 ENDIF 

 IF USED('onk_ls')
  SELECT onk_ls
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   m.recid_sl = IIF(SEEK(m.recid_usl, 'onk_usl'), onk_usl.recid_sl, 'qwert')
   m.recid_s = IIF(SEEK(m.recid_sl, 'onk_sl'), onk_sl.recid_s, 'qwert')
   m.recid   = IIF(SEEK(m.recid_s, 'talon'), talon.recid, 0)
   
   m.err = IIF(SEEK(m.recid, 'c_err'), c_err.c_err, '')

   INSERT INTO ls FROM MEMVAR 
  
  ENDSCAN 
 ENDIF 

 IF USED('cv_ls')
  SELECT cv_ls
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   m.recid   = IIF(SEEK(m.recid_s, 'talon'), talon.recid, 0)

   INSERT INTO cvls FROM MEMVAR 
  
  ENDSCAN 
 ENDIF 

 USE IN talon 
 USE IN c_err 
 USE IN merr
 IF USED('onk_sl')
  USE IN onk_sl
 ENDIF 
 IF USED('onk_usl')
  USE IN onk_usl
 ENDIF 
 IF USED('onk_diag')
  USE IN onk_diag
 ENDIF 
 IF USED('onk_ls')
  USE IN onk_ls
 ENDIF 
 IF USED('onk_cons')
  USE IN onk_cons
 ENDIF 
 IF USED('onk_napr')
  USE IN onk_napr
 ENDIF 
 IF USED('cv_ls')
  USE IN cv_ls
 ENDIF 
 
 SELECT aisoms 
 
RETURN 