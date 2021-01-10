PROCEDURE FEmptySqlId
 SET DELETED OFF 
 IF MESSAGEBOX('ÏÎÈÑÊÀÒÜ ÏÓÑÒÛÅ SQLID?',4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
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
 
 CREATE CURSOR curs_t ;
	(RecId i , ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6), pcod c(10), otd c(8), ;
	 cod n(6), tip c(1), d_u d, k_u n(4), kd_fact n(3), n_kd n(3), d_type c(1), s_all n(11,2), ;
	 s_lek n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3), kur n(5,3), ds_2 c(6), ds_3 c(6), ;
	 det n(1), k2 n(5,3), vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17), ord n(1), date_ord d, ;
	 lpu_ord n(6), recid_lpu c(7), fil_id n(6), ds_onk n(1), p_cel c(3), dn n(1), reab n(1), tal_d d, napr_v_in n(1), ;
	 c_zab n(1), napr_usl c(15), mp c(1), typ c(1), dop_r n(2), vz n(1), IsPr L)
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT talon 
  SCAN 
   *REPLACE sqlid WITH 0, sqldt WITH {}
   m.sqlid = sqlid
   IF m.sqlid>0
    LOOP 
   ENDIF 
   
   SCATTER MEMVAR 
   INSERT INTO curs_t FROM MEMVAR 
   
  ENDSCAN 
  USE IN talon 
  SELECT aisoms 

 ENDSCAN 
 USE IN aisoms 
 
 SELECT curs_t 
 COPY TO &pBase\&gcPeriod\curs_t
 BROWSE 
 
 MESSAGEBOX('OK!', 0+64, '')

RETURN 