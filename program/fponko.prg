PROCEDURE FPOnko
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“?', 4+32, '')=7
  RETURN
 ENDIF 
 m.dotname = pTempl+'\FPOnko.xls'
 IF !fso.FileExists(m.dotname)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ ‘¿…À Œ“◊≈“¿'+CHR(13)+CHR(10)+UPPER(m.dotname), 0+64, '')
  RETURN 
 ENDIF 
 IF OpenFile(pCommon+'\dspcodes', 'dspcodes', 'share', 'cod')>0
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 
 
 PUBLIC dimdata(1,20)
 dimdata = 0

 CREATE CURSOR curss (nrec i)
 INSERT INTO curss (nrec) VALUES (0)
 
 CREATE CURSOR paz (sn_pol c(25), ds c(6), ds_2 c(6), k_u n(3), d_u d, ds_onk n(1), ;
 	d_onk d, ds1_t l, dsp l, dn n(1), ch_t l, k_def n(3), ;
 	n_st n(4), s_st n(13,2), n_dst n(4), s_dst n(13,2), s_lek n(13,2))
 SELECT paz 
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 
 CREATE CURSOR ls (cod n(6), typ n(1), code_sh c(10))
 
 FOR n_mon = 1 TO m.tMonth
  m.lcperiod = STR(m.tYear,4)+PADL(n_mon,2,'0')
  IF !fso.FolderExists(m.pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  
  =OnePeriod(m.lcperiod)
  
 ENDFOR 
 
 m.docname = m.pbase+'\'+m.gcperiod+'\FPOnko.xls'
 
 SELECT paz 
 dimdata(1,1) = RECCOUNT('paz')
 COUNT FOR ds_onk=1 TO dimdata(1,2)
 dimdata(1,3) = dimdata(1,1) - dimdata(1,2)
 COUNT FOR ds1_t TO dimdata(1,4)
 COUNT FOR ds1_t AND dsp TO dimdata(1,5)
 COUNT FOR INLIST(dn,1,2) TO dimdata(1,6)
 SUM k_u TO dimdata(1,7)
 SUM k_u FOR ch_t TO dimdata(1,8)

 SUM k_def TO dimdata(1,9)
 SUM k_def FOR ch_t TO dimdata(1,10)
 
 dimdata(1,11) = INT(dimdata(1,9) * 1.5)
 dimdata(1,12) = INT(dimdata(1,10) * 1.4)
 
 SUM s_st TO dimdata(1,13)
 SUM s_dst TO dimdata(1,14)
 SUM s_lek TO dimdata(1,15)
 SUM n_st TO dimdata(1,16)
 SUM n_dst TO dimdata(1,17)
 
 dimdata(1,13) = ROUND(dimdata(1,13)/1000000,2)
 dimdata(1,14) = ROUND(dimdata(1,14)/1000000,2)
 dimdata(1,15) = ROUND(dimdata(1,15)/1000000,2)
 
 SELECT typ, coun(code_sh) as cnt FROM ls GROUP BY typ INTO CURSOR c_ls
 SELECT c_ls
 COPY TO &pBase\&gcPeriod\c_ls
 
 SUM cnt FOR typ=1 TO dimdata(1,18)
 SUM cnt FOR typ=2 TO dimdata(1,19)
 
 USE 
 
 SELECT curss  
 COPY TO &pBase\&gcPeriod\FPOnko

 m.llResult = X_Report(m.dotname, m.docname+'.xls', .T.)
 USE IN curss
 
 USE IN paz

 USE IN dspcodes

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
 IF OpenFile(m.pbase+'\'+gcperiod+'\'+'nsi'+'\profot', 'profot', 'share', 'otd')>0
  USE IN aisoms
  IF USED('profot')
   USE IN profot
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
 USE IN profot
 
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
 IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
  USE IN talon
  IF USED('err')
   USE IN err
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merr', 'shar', 'rid')>0
  USE IN talon
  USE IN err
  IF USED('merr')
   USE IN merr
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_SL'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_SL'+m.qcod, 'onk_sl', 'shar', 'recid_s')>0
   IF USED('onk_sl')
    USE IN onk_sl
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_USL'+m.qcod+'.dbf')
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\ONK_USL'+m.qcod, 'onk_usl', 'shar', 'recid_s')>0
   IF USED('onk_usl')
    USE IN onk_usl
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
 
 IF USED('onk_ls')
  SELECT onk_ls
  SCAN 
   m.s_all = s_all
   IF m.s_all<=0
    LOOP 
   ENDIF 
   m.cod = cod 
   m.usl_ok = IIF(INLIST(INT(m.cod/1000),97,197,297),'2', '1')
   DO CASE 
    CASE m.usl_ok='1'
     m.typ = 1
    CASE m.usl_ok='2'
     m.typ = 2
    OTHERWISE 
     LOOP 
   ENDCASE 
   m.code_sh = code_sh
   INSERT INTO ls FROM MEMVAR 
  ENDSCAN 
  USE IN onk_ls
 ENDIF 
 
 SELECT talon 
 SET RELATION TO recid INTO err
 SET RELATION TO recid INTO merr ADDITIVE 
 
 SCAN 
  m.ds      = ds 
  m.ds_2    = ds_2
  m.recid_s = recid_lpu
  m.ds1_t   = .F.
  m.cod     = cod
  m.recid_sl = ''
  *m.k_u     = IIF(IsMes(m.cod) or IsVmp(m.cod), 1, k_u)
  m.ch_t = .F.

  m.usl_ok = IIF(SEEK(SUBSTR(otd,2,2), 'profot'), profot.usl_ok, '0')
  m.is_gsp  = IIF(m.usl_ok='1', .T., .F.)
  m.is_dst  = IIF(m.usl_ok='2', .T., .F.)
  m.s_all   = s_all
  m.s_lek   = s_lek
  
  m.is_def = IIF(!EMPTY(err.c_err), .T., .F.)
  
  IF USED('onk_sl')
   IF SEEK(m.recid_s, 'onk_sl')
    m.ds1_t = IIF(onk_sl.ds1_t=0, .T., .F.)
    m.recid_sl = onk_sl.recid
   ENDIF  
  ENDIF 

  IF USED('onk_usl')
   IF !EMPTY(m.recid_sl) AND SEEK(m.recid_sl, 'onk_usl')
    m.ch_t = IIF(INLIST(onk_usl.usl_tip,2,4), .T., .F.)
   ENDIF  
  ENDIF 

  *m.IsOnkDs = IIF((BETWEEN(LEFT(m.ds,3), 'C00', 'C80') OR m.ds='C97') OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
  m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) OR ;
  		BETWEEN(LEFT(m.ds,3),'D00','D09') , .T., .F.)
  m.ds_onk = IIF(FIELD('ds_onk')=UPPER('ds_onk'), ds_onk, 0)

  IF !m.IsOnkDs AND m.ds_onk<>1
   LOOP 
  ENDIF 

  m.d_u    = d_u
  m.sn_pol = sn_pol
  m.ds_onk = IIF(FIELD('ds_onk')=UPPER('ds_onk'), ds_onk, 0)
  m.dn     = IIF(FIELD('dn')=UPPER('dn'), dn, 0)
  
  m.is_dsp = IIF(SEEK(m.cod, 'dspcodes'), .T., .F.)
  
  DO CASE 
   CASE m.usl_ok='1' && ÒÚ‡ˆËÓÌ‡
   CASE m.usl_ok='2' && ‰ÌÂ‚ÌÓÈ ÒÚ‡ˆËÓÌ‡
   OTHERWISE 
  ENDCASE 
  
  IF !SEEK(m.sn_pol, 'paz')
   INSERT INTO paz (sn_pol, d_u, ds, ds_2, ds_onk, d_onk, ds1_t, dsp, dn, k_u, ch_t, k_def, n_st, s_st, n_dst, s_dst, s_lek) VALUES ;
   	(m.sn_pol, m.d_u, m.ds, m.ds_2, m.ds_onk, IIF(m.ds_onk=1, m.d_u, {}), m.ds1_t, m.is_dsp, m.dn, 1, m.ch_t, IIF(m.is_def, 1, 0),;
   	IIF(m.usl_ok='1',1,0), IIF(m.usl_ok='1',m.s_all+m.s_lek,0),IIF(m.usl_ok='2',1,0),IIF(m.usl_ok='2',m.s_all+m.s_lek,0), IIF(INLIST(m.usl_ok,'1','2'),m.s_lek,0))
  ELSE 
   m.ok_u   = paz.k_u
   m.ok_def = paz.k_def
   UPDATE paz SET k_u = m.ok_u+1 WHERE sn_pol=m.sn_pol
   IF m.is_def
    UPDATE paz SET k_def = m.ok_def+1 WHERE sn_pol=m.sn_pol
   ENDIF 
   
   IF m.usl_ok='1'
    UPDATE paz SET n_st=n_st+1, s_st=s_st+(m.s_all+m.s_lek), s_lek=s_lek+m.s_lek WHERE sn_pol=m.sn_pol
   ENDIF 
   IF m.usl_ok='2'
    UPDATE paz SET n_dst=n_dst+1, s_dst=s_dst+(m.s_all+m.s_lek), s_lek=s_lek+m.s_lek WHERE sn_pol=m.sn_pol
   ENDIF 
   IF paz.ds_onk < m.ds_onk
    UPDATE paz SET ds_onk=m.ds_onk, d_onk=m.d_u WHERE sn_pol=m.sn_pol
   ENDIF 
   IF !paz.ds1_t AND m.ds1_t
    UPDATE paz SET ds1_t=m.ds1_t WHERE sn_pol=m.sn_pol
   ENDIF 
   IF !paz.dsp AND m.is_dsp
    UPDATE paz SET dsp=m.is_dsp WHERE sn_pol=m.sn_pol
   ENDIF 
   IF !INLIST(paz.dn,1,2) AND INLIST(m.dn,1,2)
    UPDATE paz SET dn=m.dn WHERE sn_pol=m.sn_pol
   ENDIF 
   IF !paz.ch_t  AND m.ch_t
    UPDATE paz SET ch_t=m.ch_t WHERE sn_pol=m.sn_pol
   ENDIF 
  ENDIF 
 
 ENDSCAN 
 SET RELATION OFF INTO merr
 SET RELATION OFF INTO err
 USE 
 USE IN err
 USE IN merr
 IF USED('onk_sl')
  USE IN onk_sl
 ENDIF 
 IF USED('onk_usl')
  USE IN onk_usl
 ENDIF 
 
 SELECT aisoms 
 
RETURN 