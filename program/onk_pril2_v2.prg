PROCEDURE Onk_pril2_v2
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“?', 4+32, '')=7
  RETURN
 ENDIF 
 m.dotname = pTempl+'\onk_pril2.xls'
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
 IF OpenFile(pCommon+'\gr_plan', 'gr_plan', 'share', 'cod')>0
  USE IN dspcodes
  IF USED('gr_plan')
   USE IN gr_plan
  ENDIF 
  RETURN 
 ENDIF 
 
 PUBLIC dimdata(1,21)
 dimdata = 0

 CREATE CURSOR curss (nrec i)
 INSERT INTO curss (nrec) VALUES (0)
 
 CREATE CURSOR paz (sn_pol c(25), ds c(6), ds_2 c(6), k_u n(3), d_u d, ds_onk n(1), ;
 	d_onk d, ds1_t l, dsp l, dn n(1), ch_t l, k_def n(3))
 SELECT paz 
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 
 CREATE CURSOR exps (tip n(1), isxt l, isdef l)
 
 FOR n_mon = 1 TO m.tMonth
  m.lcperiod = STR(m.tYear,4)+PADL(n_mon,2,'0')
  IF !fso.FolderExists(m.pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  
  =OnePeriod(m.lcperiod)
  
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
   IF USED('aisoms')
    USE IN aisoms 
   ENDIF 
   LOOP 
  ENDIF 
  
  IF fso.FolderExists(pOut+'\'+m.lcperiod)
  oMailDir        = fso.GetFolder(pOut+'\'+m.lcperiod)
  MailDirName     = oMailDir.Path
  oFilesInMailDir = oMailDir.Files
  nFilesInMailDir = oFilesInMailDir.Count

  FOR EACH oFileInMailDir IN oFilesInMailDir
   m.BFullName = oFileInMailDir.Path
   m.bname     = oFileInMailDir.Name
   m.recieved  = oFileInMailDir.DateLastModified
  
   IF LEN(m.bname)!=12
    LOOP 
   ENDIF 
  
   m.part01 = UPPER(LEFT(m.bname,2))
   m.part02 = UPPER(SUBSTR(m.bname,3,2))
   m.part03 = SUBSTR(m.bname,5,4)
   m.ext    = LOWER(RIGHT(m.bname,3))

   IF part01 != 'ME'
    LOOP 
   ENDIF 
   IF part02 != m.qcod
    LOOP 
   ENDIF 
   IF !INLIST(ext, 'dbf')
    LOOP 
   ENDIF 
   IF !SEEK(INT(VAL(m.part03)),'aisoms')
    LOOP 
   ENDIF 
   
   IF OpenFile(m.BFullName, 'mefile', 'shar')>0
    IF USED('mefile')
     USE IN mefile
    ENDIF 
    LOOP 
   ENDIF 
   
   SELECT mefile 
   SCAN 
    m.act         = act 
    m.et          = et
    *m.tip         = IIF(INLIST(SUBSTR(m.act,7,1),'2','3','7','8'), 1, 2) && Ã››/› Ãœ
    m.tip         = IIF(INLIST(m.et,'2','9','T'), 1, 2) && Ã››/› Ãœ '3', '4', 'K
    m.cod         = cod
    m.isxt        = IIF(SEEK(m.cod, 'gr_plan') and gr_plan.gr_plan='on_ı', .T., .F.)
    m.ds          = ds
    *m.IsOnkDs     = IIF((BETWEEN(LEFT(m.ds,3), 'C00', 'C80') OR m.ds='C97') OR ;
  		(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
    m.IsOnkDs     = IIF((BETWEEN(LEFT(m.ds,3), 'C00', 'C80') OR m.ds='C97') OR ;
  		(m.ds='D70') , .T., .F.)
   
    m.isdef = IIF(er_c<>'99', .T., .F.)
    
  	IF !m.IsOnkDs
  	 LOOP 
  	ENDIF 
  	
  	INSERT INTO exps FROM MEMVAR 

   ENDSCAN 
   USE IN mefile
  ENDFOR 
  ENDIF 

  USE IN aisoms 

 ENDFOR 
 
 m.docname = m.pbase+'\'+m.gcperiod+'\onk_pril2.xls'
 
 SELECT paz 
 dimdata(1,1) = RECCOUNT('paz')
 COUNT FOR ds_onk=1 TO dimdata(1,2)
 dimdata(1,3) = dimdata(1,1) - dimdata(1,2)
 COUNT FOR ds1_t TO dimdata(1,4)
 COUNT FOR ds1_t AND dsp TO dimdata(1,5)
 COUNT FOR INLIST(dn,1,2) TO dimdata(1,6)
 COUNT FOR INLIST(dn,1,2) AND MONTH(d_u)=m.tMonth TO dimdata(1,21)
 COUNT TO dimdata(1,7)
 *COUNT FOR k_def=1 TO dimdata(1,7)
 *SUM k_u FOR ch_t TO dimdata(1,8)
 COUNT FOR ch_t TO dimdata(1,8)

 SUM k_def TO dimdata(1,9)
 SUM k_def FOR ch_t TO dimdata(1,10)
 
 dimdata(1,11) = INT(dimdata(1,9) * 1.5)
 dimdata(1,12) = INT(dimdata(1,10) * 1.4)
 
 COPY TO &pBase\&gcPeriod\onk_pril2
 
 SELECT exps
 *BROWSE 
 
 *SELECT tip, isxt, coun(*) as cnt FROM exps GROUP BY tip,isxt INTO CURSOR expss
 *USE IN exps

 COUNT FOR tip=1 TO dimdata(1,13) && Ã››
 COUNT FOR tip=1 AND isxt TO dimdata(1,14) && Ã››
 COUNT FOR tip=2 TO dimdata(1,15) && › Ãœ
 COUNT FOR tip=2 AND isxt TO dimdata(1,16) && › Ãœ

 COUNT FOR tip=1 AND isdef TO dimdata(1,17) && Ã››
 COUNT FOR tip=1 AND isxt AND isdef TO dimdata(1,18) && Ã››
 COUNT FOR tip=2  AND isdef TO dimdata(1,19) && › Ãœ
 COUNT FOR tip=2 AND isxt AND isdef TO dimdata(1,20) && › Ãœ

 USE IN exps

 m.llResult = X_Report(m.dotname, m.docname+'.xls', .T.)
 USE IN curss
 
 USE IN paz

 USE IN dspcodes
 USE IN gr_plan

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

  m.IsOnkDs = IIF((BETWEEN(LEFT(m.ds,3), 'C00', 'C80') OR m.ds='C97') OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
  m.ds_onk = IIF(FIELD('ds_onk')=UPPER('ds_onk'), ds_onk, 0)

  IF !m.IsOnkDs AND m.ds_onk<>1
   LOOP 
  ENDIF 

  m.d_u    = d_u
  m.sn_pol = sn_pol
  m.ds_onk = IIF(FIELD('ds_onk')=UPPER('ds_onk'), ds_onk, 0)
  m.dn     = IIF(FIELD('dn')=UPPER('dn'), dn, 0)
  
  m.is_dsp = IIF(SEEK(m.cod, 'dspcodes'), .T., .F.)
  
  IF !SEEK(m.sn_pol, 'paz')
   INSERT INTO paz (sn_pol, d_u, ds, ds_2, ds_onk, d_onk, ds1_t, dsp, dn, k_u, ch_t, k_def) VALUES ;
   	(m.sn_pol, m.d_u, m.ds, m.ds_2, m.ds_onk, IIF(m.ds_onk=1, m.d_u, {}), m.ds1_t, m.is_dsp, m.dn, 1, m.ch_t, IIF(m.is_def, 1, 0))
  ELSE 
   m.ok_u   = paz.k_u
   m.ok_def = paz.k_def
   UPDATE paz SET k_u = m.ok_u+1 WHERE sn_pol=m.sn_pol
   IF m.is_def
    UPDATE paz SET k_def = m.ok_def+1 WHERE sn_pol=m.sn_pol
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