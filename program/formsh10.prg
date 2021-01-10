PROCEDURE FormSh10
 IF MESSAGEBOX(CHR(13)+CHR(10)+'�������� � ������� ������� ���������'+CHR(13)+CHR(10)+;
 	'��������� �������������� (I60-I64)?',4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\FormSh09.xls')
  MESSAGEBOX('����������� ���� FormSh09.xls!',0+64,'')
  RETURN 
 ENDIF 
 
 ppath = pbase+'\'+m.gcperiod
 IF !fso.FileExists(ppath+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'����������� ���� AISOMS.DBF'+CHR(13)+CHR(10),0+16,m.gcperiod)
  RETURN 
 ENDIF 
 
 IF OpenFile(ppath+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
 ENDIF 
 IF OpenFile(ppath+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  USE IN aisoms
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
 ENDIF 
 
 CREATE CURSOR crs (gcperiod c(6), mcod c(7), lpuname c(120), c_i c(25), sn_pol c(25), d_u d, k_u n(3), cod n(6), ds c(6), profil c(3),;
 	s_all n(11,2), s_1 n(11,2), s_2 n(11,2), fil_id n(6), et c(1), koeff n(4,2), straf n(4,2), e_period c(6), w0 n(3),;
 	werr n(3))
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  IF INT(VAL(SUBSTR(m.mcod,3,2)))<41
*   SKIP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+m.gcperiod)
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   USE IN talon
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod++'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('err')
    USE IN err
   ENDIF 
   USE IN people
   USE IN talon 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod++'\m'+m.mcod, 'merr', 'shar', 'rid')>0
   IF USED('merr')
    USE IN merr
   ENDIF 
   USE IN people
   USE IN talon 
   USE IN merr
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT talon 
  SET RELATION TO recid INTO err
  SET RELATION TO recid INTO merr ADDITIVE 
  SET RELATION TO sn_pol INTO people ADDITIVE 
  SCAN 
   IF !EMPTY(err.rid)
    LOOP 
   ENDIF 
   m.tip = tip 
   IF EMPTY(m.tip)
    LOOP 
   ENDIF 
   m.ds = ds
   IF !BETWEEN(m.ds, 'I60','I64')
    LOOP 
   ENDIF 
   
   m.c_i    = c_i
   m.sn_pol = sn_pol
   m.d_u    = d_u 
   m.k_u    = k_u
   m.cod    = cod
   m.profil = profil
   m.fil_id = fil_id
   m.s_all  = s_all	
   
   m.et       = merr.et
   m.err_mee  = UPPER(LEFT(merr.err_mee,2))
   m.koeff    = merr.koeff
   m.straf    = merr.straf
   m.s_1      = merr.s_1
   m.s_2      = merr.s_2
   m.e_period = merr.e_period
   
   IF !EMPTY(m.err_mee)
    m.w0   = IIF(m.err_mee='W0',1,0)
    m.werr = IIF(m.err_mee='W0',0,1)
   ENDIF 
   
   INSERT INTO crs FROM MEMVAR 
   
  ENDSCAN 
  SET RELATION OFF INTO merr
  SET RELATION OFF INTO err 
  SET RELATION OFF INTO people
  USE IN merr 
  USE IN err 
  USE IN talon 
  USE IN people 
  
  WAIT CLEAR 
  
  SELECT aisoms 

 ENDSCAN 
 USE IN aisoms
 
 SELECT crs 
 COPY TO pmee+'\SH10'+SUBSTR(m.gcperiod,3)
 
 SELECT 000000 as recid, mcod, SPACE(120) as lpuname, profil, coun(*) as cnt, SUM(s_all) as s_all, SUM(s_1) as s_1, SUM(s_2) as s_2,;
 	SUM(w0) as w0, SUM(werr) as werr FROM crs;
 	GROUP BY mcod, profil INTO CURSOR curdata READWRITE 
 SELECT curdata
 SET RELATION TO mcod INTO sprlpu
 REPLACE ALL recid WITH RECNO(), lpuname WITH sprlpu.fullname
 SET RELATION OFF INTO sprlpu 
 USE IN sprlpu
 
 USE IN crs
 
 m.dotname = ptempl+'\FormSh10.xls'
 m.docname = pmee+'\SH10'+SUBSTR(m.gcperiod,3)
 IF fso.FileExists(m.docname+'.xls')
  fso.DeleteFile(m.docname+'.xls')
 ENDIF 
 m.llResult = X_Report(m.dotname, m.docname+'.xls', .T.)
 USE IN curdata

* MESSAGEBOX('��������� ���������!',0+64,'')
RETURN 