FUNCTION OneFlk_2(ppath)
 IF !fso.FolderExists(ppath)
  RETURN 
 ENDIF 
 IF !fso.FileExists(ppath+'\people.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(ppath+'\talon.dbf')
  RETURN 
 ENDIF 

 m.cfrom = ALLTRIM(cfrom)

 m.IsPr = IsPr
 *IF IsPr
 * =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
 * RETURN 
 *ENDIF 

 m.lpuid    = lpuid
 m.mcod     = mcod
 m.IsStomat = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
 m.IsIskl   = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)
 m.lpuname  = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')+', '+sprlpu.cokr+', '+sprlpu.mcod
 m.period   = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 m.dat1     = CTOD('01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4))
 m.dat2     = GOMONTH(m.dat1,1)-1
 *m.IsStac = IIF(VAL(SUBSTR(m.mcod,3,2))<41 or m.IsLpuTpn, .F., .T.)
  
 o_rec = RECNO('aisoms')
 IF !OpBase(ppath)
  =ClBase()
  RETURN .f.
 ENDIF 
 SELECT aisoms 
 GO (o_rec)
  
 IF m.IsPr
  CREATE CURSOR AllGood (sn_pol c(25))
  SELECT AllGood
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO sn_pol 
  
  CREATE CURSOR AllBad (sn_pol c(25))
  SELECT AllBad
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO sn_pol 

  SELECT Talon
  SET RELATION TO recid INTO serror
  SCAN 
   IF !EMPTY(serror.c_err) AND serror.c_err<>'PKA'
    LOOP 
   ENDIF 

   m.sn_pol = sn_pol
   IF !SEEK(m.sn_pol, 'allgood')
    INSERT INTO AllGood FROM MEMVAR 
   ENDIF 
  ENDSCAN 
  SET RELATION OFF INTO serror
  
  
  SELECT people
  SCAN 
   m.sn_pol = sn_pol
   IF !SEEK(m.sn_pol, 'allgood')
    m.recid = recid
    IF !SEEK(m.RecId, 'rError')
     =InsError('R', 'PNA', m.recid)
    ENDIF 
    IF !SEEK(m.sn_pol, 'allbad')
     INSERT INTO AllBad FROM MEMVAR 
    ENDIF 
   ENDIF 
  ENDSCAN 
  SET ORDER TO recid

  SELECT rerror
  SET RELATION TO rid INTO people 
  SCAN 
   m.c_err  = c_err
   m.sn_pol = people.sn_pol
   IF !SEEK(m.sn_pol, 'allbad') AND m.c_err='PNA'
    DELETE 
   ENDIF 
  ENDSCAN 
  SET RELATION OFF INTO people 
  
  USE IN allgood
  USE IN allbad 

  =ClBase()
  =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
  RETURN 
 ENDIF 

 m.s_flk  = 0  
 m.s_flk2 = 0  
 m.ls_flk = 0

 SELECT talon
 SET RELATION TO sn_pol INTO people 
  
 SCAN
  DO ss_flk_2
 ENDSCAN
  
 CREATE CURSOR AllBad (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol 
 SELECT talon 
 SET ORDER TO sn_pol && !!!
 GO TOP 
  
 SELECT sn_pol DISTINCT  FROM talon WHERE recid NOT IN ;
  	(SELECT rid FROM serror WHERE f='S') INTO CURSOR AllGood READWRITE 
 SELECT AllGood
 INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  
 SELECT talon 
 SET RELATION TO sn_pol INTO AllGood ADDITIVE 
 SCAN 
  IF !EMPTY(AllGood.sn_pol)
   LOOP 
  ENDIF 
  m.polis = sn_pol
  INSERT INTO Allbad (sn_pol) VALUES (m.polis)
 ENDSCAN 
 SET RELATION OFF INTO AllGood
 USE IN AllGood
 
 SELECT People
 SCAN 
  m.polis = sn_pol
  m.recid = recid
  IF !SEEK(m.polis, 'allbad')
   LOOP 
  ENDIF 
  IF !SEEK(RecId, 'rError')
    =InsError('R', 'PNA', m.recid)
  ENDIF 
 ENDSCAN 
 USE IN AllBad
  
 SELECT talon 
 SET RELATION TO recid INTO serror ADDITIVE 
 SUM s_all FOR !EMPTY(serror.rid) TO m.s_flk
 SUM s_all FOR !EMPTY(serror.rid) AND serror.et=2 TO m.s_flk2
 SUM s_lek FOR !EMPTY(serror.rid) TO m.ls_flk
 SET RELATION OFF INTO serror
 SET RELATION OFF INTO people
  
 =ClBase()

 SELECT AisOms
 =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
 SELECT AisOms

 IF sum_flk != m.s_flk
  m.b_flk = pbase+'\'+m.gcperiod+'\'+mcod+'\b_flk_'+mcod
  IF fso.FileExists(m.b_flk)
   fso.DeleteFile(m.b_flk)
  ENDIF 
  m.b_mek = pbase+'\'+m.gcperiod+'\'+mcod+'\b_mek_'+mcod
  IF fso.FileExists(m.b_mek)
   fso.DeleteFile(m.b_mek)
  ENDIF 
  
  ** Удаляем все файлы: протокол, акт, реестр актов, табличную форму актов
   
  m.l_path = pbase+'\'+m.gcperiod+'\'+mcod
  m.mmy    = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
   
  DIMENSION dim_files(5)
  dim_files(1) = "Pr"+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
  dim_files(2) = "Mk" + STR(m.lpuid,4) + m.qcod + m.mmy
   dim_files(3) = "Mt" + STR(m.lpuid,4) + m.qcod + m.mmy
  dim_files(4) = "Mc" + STR(m.lpuid,4) + m.qcod + m.mmy
  dim_files(5) = 'pdf'+m.qcod+m.mmy
   
  FOR i=1 TO ALEN(dim_files,1)
   IF fso.FileExists(m.l_path+'\'+ALLTRIM(dim_files(i))+'.xls')
    fso.DeleteFile(m.l_path+'\'+ALLTRIM(dim_files(i))+'.xls')
   ENDIF 
   IF fso.FileExists(m.l_path+'\'+ALLTRIM(dim_files(i))+'.pdf')
    fso.DeleteFile(m.l_path+'\'+ALLTRIM(dim_files(i))+'.pdf')
   ENDIF 
  ENDFOR 
   
  RELEASE dim_files, l_path

  ** Удаляем все файлы: протокол, акт, реестр актов, табличную форму актов
 ENDIF 

 REPLACE sum_flk WITH m.s_flk, sum_flk2 WITH m.s_flk2, ls_flk WITH m.ls_flk
 
 WAIT CLEAR 

RETURN 

FUNCTION InsError(WFile, cError, cRecId, cDetail, cComment)
 IF PARAMETERS()<5
  cComment = ''
 ENDIF 
 IF PARAMETERS()<4
  cDetail = ''
 ENDIF 
 IF WFile == 'R'
  IF !SEEK(cRecId, 'rError')
   INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('R', 2, cError, cRecId, cDetail, cComment)
  ELSE 
  ENDIF !SEEK(cRecId, 'rError')
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(cRecId, 'sError')
   INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('S', 2, cError, cRecId, cDetail, cComment)
   RETURN .T.
  ELSE 
   IF cError != sError.c_err
    INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('S', 2, cError, cRecId, cDetail, cComment)
   ENDIF cError != sError.c_err 
  ENDIF !SEEK(cRecId, 'sError')
 ENDIF 
RETURN .F.

FUNCTION OpBase(ppath)
 tnresult = 0
 tnresult = tnresult + OpenFile(pcommon+'\dspcodes', 'dspcodes', 'share', 'cod')

 IF tmonth>1
  m.dspfile = pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\dsp'
 ELSE
  m.dspfile = pbase +'\'+ STR(tyear-1,4)+'12'+'\dsp'
 ENDIF 
 IF OpenFile(m.dspfile, 'dsp', 'shar', 'exptag')>0
  IF USED('dsp')
   USE IN dsp 
  ENDIF 
  RETURN .F.
 ENDIF 

 IF !fso.FileExists(pbase+'\'+gcperiod+'\disp.dbf') && аналог файла dsp, но за текущий период
  IF MakeDisp()
  ELSE 
   RETURN .F.
  ENDIF 
 ELSE 
  IF OpenFile(pbase+'\'+gcperiod+'\disp', 'disp', 'shar', 'exptag')>0
   RETURN .F.
  ENDIF 
 ENDIF 
 * файл disp остается открытым!
 
 SELECT * FROM dsp INTO CURSOR dspp READWRITE 
 USE IN dsp
 SELECT dspp
 INDEX on sn_pol+STR(tip,1) TAG exptag
 SET ORDER TO exptag
 
 APPEND FROM &pbase\&gcperiod\disp

 IF !fso.FileExists(pbase+'\'+gcperiod+'\gosp.dbf')
  IF MakeGsp()
  ELSE 
   RETURN .F.
  ENDIF 
  ** Файл остается открытым
 ELSE 
  IF OpenFile(pbase+'\'+gcperiod+'\gosp', 'gosp', 'shar', 'sn_pol')>0
   RETURN .F.
  ENDIF 
 ENDIF 

 ** Аналог polic_h, но за текущий период
 IF !fso.FileExists(pbase+'\'+gcperiod+'\p_h.dbf')
  IF MakeP_H()
  ELSE 
   RETURN .F.
  ENDIF 
  ** Файл остается открытым
 ELSE 
  IF OpenFile(pbase+'\'+gcperiod+'\p_h', 'polic_h', 'shar', 'sn_pol')>0
   RETURN .F.
  ENDIF 
 ENDIF 
 ** Аналог polic_h, но за текущий период

 tnresult = tnresult + OpenFile(ppath+'\people', 'people', 'share', 'sn_pol')
 tnresult = tnresult + OpenFile(ppath+'\talon', 'talon', 'share')
 
 *IF tmonth>1
 * m.deads = pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\deads'
 * m.stop = pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\stop'
 *ELSE
 * m.deads = pbase +'\'+ STR(tyear-1,4)+'12'+'\deads'
 * m.stop = pbase +'\'+ STR(tyear-1,4)+'12'+'\stop'
 *ENDIF 

 *IF fso.FileExists(m.deads+'.dbf')
 * tnresult = tnresult + OpenFile(m.deads, 'deads', 'share', 'sn_pol')
 *ENDIF 
 *IF fso.FileExists(m.stop+'.dbf')
 * tnresult = tnresult + OpenFile(m.stop, 'stop', 'share', 'enp')
 *ENDIF 

 IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\deads.dbf')
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\deads.dbf', 'deads', 'share', 'sn_pol')>0
   IF USED('deads')
    USE IN deads 
   ENDIF 
  ENDIF 
 ENDIF 

 tnresult = tnresult + OpenFile(ppath+'\e'+m.mcod, 'rerror', 'share', 'rrid')
 tnresult = tnresult + OpenFile(ppath+'\e'+m.mcod, 'serror', 'share', 'rid', 'again')
 *tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\polic_h', 'polic_h', 'share', 'sn_pol')
 
 IF !m.IsPr
  UPDATE rerror SET Tip=1, dt=DATETIME(), usr=m.gcUser WHERE SUBSTR(c_err,3,1)='A'
  DELETE FOR SUBSTR(c_err,3,1)='A' AND IIF(FIELD('et')='ET', et=2, 0=1) IN rerror

  UPDATE serror SET Tip=1, dt=DATETIME(), usr=m.gcUser WHERE SUBSTR(c_err,3,1)='A'
  DELETE FOR SUBSTR(c_err,3,1)='A' AND IIF(FIELD('et')='ET', et=2, 0=1) IN serror
 ENDIF 
RETURN .t.

FUNCTION ClBase()
 IF USED('talon')
  USE IN talon 
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('rerror')
  USE IN rerror
 ENDIF 
 IF USED('serror')
  USE IN serror
 ENDIF 
 IF USED('gosp')
  USE IN Gosp
 ENDIF 
 IF USED('polic_h')
  USE IN polic_h
 ENDIF 
 IF USED('disp')
  USE IN disp
 ENDIF 
 IF USED('dsp')
  USE IN dsp
 ENDIF 
 IF USED('dspp')
  USE IN dspp
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 
 IF USED('deads')
  USE IN deads 
 ENDIF 
 IF USED('stop')
  USE IN stop 
 ENDIF 
RETURN 

FUNCTION MakeDisp
 PRIVATE mcod
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ ФАЙЛ DISP-ФАЙЛ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN .F.
 ENDIF 
 IF !USED('dspcodes')
  RETURN .F.
 ENDIF 

 m.lcpath = pbase+'\'+m.gcperiod
 m.period = m.gcperiod

 CREATE TABLE &pBase\&gcPeriod\disp (recid i, period c(6), mcod c(7), sn_pol c(17), c_i c(30), ds c(6),;
 	fam c(20), im c(20), ot c(20), w n(1), dr d, ages n(2), cod n(6), rslt n(3), ;
 	d_u d, s_all n(11,2), k_u2 n(3), s_all2 n(11,2), k_u2ok n(3), s_all2ok n(11,2), er c(3), tip n(1))
 INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq && используется в formdds
 INDEX on sn_pol+PADL(tip,1,'0') TAG exptag
 INDEX on sn_pol+PADL(cod,6,'0') TAG un_tag
 INDEX on c_i TAG c_i 
 SET ORDER TO uniqq

 SELECT disp 
 SET RELATION TO cod INTO dspcodes 
 REPLACE tip WITH dspcodes.tip ALL
 SET RELATION OFF INTO  dspcodes 
 
 SELECT aisoms
 o_f = SET("Filter")
 IF !EMPTY(o_f)
  SET FILTER TO 
 ENDIF 
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF RECCOUNT('talon')<=0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('err')
    USE IN err
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
 
  WAIT "ОБРАБОТКА..." WINDOW NOWAIT 
  SELECT talon
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO err ADDITIVE 
  m.nusl = 0
  SCAN 
   m.cod = cod
   IF !SEEK(m.cod, 'dspcodes')
    LOOP
   ENDIF 
  
   m.tipofcod = dspcodes.tip
   m.rslt     = rslt
  
   DO CASE 
    CASE m.tipofcod = 1 && Диспасеризация взрослых, с июля
     IF !(INLIST(m.rslt,317,318,353,357,358) OR BETWEEN(m.rslt,355,356)) && !BETWEEN(m.rslt,316,319) AND !BETWEEN(m.rslt,352,358) 
      LOOP 
     ENDIF 
    CASE m.tipofcod = 2 && Профосмотры взрослых, с сентября
     *IF !BETWEEN(m.rslt,343,345) AND !INLIST(m.rslt,373,374)
     IF !BETWEEN(m.rslt,343,344) AND !INLIST(m.rslt,373,374)
      LOOP 
     ENDIF 
    CASE m.tipofcod = 3 && Диспасеризация детей-сирот, с июля
     IF !(BETWEEN(m.rslt,347,351) OR BETWEEN(m.rslt,369,372)) AND ;
     	!(BETWEEN(m.rslt,321,325) OR BETWEEN(m.rslt,365,368)) && !BETWEEN(m.rslt,321,325) AND m.rslt!=320 AND m.rslt!=390 AND !BETWEEN(m.rslt,347,351) && Это - усыновленные сироты!
    * IF !BETWEEN(m.rslt,321,325) AND m.rslt!=320 AND m.rslt!=390 AND !BETWEEN(m.rslt,347,351) && Это - усыновленные сироты!
      LOOP 
     ENDIF 
    CASE m.tipofcod = 4 && Профосмотры несовершеннолетних, с сентября
     *IF !BETWEEN(m.rslt,332,336) AND m.rslt!=326
     IF !(BETWEEN(m.rslt,332,336) OR BETWEEN(m.rslt,361,364)) && !BETWEEN(m.rslt,332,336) AND m.rslt!=326
      LOOP 
     ENDIF 
    *CASE m.tipofcod = 5 && Предварительные профосмотры несовершеннолетних, с сентября
    * IF !BETWEEN(m.rslt,337,341) AND !INLIST(m.rslt,326,396)
    *  LOOP 
    * ENDIF 
    *CASE m.tipofcod = 6 && Периодические профосмотры несовершеннолетних, с сентября
    * IF m.rslt!=342  AND m.rslt!=326
    *  LOOP 
    * ENDIF 
    *CASE m.tipofcod = 7 && профосмотры, с сентября
    OTHERWISE 
     LOOP 
   ENDCASE 

   m.nusl = m.nusl + 1

   m.mcod  = mcod
   m.recid = recid
   m.key   = m.gcperiod + m.mcod + PADL(m.recid,6,'0')
   IF SEEK(m.key, 'disp')
    LOOP 
   ENDIF 

   m.d_u   = d_u
   m.k_u   = k_u
   m.s_all = s_all
   m.er    = ''
   m.ds    = ds

   m.id_smo = recid
   m.sn_pol = sn_pol
   m.fam    = people.fam
   m.im     = people.im
   m.ot     = people.ot
   m.w      = people.w
   m.dr     = people.dr
   m.er     = err.c_err
   
   m.ages   = YEAR(tdat1) - YEAR(m.dr)
  
   m.c_i    = c_i 
   
   m.tip = m.tipofcod

   INSERT INTO disp FROM MEMVAR 
   
  ENDSCAN 
  SET RELATION OFF INTO people
  SET RELATION OFF INTO err

  USE IN err
  USE IN talon
  USE IN people

  SELECT aisoms
  WAIT CLEAR 
 ENDSCAN && aisoms
 
 MESSAGEBOX(CHR(13)+CHR(10)+'ОБНАРУЖЕНО '+TRANSFORM(m.nusl, '99999')+' УСЛУГ!',0+64,'')
 
 WAIT "РАСЧЕТ ВТОРОГО ЭТАПА (НОВЫЙ АЛГОРИТМ)..." WINDOW NOWAIT 
 SELECT disp
 SET ORDER TO c_i
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF RECCOUNT('talon')<=0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('err')
    USE IN err
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
 
  SELECT talon 
  SET ORDER TO c_i 
  SET RELATION TO LEFT(c_i,25) INTO disp 
  SET RELATION TO recid INTO err ADDITIVE 

  m.oc_i     = c_i 
  m.k_u2     = 0
  m.s_all2   = 0
  m.k_u2ok   = 0
  m.s_all2ok = 0

  SCAN 
   IF EMPTY(disp.c_i)
    LOOP 
   ENDIF 
   m.rslt = disp.rslt 
   IF !INLIST(m.rslt,320,326,352,353,357,358,390,396)
    LOOP 
   ENDIF 
   IF disp.k_u2>0
    LOOP 
   ENDIF 	

   IF d_u < disp.d_u 
    LOOP 
   ENDIF 
   IF cod = disp.cod 
    LOOP 
   ENDIF 
  
   IF c_i != m.oc_i
    IF m.k_u2>0
     REPLACE k_u2 WITH m.k_u2, s_all2 WITH m.s_all2, k_u2ok WITH m.k_u2ok, s_all2ok WITH m.s_all2ok IN disp 
    ENDIF 
   
    m.oc_i = c_i 
   
    m.k_u2     = 0
    m.s_all2   = 0
    m.k_u2ok   = 0
    m.s_all2ok = 0
   ENDIF 
  
   m.k_u2   = m.k_u2 + k_u
   m.s_all2 = m.s_all2 + s_all
    
   IF EMPTY(err.rid)
    m.k_u2ok   = m.k_u2ok + k_u
    m.s_all2ok = m.s_all2ok + s_all
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO disp 
  SET RELATION OFF INTO err
  WAIT CLEAR 

  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('err')
   USE IN err
  ENDIF 

 ENDSCAN 
 IF !EMPTY(o_f)
  SET FILTER TO &o_f
 ENDIF 

 SELECT disp
 SET ORDER TO exptag
 
 MESSAGEBOX('ФОРМИРОВАНИЕ ФАЙЛА DISP ЗАКОНЧЕНО!',0+64,'')

RETURN .T.

FUNCTION MakeGsp
 LOCAL mcod
 IF MESSAGEBOX('СФОРМИРОВАТЬ СВОДНЫЙ ФАЙЛ ГОСПИТАЛИЗАЦИЙ?', 4+32, '')=7
  RETURN .F.
 ENDIF 

 CREATE CURSOR Gosp (recid i, mcod c(7), sn_pol c(25), c_i c(30), cod n(6), d_u d, k_u n(3))
 
 SELECT aisoms
 o_f = SET("Filter")
 IF !EMPTY(o_f)
  SET FILTER TO 
 ENDIF 
 SCAN 
  m.mcod = mcod 
  IF !IsStac(m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FolderExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 

  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod + '...' WINDOW NOWAIT 
  
  SELECT talon 
  SCAN 
   SCATTER MEMVAR 
   IF !IsMes(m.cod) AND !IsVmp(m.cod)
    LOOP 
   ENDIF 
   IF INT(m.cod/1000)=297
    LOOP 
   ENDIF 
   INSERT INTO Gosp FROM MEMVAR 
  ENDSCAN 
  USE 
  
  WAIT CLEAR 
 
  SELECT aisoms 
  
 ENDSCAN 
 IF !EMPTY(o_f)
  SET FILTER TO &o_f
 ENDIF 
 
 SELECT mcod,sn_pol,c_i,MAX(d_u) as d_u,MAX(cod) as cod,SUM(k_u) as k_u FROM gosp ;
	GROUP BY sn_pol,c_i, mcod ORDER BY mcod,sn_pol,c_i INTO CURSOR Gsp

 USE IN Gosp
 SELECT Gsp
 INDEX ON sn_pol TAG sn_pol && FOR IsMes(cod) OR IsVMP(cod)
 COPY TO &pBase\&gcPeriod\Gosp WITH cdx 
 USE IN Gsp 
 
 SELECT aisoms
 
 MESSAGEBOX('СВОДНЫЙ ФАЙЛ ГОСПИТАЛИЗАЦИЙ СОБРАН!', 0+64, '')

RETURN .T.

PROCEDURE MakeP_H
 LOCAL mcod
 IF MESSAGEBOX('СОБРАТЬ ЛОКАЛЬНЫЙ ФАЙЛ "ЦЕНТР ЗДОРОВЬЯ"',4+32,'')=7
  RETURN .F.
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\p_h.dbf')
  CREATE TABLE &pBase\&gcPeriod\p_h (sn_pol c(25), mcod c(7), d_u d, cod n(6))
  INDEX on sn_pol TAG sn_pol
  USE 
 ENDIF 
 
 IF OpenFile(pCommon+'\lpu_cz', 'lpu_cz', 'shar', 'mcod')>0
  IF USED('lpu_cz')
   USE IN lpu_cz
  ENDIF 
  RETURN .F.
 ENDIF 

 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\p_h', 'polic_h', 'shar', 'sn_pol')>0
  IF USED('polic_h')
   USE IN polic_h
  ENDIF 
  USE IN lpu_cz
  RETURN .F.
 ENDIF 
 
 m.lAisUsed = .T.
 IF !USED('aisoms')
 m.lAisUsed = .F.
 IF !fso.FolderExists(pBase+'\'+m.gcperiod)
  USE IN polic_h
  USE IN lpu_cz
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcperiod+'\aisoms.dbf')
  USE IN polic_h
  USE IN lpu_cz
  RETURN .F.
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  USE IN polic_h
  USE IN lpu_cz
  RETURN .F.
 ENDIF 
 ENDIF 
  
 WAIT 'СОБИРАЮ...' WINDOW NOWAIT 
 SELECT aisoms
 o_f = SET("Filter")
 IF !EMPTY(o_f)
  SET FILTER TO 
 ENDIF 
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(pBase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !SEEK(m.mcod, 'lpu_cz')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
   
  SELECT talon 
  SCAN 
   m.cod = cod 
   IF !INLIST(m.cod,15001,115001)
    LOOP 
   ENDIF 
   m.sn_pol = sn_pol
   m.d_u    = d_u
   INSERT INTO polic_h FROM MEMVAR 
  ENDSCAN 
  USE IN talon 
  SELECT aisoms

 ENDSCAN
 IF !EMPTY(o_f)
  SET FILTER TO &o_f
 ENDIF 
 IF m.lAisUsed = .F.
 USE IN aisoms 
 ENDIF 
 WAIT CLEAR 
  
 USE IN lpu_cz
 
RETURN .T.