PROCEDURE MakeUDSt
 IF MESSAGEBOX(CHR(13)+	CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ UD-ÔÀÉËÛ?'+CHR(13)+CHR(10),4+32,'ÑÒÎÌÀÒÎËÎÃÈß')=7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilots', 'pilots', 'shar', 'lpu_id')>0
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpus', 'horlpus', 'shar', 'lpu_id')>0
  IF USED('horlpus')
   USE IN horlpus
  ENDIF 
  USE IN pilots
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN horlpus
  USE IN pilots
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN sprlpu
  USE IN horlpus
  USE IN pilots
  USE IN aisoms
  RETURN 
 ENDIF 

 
 dufile = pbase+'\'+m.gcperiod+'\udst'+m.qcod+m.gcperiod
 IF fso.FileExists(dufile + '.dbf')
  fso.DeleteFile(dufile + '.dbf')
 ENDIF 
 
 CREATE TABLE &dufile ;
  (mcod c(7), lpu_id n(6), prmcod c(7), prlpu_id n(6), sn_pol c(25), cod n(6), k_u n(3), pr_all n(12,2), vz n(1))
 USE 
 =OpenFile(dufile, 'dufile', 'shar')
 
 SELECT aisoms
 SCAN FOR !DELETED()

  m.lpuid    = lpuid
  m.mcod     = mcod 
  m.IsStomat = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
  m.IsIskl   = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)
  
  m.IsPilots  = SEEK(m.lpuid, 'pilots')
  m.IsHorLpus = SEEK(m.lpuid, 'horlpus')
  
  IF !m.IsPilots && Äîáàâèë 26.11.2018!
   IF !SEEK(m.lpuid, 'horlpus')
    LOOP 
   ENDIF 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('people')
    USE IN people 
   ENDIF 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')>0
   IF USED('error')
    USE IN error 
   ENDIF 
   IF USED('people')
    USE IN people 
   ENDIF 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 

  SELECT talon 
  SET RELATION TO recid INTO error
  SET RELATION TO sn_pol INTO people ADDITIVE 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  SCAN 
   IF !EMPTY(error.rid)
    LOOP 
   ENDIF 
   m.Mp  = Mp
   m.Typ = Typ
   IF !(m.Typ='2' AND m.Mp='s')
    LOOP 
   ENDIF 
   
   m.vz = vz
   IF m.vz<=0
    LOOP 
   ENDIF 
   *IF m.vz=4
   * LOOP 
   *ENDIF 
   
   m.prmcod   = people.prmcods
   m.prlpuid  = IIF(SEEK(m.prmcod, 'pilots', 'mcod'), pilots.lpu_id, 0)
   m.cod     = cod
   m.k_u     = k_u
   m.s_all   = s_all
   m.sn_pol = sn_pol

   INSERT INTO dufile (mcod, lpu_id, prmcod, prlpu_id, cod, k_u, pr_all, sn_pol, vz) VALUES ;
   	(m.mcod, m.lpuid, m.prmcod, m.prlpuid, m.cod, m.k_u, m.s_all, m.sn_pol, m.vz)

  ENDSCAN 

  SET RELATION OFF INTO people
  SET RELATION OFF INTO error
  USE IN people
  USE IN talon 
  USE IN error
  
  WAIT CLEAR 
  
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 
 SELECT dufile
 SELECT prmcod DISTINCT FROM dufile INTO CURSOR svcur
 IF _tally<=0
  USE IN svcur
  USE IN dufile
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\udqqxxxx.dbf')
  USE IN svcur
  USE IN dufile
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÒÓÒÑÒÂÓÅÒ ØÀÁËÎÍ UDQQXXXX.DBF!'+CHR(13)+CHR(10),0+16,'Templates')
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\pr4st.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\pr4st', 'pr4st', 'shar', 'mcod')<=0
   SELECT prmcod as mcod, SUM(pr_all) as s_all FROM dufile GROUP BY prmcod INTO CURSOR curstat
   SELECT curstat
   SET RELATION TO mcod INTO pr4st
   m.sumstat = 0
   m.sumpr4  = 0
   SCAN 
    m.mcod = mcod
    IF s_all != pr4st.s_others
     MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÍÀÐÓÆÅÍÎ ÐÀÑÕÎÆÄÅÍÈÅ ÌÅÆÄÓ ÔÀÉËÎÌ PR4ST È'+CHR(13)+CHR(10)+;
      'È ÑÔÎÐÌÈÐÎÂÀÍÍÛÌÈ UD-ÔÀÉËÀÌÈ!'+CHR(13)+CHR(10)+;
      'ÏÎ ÄÀÍÍÛÌ PR4ST ÑÓÌÌÀ S_OTHERS='+TRANSFORM(pr4st.s_others,'99999999.99')+CHR(13)+CHR(10)+;
      'ÏÎ ÄÀÍÍÛÌ UD-ÔÀÉËÎÂ ÑÓÌÌÀ='+TRANSFORM(s_all,'99999999.99')+CHR(13)+CHR(10),0+48, m.mcod)
    ENDIF 
    m.sumstat = m.sumstat + s_all
    *m.sumpr4  = m.sumpr4 + pr4.s_others + pr4st.s_others
    m.sumpr4  = m.sumpr4 + pr4st.s_others
   ENDSCAN 
   SET RELATION OFF INTO pr4st
   USE IN pr4st
   USE IN curstat
   IF m.sumstat!=m.sumpr4
    MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÍÀÐÓÆÅÍÎ ÐÀÑÕÎÆÄÅÍÈÅ ÌÅÆÄÓ ÔÀÉËÎÌ PR4ST È'+CHR(13)+CHR(10)+;
     'È ÑÔÎÐÌÈÐÎÂÀÍÍÛÌÈ UD-ÔÀÉËÀÌÈ!'+CHR(13)+CHR(10)+;
     'ÏÎ ÄÀÍÍÛÌ PR4ST ÑÓÌÌÀ S_OTHERS='+TRANSFORM(m.sumpr4,'99999999.99')+CHR(13)+CHR(10)+;
     'ÏÎ ÄÀÍÍÛÌ UD-ÔÀÉËÎÂ ÑÓÌÌÀ='+TRANSFORM(m.sumstat,'99999999.99')+CHR(13)+CHR(10),0+48,'')
   ELSE 
    MESSAGEBOX(CHR(13)+CHR(10)+'ÑÓÌÌÀ S_OTHERS ÔÀÉËÀ PR4ST'+CHR(13)+CHR(10)+;
     'ÑÎÎÒÂÅÒÑÒÂÓÅÒ ÑÓÌÌÅ UD-ÔÀÉËÎÂ'+CHR(13)+CHR(10)+;
     'È ÑÎÑÒÀÂËßÅÒ '+TRANSFORM(m.sumstat,'99999999.99')+CHR(13)+CHR(10),0+64,'')
   ENDIF 
  ELSE 
   IF USED('pr4st')
    USE 
   ENDIF 
   MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÓÄÀËÎÑÜ ÎÒÊÐÛÒÜ ÔÀÉË PR4ST,'+CHR(13)+CHR(10)+;
    'ÑÂÅÐÊÀ ÄÀÍÍÛÕ ÍÅ ÏÐÎÈÇÂÎÄÈÒÑß!',0+64,'')
  ENDIF 
 ELSE 
  MESSAGEBOX(CHR(13)+CHR(10)+'ÔÀÉË PR4ST ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ,'+CHR(13)+CHR(10)+;
   'ÑÂÅÐÊÀ ÄÀÍÍÛÕ ÍÅ ÏÐÎÈÇÂÎÄÈÒÑß!',0+64,'')
 ENDIF 
 
 SELECT svcur
 SCAN 
  m.prmcod  = prmcod 
  m.lppid  = IIF(SEEK(m.prmcod, 'pilots', 'mcod'), pilots.lpu_id, 0)
  IF m.lppid<=0
   LOOP
  ENDIF 
  m.uddfile = 'UDST'+UPPER(m.qcod)+PADL(m.lppid,4,'0')
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.prmcod)
   LOOP 
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
  ENDIF 

  fso.CopyFile(ptempl+'\udqqxxxx.dbf', pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
  
  =OpenFile(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile, 'udfile', 'shar')
  
  SELECT * FROM dufile WHERE prmcod = m.prmcod INTO CURSOR curmmm
  SELECT curmmm
  m.mcod   = mcod
  m.lpu_id = lpu_id
  m.period = m.gcperiod
  SCAN 
   m.cod    = cod 
   m.k_u    = k_u
   m.pr_all = pr_all
   m.lpu_id = lpu_id
   m.vz     = vz
   *IF m.vz=1
    INSERT INTO udfile FROM MEMVAR 
   *ENDIF 
  ENDSCAN 
  USE IN curmmm
  SELECT udfile
  REPLACE ALL recid WITH PADL(RECNO(),7,'0')
  USE IN udfile
  
  SELECT svcur
  
 ENDSCAN 
 USE IN svcur
 USE IN horlpus 
 USE IN pilots
 USE IN dufile
 USE IN sprlpu
 USE IN tarif

 MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 