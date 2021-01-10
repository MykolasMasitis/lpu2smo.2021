PROCEDURE MakeUDFiles
 IF MESSAGEBOX(CHR(13)+	CHR(10)+'¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹ UD-‘¿…À€?'+CHR(13)+CHR(10),4+32,'')=7
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
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilots', 'pilots', 'shar', 'lpu_id')>0
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\lputpn', 'lputpn', 'shar', 'lpu_id')>0
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  USE IN pilot
  USE IN pilots
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN lputpn
  USE IN pilot
  USE IN pilots
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpu', 'horlpu', 'shar', 'lpu_id')>0
  IF USED('horlpu')
   USE IN horlpu
  ENDIF 
  USE IN tarif
  USE IN lputpn
  USE IN pilot
  USE IN pilots
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpus', 'horlpus', 'shar', 'lpu_id')>0
  IF USED('horlpus')
   USE IN horlpus
  ENDIF 
  USE IN horlpu
  USE IN tarif
  USE IN lputpn
  USE IN pilot
  USE IN pilots
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN horlpu
  USE IN horlpus
  USE IN tarif
  USE IN lputpn
  USE IN pilot
  USE IN pilots
  USE IN aisoms
  RETURN 
 ENDIF 

 
 dufile = pbase+'\'+m.gcperiod+'\ud'+m.qcod+m.gcperiod
 IF fso.FileExists(dufile + '.dbf')
  fso.DeleteFile(dufile + '.dbf')
 ENDIF 
 
 CREATE TABLE &dufile ;
  (mcod c(7), lpu_id n(6), prmcod c(7), prlpu_id n(6), cod n(6), k_u n(3), pr_all n(12,2), vz n(1))
 USE 
 =OpenFile(dufile, 'dufile', 'shar')
 
 SELECT aisoms
 SCAN FOR !DELETED()

  m.lpuid    = lpuid
  m.mcod     = mcod 
  m.IsStomat = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
  m.IsIskl   = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
  
  m.IsPilot  = SEEK(m.lpuid, 'pilot')
  m.IsPilotS = SEEK(m.lpuid, 'pilots')
  m.IsHorLpu = SEEK(m.lpuid, 'horlpu')
  m.IsHorLpuS = SEEK(m.lpuid, 'horlpus')
  
  IF !m.IsPilot AND !m.IsPilotS AND !(m.IsHorLpuS AND !m.IsPilot)
*   IF !SEEK(m.lpuid, 'horlpu')
    LOOP 
*   ENDIF 
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
   m.prmcod   = people.prmcod
   m.lpu_prik = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   m.prmcods  = people.prmcods
   m.priks    = IIF(SEEK(m.prmcods, 'sprlpu'), sprlpu.lpu_id, 0)

   m.profil = profil
   IF !EMPTY(error.c_err)
    LOOP 
   ENDIF 
   
   IF (EMPTY(m.prmcod) AND m.IsPilot) AND (EMPTY(m.prmcods) AND m.IsPilotS)
    LOOP 
   ENDIF 
*   IF (m.prmcod=m.mcod AND m.IsPilot) AND (m.prmcods=m.mcod AND m.IsPilotS)
*    LOOP 
*   ENDIF 
   
   m.prlpuid = IIF(SEEK(m.prmcod, 'pilot', 'mcod'), pilot.lpu_id, 0)
*   IF !SEEK(m.prlpuid, 'pilot')
*    LOOP 
*   ENDIF 

   m.prlpuids = IIF(SEEK(m.prmcods, 'pilots', 'mcod'), pilots.lpu_id, 0)
*   IF !SEEK(m.prlpuids, 'pilots')
*    LOOP 
*   ENDIF 

   m.cod     = cod
   m.otd     = SUBSTR(otd,2,2)
   m.d_type  = d_type
   m.IsTpnR  = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)
   m.profil = profil
   m.ord     = ord
   m.lpu_ord = lpu_ord
   m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)

   IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod)
    LOOP 
   ENDIF 
   IF m.IsTpnR OR IsPat(m.cod) OR IsEKO(m.cod) OR m.d_type='s'
    LOOP 
   ENDIF 
   IF INLIST(m.otd,'08','70','73') AND IsStac(m.mcod)
    LOOP 
   ENDIF 
   IF INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
    LOOP 
   ENDIF 
   IF m.ord=7 AND m.lpu_ord=7665
    LOOP 
   ENDIF 
   IF (m.lIs02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))) OR m.lpu_ord>0
   ELSE  
    LOOP 
   ENDIF 


   m.cod     = cod
   m.k_u     = k_u
   m.s_all   = s_all
   m.lpu_ord = lpu_ord
   m.otd     = SUBSTR(otd,2,2)
   m.ds      = ds

   m.UslIskl      = IIF(FLOOR(m.cod/1000)=146, .T., .F.)
   m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
   m.IsStomatUsl2 = IIF(INLIST(m.cod,1101,1102,101171,101172), .T., .F.)
   
   m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)
   m.prlpuid = IIF(SEEK(m.prmcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)

   IF (m.IsPilotS OR m.IsHorLpuS) AND ;
    	(((m.IsStomat AND !m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2)) OR ;
  	    ((m.IsStomat AND m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2 OR m.IsIskl)) OR ;
  	    (!m.IsStomat AND (m.IsStomatUsl OR (m.IsStomatUsl2 AND LEFT(m.ds,2)='K0'))))
    IF m.prmcods=m.mcod OR EMPTY(m.prmcods)
     LOOP 
    ENDIF 
    IF !SEEK(m.prlpuids, 'pilots')
     LOOP 
    ENDIF 
    m.vz = 1
    INSERT INTO dufile (mcod,lpu_id,prmcod,prlpu_id,cod,k_u,pr_all,vz) VALUES ;
     (m.mcod,m.lpuid,m.prmcods,m.prlpuid,m.cod,m.k_u,m.s_all,m.vz)
   ELSE 
    IF m.prmcod=m.mcod OR EMPTY(m.prmcod)
     LOOP 
    ENDIF 
    IF !SEEK(m.prlpuid, 'pilot')
     *LOOP 
    ENDIF 
    IF (m.lIs02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))) OR m.lpu_ord>0
     m.vz = 1
    ELSE 
     m.vz = 2
    ENDIF 
    INSERT INTO dufile (mcod,lpu_id,prmcod,prlpu_id,cod,k_u,pr_all,vz) VALUES ;
     (m.mcod,m.lpuid,m.prmcod,m.prlpuid,m.cod,m.k_u,m.s_all,m.vz)
   ENDIF 

   
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
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\udqqxxxx.dbf')
  USE IN svcur
  USE IN dufile
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—“”“—“¬”≈“ ÿ¿¡ÀŒÕ UDQQXXXX.DBF!'+CHR(13)+CHR(10),0+16,'Templates')
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\pr4.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\pr4', 'pr4', 'shar', 'mcod')<=0
   SELECT prmcod as mcod, SUM(pr_all) as s_all WHERE vz=1 FROM dufile GROUP BY prmcod INTO CURSOR curstat
   SELECT curstat
   SET RELATION TO mcod INTO pr4
   m.sumstat = 0
   m.sumpr4  = 0
   SCAN 
    m.sumstat = m.sumstat + s_all
    m.sumpr4  = m.sumpr4 + pr4.s_others
   ENDSCAN 
   SET RELATION OFF INTO pr4
   USE IN pr4
   USE IN curstat
   IF m.sumstat!=m.sumpr4
    MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡Õ¿–”∆≈ÕŒ –¿—’Œ∆ƒ≈Õ»≈ Ã≈∆ƒ” ‘¿…ÀŒÃ PR4 »'+CHR(13)+CHR(10)+;
     '» —‘Œ–Ã»–Œ¬¿ÕÕ€Ã» UD-‘¿…À¿Ã»!'+CHR(13)+CHR(10)+;
     'œŒ ƒ¿ÕÕ€Ã PR4 —”ÃÃ¿ S_OTHERS='+TRANSFORM(m.sumpr4,'99999999.99')+CHR(13)+CHR(10)+;
     'œŒ ƒ¿ÕÕ€Ã UD-‘¿…ÀŒ¬ —”ÃÃ¿='+TRANSFORM(m.sumstat,'99999999.99')+CHR(13)+CHR(10),0+48,'')
   ELSE 
    MESSAGEBOX(CHR(13)+CHR(10)+'—”ÃÃ¿ S_OTHERS ‘¿…À¿ PR4'+CHR(13)+CHR(10)+;
     '—ŒŒ“¬≈“—“¬”≈“ —”ÃÃ≈ UD-‘¿…ÀŒ¬'+CHR(13)+CHR(10)+;
     '» —Œ—“¿¬Àﬂ≈“ '+TRANSFORM(m.sumstat,'99999999.99')+CHR(13)+CHR(10),0+64,'')
   ENDIF 
  ELSE 
   IF USED('pr4')
    USE 
   ENDIF 
   MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ ”ƒ¿ÀŒ—‹ Œ“ –€“‹ ‘¿…À PR4,'+CHR(13)+CHR(10)+;
    '—¬≈– ¿ ƒ¿ÕÕ€’ Õ≈ œ–Œ»«¬Œƒ»“—ﬂ!',0+64,'')
  ENDIF 
 ELSE 
  MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À PR4 Õ≈ —‘Œ–Ã»–Œ¬¿Õ,'+CHR(13)+CHR(10)+;
   '—¬≈– ¿ ƒ¿ÕÕ€’ Õ≈ œ–Œ»«¬Œƒ»“—ﬂ!',0+64,'')
 ENDIF 
 
 SELECT svcur
 SCAN 
  m.prmcod  = prmcod 
  m.lppid  = IIF(SEEK(m.prmcod, 'pilot', 'mcod'), pilot.lpu_id, 0)
  m.lppid  = IIF(m.lppid>0, m.lppid, IIF(SEEK(m.prmcod, 'pilots', 'mcod'), pilots.lpu_id, 0))
  m.uddfile = 'UD'+UPPER(m.qcod)+PADL(m.lppid,4,'0')
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
*  m.period = m.gcperiod
  SCAN 
   m.cod    = cod 
   m.k_u    = k_u
   m.pr_all = pr_all
   m.lpu_id = lpu_id
   m.vz     = vz
   IF m.vz=1
    INSERT INTO udfile FROM MEMVAR 
   ENDIF 
  ENDSCAN 
  USE IN curmmm
  SELECT udfile
  REPLACE ALL recid WITH PADL(RECNO(),7,'0')
  USE IN udfile
  
  SELECT svcur
  
 ENDSCAN 
 USE IN svcur
 USE IN horlpu 
 USE IN horlpus
 USE IN pilot 
 USE IN pilots
 USE IN lputpn
 USE IN tarif
 USE IN dufile
 USE IN sprlpu

 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 