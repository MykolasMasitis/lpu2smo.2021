PROCEDURE MakePr4n(para1, para2)
 m.NeedOpen = .t.
 m.IsSilent = .f.
 IF PARAMETERS()>0
  m.NeedOpen = para1
 ENDIF 
 IF PARAMETERS()>1
  m.IsSilent = para2
 ENDIF 

 IF !m.IsSilent
  IF MESSAGEBOX(CHR(13)+CHR(10)+'—‘Œ–Ã»–Œ¬¿“‹ œ–»ÀŒ∆≈Õ»≈ 4(NEW)?'+CHR(13)+CHR(10),4+32,'')=7
   RETURN 
  ENDIF 
 ENDIF 
 m.lPath = m.pbase+'\'+m.gcperiod
 IF !fso.FileExists(m.lPath+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF'+CHR(13)+CHR(10),0+16,m.lpath)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\pilot.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À PILOT.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\pilots.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À PILOTS.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF m.NeedOpen
  IF OpBase()>0
   RETURN .f.
  ENDIF 
 ENDIF 
 
 m.lcperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')
 
 CREATE CURSOR curpr4 (lpuid n(4), mcod c(7), ;
  pazval n(6), finval n(13,2), paz_all n(6), s_all n(11,2), s_pred_pf n(11,2), ;
  paz_empty n(6), s_empty n(11,2), ;
  paz_own n(6), s_own n(11,2), paz_guests n(6), s_guests n(11,2), paz_others n(6), s_others n(11,2),;
  paz_npilot n(6), s_npilot n(11,2), paz_bad n(6), s_bad n(11,2), s_kompl n(11,2), s_dst n(11,2))

 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 
 IF fso.FileExists(pcommon+'\pnorm.dbf')
  IF OpenFile(pcommon+'\pnorm.dbf', 'pnorm', 'shar', 'period')<=0
   SELECT pnorm 
   IF SEEK(m.gcperiod, 'pnorm')
   ELSE 
    GO BOTTOM 
   ENDIF 
   SCATTER FIELDS EXCEPT period MEMVAR 
   USE IN pnorm
  ENDIF 
 ENDIF 
 
 WAIT "—Œ«ƒ¿Õ»≈ —œ–¿¬Œ◊Õ» ¿..." WINDOW NOWAIT 
 SELECT aisoms
 SCAN 
  m.sumok = s_pred - sum_flk
  m.mcod = mcod 
  m.lpuid = lpuid

  m.pazval = pazval
  m.finval = finval

  IF !SEEK(m.lpuid, 'pilot')
   LOOP 
  ENDIF 
  IF !SEEK(m.mcod, 'curpr4')
   INSERT INTO curpr4 (lpuid, mcod, pazval, finval) VALUES ;
   	(m.lpuid, m.mcod, m.pazval, m.finval)
  ENDIF 
 ENDSCAN 
 WAIT CLEAR 

 SCAN 
  m.mcod = mcod 
  m.lpuid = lpuid
  IF !SEEK(m.lpuid, 'pilot')
   IF !SEEK(m.lpuid, 'horlpu')
    LOOP 
   ENDIF 
  ENDIF 
  
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
  m.llpath = m.lpath+'\'+m.mcod
  IF !fso.FileExists(m.llpath+'\people.dbf') OR !fso.FileExists(m.llpath+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.llpath+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llpath+'\talon', 'talon', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llpath+'\e'+m.mcod, 'errs', 'shar', 'rid')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('errs')
    USE IN errs
   ENDIF 
   LOOP 
  ENDIF 
  IF fso.FileExists(m.llpath+'\hosp.dbf')
   IF OpenFile(m.llpath+'\hosp', 'hosp', 'shar', 'c_i')>0
    IF USED('hosp')
     USE IN hosp 
    ENDIF 
   ENDIF 
  ENDIF 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  CREATE CURSOR cpazall (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpazempty (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpazown (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpazguests (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpaznopilot (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpazbad (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR curo (mcod c(7), sall n(11,2), npaz n(6))
  INDEX on mcod TAG mcod
  SET ORDER TO mcod 
  CREATE CURSOR curp (mcod c(7), sn_pol c(25))
  INDEX ON mcod TAG mcod
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol


  SELECT talon 
  SET RELATION TO recid INTO errs
  SET RELATION TO sn_pol INTO people ADDITIVE 
  
  m.s_all    = 0
  m.s_empty  = 0
  m.s_own    = 0
  m.s_guests = 0
  m.s_bad    = 0
  
  m.paz_all    = 0
  m.paz_empty  = 0
  m.paz_own    = 0
  m.paz_guests = 0
  m.paz_bad    = 0
  
  m.s_pred_pf  = 0
  m.s_mek = 0

  SCAN 
   IF !EMPTY(errs.rid)
    m.s_mek = m.s_mek + s_all
    LOOP 
   ENDIF 

   *m.cod    = cod
   *m.ds     = ds
   *m.otd    = SUBSTR(otd,2,2)
   *m.d_type = d_type
   *m.IsTpnR = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)
   *m.ord    = ord
   *m.lpu_ord = lpu_ord
   *m.profil = profil
   *m.c_i    = c_i
   
   m.Mp  = Mp
   m.Typ = Typ
   m.vz  = vz

   *m.IsUslGosp = .F.
   *IF USED('hosp')
   * m.IsUslGosp = IIF(IsUsl(m.cod) AND SEEK(m.c_i, 'hosp'), .T., .F.)
   *ENDIF 

   IF m.Mp<>'p'
    LOOP 
   ENDIF 
   m.cod = cod
   IF m.mcod<>'0343003' AND (IsMes(m.cod) OR IsVmp(m.cod))
   *IF (IsMes(m.cod) OR IsVmp(m.cod))
    LOOP 
   ENDIF 
   
   *IF SEEK(m.lpuid, 'pilots') && OR SEEK(m.lpuid, 'horlpus')
   * m.IsStomat     = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
   * m.IsIskl       = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)
   * m.UslIskl      = IIF(FLOOR(m.cod/1000)=146, .T., .F.)
   * m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
   * m.IsStomatUsl2 = IIF(INLIST(m.cod,1101,1102,101171,101172), .T., .F.)
   
   * IF m.IsStomat
   *  IF !m.IsIskl
   *   IF (m.IsStomatUsl OR m.IsStomatUsl2)
   * MESSAGEBOX('(m.IsStomatUsl OR m.IsStomatUsl2)',0+64,m.mcod)
   *    LOOP 
   *   ENDIF 
   *  ELSE 
   *   IF (m.IsStomatUsl OR m.IsStomatUsl2 OR m.UslIskl)
   * MESSAGEBOX('(m.IsStomatUsl OR m.IsStomatUsl2 OR m.UslIskl)',0+64,m.mcod)
   *    LOOP 
   *   ENDIF 
   *  ENDIF 
   * ELSE
   *  IF (m.IsStomatUsl OR (m.IsStomatUsl2 AND LEFT(m.ds,2)='K0'))
   * MESSAGEBOX('(m.IsStomatUsl OR (m.IsStomatUsl2 AND LEFT(m.ds,2)=K0))',0+64,m.mcod)
   *   LOOP 
   *  ENDIF 
   * ENDIF 
   *ENDIF 
   
   *m.otd    = SUBSTR(otd,2,2)

   *m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)

   m.sn_pol  = sn_pol
   m.s_all   = m.s_all + s_all
   *m.lpu_ord = IIF(!EMPTY(FIELD('lpu_ord')), lpu_ord, 0)
   *m.otd     = SUBSTR(otd,2,2)

   *m.fil_id = fil_id

   IF !SEEK(m.sn_pol, 'cpazall')
    INSERT INTO cpazall (sn_pol) VALUES (m.sn_pol)
   ENDIF 

   m.mcod2  = people.prmcod
   m.prlpuid = IIF(SEEK(m.mcod2, 'pilot', 'mcod'), pilot.lpu_id, 0)

   m.paztip = TipOfPaz(m.mcod, m.mcod2) && 0 (ÌÂ ÔËÍÂÔÎÂÌ),1 (ÔËÍÂÔÎÂÌ ÔÓ ÏÂÒÚÛ Ó·‡˘ÂÌËˇ),2 (Í ÔËÎÓÚÛ),3 (ÌÂ Í ÔËÎÓÚÛ)
   
   IF m.paztip=0
    m.s_empty  = m.s_empty + s_all
    IF !SEEK(m.sn_pol, 'cpazempty')
     INSERT INTO cpazempty (sn_pol) VALUES (m.sn_pol)
    ENDIF 
    LOOP 
   ENDIF 
   
   DO CASE 
    CASE m.paztip = 1 && Ò‚ÓÈ Û ÒÂ·ˇ
     m.s_own = m.s_own + s_all
     m.s_pred_pf = m.s_pred_pf
     IF !SEEK(m.sn_pol, 'cpazown')
      INSERT INTO cpazown (sn_pol) VALUES (m.sn_pol)
     ENDIF 

    CASE m.paztip = 2 && ˜ÛÊÓÈ ÔËÎÓÚ
     IF m.vz>0

      IF !SEEK(m.sn_pol, 'cpazguests')
       INSERT INTO cpazguests (sn_pol) VALUES (m.sn_pol)
      ENDIF 

      IF !SEEK(m.mcod2, 'curo')
       INSERT INTO curo (mcod) VALUES (m.mcod2)
      ENDIF 
      IF !SEEK(m.sn_pol, 'curp')
       INSERT INTO curp (mcod, sn_pol) VALUES (m.mcod2, m.sn_pol)
      ENDIF 

      m.s_guests = m.s_guests + s_all

      =SEEK(m.mcod2, 'curo')
      m.osall = curo.sall
      m.nsall = m.osall + s_all
      
      UPDATE curo SET sall = m.nsall WHERE mcod = m.mcod2
      
     ELSE && ÂÒÎË ·ÂÁ Ì‡Ô‡‚ÎÂÌËˇ Ë ÌÂ ÒÍÓ‡ˇ ÔÓÏÓ˘¸
     
      m.s_bad = m.s_bad + s_all
      IF !SEEK(m.sn_pol, 'cpazbad')
       INSERT INTO cpazbad (sn_pol) VALUES (m.sn_pol)
      ENDIF 
      *REPLACE Mm WITH 'Y'

     ENDIF 

   ENDCASE 

  ENDSCAN 
  
  m.paz_all     = RECCOUNT('cpazall')
  m.paz_empty   = RECCOUNT('cpazempty')
  m.paz_own     = RECCOUNT('cpazown')
  m.paz_guests  = RECCOUNT('cpazguests')
  m.paz_nopilot = RECCOUNT('cpaznopilot')
  m.paz_bad     = RECCOUNT('cpazbad')

  UPDATE curpr4 SET paz_all=m.paz_all, s_all=m.s_all, paz_empty=m.paz_empty,;
   s_empty=m.s_empty, paz_own=m.paz_own, s_own=m.s_own, s_pred_pf=m.s_pred_pf,;
   paz_guests=m.paz_guests, s_guests=m.s_guests, paz_bad = m.paz_bad, s_bad=m.s_bad ;
   WHERE mcod = m.mcod
  
  SET RELATION OFF INTO people
  SET RELATION OFF INTO errs
  
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('errs')
   USE IN errs
  ENDIF 
  IF USED('hosp')
   USE IN hosp 
  ENDIF 

  IF USED('curo')  
   SELECT curo
   SCAN 
    m.mcodd = mcod 
    IF !SEEK(m.mcodd, 'pilot', 'mcod')
     IF !SEEK(m.mcodd, 'horlpu', 'mcod')
      MESSAGEBOX('!!!',0+64,'')
      LOOP 
     ENDIF 
    ENDIF 
    m.sall = sall
    =SEEK(m.mcodd, 'curpr4')
    m.os_others = curpr4.s_others
    UPDATE curpr4 SET s_others = m.os_others + m.sall WHERE mcod = m.mcodd
   ENDSCAN 
   *COPY TO &pbase\&gcperiod\curo
   USE IN curo
  ENDIF

  IF USED('curp')
   SELECT curp
   SCAN 
    IF !SEEK(m.mcodd, 'pilot', 'mcod')
     IF !SEEK(m.mcodd, 'horlpu', 'mcod')
      MESSAGEBOX('!!!',0+64,'')
      LOOP 
     ENDIF 
    ENDIF 
    m.mcodd = mcod
    =SEEK(m.mcodd, 'curpr4')
    m.opaz_others = curpr4.paz_others
    UPDATE curpr4 SET paz_others = m.opaz_others + 1 WHERE mcod = m.mcodd
   ENDSCAN 
   USE IN curp
  ENDIF 
  
  IF USED('cpazall')
   USE IN cpazall
  ENDIF  
  IF USED('cpazempty')
   USE IN cpazempty
  ENDIF  
  IF USED('cpazown')
   USE IN cpazown
  ENDIF  
  IF USED('cpazguests')
   USE IN cpazguests
  ENDIF  
  IF USED('cpaznopilot')
   USE IN cpaznopilot
  ENDIF  
  IF USED('cpazbad')
   USE IN cpazbad
  ENDIF  

  IF m.s_mek>0
  ENDIF 

  SELECT aisoms
  
 ENDSCAN 

 WAIT CLEAR 
 
 IF m.NeedOpen
  =ClBase(m.NeedOpen)
 ENDIF 
 
IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\pnorm_iskl.dbf')
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\pnorm_iskl', 'pn', 'shar', 'mcod')>0
  IF USED('pn')
   USE IN pn
  ENDIF 
 ENDIF 
ENDIF 

 m.IsAttPplOk = .F.
 
 SELECT curpr4
 IF fso.FileExists(m.lPath+'\pr4.dbf')
  fso.DeleteFile(m.lPath+'\pr4.dbf')
 ENDIF 
 COPY TO m.lPath+'\pr4'
 USE 
 
 IF OpenFile(m.lPath+'\pr4', 'pr4', 'excl')>0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
 ELSE 
  SELECT pr4
  INDEX on lpuid TAG lpuid
  INDEX on mcod TAG mcod
  USE 
 ENDIF 
 
 IF !m.IsSilent
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
 ENDIF 
 
RETURN 

FUNCTION OpBase()
 tnResult = 0
 tnResult = tnResult + OpenFile(m.lPath+'\aisoms', 'aisoms', 'shar', 'mcod')
 tnResult = tnResult + OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')
 tnResult = tnResult + OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilots', 'pilots', 'shar', 'lpu_id')
 tnResult = tnResult + OpenFile(pbase+'\'+m.gcperiod+'\nsi\lputpn', 'lputpn', 'shar', 'lpu_id')
 tnResult = tnResult + OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')
 tnResult = tnResult + OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpu', 'horlpu', 'shar', 'lpu_id')
 tnResult = tnResult + OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')
RETURN tnresult

FUNCTION ClBase(para1)
 m.IsNeedOpen = para1
* IF m.IsNeedOpen
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
* ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('horlpu')
  USE IN horlpu
 ENDIF 
 IF USED('tarif')
  USE IN tarif
 ENDIF 
 IF USED('lputpn')
  USE IN lputpn
 ENDIF 
 IF USED('pilot')
  USE IN pilot
 ENDIF 
 IF USED('pilots')
  USE IN pilots
 ENDIF 
 IF USED('pn')
  USE IN pn
 ENDIF 
 *SELECT allpr4 
 *COPY TO &pBase\&gcPeriod\AllPr4
 *USE 
RETURN 