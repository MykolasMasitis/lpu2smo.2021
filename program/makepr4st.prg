PROCEDURE MakePr4St(para1, para2)
 m.NeedOpen = .t.
 m.IsSilent = .f.
 IF PARAMETERS()>0
  m.NeedOpen = para1
 ENDIF 
 IF PARAMETERS()>1
  m.IsSilent = para2
 ENDIF 

 IF !m.IsSilent
 IF MESSAGEBOX(CHR(13)+CHR(10)+'СФОРМИРОВАТЬ ПРИЛОЖЕНИЕ 4(СТОМАТ)?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 ENDIF 
 m.lPath = m.pbase+'\'+m.gcperiod
 IF !fso.FileExists(m.lPath+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ AISOMS.DBF'+CHR(13)+CHR(10),0+16,m.lpath)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\pilots.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ PILOTS.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF m.NeedOpen
  IF OpBase()>0
   RETURN .f.
  ENDIF 
 ENDIF 
 
 m.lcperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')
 
 *CREATE TABLE AllPr4 ;
	(RecId i , ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6), pcod c(10), otd c(8), ;
	 cod n(6), tip c(1), d_u d, k_u n(4), kd_fact n(3), n_kd n(3), d_type c(1), s_all n(11,2), ;
	 s_lek n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3), kur n(5,3), ds_2 c(6), ds_3 c(6), ;
	 det n(1), k2 n(5,3), vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17), ord n(1), date_ord d, ;
	 lpu_ord n(6), recid_lpu c(7), fil_id n(6), ds_onk n(1), p_cel c(3), dn n(1), reab n(1), tal_d d, napr_v_in n(1), ;
	 c_zab n(1), napr_usl c(15), mp c(1), typ c(1), dop_r n(2), vz n(1), IsPr L) && Убрал поля codnom, napr_usl, vid_vme, tipgr, mm, vz, f_type 01.06.2019

 *CREATE CURSOR curpr4 (lpuid n(4), mcod c(7), adnorm n(11,2), chnorm n(11,2), adults n(6), childs n(6),;
  adsum n(11,2), chsum n(11,2), pazval n(6), finval n(13,2), paz_all n(6), s_all n(11,2), s_pred_pf n(11,2), ;
  paz_empty n(6), s_empty n(11,2), ;
  paz_own n(6), s_own n(11,2), paz_guests n(6), s_guests n(11,2), paz_others n(6), s_others n(11,2),;
  paz_npilot n(6), s_npilot n(11,2), paz_bad n(6), s_bad n(11,2), s_kompl n(11,2), s_dst n(11,2))

 CREATE CURSOR curpr4 (lpuid n(4), mcod c(7), ;
  pazval n(6), finval n(13,2), paz_all n(6), s_all n(11,2), s_pred_pf n(11,2), ;
  paz_empty n(6), s_empty n(11,2), ;
  paz_own n(6), s_own n(11,2), paz_guests n(6), s_guests n(11,2), paz_others n(6), s_others n(11,2),;
  paz_npilot n(6), s_npilot n(11,2), paz_bad n(6), s_bad n(11,2), s_kompl n(11,2), s_dst n(11,2))

 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 
 *m.adnorm = 0
 *m.chnorm = 0
 IF fso.FileExists(pcommon+'\pnorms.dbf')
  IF OpenFile(pcommon+'\pnorms.dbf', 'pnorm', 'shar', 'period')<=0
   SELECT pnorm 
   IF SEEK(m.gcperiod, 'pnorm')
   ELSE 
    GO BOTTOM 
   ENDIF 
   SCATTER FIELDS EXCEPT period MEMVAR 
   USE IN pnorm
  ENDIF 
 ENDIF 
 
 WAIT "СОЗДАНИЕ СПРАВОЧНИКА..." WINDOW NOWAIT 
 SELECT aisoms
 SCAN 
  m.sumok = s_pred - sum_flk
  m.mcod = mcod 
  m.lpuid = lpuid

  m.pazval = pazvals
  m.finval = finvals

  IF !SEEK(m.lpuid, 'pilots')
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
  IF !SEEK(m.lpuid, 'pilots')
   *IF !SEEK(m.lpuid, 'horlpus') && С 01.10.2018 их нет!
    LOOP 
   *ENDIF 
  ENDIF 
*  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
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
  m.s_npilot = 0
  m.s_bad    = 0
  m.s_kompl  = 0
  m.s_dst    = 0
  
  m.paz_all    = 0
  m.paz_empty  = 0
  m.paz_own    = 0
  m.paz_guests = 0
  m.paz_npilot = 0
  m.paz_bad    = 0
  
  m.s_pred_pf  = 0
  m.s_mek = 0
  SCAN 
   IF !EMPTY(errs.rid)
    m.s_mek = m.s_mek + s_all
    LOOP 
   ENDIF 
   m.cod    = cod
   m.ds     = ds
   m.otd    = SUBSTR(otd,2,2)
   m.d_type = d_type
   m.IsTpnR = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)
   m.ord    = ord
   m.lpu_ord = lpu_ord
   m.profil = profil
   *m.f_type = f_type
   m.Mp = Mp
   
   *IF m.IsTpnR OR INLIST(m.otd,'08','70','73','93') OR m.d_type='s' && допуслуги
   *IF (INLIST(m.otd,'01','90') AND IsStac(m.mcod)) AND people.mcod!=people.prmcods && допуслуги
   * LOOP 
   *ENDIF 
   *IF (m.ord=7 AND m.lpu_ord=7665) AND people.mcod!=people.prmcods && допуслуги
   * LOOP 
   *ENDIF 
  
   m.IsStomat     = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
   m.IsIskl       = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)
   m.UslIskl      = IIF(FLOOR(m.cod/1000)=146, .T., .F.)
   m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
   m.IsStomatUsl2 = IIF(INLIST(m.cod,1101,1102,101171,101172), .T., .F.)
   
   *IF m.IsStomat
   * IF !m.IsIskl
   *  IF !(m.IsStomatUsl OR m.IsStomatUsl2)
   *   LOOP 
   *  ENDIF 
   * ELSE 
   *  IF !(m.IsStomatUsl OR m.IsStomatUsl2 OR m.UslIskl)
   *   LOOP 
   *  ENDIF 
   * ENDIF 
   *ELSE
   * IF !(m.IsStomatUsl OR (m.IsStomatUsl2 AND LEFT(m.ds,2)='K0'))
   *  LOOP 
   * ENDIF 
   *ENDIF 
   
   * Невозможное условие
   *IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod) && стационар
   * MESSAGEBOX(STR(m.cod,6),0+64,m.mcod)
   * LOOP 
   *ENDIF 
   * Невозможное условие

   * Это неправильно!   
   *IF m.IsTpnR && допуслуги такие нашлись! штук пять, почему?!
   * MESSAGEBOX('m.IsTpnR'+STR(m.cod,6),0+64,m.mcod)
   * LOOP 
   *ENDIF 
   * Это неправильно!   
   IF m.Mp<>'s' && это единственное значимое условие!!!
    *MESSAGEBOX('m.Mp<>s'+STR(m.cod,6)+', Mp='+m.mp,0+64,m.mcod)
    LOOP 
   ENDIF 
   IF !IsDental(m.cod, m.lpuid, m.mcod, m.ds)
    MESSAGEBOX('!IsDental'+STR(m.cod,6)+', Mp='+m.mp,0+64,m.mcod)
    LOOP 
   ENDIF 

   *SCATTER MEMVAR 
   *INSERT INTO AllPr4 FROM MEMVAR 

   m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)

   m.sn_pol  = sn_pol
   m.s_all   = m.s_all + s_all
   m.lpu_ord = IIF(!EMPTY(FIELD('lpu_ord')), lpu_ord, 0)
   m.otd     = SUBSTR(otd,2,2)

   m.fil_id = fil_id

   IF !SEEK(m.sn_pol, 'cpazall')
    INSERT INTO cpazall (sn_pol) VALUES (m.sn_pol)
   ENDIF 

   m.mcod2  = people.prmcods
   m.prlpuid = IIF(SEEK(m.mcod2, 'pilots', 'mcod'), pilots.lpu_id, 0)

   m.paztip = TipOfPazS(m.mcod, m.mcod2) && 0 (не прикреплен),1 (прикреплен по месту обращения),2 (к пилоту)
   
   IF m.paztip=0
    m.s_empty  = m.s_empty + s_all
    IF !SEEK(m.sn_pol, 'cpazempty')
     INSERT INTO cpazempty (sn_pol) VALUES (m.sn_pol)
    ENDIF 
    LOOP 
   ENDIF 
   
   DO CASE 
    CASE m.paztip = 1 && свой
     m.s_own = m.s_own + s_all
     m.s_pred_pf = m.s_pred_pf
     IF !SEEK(m.sn_pol, 'cpazown')
      INSERT INTO cpazown (sn_pol) VALUES (m.sn_pol)
     ENDIF 

    CASE m.paztip = 2 && чужой
     IF (m.lIs02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))) OR m.lpu_ord>0

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
      
      *REPLACE Mm WITH 'S'
      
     ELSE && если без направления и не скорая помощь
     
      m.s_bad = m.s_bad + s_all
      IF !SEEK(m.sn_pol, 'cpazbad')
       INSERT INTO cpazbad (sn_pol) VALUES (m.sn_pol)
      ENDIF 
      *REPLACE Mm WITH 'Z'

     ENDIF 

    *CASE m.paztip = 3 && чужой непилот
    * IF (m.lIs02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))) OR m.lpu_ord>0

    *  IF !SEEK(m.sn_pol, 'cpaznopilot')
    *   INSERT INTO cpaznopilot (sn_pol) VALUES (m.sn_pol)
    *  ENDIF 
    *  m.s_npilot  = m.s_npilot  + s_all
     
    * ELSE  && если без направления и не скорая помощь
     
    *  m.s_bad = m.s_bad + s_all
    *  IF !SEEK(m.sn_pol, 'cpazbad')
    *   INSERT INTO cpazbad (sn_pol) VALUES (m.sn_pol)
    *  ENDIF 
    *  *REPLACE Mm WITH 'Z'

    * ENDIF 

   ENDCASE 

  ENDSCAN 
  
  m.paz_all     = RECCOUNT('cpazall')
  m.paz_empty   = RECCOUNT('cpazempty')
  m.paz_own     = RECCOUNT('cpazown')
  m.paz_guests  = RECCOUNT('cpazguests')
  m.paz_nopilot = RECCOUNT('cpaznopilot')
  m.paz_bad     = RECCOUNT('cpazbad')

  *UPDATE curpr4 SET paz_all=m.paz_all, s_all=m.s_all, paz_empty=m.paz_empty,;
   s_empty=m.s_empty, paz_own=m.paz_own, s_own=m.s_own, s_pred_pf=m.s_pred_pf,;
   paz_guests=m.paz_guests, s_guests=m.s_guests, paz_npilot=m.paz_nopilot,;
   paz_bad = m.paz_bad, s_bad=m.s_bad,;
   s_npilot=m.s_npilot, s_kompl=m.s_kompl, s_dst=m.s_dst WHERE mcod = m.mcod
  UPDATE curpr4 SET paz_all=m.paz_all, s_all=m.s_all, paz_empty=m.paz_empty,;
   s_empty=m.s_empty, paz_own=m.paz_own, s_own=m.s_own, s_pred_pf=m.s_pred_pf,;
   paz_guests=m.paz_guests, s_guests=0, paz_npilot=m.paz_nopilot,;
   paz_bad = m.paz_bad, s_bad=m.s_bad,;
   s_npilot=m.s_npilot, s_kompl=m.s_kompl, s_dst=m.s_dst WHERE mcod = m.mcod 
  
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

  IF USED('curo')  
   SELECT curo
*   BROWSE 
   SCAN 
    m.mcodd = mcod 
    IF !SEEK(m.mcodd, 'pilots', 'mcod')
     *IF !SEEK(m.mcodd, 'horlpus', 'mcod')
      LOOP 
     *ENDIF 
    ENDIF 
    m.sall = sall
    =SEEK(m.mcodd, 'curpr4')
    m.os_others = curpr4.s_others
    *UPDATE curpr4 SET s_others = m.os_others + m.sall WHERE mcod = m.mcodd
   ENDSCAN 
   USE IN curo
  ENDIF

  IF USED('curp')
   SELECT curp
   *BROWSE 
   SCAN 
    m.mcodd = mcod
    IF !SEEK(m.mcodd, 'pilots', 'mcod')
     *IF !SEEK(m.mcodd, 'horlpus', 'mcod')
      LOOP 
     *ENDIF 
    ENDIF 
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
 
 m.IsAttPplOk = .F.

 IF 1=2
 m.IsAttPplOk = .T.
 m.attbase = ''
 IF EMPTY(m.pattst)
  m.IsAttPplOk = .F.
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FolderExists(m.pattst)
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FileExists(m.pattst+'\attst.cfg')
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF OpenFile(m.pattst+'\attst.cfg', 'attcfg', 'shar')>0
   IF USED('attcfg')
    USE IN attcfg
   ENDIF 
   m.IsAttPplOk = .F.
  ELSE 
   SELECT attcfg
   m.attbase = ALLTRIM(pbase)
   USE IN attcfg
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FolderExists(m.attbase)
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FolderExists(m.attbase+'\'+m.lcperiod)
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FileExists(m.attbase+'\'+m.lcperiod+'\aisoms.dbf')
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF OpenFile(m.attbase+'\'+m.lcperiod+'\aisoms', 'attais', 'shar', 'lpuid')>0
   IF USED('attais')
    USE IN attais
   ENDIF 
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 

 ENDIF 
 
 SELECT curpr4
 *REPLACE ALL adnorm WITH m.adnorm, chnorm WITH m.chnorm
 IF m.IsAttPplOk = .T.
  SET RELATION TO lpuid INTO attais
  *REPLACE ALL adults WITH attais.ad_mgf, childs WITH attais.ch_mgf
  *REPLACE ALL adsum WITH adnorm*adults, chsum WITH chnorm*childs

  SCAN 

   m.pazval = attais.ch01m+attais.ch01f+attais.ch14m+attais.ch14f+;
    attais.ch514m+attais.ch514f+attais.ch1517m+attais.ch1517f+;
    attais.m1824+attais.f1824+attais.m2534+attais.f2534+;
    attais.m3544+attais.f3544+attais.m4559+attais.f4559+;
    attais.m6068+attais.f5564+attais.m69+attais.f65

   m.finval = attais.ch01m*m.m0001+attais.ch01f*m.f0001+attais.ch14m*m.m0104+attais.ch14f*m.f0104+;
    attais.ch514m*m.m0514+attais.ch514f*m.f0514+attais.ch1517m*m.m1517+attais.ch1517f*m.f1517+;
    attais.m1824*m.m1824+attais.f1824*m.f1824+attais.m2534*m.m2534+attais.f2534*m.f2534+;
    attais.m3544*m.m3544+attais.f3544*m.f3544+attais.m4559*m.m4559+attais.f4559*m.f4554+;
    attais.m6068*m.m6068+attais.f5564*m.f5564+attais.m69*m.m6999+attais.f65*m.f6599
*   m.finval = m.finval * m.koeff

   REPLACE finval WITH m.finval, pazval WITH m.pazval

  ENDSCAN 
  SET RELATION OFF INTO attais
  USE IN attais
 ENDIF 
 IF fso.FileExists(m.lPath+'\pr4st.dbf')
  fso.DeleteFile(m.lPath+'\pr4st.dbf')
 ENDIF 
 COPY TO m.lPath+'\pr4st'
 USE 
 
 *SELECT allpr4
 *COPY TO m.lPath+'\allpr4st'
 *USE 
 
 IF OpenFile(m.lPath+'\pr4st', 'pr4', 'excl')>0
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
  MESSAGEBOX(CHR(13)+CHR(10)+'ОБРАБОТКА ЗАКОНЧЕНА!'+CHR(13)+CHR(10),0+64,'')
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
 *tnResult = tnResult + OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpus', 'horlpus', 'shar', 'lpu_id')
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
 *IF USED('horlpus')
 * USE IN horlpus
 *ENDIF 
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
RETURN 