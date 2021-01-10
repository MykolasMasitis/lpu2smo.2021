FUNCTION OneFlkNR(ppath)

 m.cfrom = ALLTRIM(cfrom)

* IF IsPr
*  =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
*  RETURN 
* ENDIF 

 LocalErrIniFile    = ppath + '\errors.ini'
 IsLocalErrIniFile  = fso.FileExists(LocalErrIniFile)
 GlobalErrIniFile   = pbin + '\errors.ini'
 IsGlobalErrIniFile = fso.FileExists(GlobalErrIniFile)
 
 WorkIniFile = ''
 IF IsLocalErrIniFile
  WorkIniFile = LocalErrIniFile
 ELSE 
  IF IsGlobalErrIniFile
   WorkIniFile = GlobalErrIniFile
  ENDIF 
 ENDIF 
 
  M.PSA = .F.
  M.ERA = .F.
  M.ECA = .F.
  M.E1A = .F.
  M.E2A = .F.
  M.E4A = .F.
  M.E7A = .F.
  M.E8A = .F.

  M.H6A = .F.
  M.COA = .F.
  M.HCA = .F.
  M.DUA = .F.
  M.H8A = .F.
  M.HEA = .F.
  M.CSA = .F.
  M.TVA = .F.
  M.NLA = .F.
  M.MDA = .F.
  M.H3A = .F.
  M.SOA = .F.
  M.R1A = .F.
  M.R2A = .F.
  M.R3A = .F. 
  M.UVA = .F.
  M.DVA = .F.
  M.UOA = .F.
  M.NOA = .F.
  M.NMA = .F.
  M.NUA = .F.
  M.NSA = .F.
  M.SMA = .F.
  M.DIA = .F.
  M.DDA = .F.
  M.HNA = .F.
  M.DLA = .F.
  M.DRA = .F.
  M.POA = .F.
  M.VDA = .F.
  M.TFA = .F.
  M.PPA = .F.
  M.G1A = .F.
  M.G2A = .F.
  M.G3A = .F.
  M.G4A = .F.
  M.NRA = .T.
  M.KEA = .F.
  M.D2A = .F.
  M.THA = .F.
  M.TLA = .F.
  M.HOA = .F.
  M.SKA = .F.
  M.FSA = .F.
  M.W2A = .F.

  M.O0A = .F.
  M.O1A = .F.
  M.O2A = .F.
  M.O3A = .F.
  M.O4A = .F.
  M.O5A = .F.
  M.O6A = .F.
  M.O7A = .F.
  M.O8A = .F.
  M.O8A = .F.
  M.OAA = .F.
  M.OBA = .F.
  M.OCA = .F.
  M.ODA = .F.
  M.OEA = .F.
  M.OFA = .F.
  M.OGA = .F.
  M.OHA = .F.
  M.OIA = .F.
  M.OJA = .F.
  M.OKA = .F.
  M.OLA = .F.
  M.OMA = .F.
  M.ONA = .F.
  M.OOA = .F.
  M.OPA = .F.
  M.OQA = .F.
  M.ORA = .F.
  M.OSA = .F.
  M.OTA = .F.
  M.OUA = .F.
  M.OVA = .F.
  M.OWA = .F.
  M.OXA = .F.
  M.OYA = .F.
  M.OZA = .F.
  M.ENA = .F.

  M.X1B = IIF(m.qcod='S7', .F., .F.)
  M.X2B = IIF(m.qcod='S7', .F., .F.)
  M.X3B = IIF(m.qcod='S7', .F., .F.)
  M.X4B = IIF(m.qcod='S7', .F., .F.)
  M.X5B = IIF(m.qcod='S7', .F., .F.)
  M.X6B = IIF(m.qcod='S7', .F., .F.)
  M.X7B = IIF(m.qcod='S7', .F., .F.)
  M.X8B = IIF(m.qcod='S7', .F., .F.)
  M.X9B = IIF(m.qcod='S7', .F., .F.)
 
 M.O0A = IIF(!EMPTY(cfrom), .F., M.O1A) && Заглушка для cfrom=oms@spuemias.msk.oms
 M.O0A = IIF(m.tdat1<{01.01.2019}, .F., M.O0A) && Заглушка для cfrom=oms@spuemias.msk.oms
 M.ENA = IIF(!EMPTY(cfrom), .F., M.ENA) && Заглушка для cfrom=oms@spuemias.msk.oms
 
  m.lpuid   = lpuid
  m.mcod    = mcod
  m.IsStomat   = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
  m.IsIskl     = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)
  m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')+', '+sprlpu.cokr+', '+sprlpu.mcod
  m.period  = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
  m.dat1 = CTOD('01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4))
  m.dat2 = GOMONTH(m.dat1,1)-1
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
  m.IsStac = IIF(VAL(SUBSTR(m.mcod,3,2))<41 or m.IsLpuTpn, .F., .T.)
  m.IsPilot = IIF(SEEK(m.lpuid, 'pilot'), .t., .f.)
  m.IsPilotS = IIF(SEEK(m.lpuid, 'pilots'), .t., .f.)
  m.IsHor  = IIF(SEEK(m.lpuid, 'horlpu'), .t., .f.)
  m.IsHorS = IIF(SEEK(m.lpuid, 'horlpus'), .t., .f.)
  
*  m.UslIskl      = IIF(FLOOR(m.cod/1000)=146, .T., .F.)
*  m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
*  m.IsStomatUsl2 = IIF(INLIST(m.cod,1101,1102,101171,101172), .T., .F.)
  
  M.G1A = IIF(m.IsHor=.t., M.G1A, .F.)
  M.G2A = IIF(m.IsHor=.t., M.G2A, .F.)
  M.G3A = IIF(m.IsHor=.t., M.G3A, .F.)
  M.G4A = IIF(m.IsHor=.t., M.G4A, .F.)
  M.NRA = IIF(m.IsHor=.t. or m.IsPilot=.t., M.NRA, .F.) && Закомментировано 06.02.2018 && Пробуем отключить! 
  *M.NRA = IIF(m.qcod='S7', IIF(M.NRA or m.IsPilot=.t. or m.IsPilotS=.t., M.NRA, .F.), M.NRA)
  
  m.lIsDspExists = .f.
  m.dspfile1 = pbase +'\'+ STR(tyear-1,4)+'12'+'\dsp'
  m.dspfile2 = pbase +'\'+ STR(tyear-2,4)+'12'+'\dsp'
  m.dspfile3 = pbase +'\'+ STR(tyear-3,4)+'12'+'\dsp'
  IF tmonth>1
   m.dspfile = pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\dsp'
  ELSE
   m.dspfile = pbase +'\'+ STR(tyear-1,4)+'12'+'\dsp'
  ENDIF 

*  m.dspfile = m.dspfile1

  IF fso.FileExists(m.dspfile+'.dbf')
   m.lIsDspExists = .t.
  ELSE 
   m.lIsDspExists = .f.
  ENDIF 
  
  IF m.lIsDspExists
  oal = ALIAS()
  IF OpenFile(m.dspfile, 'cdsp', 'shar')>0 && mcod+sn_pol+padl(cod,6,'0')
   IF USED('cdsp')
    USE IN cdsp
   ENDIF 
   SELECT (oal)
   m.lIsDspExists = .f.
  ELSE 
   SELECT * FROM cdsp INTO CURSOR dspp READWRITE 
   USE IN cdsp
   SELECT dspp
*   INDEX on mcod+sn_pol+PADL(cod,6,'0') TAG exptag
   INDEX on mcod+sn_pol+PADL(tip,1,'0') TAG exptag
   SET ORDER TO exptag
   IF fso.FileExists(m.dspfile1+'.dbf') AND tmonth>1
    APPEND FROM &dspfile1
   ENDIF 
   IF fso.FileExists(m.dspfile2+'.dbf')
    APPEND FROM &dspfile2
   ENDIF 
   IF fso.FileExists(m.dspfile3+'.dbf')
    APPEND FROM &dspfile3
   ENDIF 
   SELECT (oal)
   m.lIsDspExists = .t.
  ENDIF  
  ENDIF 
  
  IF m.lIsDspExists
  IF OpenFile(m.dspfile, 'dspyear', 'shar')>0
   IF USED('dspyear')
    USE IN dspyear
   ENDIF 
   SELECT (oal)
   m.lIsDspExists = .f.
  ELSE 
  SELECT dspyear
  =ATAGINFO(taginf)
  IF ASCAN(taginf,'EXPTAG')>0
   SET ORDER TO exptag
   m.lIsDspExists = .t.
  ELSE 
  IF OpenFile(m.dspfile, 'dspyear', 'excl')>0 && mcod+sn_pol+padl(cod,6,'0')
   IF USED('dspyear')
    USE IN dspyear
   ENDIF 
   SELECT (oal)
   m.lIsDspExists = .f.
  ELSE 
   SELECT dspyear
   INDEX on mcod+sn_pol+PADL(cod,6,'0') TAG exptag
   SET ORDER TO exptag
   USE IN dspyear
   =OpenFile(m.dspfile, 'dspyear', 'shar', 'exptag')
   m.lIsDspExists = .t.
  ENDIF 
  ENDIF 
  ENDIF 
  ENDIF 

  M.D2A = IIF(m.lIsDspExists = .t., M.D2A, .f.)
  
  *MESSAGEBOX(IIF(m.d2a=.t., '.T.','.F.'), 0+64,'')

  lcError = ppath+'\e'+m.mcod
  IF !fso.FileExists(lcError+'.dbf')
   CREATE TABLE (lcError) (f c(1), c_err c(3), rid i)
   INDEX FOR UPPER(f)='R' ON rid TAG rrid
   INDEX FOR UPPER(f)='S' ON rid TAG rid
   USE 
  ENDIF 
  
  IF !OpBase(ppath)
   RETURN .f.
  ENDIF 

  m.IsExHorS = IIF(SEEK(m.lpuid, 'exclhors'), .t., .f.)
  m.IsSkp  = IIF(SEEK(m.lpuid, 'lpuskp'), .t., .f.)
  M.SKA = IIF(m.IsSkp=.t., M.SKA, .F.)
  
  IF m.IsSkp
   SELECT c_i, sn_pol, cod, ds, SPACE(3) as prv WHERE SUBSTR(otd,2,2)='09' FROM talon INTO CURSOR curskp READWRITE 
   SELECT curskp
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol
   SET RELATION TO cod INTO profus
   REPLACE ALL prv WITH profus.profil
   SET RELATION OFF INTO profus
  ENDIF 

  SELECT sn_pol, cod, SUM(k_u) AS k_u, d_u, SUM(s_all) AS s_all ;
   FROM Talon GROUP BY sn_pol, d_u, cod WHERE d_type != '2' ;
   INTO CURSOR day_gr

*  SELECT sn_pol AS sn_pol, a.cod AS cod, ;
   k_u AS k_u, 000 as cntr, d_u AS d_u, s_all AS s_all, IIF(!IsStac, mdayp, mdays) AS in_day,;
   IIF(!IsStac, mdayp, mdays) as k_norm;
   FROM day_gr a, codku b ;
   WHERE a.cod=b.cod AND k_u > IIF(!IsStac, mdayp, mdays) ;
   INTO CURSOR e_day ORDER BY a.sn_pol, a.d_u, a.cod READWRITE 
  SELECT sn_pol AS sn_pol, a.cod AS cod, ;
   k_u AS k_u, 000 as cntr, d_u AS d_u, s_all AS s_all, mdayp AS in_day,;
   mdayp as k_norm;
   FROM day_gr a, codku b ;
   WHERE a.cod=b.cod AND k_u > mdayp ;
   INTO CURSOR e_day ORDER BY a.sn_pol, a.d_u, a.cod READWRITE 
  SELECT e_day
  INDEX ON sn_pol + STR(cod,6) + DTOS(d_u) TAG ExpTag
  SET ORDER TO ExpTag

  SELECT sn_pol, cod, SUM(k_u) AS k_u, MIN(d_u) AS d_u, SUM(s_all) AS s_all ;
   FROM Talon  WHERE d_type!='2' GROUP BY sn_pol, cod ;
   INTO CURSOR month_gr

*  SELECT sn_pol as sn_pol, a.cod as cod, k_u as k_u, 000 as cntr, IIF(!IsStac, mmsp, mmss) as k_norm, s_all as s_all, ;
   IIF(!IsStac, b.mmsp, b.mmss) as in_month ;
   FROM month_gr a, codku b ;
   WHERE a.cod=b.cod and k_u > IIF(!IsStac, mmsp, mmss) ;
   INTO CURSOR e_month ORDER BY sn_pol, a.cod READWRITE 
  SELECT sn_pol as sn_pol, a.cod as cod, k_u as k_u, 000 as cntr, mmsp as k_norm, s_all as s_all, ;
   b.mmsp as in_month ;
   FROM month_gr a, codku b ;
   WHERE a.cod=b.cod and k_u > mmsp ;
   INTO CURSOR e_month ORDER BY sn_pol, a.cod READWRITE 
  SELECT e_month
  INDEX ON sn_pol + STR(cod,6) TAG ExpTag
  SET ORDER TO ExpTag

  *SELECT * FROM Talon ORDER BY c_i, d_u DESC, k_u DESC INTO CURSOR Gosp
  SELECT * FROM Talon ORDER BY c_i, d_u DESC INTO CURSOR Gosp
  
  m.s_flk = 0  

  IF M.DRA == .T.
   SELECT recid, sn_pol FROM people WHERE sn_pol IN ;
   (SELECT sn_pol FROM people GROUP BY sn_pol HAVING coun(*)>1) INTO CURSOR dblppl
   IF _tally>0
    SELECT dblppl
    INDEX on sn_pol TAG sn_pol UNIQUE 
    SET ORDER TO sn_pol
    SELECT talon
    SET RELATION TO sn_pol INTO dblppl
    SCAN 
     IF !EMPTY(dblppl.sn_pol)  
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval = InsError('S', 'PKA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
     ENDIF 
    ENDSCAN 
    SET RELATION OFF INTO dblppl
    SELECT dblppl
    SET ORDER TO 
    SCAN 
     m.recid = recid
     =InsError('R', 'DRA', m.recid)
*      InsErrorSV(m.mcod, 'R', 'DRA', m.recid)
    ENDSCAN 
   ENDIF 
  USE IN dblppl
  ENDIF 
  
  SELECT talon
  SET RELATION TO sn_pol INTO people 
  
  SCAN
  
   m.IsOtdSkp = IIF(m.IsSkp AND SUBSTR(otd,2,2)='09', .T., .F.)

   IF EMPTY(people.sn_pol)               && Алгоритм PS
    m.polis = sn_pol
    DO WHILE sn_pol == m.polis
     m.recid = recid
     rval = InsError('S', 'PSA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'PSA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     SKIP +1 
    ENDDO 
   ENDIF 

  IF people.IsPr==.F. && Глобальное отключение ошибок регистра!
  
*   DO CASE 
*    CASE M.ERA == .T. AND (!EMPTY(people.sv) AND (SEEK(people.sv, 'osoerz') AND osoerz.kl == 'y'))
*   ENDCASE 
  

   IF M.ERA == .T. && Алгоритм ER
    IF !EMPTY(people.sv)  
     m.IsGood = IIF(SEEK(people.sv, 'osoerz') AND osoerz.kl == 'y', .T., .F.)
     IF IsVS(people.sn_pol) AND LEFT(people.sn_pol,2)=m.qcod
      IF USED('kms')
       m.vvs = INT(VAL(SUBSTR(ALLTRIM(people.sn_pol),7)))
       IF SEEK(m.vvs, 'kms')
        m.IsGood = .t.
       ENDIF 
      ENDIF 
     ENDIF 
     IF IsGood == .f.
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval = InsError('S', 'PKA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'ERA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.ECA == .T. && Алгоритм EC
    IF !EMPTY(people.sv)
     m.IsGood = IIF(people.qq = m.qcod, .T., .F.)
     IF IsVS(people.sn_pol) AND LEFT(people.sn_pol,2)=m.qcod
      IF USED('kms')
       m.vvs = INT(VAL(SUBSTR(ALLTRIM(people.sn_pol),7)))
       IF SEEK(m.vvs, 'kms')
        m.IsGood = .t.
       ENDIF 
      ENDIF 
     ENDIF 
     IF IsGood == .f.                 
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval =InsError('S', 'PKA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'ECA', m.recid)
*      InsErrorSV(m.mcod, 'R', 'ECA', m.recid)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.E1A == .T.  && Алгоритм E1
    IF !SEEK(people.d_type, 'osoree')
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E1A', m.recid)
*     InsErrorSV(m.mcod, 'R', 'E1A', m.recid)
    ENDIF 
   ENDIF 
   
   IF M.E2A == .T. && Алгоритм E2
    IF (!IsKms(people.sn_pol) AND !IsVS(people.sn_pol) AND !IsVSN(people.sn_pol) AND !IsENP(people.sn_pol))
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E2A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E2A', m.recid)
    ENDIF 
   ENDIF 

   IF  M.E4A == .T. && Алгоритм E4
    IF ((INLIST(RIGHT(PADL(ALLTRIM(People.fam),25),2),'ва','на','ая') AND INLIST(RIGHT(PADL(ALLTRIM(People.ot),20),2),'на','зы') AND People.w!=2) OR ;
       (INLIST(RIGHT(PADL(ALLTRIM(People.fam),25),2),'ов','ев','ин')  AND INLIST(RIGHT(PADL(ALLTRIM(People.ot),20),2),'ич','лы') AND People.w!=1))
     m.polis = sn_pol 
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E4A', m.recid)
*     InsErrorSV(m.mcod, 'R', 'E4A', m.recid)
    ENDIF 
   ENDIF 
   
   IF M.E7A == .T. && Алгоритм E7
    IF (!INLIST(people.w,1,2) OR (IsKms(people.sn_pol) AND SUBSTR(people.sn_pol,5,2)!='77' AND (people.w != IIF(VAL(SUBSTR(people.sn_pol,12,2))>50, 1, 2))))
     m.polis = sn_pol
     m.recsproc = 0 
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E7A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E7A', m.recid)
    ENDIF 
   ENDIF 

   IF M.E7A == .T.
    m.sn_pol = people.sn_pol                && Алгоритм E7
    Dtt = CTOD(IIF(VAL(SUBSTR(m.sn_pol,12,2))>50, ;
         PADL(INT(VAL(SUBSTR(m.sn_pol,12,2))-50),2,'0'), ;
         SUBSTR(m.sn_pol,12,2))+'.'+IIF(VAL(SUBSTR(m.sn_pol,14,2))>40, ;
         PADL(INT(VAL(SUBSTR(m.sn_pol,14,2))-40),2,'0')+'.20', ;
         SUBSTR(m.sn_pol,14,2)+'.19')+SUBSTR(m.sn_pol,16,2))
    IF (IsKms(m.sn_pol) AND !INLIST(SUBSTR(m.sn_pol,5,2),'50','51') AND (people.dr != Dtt))
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E7A', m.recid)
*     InsErrorSV(m.mcod, 'R', 'E7A', m.recid)
    ENDIF 
   ENDIF 

   IF M.E8A == .T.
    m.sn_pol = people.sn_pol                && Алгоритм E8
    IF (people.dr=={} OR (dat1-IIF(!EMPTY(people.dr), people.dr, {01.01.1850}))/365.25>120 OR ;
     IIF(!EMPTY(people.dr), people.dr, {01.01.1850}) > m.dat2)
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E8A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E8A', m.recid)
    ENDIF 
   ENDIF 
   
  ENDIF && Глобальное отключение ошибок регистра!

  IF talon.IsPr == .F. && Глобальное отключение ошибок счета!
   
   && Далее следуют алгоритмы проверки счета!

   IF M.H6A == .T. && Алгоритм H6
    m.polis=''
    DO CASE 
     CASE IsKms(sn_pol)
      m.polis = SUBSTR(sn_pol,8)
     CASE IsVs(sn_pol)
      m.polis = SUBSTR(sn_pol,7)
     OTHERWISE 
      m.polis = sn_pol
    ENDCASE 
    IF EMPTY(c_i)
     m.recid = recid
     rval =InsError('S', 'H6A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'H6A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
    IF (INLIST(cod,101927,101928,101951) OR BETWEEN(cod,101933,101945))
     IF SUBSTR(c_i,1,6)!='ПРОФД_' OR ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    IF INLIST(cod,101946,101947,101948)
     IF SUBSTR(c_i,1,6)!='ПРЕДД_' OR ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    IF INLIST(cod,101949,101950)
     IF SUBSTR(c_i,1,4)!='ПОД_' OR ALLTRIM(SUBSTR(c_i,5))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    IF (BETWEEN(cod,1900,1905) OR BETWEEN(cod,101929,101932))
     IF !INLIST(SUBSTR(c_i,1,3),'ДД_','ДС_','ДУ_') OR ALLTRIM(SUBSTR(c_i,4))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    IF BETWEEN(cod,1906,1909)
     IF SUBSTR(c_i,1,6)!='ПРОФВ_' OR ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    
    IF people.d_type='9' AND OCCURS('#', ALLTRIM(c_i))>0
     m.c_i = ALLTRIM(c_i)
     IF OCCURS('#', m.c_i)!=3
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX('H6A',0+64,m.c_i)
     ELSE 
      DO CASE 
       CASE !INLIST(SUBSTR(m.c_i,AT('#',m.c_i)+1,1),'1','2')
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*        MESSAGEBOX('H6A',0+64,m.c_i)
       CASE EMPTY(CTOD(SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),1,4)))
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*        MESSAGEBOX('H6A',0+64,m.c_i)
       CASE !INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6')
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*        MESSAGEBOX('H6A',0+64,m.c_i)
      ENDCASE 
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.COA == .T. && Алгоритм CO 
*    IF EMPTY(otd) OR LEN(ALLTRIM(otd))!=4 OR ;
     (!ISDIGIT(SUBSTR(otd,1,1)) OR !ISDIGIT(SUBSTR(otd,2,1)) OR !ISDIGIT(SUBSTR(otd,3,1)))
    IF EMPTY(otd) OR ;
     (!ISDIGIT(SUBSTR(otd,1,1)) OR !ISDIGIT(SUBSTR(otd,2,1)) OR !ISDIGIT(SUBSTR(otd,3,1)))
     m.recid = recid
     rval =InsError('S', 'COA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'COA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.HCA == .T. && Алгоритм HC
    IF k_u <= 0
     m.recid = recid
     rval = InsError('S', 'HCA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'HCA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   m.o_otd   = SUBSTR(otd,2,2)
   m.usl_ok  = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')
   m.is_gsp  = IIF(m.usl_ok='1', .T., .F.)

   IF M.OGA == .T. AND M.O0A == .T. AND m.is_gsp && Алгоритм OG
    m.recid  = recid
    m.ds_onk = ds_onk
    IF !INLIST(m.ds_onk,0,1)
     m.recid = recid
     rval = InsError('S', 'OGA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.O4A == .T. AND M.O0A == .T.  AND m.is_gsp && Алгоритм O4
    m.recid  = recid
    m.p_cel  = p_cel
    m.o_otd  = SUBSTR(otd,2,2)
    m.usl_ok = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')
    m.is_amb = IIF(m.usl_ok='3', .T., .F.)

    IF m.is_amb AND !SEEK(m.p_cel, 'onpcel')
     m.recid = recid
     rval = InsError('S', 'O4A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.OHA == .T. AND M.O0A == .T. AND m.is_gsp  && Алгоритм OH
    m.recid = recid
    m.p_cel = p_cel
    m.dn    = dn
    IF m.p_cel = '1.3' AND !INLIST(m.dn,1,2,3,4)
     m.recid = recid
     rval = InsError('S', 'OHA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.OIA == .T. AND M.O0A == .T. AND m.is_gsp  && Алгоритм OI
    m.recid = recid
    m.reab  = reab
    IF !INLIST(m.reab,0,1)
     m.recid = recid
     rval = InsError('S', 'OIA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.ENA == .T. && Алгоритм EN
    m.recid = recid
    m.cod   = cod
    m.tal_d = tal_d
    IF IsVmp(m.cod) AND EMPTY(m.tal_d)
     m.recid = recid
     rval = InsError('S', 'ENA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.O3A == .T. AND M.O0A == .T. AND m.is_gsp  && Алгоритм O3
    m.recid = recid
    m.napr_v_in =  napr_v_in
    IF !EMPTY(m.napr_v_in) AND !SEEK(m.napr_v_in, 'onnapr')
     m.recid = recid
     rval = InsError('S', 'O3A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.O5A == .T. AND M.O0A == .T. AND m.is_gsp  && Алгоритм O3
    m.recid = recid
    m.c_zab = c_zab
    IF !EMPTY(m.c_zab) AND !SEEK(m.c_zab, 'onczab')
     m.recid = recid
     rval = InsError('S', 'O5A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.DUA == .T. && Част алгоритма DU
    IF (people.tip_p==1 AND MONTH(d_u)!=tMonth)
     m.recid = recid
     rval = InsError('S', 'DUA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'DUA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.H8A == .T. && Част алгоритма H8
    IF !SEEK(ds, 'mkb10')
     m.recid = recid
     rval = InsError('S', 'H8A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'H8A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE 
     IF INLIST(LEFT(ds,3),'B95','B96','B97')
      m.recid = recid
      rval = InsError('S', 'H8A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'H8A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     m.cod = cod
     m.ds = ds
     IF (LEFT(m.ds,1)='Z' AND !INLIST(m.ds,'Z13.8','Z01.7','Z20','Z34','Z35')) AND INLIST(FLOOR(m.cod/1000),25,26,27,28,29,30,125,126,127,128,129,130)
      m.recid = recid
      rval = InsError('S', 'H8A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'H8A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF
   ENDIF 

   IF M.HEA == .T.  && Алгоритм HE
    m.sex = IIF(OCCURS('#',c_i)==2, SUBSTR(c_i, AT('#',c_i,1)+1, 1), STR(people.w,1))
    IF (SEEK(ds, 'mkb10') AND !EMPTY(mkb10.sex)) AND m.sex != mkb10.sex
     m.recid = recid
     rval = InsError('S', 'HEA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'HEA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 
   
   IF M.CSA == .T.
    IF !SEEK(cod, 'tarif')
     m.recid = recid
     rval = InsError('S', 'CSA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'CSA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.TFA == .T.
    IF m.IsOtdSkp AND !INLIST(Tip, '0', 'A', 'v')
     m.recid = recid
     rval = InsError('S', 'TFA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF SEEK(cod, 'tarif')
     IF !EMPTY(Tarif.n_kd) AND !SEEK(Tip, 'kpresl')
      m.recid = recid
      rval = InsError('S', 'TFA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
   ENDIF 

   IF M.TLA == .T.
    m.cod = cod
    m.k_u = k_u
    IF (INLIST(m.cod,97041,97013,197013) OR BETWEEN(cod, 84000, 84999)) AND m.k_u>1
     m.recid = recid
     rval = InsError('S', 'TLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 

   IF M.THA == .T.
    m.cod = cod
    m.otd = otd
    m.c_i    = c_i
    m.sn_pol = sn_pol
    m.tip = tip

    IF m.IsOtdSkp
*     m.IsWithOper = IIF(SEEK(m.cod, 'reeskp', 'cod'), .t., .f.)
*     IF m.IsWithOper
      IF !USED('ho')
       m.recid = recid
       rval = InsError('S', 'THA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ELSE 
       m.vir = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
       IF !SEEK(m.vir, 'ho')
        m.recid = recid
        rval = InsError('S', 'THA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
*     ENDIF 
    ENDIF 

    IF INLIST(m.tip,'0', '8', 'А', 'v')
     IF IsHOOtd(m.otd) AND !SEEK(m.cod, 'noth')
      m.vir = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
      IF !USED('ho')
       m.recid = recid
       rval = InsError('S', 'THA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ELSE 
       IF !SEEK(m.vir, 'ho')
        m.recid = recid
        rval = InsError('S', 'THA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.HOA == .T.
    IF m.IsOtdSkp
     IF USED('ho')
      m.sn_pol = sn_pol
      m.c_i    = c_i
      m.cod    = cod
      
      m.vir = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
      IF SEEK(m.vir, 'ho')
       m.cod   = cod
       m.codho = ho.codho
       m.ds    = ds
       m.vir   = PADL(m.cod,6,'0') + m.codho + LEFT(m.ds,5)
*       MESSAGEBOX(m.vir, 0+64, '')
       IF !SEEK(m.vir, 'reeskp', 'unik')
        m.recid = recid
        rval = InsError('S', 'HOA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ELSE 
       m.recid = recid
       rval = InsError('S', 'HOA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ELSE 
      m.recid = recid
      rval = InsError('S', 'HOA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.TVA == .T.
    m.tiplpu = SUBSTR(m.mcod,2,1)
    IF m.tiplpu!='3'
     IF SEEK(cod, 'codwdr') AND (!EMPTY(codwdr.kp) AND m.tiplpu!=codwdr.kp)
      m.recid = recid
      rval = InsError('S', 'TVA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'TVA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
   ENDIF 

   IF M.TVA == .T.
    IF INLIST(m.lpuid,2202,1873,1872,1871,1874) && Св. Владимира, Сперанского, Морозовская, Башляевой, Филатова (13)
     m.profil  = profil
*     IF cod<=99999 AND m.profil!='034'
     IF (cod>=61000 AND cod<=99999) AND m.profil!='034'
      m.recid = recid
      rval = InsError('S', 'TVA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'TVA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
   ENDIF 

   IF M.FSA == .T.
    *IF !IsPilots AND !INLIST(m.mcod, '4344623','4344700','4344621')
    IF !IsPilots AND !INLIST(m.mcod, '4344623','4344700','4344621','4344640','4344613','4134752') && С 01.02.2019 
    * IF m.IsExHorS
    
*     IF m.IsStomatUsl AND (EMPTY(people.prmcods) AND people.prmcods<>people.mcod)
     m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
     IF m.IsStomatUsl
      m.recid = recid
      rval = InsError('S', 'FSA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.NLA == .T.
    m.tiplpu = IIF(VAL(SUBSTR(m.mcod,3,2))<41, 'p', IIF(VAL(SUBSTR(m.mcod,3,2))!=71, 's', 'b'))
    IF SEEK(cod, 'codwdr') AND !EMPTY(codwdr.stac) AND LOWER(codwdr.stac) != m.tiplpu
     IF INLIST(cod,29006,29007) AND INLIST(m.lpuid,1863,1891,1842,5009)
*      MESSAGEBOX('NLA!',0+64,'1')
     ELSE 
*      MESSAGEBOX('NLA!',0+64,'2')
      m.recid = recid
      rval = InsError('S', 'NLA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF

    * Диспансеризация детей-сирот
    IF INLIST(cod,101929,101930,101931,101932) AND 1=2
    IF !SEEK(m.lpuid, 'dsdisp')
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
    ENDIF 
    * Диспансеризация детей-сирот
    
    IF IsVMP(cod) AND !SEEK(m.lpuid, 'movmp') AND 1=2
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    * Диспансеризация взрослых
    IF INLIST(cod,1900,1901,1902,1903,1904,1905) AND 1=2
     IF !SEEK(m.lpuid, 'spidd')
      m.recid = recid
      rval = InsError('S', 'NLA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
    * Диспансеризация взрослых

    * Неотложка в ЛПУ
    IF (INLIST(cod, 56031, 156002) AND m.tdat1>={01.10.2017}) AND LEFT(m.mcod,1)='0'
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    * Неотложка в ЛПУ

   ENDIF 

   IF M.MDA == .T.
    m.sex = IIF(OCCURS('#',c_i)==2, SUBSTR(c_i, AT('#',c_i,1)+1, 1), STR(people.w,1))
    IF (SEEK(cod, 'codwdr') AND !EMPTY(codwdr.sex)) AND m.sex != codwdr.sex
     m.recid = recid
     rval = InsError('S', 'MDA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'MDA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.H3A == .T.
    IF !SEEK(d_type, 'ososch')
     m.recid = recid
     rval = InsError('S', 'H3A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'H3A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE 
*     IF (IsMes(cod) AND ((Tip='5' AND d_type!='5') OR (Tip!='5' AND  d_type='5'))) OR
     IF (IsMes(cod) AND (Tip!='5' AND  d_type='5')) OR ;
      cod=1561 AND d_type!='5'
      m.recid = recid
      rval = InsError('S', 'H3A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'H3A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF
   ENDIF 
   
   IF M.SOA == .T.
    m.tiplpu = IIF(VAL(SUBSTR(m.mcod,3,2))<41, 'p', 's')
    IF !SEEK(SUBSTR(otd,2,2), 'profot')
     m.recid = recid
     rval = InsError('S', 'SOA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'SOA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
*     IF profot.stac='s' AND m.tiplpu=='p'
*      m.recid = recid
*      rval = InsError('S', 'SOA', m.recid)
**      InsErrorSV(m.mcod, 'S', 'SOA', m.recid)
*      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     ENDIF 
    ENDIF 
   ENDIF 

   IF M.R1A == .T.
    IF EMPTY(ishod)
     m.recid = recid
     rval = InsError('S', 'R1A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(ishod, 'isv012')
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     && Проверка 14.01.2019 соответствия условиям оказания
     m.otd = otd
     DO CASE 
      CASE !(IsDstOtd(m.otd) OR IsPlkOtd(m.otd)) AND LEFT(STR(ishod,3),1) != '1'
      *CASE IsGsp(cod) AND LEFT(STR(ishod,3),1) != '1'
       m.recid = recid
       rval = InsError('S', 'R1A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE IsDstOtd(m.otd) AND LEFT(STR(ishod,3),1) != '2'
      *CASE IsDst(cod) AND LEFT(STR(ishod,3),1) != '2'
       m.recid = recid
       rval = InsError('S', 'R1A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE IsPlkOtd(m.otd) AND LEFT(STR(ishod,3),1) != '3'
      *CASE IsPlk(cod) AND LEFT(STR(ishod,3),1) != '3'
       m.recid = recid
       rval = InsError('S', 'R1A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      OTHERWISE 
     ENDCASE 
     && Проверка 14.01.2019 соответствия условиям оказания
    ENDIF 

    IF 3=2
    IF BETWEEN(cod,101927,101928) OR BETWEEN(cod,101933,101945) OR cod=101951 && Профилактические
     IF !BETWEEN(rslt,332,336) AND rslt!=326
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF BETWEEN(cod,101946,101948) && Предварительные
     IF !BETWEEN(rslt,337,341) AND rslt!=396
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF BETWEEN(cod,1900,1905) && Диспансеризация взрослых
     IF !BETWEEN(rslt,317,318) AND !BETWEEN(rslt,352,353) AND !BETWEEN(rslt,355,358) AND !BETWEEN(rslt,321,325) AND rslt!=320
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF BETWEEN(cod,101929,101932) && Диспансеризация детей-сирот в стационаре
     IF !BETWEEN(rslt,317,318) AND !BETWEEN(rslt,352,353) AND !BETWEEN(rslt,355,358) AND ;
     !BETWEEN(rslt,321,325) AND !BETWEEN(rslt,347,351) AND rslt!=320 AND rslt!=390
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF
    ENDIF  

   ENDIF 

   IF M.R2A == .T.
    IF EMPTY(rslt)
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(rslt, 'rsv009')
      m.recid = recid
      rval = InsError('S', 'R2A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    IF BETWEEN(cod,1900,1905) AND !(BETWEEN(rslt,317,319) OR BETWEEN(rslt,352,354) OR BETWEEN(rslt,355,358))
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF BETWEEN(cod,101929,101932) AND !(BETWEEN(rslt,321,325) OR BETWEEN(rslt,347,351) OR INLIST(rslt,320,390))
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF BETWEEN(cod,1906,1909) AND !BETWEEN(rslt,343,345)
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF (INLIST(cod,101927,101928,101951) OR BETWEEN(cod,101933,101945)) AND rslt=304
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF BETWEEN(cod,101946,101948) AND rslt=304
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    

    IF BETWEEN(cod,101949,101950) AND rslt!=342
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 

   IF M.R3A == .T.
    IF EMPTY(prvs)
     m.recid = recid
     rval = InsError('S', 'R3A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'R3A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(prvs, 'kspec')
      m.recid = recid
      rval = InsError('S', 'R3A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'R3A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.NRA == .T.
    m.cod = cod
    IF IsUsl(m.cod) AND (SEEK(m.cod, 'tarif') AND tarif.tpn!='q')
      m.recid   = recid
      m.ord     = ord
      m.lpu_ord = lpu_ord
      m.sn_pol  = sn_pol
      IF (SEEK(m.sn_pol, 'people') AND !EMPTY(people.prmcod) AND people.prmcod!=m.mcod) AND (EMPTY(m.lpu_ord) OR m.ord=0)
       rval = InsError('S', 'NRA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
    ENDIF 
   ENDIF 
   
   IF m.IsOtdSkp AND 1=2
    m.ord   = ord
    IF !INLIST(m.ord,1)
     rval = InsError('S', 'NRA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 
   
   IF M.NRA == .T. AND 1=2  && Версия до 15.01.2019!
    IF !EMPTY(m.cfrom) && Версия для ЕМИАС
     m.cod     = cod 
     m.sn_pol  = sn_pol
     m.facotd  = SUBSTR(otd,2,2)
     m.profil  = profil
     m.lpu_ord = lpu_ord

     DO CASE 
      CASE IsMes(m.cod)
       m.recid = recid
       m.ord   = ord
       m.lpu_ord = lpu_ord
       IF !INLIST(m.ord,1,2,3,5,6)
        rval = InsError('S', 'NRA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 

      CASE IsVmp(m.cod)
       m.recid = recid
       m.ord   = ord
       m.lpu_ord = lpu_ord
       IF !INLIST(m.ord,1,2,3,5)
        rval = InsError('S', 'NRA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 

      OTHERWISE 
      
       m.ord = ord
       m.isstomatusl  = IIF(INLIST(FLOOR(m.cod / 1000), 9, 109), .T., .F.)
       m.isstomatusl2 = IIF(INLIST(m.cod, 1101, 1102, 101171, 101172), .T., .F.)
       IF (SEEK(m.cod, 'tarif') AND tarif.tpn='q') AND (m.isstomatusl OR m.isstomatusl2) AND m.ihors
        IF m.ord=0
         rval = inserror('S','NRB',m.recid)
         m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
        ENDIF 
       ENDIF 

       IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p') AND !INLIST(m.facotd,'01','08', '91', '92', '70', '73') ;
        AND m.profil!='100' AND !IsPat(m.cod) AND !IsSimult(m.cod) AND !INLIST(m.cod, 49001,149002)
		
      DO CASE
           CASE (m.ishor .AND.  .NOT. m.ishors) && .OR. (m.ispilot .AND.  .NOT. m.ispilots)
                IF (SEEK(m.sn_pol, 'people') .AND.  .NOT. EMPTY(people.prmcod) .AND. people.prmcod <> people.mcod)
                     m.recid = recid
                     m.ord = ord
                     IF  .NOT. INLIST(m.ord, 4, 6, 8, 9) .AND.  .NOT. (m.ord = 7 .AND. m.lpu_ord = 7665)
                          rval = inserror('S','NRA',m.recid)
                          m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
*                          MESSAGEBOX('1',0+64,'')
                     ENDIF
                ELSE
                     m.recid = recid
                     m.ord = ord
                     IF  .NOT. INLIST(m.ord, 0, 4, 6, 8, 9) .AND.  .NOT. (m.ord = 7 .AND. m.lpu_ord = 7665)
                          rval = inserror('S','NRA',m.recid)
                          m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
*                          MESSAGEBOX('2',0+64,'')
                     ENDIF
                ENDIF
           CASE (m.ishors .AND.  .NOT. m.ishor) && .OR. (m.ispilots .AND.  .NOT. m.ispilot)
                m.usliskl = IIF(FLOOR(m.cod / 1000) = 146, .T., .F.)
                m.isstomatusl = IIF(INLIST(FLOOR(m.cod / 1000), 9, 109), .T., .F.)
                m.isstomatusl2 = IIF(INLIST(m.cod, 1101, 1102, 101171, 101172), .T., .F.)
                IF ((m.isstomat .AND.  .NOT. m.isiskl) .AND. (m.isstomatusl .OR. m.isstomatusl2)) .OR. ((m.isstomat .AND. m.isiskl) .AND. (m.isstomatusl .OR. m.isstomatusl2 .OR. m.isiskl)) .OR. ( .NOT. m.isstomat .AND. (m.isstomatusl .OR. (m.isstomatusl2 .AND. LEFT(m.ds, 2) = 'K0')))
                     IF (SEEK(m.sn_pol, 'people') .AND.  .NOT. EMPTY(people.prmcods) .AND. people.prmcods <> people.mcod)
                          m.recid = recid
                          m.ord = ord
                          IF  .NOT. INLIST(m.ord, 4, 6, 8, 9) .AND.  .NOT. (m.ord = 7 .AND. m.lpu_ord = 7665)
                               rval = inserror('S','NRA',m.recid)
                               m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
*	                          MESSAGEBOX('3',0+64,'')
                          ENDIF
                     ELSE
                          m.recid = recid
                          m.ord = ord
                          IF  .NOT. INLIST(m.ord, 0, 4, 6, 8, 9) .AND.  .NOT. (m.ord = 7 .AND. m.lpu_ord = 7665)
                               rval = inserror('S','NRA',m.recid)
                               m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
 *                         MESSAGEBOX('4',0+64,'')
                          ENDIF
                     ENDIF
                ELSE
                ENDIF
           CASE (m.ishor .AND. m.ishors) && .OR. (m.ispilot .AND. m.ispilots)
                m.usliskl = IIF(FLOOR(m.cod / 1000) = 146, .T., .F.)
                m.isstomatusl = IIF(INLIST(FLOOR(m.cod / 1000), 9, 109), .T., .F.)
                m.isstomatusl2 = IIF(INLIST(m.cod, 1101, 1102, 101171, 101172), .T., .F.)
                IF ((m.isstomat .AND.  .NOT. m.isiskl) .AND. (m.isstomatusl .OR. m.isstomatusl2)) .OR. ((m.isstomat .AND. m.isiskl) .AND. (m.isstomatusl .OR. m.isstomatusl2 .OR. m.isiskl)) .OR. ( .NOT. m.isstomat .AND. (m.isstomatusl .OR. (m.isstomatusl2 .AND. LEFT(m.ds, 2) = 'K0')))
                     IF (SEEK(m.sn_pol, 'people') .AND.  .NOT. EMPTY(people.prmcods) .AND. people.prmcods <> people.mcod)
                          m.recid = recid
                          m.ord = ord
                          IF  .NOT. INLIST(m.ord, 4, 6, 8, 9) .AND.  .NOT. (m.ord = 7 .AND. m.lpu_ord = 7665)
                               rval = inserror('S','NRA',m.recid)
                               m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
*                          		MESSAGEBOX('5',0+64,'')
                          ENDIF
                     ELSE
                          m.recid = recid
                          m.ord = ord
                          IF  .NOT. INLIST(m.ord, 0, 4, 6, 8, 9) .AND.  .NOT. (m.ord = 7 .AND. m.lpu_ord = 7665)
                               rval = inserror('S','NRA',m.recid)
                               m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
*                          		MESSAGEBOX('6',0+64,'')
                          ENDIF
                     ENDIF
                ELSE
                     IF (SEEK(m.sn_pol, 'people') .AND.  .NOT. EMPTY(people.prmcod) .AND. people.prmcod <> people.mcod)
                          m.recid = recid
                          m.ord = ord
                          IF  .NOT. INLIST(m.ord, 4, 6, 8, 9) .AND.  .NOT. (m.ord = 7 .AND. m.lpu_ord = 7665)
                               rval = inserror('S','NRA',m.recid)
                               m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
*		                          MESSAGEBOX('7',0+64,'')
                          ENDIF
                     ELSE
                          m.recid = recid
                          m.ord = ord
                          IF  .NOT. INLIST(m.ord, 0, 4, 6, 8, 9) .AND.  .NOT. (m.ord = 7 .AND. m.lpu_ord = 7665)
                               rval = inserror('S','NRA',m.recid)
                               m.s_flk = m.s_flk + IIF(rval == .T., s_all, 0)
*	                          MESSAGEBOX('8',0+64,'')
                          ENDIF
                     ENDIF
                ENDIF
      ENDCASE


       ENDIF 

     ENDCASE 

   ELSE  && IF !EMPTY(cfrom)

   *IF M.NRA == .T. && Версия после 15.01.2019!
    m.cod     = cod 
    m.sn_pol  = sn_pol
    m.facotd  = SUBSTR(otd,2,2)
    m.profil  = profil
    m.lpu_ord = lpu_ord
    m.otd = otd

   DO CASE 
    CASE IsPlkOtd(m.otd)
     m.recid = recid
     m.ord   = ord
     m.lpu_ord = lpu_ord
     IF !INLIST(m.ord,0,4,6,7,8,9)
      rval = InsError('S', 'NRA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 

    CASE IsDstOtd(m.otd)
     m.recid = recid
     m.ord   = ord
     m.lpu_ord = lpu_ord
     IF !INLIST(m.ord,0,1,2,3,5,6)
      rval = InsError('S', 'NRA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 

    OTHERWISE 
     m.recid = recid
     m.ord   = ord
     m.lpu_ord = lpu_ord
     IF !INLIST(m.ord,1,2,3,5,6)
      rval = InsError('S', 'NRA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 

   ENDCASE 
   ENDIF && IF !EMPTY(cfrom)
   ENDIF 

   IF M.G1A == .T.
    m.cod = cod 
    m.sn_pol = sn_pol
    m.lpu_ord = lpu_ord
    m.ordmcod = IIF(SEEK(m.lpu_ord, 'sprlpu'), sprlpu.mcod, '')
    m.facotd  = SUBSTR(otd,2,2)
    m.profil  = profil
    IF !(IsMes(m.cod) OR IsVmp(m.cod))
     IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p') AND !INLIST(m.facotd,'01','08', '91', '70', '73') ;
        AND m.profil!='100' AND !IsPat(m.cod) AND !IsSimult(m.cod)
      m.ord     = ord
      m.lpu_ord = lpu_ord
      m.recid   = recid
      DO CASE 
       CASE m.ord = 4
        IF IIF(m.qcod='P2', m.ordmcod!=people.prmcod, !SEEK(m.lpu_ord, 'sprlpu')) AND m.lpu_ord!=4708
*        IF !SEEK(m.lpu_ord, 'sprlpu') AND m.lpu_ord!=4708
         rval = InsError('S', 'G1A', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       CASE INLIST(m.ord,6,9)
        IF m.lpu_ord!=9999
         rval = InsError('S', 'G1A', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       CASE m.ord =8
        IF m.lpu_ord!=8888
         rval = InsError('S', 'G1A', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       OTHERWISE 
      ENDCASE 
     ENDIF 
    ELSE  && Если МЭС или ВМП
     m.ord     = ord
     m.lpu_ord = lpu_ord
     m.recid   = recid
     DO CASE 
      CASE m.ord = 1
*       IF IIF(m.qcod='P2', m.ordmcod!=people.prmcod, !SEEK(m.lpu_ord, 'sprlpu')) AND m.lpu_ord!=4708
       IF !SEEK(m.lpu_ord, 'sprlpu') AND  m.lpu_ord!=9999
        rval = InsError('S', 'G1A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      CASE m.ord = 2
       IF !INLIST(m.lpu_ord,4708,9999)
        rval = InsError('S', 'G1A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      CASE m.ord=6
       IF m.lpu_ord!=9999
        rval = InsError('S', 'G1A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      OTHERWISE 
     ENDCASE 
    ENDIF 
   ENDIF 

   IF M.G2A == .T.
    m.cod = cod 
    m.sn_pol = sn_pol
    
    IF qcod!='P2'
    
    IF !(IsMes(m.cod) OR IsVmp(m.cod))
     IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p')
      m.ord      = ord
      m.date_ord = date_ord
      m.recid    = recid
      m.d_u      = d_u
      IF INLIST(m.ord,4,6,8,9)
       IF EMPTY(m.date_ord) OR (!EMPTY(m.date_ord) AND m.date_ord>m.d_u)
        rval = InsError('S', 'G2A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
     ENDIF 
    ELSE 
     m.ord      = ord
     m.date_ord = date_ord
     m.recid    = recid
     m.d_u      = d_u
     IF INLIST(m.ord,1,2,6)
      IF EMPTY(m.date_ord) OR (!EMPTY(m.date_ord) AND m.date_ord>m.d_u)
       rval = InsError('S', 'G2A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
    
    ELSE 
    
     m.ord      = ord
     m.date_ord = date_ord
     m.recid    = recid
     m.d_u      = d_u
     IF INLIST(m.ord,1,4,6,8,9)
      IF EMPTY(m.date_ord) OR (!EMPTY(m.date_ord) AND m.date_ord>m.d_u)
       rval = InsError('S', 'G2A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 

    ENDIF 
    
   ENDIF 

*   IF M.G3A == .T.
    m.ord     = ord
    m.cod     = cod 
    m.sn_pol  = sn_pol
    m.recid   = recid
    m.lpu_ord = lpu_ord
    m.n_u     = ALLTRIM(n_u)
    
    IF FIELD('n_vmp')='N_VMP'
     IF IsVmp(m.cod)
      m.n_vmp   = ALLTRIM(n_vmp)
      IF m.ord=1 AND EMPTY(m.n_vmp)
       rval = InsError('S', 'G3A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
     
     m.n_vmp = ALLTRIM(n_vmp)
     IF m.cod=97041 AND !IsOkNVmpForEco(m.n_vmp)
      rval = InsError('S', 'G3A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 

    ENDIF 
*   ENDIF 

	

   IF M.G3A == .T.
    m.cod    = cod 
    m.sn_pol = sn_pol
    m.recid   = recid
    m.lpu_ord = lpu_ord
    m.n_u     = ALLTRIM(n_u)
    
*    IF FIELD('n_vmp')='N_VMP'
*     IF IsVmp(m.cod)
*      m.n_vmp   = ALLTRIM(n_vmp)
*      IF m.ord=1 AND EMPTY(m.n_vmp)
*       rval = InsError('S', 'G3A', m.recid)
*       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      ENDIF 
*     ENDIF 
*    ENDIF 

*    IF qcod!='P2'

    IF !(IsMes(m.cod) OR IsVmp(m.cod))
     IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p')
      IF INLIST(m.ord,4,6,8,9) AND m.lpu_ord = 4708 AND EMPTY(m.n_u)
       rval = InsError('S', 'G3A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ELSE 
     IF INLIST(m.ord,1,2,6) AND m.lpu_ord = 4708 AND EMPTY(m.n_u)
      rval = InsError('S', 'G3A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 

*    ELSE 

*    IF IsMes(m.cod) OR IsVmp(m.cod)
*     IF m.lpu_ord = 4708 AND EMPTY(m.n_u)
*      rval = InsError('S', 'G3A', m.recid)
*      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     ENDIF 
*    ENDIF 

*    ENDIF 

   ENDIF 

   IF M.G4A == .T.
    m.cod = cod 
    m.sn_pol = sn_pol
    IF IsMes(m.cod) OR IsVmp(m.cod)
     m.recid = recid
     m.ord   = ord
     m.ds_0 = ALLTRIM(ds_0)
     IF INLIST(m.ord,1,2,6) AND EMPTY(m.ds_0)
      rval = InsError('S', 'G4A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.W2A == .T.
    m.cod = cod 
    m.sn_pol = sn_pol
    
    IF FLOOR(m.cod/1000)=300 OR INLIST(m.cod,1719,8050,8051,8052,26281,28210,31001,31002,31003,40040,40041,40042,40043,40044)
     m.recid = recid
     rval = InsError('S', 'W2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 

   IF m.mcod = '0141045' AND m.gcPeriod='201805'
    m.cod = cod 
    m.sn_pol = sn_pol
    m.d_u = d_u
    
    IF m.d_u>{14.05.2018}
     m.recid = recid
     rval = InsError('S', 'W2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 

   IF M.UVA == .T.
    m.d_u = d_u
    m.ldr = people.dr
    m.ddr = IIF(OCCURS('#',c_i)>=2, ;
     CTOD(SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),1,4)), ;
     people.dr)
    IF !EMPTY(m.ddr)
     nmonthes = ((m.d_u-m.ddr)/365.25)*12 && Переделал с m.dat2 на m.d_u 13.03.2017 по просьбе УралСиба
*     nmonthes = ((m.dat2-m.ddr)/365.25)*12 && Переделал с m.dat2 на m.d_u 13.03.2017 по просьбе УралСиба
*     nmonthes = ((m.ldr-m.ddr)/365.25)*12
     IF SEEK(cod, 'codwdr') AND (!BETWEEN(nmonthes, IIF(BETWEEN(cod,1821,1825), 0, CodWDr.min_ms), CodWDr.max_ms) AND ;
       (!INLIST(d_type,'1','2','5') AND ;
       (!SEEK(ds, 'nocodr', 'ds1') AND !SEEK(ds, 'nocodr', 'ds2') AND !SEEK(ds, 'nocodr', 'ds3'))))
      m.recid = recid
      rval = InsError('S', 'UVA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'UVA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.DVA == .T.
    IF !EMPTY(Tip) AND !INLIST(SUBSTR(PADL(Cod,6,'0'),2,2), '83', '84')
     IF IsMes(Cod)
      m.IsOtdSkp = IIF(m.IsSkp AND SUBSTR(otd,2,2)='09', .T., .F.)
      m.perem = IIF(!ISDIGIT(SUBSTR(Ds,5,1)), STR(Cod,6)+' '+LEFT(Ds,3)+'   ', STR(Cod,6)+' '+Ds)
      IF (!SEEK(IIF(!m.IsOtdSkp, m.perem, LEFT(m.perem,5)), IIF(m.IsOtdSkp, 'ReesKp', 'MesMkb'), 'ds_ms') AND !INLIST(d_type,'1','5'))
       m.recid = recid
       rval = InsError('S', 'DVA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ELSE && Если ВМП
      m.IsErr=.t.
      FOR m.opl=0 TO 3
       m.perem = STR(Cod,6)+' '+SUBSTR(Ds,1,6-m.opl)+SPACE(m.opl)
       IF SEEK(m.perem,'MesMkb')
        m.IsErr=.f.
        EXIT 
       ENDIF 
      ENDFOR 
      IF m.IsErr=.t.
       m.recid = recid
       rval = InsError('S', 'DVA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 

   && Проверка по "y" && Только такие услуги в этом отделении
   IF M.UOA == .T.
    SET ORDER TO notd IN CodOtd
    IsCheck = IIF(SEEK(SUBSTR(otd,4,6), 'CodOtd', 'notd'), .T., .F.)
    IF IsCheck AND d_type!='2'
     IsOk = .f.
     DO WHILE SUBSTR(otd,4,6) = CodOtd.otd
      IF cod = CodOtd.Cod
       IsOk = .t.
       EXIT
      ENDIF
      SKIP IN CodOtd
     ENDDO 
    
     IF IsOk = .f.
      m.recid = recid
      rval = InsError('S', 'UOA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'UOA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    && Добавлено 12.02.2018
    && Исключено 18.01.2019
    IF 1=2
    m.cod = cod
    m.facotd  = SUBSTR(otd,2,2)
    m.recid = recid
    IF BETWEEN(m.cod, 76411, 76570)
     IF INLIST(m.cod, 76431, 76521, 76530)
      IF !INLIST(m.facotd,'39','40')
       rval = InsError('S', 'UOA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ELSE 
      IF !INLIST(m.facotd,'39')
       rval = InsError('S', 'UOA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 	

    IF BETWEEN(m.cod, 76581, 76891)
     IF !INLIST(m.facotd,'38','39','40')
      rval = InsError('S', 'UOA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 	
    ENDIF 
    && Исключено 18.01.2019
    && Добавлено 12.02.2018
    
   ENDIF 
   
   IF !(INLIST(m.cod,25050,25203,25243,25271,25268,26003,26087,26158,26180,26229,26243,26275,27005,27009,27010,27015,27016,27019,27020,27024,27027) OR ;
   	  INLIST(m.cod,27028,27030,27032,27050,27061,27071,28019,28021,28024,28050,28067,28077,28093,28097,28119,28126,28144,28147,28164,28188,28190,28208) OR ;
   	  INLIST(m.cod,125050,125203,125243,125271,125268,126003,126087,126158,126180,126229,126243,126275,127005,127009,127010,127015,127016,127019,127020,127024,127027) OR ;
   	  INLIST(m.cod,127028,127030,127032,127050,127061,127071,128019,128021,128024,128050,128067,128077,128093,128097,128119,128126,128144,128147,128164,128188,128190,128208))
      
   IF M.NOA == .T.
    m.perem = sn_pol+str(cod,6)+dtos(d_u)
    IF SEEK(m.perem, 'e_day') AND d_type!='2'
     m.ocntr = e_day.cntr
     m.ncntr = m.ocntr + k_u
     IF m.ncntr<=e_day.k_norm
      REPLACE e_day.cntr WITH m.ncntr IN e_day
     ELSE 
     m.recid = recid
     rval = InsError('S', 'NOA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'NOA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.NMA == .T.
    m.perem = sn_pol+str(cod,6)
    IF SEEK(m.perem, 'e_month') AND d_type!='2'
     m.ocntr = e_month.cntr
     m.ncntr = m.ocntr + k_u
     IF m.ncntr<=e_month.k_norm
      REPLACE e_month.cntr WITH m.ncntr IN e_month
     ELSE 
     m.recid = recid
     rval = InsError('S', 'NMA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'NMA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    ENDIF 
   ENDIF 

   ENDIF 

   IF M.D2A == .T.
    m.cod = cod
    m.dsptip = IIF(SEEK(m.cod,'dspcodes'), dspcodes.tip, 0)
    m.oldcod = 0 

    IF m.dsptip > 0
    
    DO CASE 
     CASE m.dsptip = 1  && Диспансеризция взрослых, tip=1
      *m.lastt = 12*3
      m.lastt = dspcodes.last
     CASE m.dsptip = 2 && Профосмотры взрослых, tip=2
      *m.lastt = 12*2
      m.lastt = dspcodes.last
     CASE m.dsptip = 3 && Диспансеризация детей, tip=3
      *m.lastt = 12
      m.lastt = dspcodes.last
     CASE m.dsptip = 4 && Профосмотры детей, tip=4
      *m.lastt = IIF(m.qcod != 'P2', 12, 3)
      *m.lastt = IIF(m.qcod != 'P2', 12, 3)
      *m.lastt = dspcodes.last
      m.lastt = 0
     CASE m.dsptip = 5 && Предварительные, tip=5
      *m.lastt = 3
      m.lastt = dspcodes.last
     CASE m.dsptip = 6 && Периодические, tip=6
      *m.lastt = 3
      m.lastt = dspcodes.last
     OTHERWISE 
      m.lastt = 0
    ENDCASE
    
    m.lastt = IIF(!INLIST(m.cod, 25204, 35401), m.lastt, 0)

    
    IF m.dsptip=4 AND 1=2
      m.perem = m.mcod+LEFT(sn_pol,17)+PADL(3,1,'0')
      IF m.qcod!='P2'
       IF SEEK(m.perem, 'dspyear')
        m.recid = recid
        rval = InsError('S', 'D2A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ELSE
       IF SEEK(m.perem, 'dspp') AND (d_u - dspp.d_u)/30<m.lastt
        m.recid = recid
        rval = InsError('S', 'D2A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
    ENDIF 

    IF m.lastt>0
*     m.perem = m.mcod+LEFT(sn_pol,17)+PADL(cod,6,'0')
     m.perem = m.mcod+LEFT(sn_pol,17)+PADL(m.dsptip,1,'0')
     IF SEEK(m.perem, 'dspp') AND (d_u - dspp.d_u)/30<m.lastt
      m.recid = recid
      rval = InsError('S', 'D2A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     IF m.oldcod>0
      m.perem = m.mcod+LEFT(sn_pol,17)+PADL(m.oldcod,6,'0')
      IF SEEK(m.perem, 'dspp') AND (d_u - dspp.d_u)/30<m.lastt
       m.recid = recid
       rval = InsError('S', 'D2A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
    
   ENDIF && IF m.dsptip>0
   
   ENDIF && IF M.D2A == .T.

   IF M.NUA == .T.
    SET ORDER TO ncod IN sovmno
    IF SEEK(cod, 'sovmno') && Алгоритм NU - несовместимые услуги 
     DO WHILE sovmno.cod == cod
      IF SEEK(sn_pol+STR(sovmno.cod_1,6)+DTOS(d_u), 'talon_exp')
       IF (EMPTY(UPPER(sovmno.Stac)) OR (!IsStac AND UPPER(sovmno.Stac)='P') OR ;
        (IsStac AND UPPER(sovmno.Stac)='S')) AND (d_type != '2' OR talon_exp.d_type != '2')
        m.recid = recid
        rval = InsError('S', 'NUA', m.recid)
*        InsErrorSV(m.mcod, 'S', 'NUA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        m.recid = talon_exp.recid
        rval = InsError('S', 'NUA', m.recid)
*        InsErrorSV(m.mcod, 'S', 'NUA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
      SKIP +1 IN sovmno 
     ENDDO 
    ENDIF
   ENDIF  

   IF M.NSA == .T.
    SET ORDER TO scod IN sovmno
    IF SEEK(cod, 'sovmno') && Алгоритм NS - несовместимые услуги 
     IsSovmUsl = .F.
     DO WHILE sovmno.cod == cod
      IF SEEK(sn_pol+STR(sovmno.cod_1,6)+DTOS(d_u), 'talon_exp')
       IsSovmUsl = .T.
       EXIT 
      ENDIF 
      SKIP +1 IN sovmno 
     ENDDO 
     IF !IsSovmUsl
      IF (EMPTY(UPPER(sovmno.Stac)) OR (!IsStac AND UPPER(sovmno.Stac)='P') OR ;
         (IsStac AND UPPER(sovmno.Stac)='S')) AND d_type != '2'
       m.recid = recid
       rval = InsError('S', 'NSA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'NSA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.HNA == .T.
    m.cod    = cod
    m.sn_pol = sn_pol
    m.d_u    = d_u
    IF INLIST(m.cod,15001,115001) AND SEEK(m.sn_pol, 'polic_h') AND m.d_u - polic_h.d_u < 365
     m.recid = recid
     rval = InsError('S', 'HNA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'HNA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    RELEASE cod, sn_pol, d_u
   ENDIF 

   IF M.HNA == .T.
    m.cod    = cod
    m.sn_pol = sn_pol
    m.d_u    = d_u
    IF INLIST(m.cod,101927,101928) AND SEEK(m.sn_pol, 'polic_dp')
     m.recid = recid
     rval = InsError('S', 'HNA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'HNA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    RELEASE cod, sn_pol, d_u
   ENDIF 

   IF M.POA == .T. && Диспансеризация в непрофильном ЛПУ
    m.cod    = cod
*    IF INLIST(m.cod, 101927, 101928) AND !SEEK(m.mcod, 'lpu_m')
    IF INLIST(m.cod, 101927, 101928)
     m.recid = recid
     rval = InsError('S', 'POA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'POA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    RELEASE cod, sn_pol, d_u
   ENDIF 

   IF M.VDA == .T. AND m.tdat1>={01.05.2014} && Просроченный сертификат
    m.pcod = pcod
    m.prvs  = prvs
    m.d_ser = {}
    m.d_u   = d_u
    IF SEEK(m.pcod, 'doctor')
     m.d_ser  = doctor.d_ser
     m.d_ser2 = IIF(FIELD('d_ser2', 'doctor')=UPPER('d_ser2'), IIF(!EMPTY(doctor.d_ser2), doctor.d_ser2, {01.01.0001}), {01.01.0001})
    ENDIF 
    IF !EMPTY(m.d_ser)
     IF m.d_u-m.d_ser > 365.25*5 AND m.d_u-m.d_ser2 > 365.25*5
      m.recid = recid
      rval = InsError('S', 'VDA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.SKA = .T.
    IF USED('curskp')
     SET RELATION TO sn_pol INTO curskp ADDITIVE 
     SET RELATION TO cod INTO profus ADDITIVE 

     IF !EMPTY(Tip) OR SUBSTR(otd,2,2)='09' OR EMPTY(curskp.sn_pol) OR curskp.prv!=profus.profil
     ELSE 
      
      m.recid = recid
      rval    = InsError('S', 'SKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     
     ENDIF 

     SET RELATION OFF INTO curskp
     SET RELATION OFF INTO profus
    ENDIF 
   ENDIF 
   
   IF M.O0A = .T. && M.O0A - отключалка для всей онкологии!
    m.cod     = cod
    m.ds      = ds
    m.ds_2    = ds_2
    m.sn_pol  = sn_pol
    m.recid_s = recid_lpu
    m.recid_sl = ''
    m.recid_usl = ''
    m.o_otd   = SUBSTR(otd,2,2)
    m.usl_ok  = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')
    m.is_gsp  = IIF(m.usl_ok='1', .T., .F.)
	m.usl_tip = 0
    *m.reab    = reab
    
    m.IsOnkDs = IIF((BETWEEN(LEFT(m.ds,3), 'C00', 'C80') OR m.ds='C97') OR ;
    	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
    IF m.IsOnkDs AND m.is_gsp && IsGsp(m.cod)
     IF !USED('onk_sl')
      m.recid = recid
      rval    = InsError('S', 'O6A', m.recid)
     ELSE 
      IF !SEEK(m.recid_s, 'onk_sl')
       m.recid = recid
       rval    = InsError('S', 'O6A', m.recid)
      ELSE 
       m.recid_sl = onk_sl.recid
       IF !SEEK(onk_sl.ds1_t, 'onreas')
        m.recid = recid
        rval    = InsError('S', 'O6A', m.recid)
       ENDIF 

       IF !IsVMP(m.cod)
        IF INLIST(onk_sl.ds1_t,0,1,2,3,4)
         IF !SEEK(onk_sl.stad, 'onstad')
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ELSE && 5,6
         IF !EMPTY(onk_sl.stad)
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ENDIF 
       ELSE && IF IsVMP(m.cod)
        IF INLIST(onk_sl.ds1_t,0,1,2)
         IF !SEEK(onk_sl.stad, 'onstad')
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ELSE && 5,6
         IF !EMPTY(onk_sl.stad)
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 
       
       IF SEEK(onk_sl.stad, 'onstad')
        IF !EMPTY(onstad.ds)
         m.c_len = LEN(ALLTRIM(onstad.ds))
         IF LEFT(m.ds, m.c_len) != LEFT(onstad.ds, m.c_len)
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ELSE 
         IF SEEK(m.ds, 'onstad', 'ds')
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 

       IF onk_sl.ds1_t=0 AND (m.tdat1-people.dr)/365.25>=18 AND !SEEK(onk_sl.onk_t, 'ontum')
       	*MESSAGEBOX('OVA'+STR(m.recid,6),0+64,'1')
        m.recid = recid
        rval    = InsError('S', 'OVA', m.recid)
       ELSE 
        IF (onk_sl.ds1_t!=0 OR (m.tdat1-people.dr)/365.25<18) AND !EMPTY(onk_sl.onk_t)
       	*MESSAGEBOX('OVA'+STR(m.recid,6),0+64,IIF(onk_sl.ds1_t!=0,'.T.','.F.'))
       	*MESSAGEBOX('OVA'+STR(m.recid,6),0+64,IIF((m.tdat1-people.dr)/365.25<18,'.T.','.F.'))
         m.recid = recid
         rval    = InsError('S', 'OVA', m.recid)
        ENDIF 
       ENDIF 
       
       IF SEEK(onk_sl.onk_t, 'ontum')
        IF !EMPTY(ontum.ds)
         m.c_len = LEN(ALLTRIM(ontum.ds))
         IF LEFT(m.ds, m.c_len) != LEFT(ontum.ds, m.c_len)
       	*MESSAGEBOX('OVA'+STR(m.recid,6),0+64,'3')
          m.recid = recid
          rval    = InsError('S', 'OVA', m.recid)
         ENDIF 
        ELSE 
         IF SEEK(m.ds, 'ontum', 'ds')
       	*MESSAGEBOX('OVA'+STR(m.recid,6),0+64,'4')
          m.recid = recid
          rval    = InsError('S', 'OVA', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 
       
       IF onk_sl.ds1_t=0 AND (m.tdat1-people.dr)/365.25>=18 AND !SEEK(onk_sl.onk_n, 'onnod')
       	*MESSAGEBOX('OWA'+STR(m.recid,6),0+64,'1')
        m.recid = recid
        rval    = InsError('S', 'OWA', m.recid)
       ELSE 
        IF (onk_sl.ds1_t!=0 OR (m.tdat1-people.dr)/365.25<18) AND !EMPTY(onk_sl.onk_n)
        *IF !EMPTY(onk_sl.onk_n)
       	*MESSAGEBOX('OWA'+STR(m.recid,6),0+64,IIF(onk_sl.ds1_t!=0,'.T.','.F.'))
       	*MESSAGEBOX('OWA'+STR(m.recid,6),0+64,IIF((m.tdat1-people.dr)/365.25<18,'.T.','.F.'))
       	*MESSAGEBOX('OWA'+STR(m.recid,6),0+64,'2')
         m.recid = recid
         rval    = InsError('S', 'OWA', m.recid)
        ENDIF 
       ENDIF 
       
       IF SEEK(onk_sl.onk_n, 'onnod')
        IF !EMPTY(onnod.ds)
         m.c_len = LEN(ALLTRIM(onnod.ds))
         IF LEFT(m.ds, m.c_len) != LEFT(onnod.ds, m.c_len)
       	*MESSAGEBOX('OWA'+STR(m.recid,6),0+64,'3')
          m.recid = recid
          rval    = InsError('S', 'OWA', m.recid)
         ENDIF 
        ELSE 
         IF SEEK(m.ds, 'onnod', 'ds')
       	*MESSAGEBOX('OWA'+STR(m.recid,6),0+64,'4')
          m.recid = recid
          rval    = InsError('S', 'OWA', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 
       
       IF onk_sl.ds1_t=0 AND (m.tdat1-people.dr)/365.25>=18 AND !SEEK(onk_sl.onk_m, 'onmet')
        m.recid = recid
        rval    = InsError('S', 'OXA', m.recid)
       ELSE 
        *IF !EMPTY(onk_sl.onk_m)
        IF (onk_sl.ds1_t!=0 OR (m.tdat1-people.dr)/365.25<18) AND !EMPTY(onk_sl.onk_m)
         m.recid = recid
         rval    = InsError('S', 'OXA', m.recid)
        ENDIF 
       ENDIF 
       
       IF SEEK(onk_sl.onk_m, 'onmet')
        IF !EMPTY(onmet.ds)
         m.c_len = LEN(ALLTRIM(onmet.ds))
         IF LEFT(m.ds, m.c_len) != LEFT(onmet.ds, m.c_len)
          m.recid = recid
          rval    = InsError('S', 'OXA', m.recid)
         ENDIF 
        ELSE 
         IF SEEK(m.ds, 'onmet', 'ds')
          m.recid = recid
          rval    = InsError('S', 'OXA', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 

       IF INLIST(onk_sl.ds1_t,1,2)
        IF !INLIST(onk_sl.mtstz,0,1)
         m.recid = recid
         rval    = InsError('S', 'OJA', m.recid)
        ENDIF 
       ELSE 
        IF onk_sl.mtstz!=0
         m.recid = recid
         rval    = InsError('S', 'OJA', m.recid)
        ENDIF 
       ENDIF 

      ENDIF && IF !SEEK(m.recid_s, 'onk_sl')
     ENDIF && IF !USED('onk_sl')
     
     IF USED('onk_sl') AND USED('onk_diag') && проверка файла onk_diag
      IF !EMPTY(m.recid_sl) && m.recid_sl = onk_sl.recid
       IF SEEK(m.recid_sl, 'onk_diag')
        IF (!EMPTY(onk_diag.diag_date) OR !EMPTY(onk_diag.diag_code) OR !EMPTY(onk_diag.rec_rslt)) AND ;
        	!INLIST(onk_diag.diag_tip,1,2)
         m.recid = recid
         rval    = InsError('S', 'OLA', m.recid)
        ENDIF 
        IF (EMPTY(onk_diag.diag_date) AND EMPTY(onk_diag.diag_code) AND EMPTY(onk_diag.rec_rslt)) AND ;
        	onk_diag.diag_tip!=0
         m.recid = recid
         rval    = InsError('S', 'OLA', m.recid)
        ENDIF 
        
        IF onk_diag.diag_tip=1 AND !SEEK(onk_diag.diag_code, 'onmrf')
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=1 AND onk_diag.diag_code=0
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=1 AND !SEEK(onk_diag.diag_code, 'onmrds')
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=2 AND !SEEK(onk_diag.diag_code, 'onigh')
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=2 AND !SEEK(onk_diag.diag_code, 'onigds')
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=2 AND onk_diag.diag_code=0
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF !EMPTY(onk_diag.diag_rslt) AND onk_diag.diag_code=0
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=1 AND !SEEK(onk_diag.diag_rslt, 'onmrfr')
         m.recid = recid
         rval    = InsError('S', 'OEA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=2 AND !SEEK(onk_diag.diag_rslt, 'onigrt')
         m.recid = recid
         rval    = InsError('S', 'OEA', m.recid)
        ENDIF 
        
        IF onk_diag.rec_rslt=1 AND onk_diag.diag_code=0
         m.recid = recid
         rval    = InsError('S', 'OEA', m.recid)
        ENDIF 

        IF EMPTY(onk_diag.diag_date) AND ;
        	(!EMPTY(onk_diag.diag_code) OR !EMPTY(onk_diag.diag_tip) OR !EMPTY(onk_diag.rec_rslt))
         m.recid = recid
         rval    = InsError('S', 'OQA', m.recid)
        ENDIF 

        IF onk_diag.rec_rslt!=1 AND !EMPTY(onk_diag.diag_rslt)
         m.recid = recid
         rval    = InsError('S', 'OKA', m.recid)
        ENDIF 
        IF onk_diag.rec_rslt!=0 AND EMPTY(onk_diag.diag_rslt)
         m.recid = recid
         rval    = InsError('S', 'OKA', m.recid)
        ENDIF 
        
        
       ENDIF 
      ENDIF 
     ENDIF 
	 
	 IF IsGsp(m.cod) OR IsDst(m.cod)
     IF !USED('onk_usl')
      m.recid = recid
      rval    = InsError('S', 'O8A', m.recid)
     ELSE 
      IF EMPTY(m.recid_sl) OR (!EMPTY(m.recid_sl) AND !SEEK(m.recid_sl, 'onk_usl'))
       m.recid = recid
       rval    = InsError('S', 'O8A', m.recid)
      ELSE 
       m.recid_usl = onk_usl.recid
       IF !SEEK(onk_usl.usl_tip, 'onlech')
        m.recid = recid
        rval    = InsError('S', 'O8A', m.recid)
       ELSE 
	    m.usl_tip = onk_usl.usl_tip
       ENDIF 
	   
       IF onk_usl.usl_tip!=1 AND !EMPTY(onk_usl.hir_tip)
        m.recid = recid
        rval    = InsError('S', 'O9A', m.recid)
       ENDIF 
       IF onk_usl.usl_tip=1 AND EMPTY(onk_usl.hir_tip)
        m.recid = recid
        rval    = InsError('S', 'O9A', m.recid)
       ENDIF 
       IF !EMPTY(onk_usl.hir_tip) AND !SEEK(onk_usl.hir_tip, 'onhir')
        m.recid = recid
        rval    = InsError('S', 'O9A', m.recid)
       ENDIF 
       
       IF onk_usl.usl_tip!=2 AND !EMPTY(onk_usl.lek_tip_l)
        *MESSAGEBOX('OAA'+STR(m.recid,6),0+64,'1')
        m.recid = recid
        rval    = InsError('S', 'OAA', m.recid)
       ENDIF 
       IF onk_usl.usl_tip=2 AND EMPTY(onk_usl.lek_tip_l)
        *MESSAGEBOX('OAA'+STR(m.recid,6),0+64,'2')
        m.recid = recid
        rval    = InsError('S', 'OAA', m.recid)
       ENDIF 
       IF !EMPTY(onk_usl.lek_tip_l) AND !SEEK(onk_usl.lek_tip_l, 'onlekl')
        *MESSAGEBOX('OAA'+STR(m.recid,6),0+64,'3')
        m.recid = recid
        rval    = InsError('S', 'OAA', m.recid)
       ENDIF 

       IF onk_usl.usl_tip!=2 AND !EMPTY(onk_usl.lek_tip_v)
        m.recid = recid
        rval    = InsError('S', 'OBA', m.recid)
       ENDIF 
       IF onk_usl.usl_tip=2 AND EMPTY(onk_usl.lek_tip_v)
        m.recid = recid
        rval    = InsError('S', 'OBA', m.recid)
       ENDIF 
       IF !EMPTY(onk_usl.lek_tip_v) AND !SEEK(onk_usl.lek_tip_v, 'onlekv')
        m.recid = recid
        rval    = InsError('S', 'OBA', m.recid)
       ENDIF 

       IF !INLIST(onk_usl.usl_tip,3,4) AND !EMPTY(onk_usl.luch_tip)
        m.recid = recid
        rval    = InsError('S', 'OCA', m.recid)
       ENDIF 
       IF INLIST(onk_usl.usl_tip,3,4) AND EMPTY(onk_usl.luch_tip)
        m.recid = recid
        rval    = InsError('S', 'OCA', m.recid)
       ENDIF 
       IF !EMPTY(onk_usl.luch_tip) AND !SEEK(onk_usl.luch_tip, 'onluch')
        m.recid = recid
        rval    = InsError('S', 'OCA', m.recid)
       ENDIF 

      ENDIF && IF !SEEK(m.recid_s, 'onk_usl')
     ENDIF && IF !USED('onk_usl')

     IF INLIST(m.usl_tip,2,4)
     IF !USED('onk_ls')
      m.recid = recid
      rval    = InsError('S', 'OSA', m.recid)
     ELSE 
      IF EMPTY(m.recid_usl) OR (!EMPTY(m.recid_usl) AND !SEEK(m.recid_usl, 'onk_ls'))
       *MESSAGEBOX(m.recid_usl,0+64,'!')
       m.recid = recid
       rval    = InsError('S', 'OSA', m.recid)
      ELSE 
      ENDIF 

      ENDIF && IF !SEEK(m.recid_s, 'onk_usl')
     ENDIF && IF !USED('onk_usl')

     ENDIF && IsGsp or IsDst

    ENDIF && IF m.IsOnkDs
   ENDIF && IF M.O0A = .T.

  ENDIF && Глобальное отключение ошибок счета!

  ENDSCAN

  IF talon.IsPr == .F. && Глобальное отключение ошибок счета!
  IF M.SMA = .T.
  SELECT Gosp && Проверка по алгоритмам DU,SM,DI,DD.
  DO WHILE !EOF()
   Karta = c_i
   DO WHILE c_i = Karta
    DO CASE
     CASE !BETWEEN(d_u, dat1, dat2) And !EMPTY(Tip) && Дата выписки вне периода
       DO WHILE c_i = Karta
        m.recid = recid
        rval = InsError('S', 'DUA', m.recid)
*        InsErrorSV(m.mcod, 'S', 'DUA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        SKIP
       ENDDO 

     CASE BETWEEN(d_u, dat1, dat2) And !EMPTY(Tip) && Дата выписки в периоде
      DVip  = d_u
      DPost = d_u
      kMS   = 1
      DO WHILE c_i = Karta
       DO CASE
        CASE !EMPTY(Tip) And d_u = DPost And (!kMs=1 And !(kMS=2 And Cod=83010 And k_u=1)) AND (!kMs=1 And !(kMS=2 And Cod=83010 And k_u=1)) && Два МЭС на дату выписки
         m.recid = recid
         rval = InsError('S', 'SMA', m.recid)
*         InsErrorSV(m.mcod, 'S', 'SMA', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        
        CASE !EMPTY(Tip) And d_u < DVip && Разрыв
         m.recid = recid
         rval = InsError('S', 'DIA', m.recid)
*         InsErrorSV(m.mcod, 'S', 'DIA', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

*        CASE !EMPTY(Tip) And !INLIST(Cod,83010,83020,83030,83040,83050,183010,183020) And d_u > DVip && Пересечение
        CASE !EMPTY(Tip) And !INLIST(INT(Cod/1000),83,183) And d_u > DVip && Пересечение
         m.recid = recid
         rval = InsError('S', 'DDA', m.recid)
*         InsErrorSV(m.mcod, 'S', 'DDA', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

        OTHERWISE
         DVip  = DVip - k_u
         DPost = d_u

       ENDCASE
       SKIP
       kMS = kMS + 1
     ENDDO 
    OTHERWISE 
     SKIP 
    ENDCASE
   ENDDO 
  ENDDO  
  SELECT Talon
  ENDIF  && Отключение ошибки SMA
  ENDIF && Глобальное отключение ошибок счета!
  
  IF talon.IsPr == .F. && Глобальное отключение ошибок счета!
   IF M.DLA == .T.
    SET ORDER TO Unik
    GO TOP
    DO WHILE !EOF()
     m.c_i     = c_i
     m.add_vir = '0'
     IF OCCURS('#', m.c_i)>=3
      m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1), '0')
     ENDIF 
     Vir = m.add_vir+sn_pol+otd+ds+Padl(cod,6,'0')+DToC(d_u)
     SKIP
     jjj = .T.
     m.c_i     = c_i
     m.add_vir = '0'
     IF OCCURS('#', m.c_i)>=3
      m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1), '0')
     ENDIF 
     DO WHILE m.add_vir+sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u) = vir
      IF jjj = .T.
       jjj = .F.
       SKIP -1
       IF d_type != '2'
        m.recid = recid
        rval = InsError('S', 'DLA', m.recid)
*        InsErrorSV(m.mcod, 'S', 'DLA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
       SKIP 1
      ENDIF
      IF d_type != '2'
       m.recid = recid
       rval = InsError('S', 'DLA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'DLA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      SKIP 
      m.c_i     = c_i
      m.add_vir = '0'
      IF OCCURS('#', m.c_i)>=3
       m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1), '0')
      ENDIF 
     ENDDO  
    ENDDO 
   ENDIF 
  ENDIF 

  IF talon.IsPr == .F. && Глобальное отключение ошибок счета!
*   MESSAGEBOX('OK',0+64,'')
   IF M.PPA == .T.
    IF SEEK(m.lpuid, 'lpudogs')
     m.syear = 0 
     m.syear = lpudogs.kv01+lpudogs.kv02+lpudogs.kv03+lpudogs.kv04
     IF m.syear>0
      m.sfact = 0
      m.beginm = 1
      DO CASE 
       CASE BETWEEN(m.tmonth,1,3)
*        m.beginm = 1
        m.kvlimit = lpudogs.kv01
       CASE BETWEEN(m.tmonth,4,6)
*        m.beginm = 4
        m.kvlimit = lpudogs.kv01 + lpudogs.kv02
       CASE BETWEEN(m.tmonth,7,9)
*        m.beginm = 7
        m.kvlimit = lpudogs.kv01 + lpudogs.kv02 + lpudogs.kv03
       CASE BETWEEN(m.tmonth,10,12)
*        m.beginm = 10
        m.kvlimit = lpudogs.kv01 + lpudogs.kv02 + lpudogs.kv03 + lpudogs.kv04
      ENDCASE 
      FOR m.nmon = m.beginm TO m.tmonth-1
       m.lcperiod = STR(m.tyear,4)+PADL(m.nmon,2,'0')
       m.lppath = pbase+ '\'+m.lcperiod
       IF fso.FolderExists(m.lppath)
        IF fso.FileExists(m.lppath+'\aisoms.dbf')
         IF OpenFile(m.lppath+'\aisoms', 'lcais', 'shar', 'lpuid')=0
          IF SEEK(m.lpuid, 'lcais')
           m.sfact = m.sfact + (lcais.s_pred-lcais.sum_flk)
          ENDIF 
          USE IN lcais
         ELSE 
          IF USED('lcais')
           USE IN lcais
          ENDIF 
         ENDIF 
        ENDIF 
       ENDIF 
      NEXT 

*      m.sfact = m.sfact + (aisoms.s_pred-aisoms.sum_flk)
      
*      IF m.sfact > m.kvlimit
       oord = ORDER()
       SET ORDER TO d_u
       SCAN 
        m.sfact = m.sfact + s_all
        IF m.sfact<=m.kvlimit
         LOOP 
        ENDIF 
        m.recid = recid
        rval = InsError('S', 'PPA', m.recid)
       ENDSCAN
       SET ORDER TO &oord

*      ENDIF 
     
     ENDIF 
    ENDIF 
   ENDIF 
  ENDIF 
  
  
  CREATE CURSOR AllBad (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol 
  SELECT talon 
  SET ORDER TO sn_pol
  GO TOP 
  
  DO WHILE !EOF()
   m.polis = sn_pol
   m.lAllBad=.t.
   DO WHILE sn_pol = m.polis
    m.recid = recid
    IF !SEEK(m.RecId, 'sError')
     m.lAllBad = .f.
*     MESSAGEBOX(TRANSFORM(m.RecId,'999999'),0+64,sn_pol)
    ENDIF 
    SKIP 
   ENDDO 
   IF m.lAllBad
    IF !SEEK(m.polis, 'Allbad')
     INSERT INTO Allbad (sn_pol) VALUES (m.polis)
    ENDIF 
   ENDIF 
  ENDDO 
  
*  SELECT allbad
*  BROWSE 
  
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
  SUM(s_all) FOR SEEK(RecId, 'sError') TO m.s_flk
  SET RELATION OFF INTO people
  
  =ClBase()

  SELECT AisOms
  =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
  SELECT AisOms
  
  REPLACE sum_flk WITH m.s_flk

 WAIT CLEAR 
* MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!', 0+64, '')

RETURN 

FUNCTION InsError(WFile, cError, cRecId)
 IF WFile == 'R'
  IF 1=2
  IF !SEEK(cRecId, 'rError')
   INSERT INTO rError (f, c_err, rid) VALUES ('R', cError, cRecId)
  ELSE 
*   IF cError != rError.c_err
*    INSERT INTO rError (f, c_err, rid) VALUES ('R', cError, cRecId)
*   ENDIF cError != rError.c_err
  ENDIF !SEEK(cRecId, 'rError')
 ENDIF 
 ENDIF 
 IF WFile == 'S'
 IF cError='NRA'
  IF !SEEK(cError, 'sError')
   INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
   RETURN .T.
  ELSE 
   IF cError != sError.c_err
    INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
   ENDIF cError != sError.c_err 
  ENDIF !SEEK(cRecId, 'sError')
 ENDIF 
 ENDIF 
RETURN .F.

FUNCTION InsErrorSV(mmcod, WFile, cError, cRecId)
 IF WFile == 'R'
  IF !SEEK(mmcod+STR(cRecId,9), 'resv')
   INSERT INTO resv (mcod, f, c_err, rid) VALUES (mmcod, 'R', cError, cRecId)
  ELSE 
   IF cError != resv.c_err
    INSERT INTO resv (mcod, f, c_err, rid) VALUES (mmcod, 'R', cError, cRecId)
   ENDIF
  ENDIF
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(mcod+STR(cRecId,9), 'sesv')
   INSERT INTO sesv (mcod, f, c_err, rid) VALUES (mmcod, 'S', cError, cRecId)
   RETURN .T.
  ELSE 
   IF cError != sesv.c_err
    INSERT INTO sesv (mcod, f, c_err, rid) VALUES (mmcod, 'S', cError, cRecId)
   ENDIF
  ENDIF
 ENDIF 
RETURN .F.

FUNCTION OpBase(ppath)
 tnresult = 0
 tnresult = tnresult + OpenFile(ppath+'\people', 'people', 'share', 'sn_pol')
 tnresult = tnresult + OpenFile(ppath+'\talon', 'talon', 'share')
 tnresult = tnresult + OpenFile(ppath+'\doctor', 'doctor', 'share', 'pcod')
 IF fso.FileExists(ppath+'\ho'+m.qcod+'.dbf')
  tnresult = tnresult + OpenFile(ppath+'\ho'+m.qcod, 'ho', 'share', 'unik')
 ENDIF 
 IF fso.FileExists(ppath+'\ONK_SL'+m.qcod+'.dbf')
  tnresult = tnresult + OpenFile(ppath+'\ONK_SL'+m.qcod, 'onk_sl', 'share', 'recid_s')
 ENDIF 
 IF fso.FileExists(ppath+'\ONK_USL'+m.qcod+'.dbf')
  tnresult = tnresult + OpenFile(ppath+'\ONK_USL'+m.qcod, 'onk_usl', 'share', 'recid_s')
 ENDIF 
 IF fso.FileExists(ppath+'\ONK_LS'+m.qcod+'.dbf')
  tnresult = tnresult + OpenFile(ppath+'\ONK_LS'+m.qcod, 'onk_ls', 'share', 'recid_s')
 ENDIF 
 IF fso.FileExists(ppath+'\ONK_DIAG'+m.qcod+'.dbf')
  tnresult = tnresult + OpenFile(ppath+'\ONK_DIAG'+m.qcod, 'onk_diag', 'share', 'recid')
 ENDIF 

 tnresult = tnresult + OpenFile(ppath+'\e'+m.mcod, 'rerror', 'share', 'rrid')
 tnresult = tnresult + OpenFile(ppath+'\e'+m.mcod, 'serror', 'share', 'rid', 'again')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoree', 'osoree', 'share', 'd_type')
* tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\profus', 'profus', 'share', 'cod')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\mkb10', 'mkb10', 'share', 'ds')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'share', 'cod')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\CodWDr', 'CodWDr', 'share', 'cod')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\ososch', 'ososch', 'share', 'd_type')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\profot', 'profot', 'share', 'otd')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\isv012', 'isv012', 'share', 'ishod')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\rsv009', 'rsv009', 'share', 'rslt')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\'+IIF(m.tdat1>={01.05.2014}, 'spv015', 'kspec'), 'kspec', 'share', IIF(m.tdat1>={01.05.2014}, 'code', 'prvs'))
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\NoCodR', 'NoCodR', 'share')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\ms_mkb', 'MesMkb', 'share', 'ds_ms')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\CodOtd', 'CodOtd', 'share')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\CodKU', 'CodKU', 'share')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoerzxx', 'OsoERZ', 'Shar', 'ans_r')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sovmno', 'sovmno', 'Shar', 'ncod')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\reeskp', 'reeskp', 'Shar', 'unik')
 tnresult = tnresult + OpenFile(ppath+'\talon', 'talon_exp', 'share', 'ExpTag', 'again')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\polic_h', 'polic_h', 'share', 'sn_pol')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\polic_dp', 'polic_dp', 'share', 'sn_pol')
 tnresult = tnresult + OpenFile(pcommon+'\dsdisp', 'dsdisp', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\kpresl', 'kpresl', 'share', 'tip')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\mo_vmp', 'movmp', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spi_lpu_dd', 'spidd', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\noth', 'noth', 'share', 'cod')
 tnresult = tnresult + OpenFile(pcommon+'\lpudogs', 'lpudogs', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pcommon+'\dspcodes', 'dspcodes', 'share', 'cod')
 tnresult = tnresult + OpenFile(pcommon+'\lpuskp', 'lpuskp', 'share', 'lpuid')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\exclhors', 'exclhors', 'share', 'lpu_id')
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onmet_xx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onmet_xx', 'onmet', 'share', 'cod_m')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onnod_xx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onnod_xx', 'onnod', 'share', 'cod_n')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onreasxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onreasxx', 'onreas', 'share', 'cod_reas')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onstadxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onstadxx', 'onstad', 'share', 'cod_st')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\ontum_xx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\ontum_xx', 'ontum', 'share', 'cod_t')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onlechxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onlechxx', 'onlech', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onhir_xx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onhir_xx', 'onhir', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onleklxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onleklxx', 'onlekl', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onlekvxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onlekvxx', 'onlekv', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onluchxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onluchxx', 'onluch', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onmrf_xx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onmrf_xx', 'onmrf', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onmrdsxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onmrdsxx', 'onmrds', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onigh_xx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onigh_xx', 'onigh', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onigdsxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onigdsxx', 'onigds', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onmrfrxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onmrfrxx', 'onmrfr', 'share', 'id_r_m')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onigrtxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onigrtxx', 'onigrt', 'share', 'id_r_i')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onconsxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onconsxx', 'oncons', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onpcelxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onpcelxx', 'onpcel', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onnaprxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onnaprxx', 'onnapr', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onczabxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onczabxx', 'onczab', 'share', 'cod')
 ENDIF 
 
 IF 1=2
 DELETE FOR SUBSTR(c_err,3,1)='A' IN rerror
 DELETE FOR SUBSTR(c_err,3,1)='A' IN serror
 DELETE FOR SUBSTR(c_err,3,1)='B' IN rerror
 DELETE FOR SUBSTR(c_err,3,1)='B' IN serror
 ENDIF 
RETURN .t.

FUNCTION ClBase()
 IF USED('talon')
  USE IN talon 
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('doctor')
  USE IN doctor
 ENDIF 
 IF USED('rerror')
  USE IN rerror
 ENDIF 
 IF USED('serror')
  USE IN serror
 ENDIF 
 IF USED('osoree')
  USE IN osoree
 ENDIF 
 IF USED('mkb10')
  USE IN mkb10
 ENDIF 
 IF USED('tarif')
  USE IN tarif
 ENDIF 
 IF USED('codwdr')
  USE IN CodWDr
 ENDIF 
 IF USED('ososch')
  USE IN OsoSch
 ENDIF 
 IF USED('profot')
  USE IN ProfOt
 ENDIF 
 IF USED('isv012')
  USE IN isv012
 ENDIF 
 IF USED('rsv009')
  USE IN rsv009
 ENDIF 
 IF USED('kspec')
  USE IN kspec
 ENDIF 
 IF USED('nocodr')
  USE IN NoCodR
 ENDIF 
 IF USED('mesmkb')
  USE IN MesMkb
 ENDIF 
 IF USED('codotd')
  USE IN CodOtd
 ENDIF 
 IF USED('codku')
  USE IN CodKU
 ENDIF 
 IF USED('day_gr')
  USE IN Day_gr
 ENDIF 
 IF USED('e_day')
  USE IN e_day 
 ENDIF 
 IF USED('month_gr')
  USE IN Month_Gr
 ENDIF 
 IF USED('e_month')
  USE IN e_month
 ENDIF 
 IF USED('osoerz')
  USE IN OsoERZ
 ENDIF 
 IF USED('sovmno')
  USE IN sovmno
 ENDIF 
 IF USED('talon_exp')
  USE IN talon_exp
 ENDIF 
 IF USED('gosp')
  USE IN Gosp
 ENDIF 
 IF USED('polic_h')
  USE IN polic_h
 ENDIF 
 IF USED('polic_dp')
  USE IN polic_dp
 ENDIF 
 IF USED('dsdisp')
  USE IN dsdisp
 ENDIF 
 IF USED('kpresl')
  USE IN kpresl
 ENDIF 
 IF USED('movmp')
  USE IN movmp
 ENDIF 
 IF USED('spidd')
  USE IN spidd
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 IF USED('dspp')
  USE IN dspp
 ENDIF 
 IF USED('dspyear')
  USE IN dspyear
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 
 IF USED('noth')
  USE IN noth
 ENDIF 
 IF USED('ho')
  USE IN ho
 ENDIF 
 IF USED('lpuskp')
  USE IN lpuskp
 ENDIF 
 IF USED('reeskp')
  USE IN reeskp
 ENDIF 
* IF USED('profus')
*  USE IN profus
* ENDIF 
 IF USED('curskp')
  USE IN curskp
 ENDIF 
 IF USED('exclhors')
  USE IN exclhors
 ENDIF 
 IF USED('onk_sl')
  USE IN onk_sl
 ENDIF 
 IF USED('onk_usl')
  USE IN onk_usl
 ENDIF 
 IF USED('onk_ls')
  USE IN onk_ls
 ENDIF 
 IF USED('onk_diag')
  USE IN onk_diag
 ENDIF 
 IF USED('onmet')
  USE IN onmet
 ENDIF 
 IF USED('onnod')
  USE IN onnod
 ENDIF 
 IF USED('onreas')
  USE IN onreas
 ENDIF 
 IF USED('onstad')
  USE IN onstad
 ENDIF 
 IF USED('ontum')
  USE IN ontum
 ENDIF 
 IF USED('onlech')
  USE IN onlech
 ENDIF 
 IF USED('onhir')
  USE IN onhir
 ENDIF 
 IF USED('onlekl')
  USE IN onlekl
 ENDIF 
 IF USED('onlekv')
  USE IN onlekv
 ENDIF 
 IF USED('onluch')
  USE IN onluch
 ENDIF 
 IF USED('onmrf')
  USE IN onmrf
 ENDIF 
 IF USED('onmrds')
  USE IN onmrds
 ENDIF 
 IF USED('onigh')
  USE IN onigh
 ENDIF 
 IF USED('onigds')
  USE IN onigds
 ENDIF 
 IF USED('onmrfr')
  USE IN onmrfr
 ENDIF 
 IF USED('onigrt')
  USE IN onigrt
 ENDIF 
 IF USED('oncons')
  USE IN oncons
 ENDIF 
 IF USED('onpcel')
  USE IN onpcel
 ENDIF 
 IF USED('onnapr')
  USE IN onnapr
 ENDIF 
 IF USED('onczab')
  USE IN onczab
 ENDIF 

RETURN 