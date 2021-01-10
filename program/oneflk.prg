FUNCTION OneFlk(ppath)

 m.cfrom = ALLTRIM(cfrom)
 m.t_a = SECONDS()
 m.t_0   = SECONDS()
 m.t_t   = m.t_0
 
 m.IsPr = IsPr

 *IF IsPr 
 * =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
 * RETURN 
 *ENDIF 

 IF m.istestmode
  CREATE CURSOR c_test (et c(10), timing n(10))
  INSERT INTO c_test VALUES ('Start: ', 0)
  SELECT aisoms 
 ENDIF 

 PUBLIC M.PSA, M.ERA, M.ECA, M.E1A, M.E2A, M.E4A, M.E5A, M.E6A, M.E7A, M.E8A, M.H6A, M.COA, M.HCA, M.DUA, M.H8A, M.HEA, M.CSA, M.TVA, M.NLA,;
 	M.MDA, M.H3A, M.SOA, M.R1A, M.R2A, M.R3A, M.UVA, M.DVA, M.UOA, M.NOA, M.NMA, M.NUA, M.NSA, M.SMA, M.DIA, M.DDA, M.HNA, M.DLA,;
    M.DRA, M.POA, M.VDA, M.TFA, M.PPA, M.G1A, M.G2A, M.G3A, M.G4A, M.NRA, M.KEA, M.D2A, M.THA, M.TLA, M.HOA, M.SKA, M.FSA, M.PFA,;
    M.W2A, M.O0A, M.O1A, M.O2A, M.O3A, M.O4A, M.O5A, M.O6A, M.O7A, M.O8A, M.OAA, M.OBA, M.OCA, M.ODA, M.OEA, M.OFA, M.OGA,;
    M.OHA, M.OIA, M.OJA, M.OKA, M.OLA, M.OMA, M.ONA, M.OOA, M.OPA, M.OQA, M.ORA, M.OSA, M.OTA, M.OUA, M.OVA, M.OWA, M.OXA, M.OYA,;
    M.OZA, M.ENA, M.IPA, M.EGA, M.X1A, M.X2A, M.X3A, M.X4A, M.X5A, M.X6A, M.X7A, M.X8A, M.X9A, M.PLA, M.H7A, M.DKA, M.UMA, M.D4A, ;
    M.S1A, M.MPA, M.MZA, M.NDA, M.D1A, M.PGA, M.NWA, M.HRA, M.D6A, M.DNA, M.NVA, M.TPA, M.R4A, M.PHA, M.WEA, M.MMA, M.CVA, M.CPA
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
 
 =ReadErrorStatus(WorkIniFile) && Чтение включения/выключения ошибок
 
 M.O0A = IIF(!EMPTY(cfrom), .F., M.O0A) && Заглушка для cfrom=oms@spuemias.msk.oms
 M.O0A = IIF(m.tdat1<{01.01.2019}, .F., M.O0A) && Заглушка для cfrom=oms@spuemias.msk.oms
 M.ENA = IIF(!EMPTY(cfrom), .F., M.ENA) && Заглушка для cfrom=oms@spuemias.msk.oms
 
 M.PPA = IIF(m.tdat1<{01.01.2020}, .F., M.PPA) 


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
  
 m.IsStPilot  = IIF(SEEK(m.lpuid, 'stpilot'), .T., .F.) && Кончаловского

 m.IsSprNCO   = IIF(SEEK(m.lpuid, 'sprnco'), .T., .F.)
 m.IsIG       = IIF(SEEK(m.lpuid, 'sprnco') AND sprnco.ig=1, .T., .F.)
 *m.d_b        = IIF(SEEK(m.lpuid, 'sprnco'), sprnco.date_b, {}) && {23.03.2020}

 m.d_b        = IIF(SEEK(m.lpuid, 'sprnco'), ;
    IIF(FIELD('DATEBEG_3','sprnco')='DATEBEG_3', sprnco.datebeg_3, sprnco.datebeg), ;
 	{}) && {23.03.2020}
 m.d_e        = IIF(SEEK(m.lpuid, 'sprnco'), ;
    IIF(FIELD('DATEEND_3','sprnco')='DATEEND_3', sprnco.dateend_3, sprnco.dateend), ;
 	{}) && {23.03.2020}

 m.d_b2        = IIF(SEEK(m.lpuid, 'sprnco'), ;
    IIF(FIELD('DATEBEG_4','sprnco')='DATEBEG_4', sprnco.datebeg_4, sprnco.datebeg), ;
 	{}) && {23.03.2020}
 m.d_e2        = IIF(SEEK(m.lpuid, 'sprnco'), ;
    IIF(FIELD('DATEEND_4','sprnco')='DATEEND_4', sprnco.dateend_4, sprnco.dateend), ;
 	{}) && {23.03.2020}

 M.WEA = IIF(m.lpuid=5139, M.WEA, .F.) 
 M.PPA = IIF(INLIST(m.lpuid,1989,4963,4708), .F., M.PPA) 
 *M.UMA = IIF(m.IsStac, M.UMA, .F.)
  
 m.lIsDspExists = .f.
 m.dspfile1 = pbase +'\'+ STR(tyear-1,4)+'12'+'\dsp'
 m.dspfile2 = pbase +'\'+ STR(tyear-2,4)+'12'+'\dsp'
 m.dspfile3 = pbase +'\'+ STR(tyear-3,4)+'12'+'\dsp'
 ** Ищем файл за предыдущий период!
 IF tmonth>1
  m.dspfile = pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\dsp'
 ELSE
  m.dspfile = pbase +'\'+ STR(tyear-1,4)+'12'+'\dsp'
 ENDIF 
 ** Ищем файл за предыдущий период!

 IF fso.FileExists(m.dspfile+'.dbf')
  m.lIsDspExists = .t.
 ELSE 
  m.lIsDspExists = .f.
 ENDIF 
 
 IF m.qcod<>'R2'
  IF m.lIsDspExists
   oal = ALIAS()
   IF OpenFile(m.dspfile, 'cdsp', 'shar')>0 && mcod+sn_pol+padl(cod,6,'0')
    IF USED('cdsp')
     USE IN cdsp
    ENDIF 
    *SELECT (oal)
    m.lIsDspExists = .f.
   ELSE 
    SELECT * FROM cdsp INTO CURSOR dspp READWRITE 
    USE IN cdsp
    SELECT dspp
    *INDEX on mcod+sn_pol+PADL(tip,1,'0') TAG exptag
    *INDEX on mcod+sn_pol+PADL(cod,6,'0') TAG un_tag
    INDEX on sn_pol+PADL(tip,1,'0') TAG exptag
    INDEX on sn_pol+PADL(cod,6,'0') TAG un_tag
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
    *SELECT (oal)
    m.lIsDspExists = .t.
   ENDIF  
  ENDIF 
 ELSE && Если ВТБ то просто открываем файл!
  IF OpenFile(m.dspfile, 'dspp', 'shar', 'exptag')>0
   IF USED('ddsp')
    USE IN ddsp
   ENDIF 
   *SELECT (oal)
   m.lIsDspExists = .f.
  ELSE 
   && все ок!
  ENDIF 
 ENDIF 

 IF m.istestmode
  INSERT INTO c_test VALUES ('dsp ', SECONDS()-m.t_t)
  m.t_t = SECONDS()
 ENDIF 

 M.D2A = IIF(m.lIsDspExists = .t., M.D2A, .f.)
  
 lcError = ppath+'\e'+m.mcod
 IF !fso.FileExists(lcError+'.dbf')
  CREATE TABLE (lcError) (f c(1), c_err c(3), rid i)
  INDEX FOR UPPER(f)='R' ON rid TAG rrid
  INDEX FOR UPPER(f)='S' ON rid TAG rid
  USE 
 ENDIF 
  
 IF !OpBase(ppath)
  RELEASE M.PSA, M.ERA, M.ECA, M.E1A, M.E2A, M.E4A, M.E5A, M.E6A, M.E7A, M.E8A, M.H6A, M.COA, M.HCA, M.DUA, M.H8A, M.HEA, M.CSA, M.TVA, M.NLA,;
  	M.MDA, M.H3A, M.SOA, M.R1A, M.R2A, M.R3A, M.UVA, M.DVA, M.UOA, M.NOA, M.NMA, M.NUA, M.NSA, M.SMA, M.DIA, M.DDA, M.HNA, M.DLA,;
  	M.DRA, M.POA, M.VDA, M.TFA, M.PPA, M.G1A, M.G2A, M.G3A, M.G4A, M.NRA, M.KEA, M.D2A, M.THA, M.TLA, M.HOA, M.SKA, M.FSA, M.PFA,;
 	M.W2A, M.O0A, M.O1A, M.O2A, M.O3A, M.O4A, M.O5A, M.O6A, M.O7A, M.O8A, M.OAA, M.OBA, M.OCA, M.ODA, M.OEA, M.OFA, M.OGA,;
  	M.OHA, M.OIA, M.OJA, M.OKA, M.OLA, M.OMA, M.ONA, M.OOA, M.OPA, M.OQA, M.ORA, M.OSA, M.OTA, M.OUA, M.OVA, M.OWA, M.OXA, M.OYA,;
  	M.OZA, M.ENA, M.IPA, M.EGA, M.X1A, M.X2A, M.X3A, M.X4A, M.X5A, M.X6A, M.X7A, M.X8A, M.X9A, M.PLA, M.H7A, M.DKA, M.UMA, M.D4A,;
  	M.S1A, M.MPA, M.MZA, M.NDA, M.D1A, M.PGA, M.NWA, M.HRA, M.D6A, M.DNA, M.NVA, M.TPA, M.R4A, M.PHA, M.WEA, M.MMA, M.CVA, M.CPA
  =ClBase()
  RETURN .f.
 ENDIF 
 
 SELECT aisoms 
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
   * Здесь перечислить все реестровые ошибки!
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

 IF m.istestmode
  INSERT INTO c_test VALUES ('OpBase', SECONDS()-m.t_t)
  m.t_t = SECONDS()
 ENDIF 

 m.IsExHorS = IIF(SEEK(m.lpuid, 'exclhors'), .t., .f.)
  
 CREATE CURSOR c_pp (c_i c(30))
 SELECT c_pp
 INDEX on c_i TAG c_i 
 SET ORDER TO c_i

 SELECT c_i, sn_pol, cod, otd, d_u, k_u, ds, SPACE(3) as prv WHERE SUBSTR(otd,2,2)='09' FROM talon INTO CURSOR curskp READWRITE 
 SELECT curskp
 IF RECCOUNT('curskp')>0
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  SET RELATION TO cod INTO profus
  REPLACE ALL prv WITH profus.profil
  SET RELATION OFF INTO profus
 ELSE 
  USE IN curskp
 ENDIF 
 
 * Select для алгоритма NO 
 *SELECT sn_pol, MIN(VAL(SUBSTR(otd,2,2))) as otd, cod, MIN(IIF(INLIST(d_type,'2','8'),1,0)) as d_type, ;
 	SUM(k_u) AS k_u, SUM(kd_fact) AS kd_fact, d_u, SUM(s_all) AS s_all ;
 	FROM Talon GROUP BY sn_pol, d_u, cod;
 	INTO CURSOR day_gr
 SELECT sn_pol, MIN(VAL(SUBSTR(otd,2,2))) as otd, cod, MIN(IIF(INLIST(d_type,'8'),1,0)) as d_type, ;
 	SUM(k_u) AS k_u, SUM(kd_fact) AS kd_fact, d_u, SUM(s_all) AS s_all ;
 	FROM Talon GROUP BY sn_pol, d_u, cod;
 	INTO CURSOR day_gr
  
 SELECT sn_pol AS sn_pol, a.cod AS cod, ;
  IIF(!INLIST(otd,80,81), k_u, kd_fact) AS k_u, 0000 as cntr, d_u AS d_u, s_all AS s_all, IIF(INLIST(otd,0,1,8,22,85,90,91,92,93), mdayp, mdays) AS in_day,;
  IIF(INLIST(otd,0,1,8,22,85,90,91,92,93), mdayp, mdays) as k_norm, d_type;
  FROM day_gr a, codku b ;
  WHERE a.cod=b.cod AND a.d_type=0 AND IIF(!INLIST(otd,80,81), k_u, kd_fact) > IIF(INLIST(otd,0,1,8,22,85,90,91,92,93), mdayp, mdays) ;
  INTO CURSOR e_day ORDER BY a.sn_pol, a.d_u, a.cod READWRITE 

 SELECT e_day
 INDEX ON sn_pol + STR(cod,6) + DTOS(d_u) TAG ExpTag
 SET ORDER TO ExpTag
 * Select для алгоритма NO 
   
 * Select для алгоритма NM
 *SELECT sn_pol, cod,  ;
  MIN(VAL(SUBSTR(otd,2,2))) as otd, SUM(k_u) AS k_u, SUM(kd_fact) AS kd_fact, MIN(d_u) AS d_u, SUM(s_all) AS s_all,;
  MIN(IIF(INLIST(d_type,'2','8'),1,0)) as d_type ;
  FROM Talon  GROUP BY sn_pol, cod ;
  INTO CURSOR month_gr
 SELECT sn_pol, cod,  ;
  MIN(VAL(SUBSTR(otd,2,2))) as otd, SUM(k_u) AS k_u, SUM(kd_fact) AS kd_fact, MIN(d_u) AS d_u, SUM(s_all) AS s_all,;
  MIN(IIF(INLIST(d_type,'8'),1,0)) as d_type ;
  FROM Talon  GROUP BY sn_pol, cod ;
  INTO CURSOR month_gr

 SELECT sn_pol as sn_pol, a.cod as cod, IIF(!INLIST(otd,80,81), k_u, kd_fact) as k_u, 0000 as cntr, ;
  IIF(INLIST(otd,0,1,8,22,85,90,91,92,93), mmsp, mmss) as k_norm, s_all as s_all, ;
  IIF(INLIST(otd,0,1,8,22,85,90,91,92,93), b.mmsp, b.mmss) as in_month ;
  FROM month_gr a, codku b ;
  WHERE a.cod=b.cod AND a.d_type=0 AND IIF(!INLIST(otd,80,81), k_u, kd_fact) > IIF(INLIST(otd,0,1,8,22,85,90,91,92,93), mmsp, mmss) ;
  INTO CURSOR e_month ORDER BY sn_pol, a.cod READWRITE 

 SELECT e_month
 INDEX ON sn_pol + STR(cod,6) TAG ExpTag
 SET ORDER TO ExpTag
 * Select для алгоритма NM
 
 IF IsStac(m.mcod)
  SELECT c_i DISTINCT  FROM Talon WHERE IsMes(Cod) OR IsVmp(Cod) OR INLIST(cod,56029,156003) INTO CURSOR t_tst READWRITE 
  SELECT t_tst
  INDEX on c_i TAG c_i 
  SET ORDER TO c_i 
  
  SELECT *, .f. as IsMsExt FROM Talon WHERE c_i IN (SELECT c_i FROM t_tst);
 	ORDER BY c_i, d_u DESC INTO CURSOR Gosp READWRITE 
  SELECT *, d_u-k_u as d_pos FROM Gosp WHERE IsMes(Cod) OR IsVmp(Cod) ORDER BY c_i, d_pos ASC ;
  	INTO CURSOR Gosp_d
  SELECT Gosp_d
  INDEX on c_i TAG c_i 
  SET ORDER TO c_i
 	
  USE IN t_tst

  *SELECT *, .f. as IsMsExt FROM Talon WHERE c_i IN (SELECT c_i FROM Talon WHERE IsMes(Cod) OR IsVmp(Cod) OR INLIST(cod,56029,156003));
 	ORDER BY c_i, d_u DESC INTO CURSOR Gosp READWRITE 
  *SELECT *, .f. as IsMsExt FROM Talon WHERE IsMes(Cod) OR IsVmp(Cod) OR INLIST(cod,56029,156003) ;
 	ORDER BY c_i, d_u DESC INTO CURSOR Gosp READWRITE 
 ELSE 
  SELECT *, .f. as IsMsExt FROM Talon WHERE c_i="*#@";
 	ORDER BY c_i, d_u DESC INTO CURSOR Gosp READWRITE 
 ENDIF 
 
 SELECT Gosp
 REPLACE FOR SEEK(cod, 'msext') OR INLIST(INT(cod/1000),83,183) OR INLIST(cod,56029,156003) IsMsExt WITH .T.
 INDEX on c_i TAG c_i FOR IsMsExt
 INDEX on c_i TAG karta
 INDEX ON sn_pol TAG sn_pol FOR IsMes(cod) OR IsVMP(cod)
 
 *COPY TO &pBase\&gcPeriod\&mcod\Gosp
 
 * Select для части алгоритма H6 - поиск дублей по коду карты 
 SELECT a.recid as recid FROM talon a JOIN talon b ;
	ON a.c_i=b.c_i AND a.sn_pol<>b.sn_pol WHERE (isgsp(a.cod) OR isdst(a.cod)) AND (isgsp(a.cod) OR isdst(a.cod));
	INTO CURSOR curs_h6
 INDEX on recid TAG recid
 SET ORDER TO recid
 IF _tally=0
  USE IN curs_h6
 ENDIF 

 SELECT a.recid DISTINCT FROM talon a JOIN talon b ;
	ON a.c_i=b.c_i AND a.sn_pol<>b.sn_pol WHERE !INLIST(SUBSTR(a.otd,2,2),'00','01','08','85','90','91','92') AND ;
	!INLIST(SUBSTR(b.otd,2,2),'00','01','08','85','90','91','92') INTO CURSOR curs_h6p
 INDEX on recid TAG recid
 SET ORDER TO recid
 IF _tally=0
  USE IN curs_h6p
 ENDIF 
 * Select для части алгоритма H6 - поиск дублей по коду карты 

 IF m.istestmode
  INSERT INTO c_test VALUES ('Selects', SECONDS()-m.t_t)
  m.t_t = SECONDS()
 ENDIF 
 *COPY TO &pbase\&gcperiod\Gosp
  
  m.s_flk = 0  
  m.ls_flk = 0

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
  
 IF m.istestmode
  INSERT INTO c_test VALUES ('DRA', SECONDS()-m.t_t)
  m.t_t = SECONDS()
 ENDIF 

 IF M.PPA && сброс начальных значений
  IF USED('nsif')

   IF m.tmonth=1
    UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   		n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
   ELSE 
    m.lcperiod = STR(m.tyear,4)+PADL(m.tmonth-1,2,'0')
    m.lppath = pbase+ '\'+m.lcperiod
    IF fso.FolderExists(m.lppath+'\nsi')
     IF fso.FileExists(m.lppath+'\nsi\nsif.dbf')
      IF OpenFile(m.lppath+'\nsi\nsif', 'p_nsif', 'shar', 'lpu_id')=0
       SELECT nsif 
       SET RELATION TO lpu_id INTO p_nsif
       
       REPLACE n_ks WITH p_nsif.n_ks, ks_fact WITH p_nsif.ks_fact, n_ds WITH p_nsif.n_ds, ds_fact WITH p_nsif.ds_fact, ;
       	app_fact WITH p_nsif.app_fact, n_kt WITH p_nsif.n_kt, ptkt_fact WITH p_nsif.ptkt_fact, n_gem WITH p_nsif.n_gem, ;
       	gem_fact WITH p_nsif.gem_fact, n_eco WITH p_nsif.n_eco, eco_fact WITH p_nsif.eco_fact FOR lpu_id = m.lpuid 
       
       SET RELATION OFF INTO p_nsif
       USE IN p_nsif
      ELSE 
       IF USED('p_nsif')
        USE IN p_nsif
       ENDIF 
       UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   		n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
      ENDIF 
     ELSE 
      UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   		n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
     ENDIF 
    ELSE 
     UPDATE nsif SET ks_fact=0, ds_fact=0, app_fact=0, ptkt_fact=0, gem_fact=0, eco_fact=0 ;
   		WHERE lpu_id=m.lpuid
    ENDIF 
   ENDIF 
  ENDIF 
 ENDIF 
 
 IF m.istestmode
  INSERT INTO c_test VALUES ('PPA (step 1)', SECONDS()-m.t_t)
  m.t_t = SECONDS()
 ENDIF 

 IF people.IsPr==.F. && Глобальное отключение ошибок регистра!
  m.t_1 = SECONDS()
  SELECT people
  SCAN 
   DO r_flkn IN r_flkn
  ENDSCAN 
  m.t_2 = SECONDS()
  m.t_r_flk = (m.t_2-m.t_1)
 ENDIF 

 IF m.istestmode
  INSERT INTO c_test VALUES ('r_flkn', SECONDS()-m.t_t)
  m.t_t = SECONDS()
 ENDIF 

 CREATE CURSOR n_ds (c_i c(30))
 SELECT n_ds
 INDEX on c_i TAG c_i 
 SET ORDER TO c_i
  
 CREATE CURSOR n_ks (c_i c(30))
 SELECT n_ks
 INDEX on c_i TAG c_i 
 SET ORDER TO c_i

 CREATE CURSOR n_eco (c_i c(30))
 SELECT n_eco
 INDEX on c_i TAG c_i 
 SET ORDER TO c_i

 CREATE CURSOR n_gem (c_i c(30))
 SELECT n_gem
 INDEX on c_i TAG c_i 
 SET ORDER TO c_i

 SELECT c_talon
 SET RELATION TO sn_pol INTO people 
 
 m.t_s_flk = 0
 SCAN
  
  m.IsOtdSkp = IIF(SUBSTR(otd,2,2)='09', .T., .F.)

  IF EMPTY(people.sn_pol)               && Алгоритм PS
   m.polis = sn_pol
   DO WHILE sn_pol == m.polis
    m.recid = recid
    rval = InsError('S', 'PSA', m.recid)
    m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    SKIP +1 
   ENDDO 
  ENDIF 
  
  m.r_id = people.recid
  m.recid = recid
  IF SEEK(m.r_id, 'rerror')
   rval = InsError('S', 'PKA', m.recid, '',;
   	'Запись счета забракована по регистровой ошибке '+rerror.c_err)
   m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
  ENDIF 
  m.t_1 = SECONDS()
  IF c_talon.IsPr == .F. && Глобальное отключение ошибок счета!
   DO ss_flk IN ss_flk 
  ENDIF && Глобальное отключение ошибок счета!
  m.t_2 = SECONDS()
  m.t_s_flk = m.t_s_flk + (m.t_2-m.t_1)

 ENDSCAN
  
 IF USED('n_ds')
  USE IN n_ds
 ENDIF 
 IF USED('n_ks')
  *SELECT n_ks
  *COPY TO &pBase\&gcPeriod\n_ks
  USE IN n_ks
 ENDIF 
 IF USED('n_eco')
  USE IN n_eco
 ENDIF 
 IF USED('n_gem')
  USE IN n_gem
 ENDIF 

 IF m.istestmode
  INSERT INTO c_test VALUES ('s_flk:', m.t_s_flk)
  m.t_t = SECONDS()
 ENDIF 
 IF c_talon.IsPr == .F. && Глобальное отключение ошибок счета!

  IF M.SMA = .T.
  SELECT Gosp && Проверка по алгоритмам DU,SM,DI,DD.
  SET ORDER TO 
  *COPY TO &pBase\&gcPeriod\gGosp
  GO TOP 
  *BROWSE 
  DO WHILE !EOF()
   Karta = c_i
   m.recid = recid
   m.cod = cod
   DO WHILE c_i = Karta

    *DO CASE
     *CASE !BETWEEN(d_u, dat1, dat2) AND !EMPTY(Tip) && Дата выписки вне периода
     *  DO WHILE c_i = Karta
     *   m.recid = recid
     *   rval = InsError('S', 'DUA', m.recid, '',;
     *   	'Дата выписка стационарного пациента не в периоде ('+DTOC(d_u)+')')
     *   m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     *   SKIP
     *  ENDDO 

     *CASE BETWEEN(d_u, dat1, dat2) AND !EMPTY(Tip)
     IF !EMPTY(Tip)

      IF Tip='7' && AND !INLIST(INT(m.cod/1000),83,183)

       m.Is56029 = .F.
       m.IsOneLast = .T.
       SKIP -1 IN Gosp
       IF c_i=Karta AND INLIST(cod,56029,156003)
        m.Is56029 = .T.
       ENDIF 
       IF c_i=Karta
        m.IsOneLast = .F.
       ENDIF 
       SKIP +1 IN Gosp
       IF m.Is56029 = .F.
        SKIP +1 IN Gosp
        IF c_i=Karta AND INLIST(cod,56029,156003)
         m.Is56029 = .T.
        ENDIF 
        IF c_i=Karta
         m.IsOneLast = .F.
        ENDIF 
        SKIP -1 IN Gosp
       ENDIF 
       
       *MESSAGEBOX(sn_pol,0+64,STR(cod,6))
       
        *DO WHILE c_i = Karta
         IF !EMPTY(Tip)
          IF !m.Is56029 AND m.IsOneLast
           m.recid = recid
           rval = InsError('S', 'TFA', m.recid, '',;
        	'Tip=7 при отсутствии последующих МЭСов')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 
         ENDIF 
         SKIP
        *ENDDO 

      ELSE 
       DVip  = d_u
       DPost = d_u
       kMS   = 1
       DO WHILE c_i = Karta
        DO CASE
         CASE !EMPTY(Tip) And d_u = DPost And ;
        	(!kMs=1 And !(kMS=2 And Cod=83010 And k_u=1)) AND (!kMs=1 And !(kMS=2 And Cod=83010 And k_u=1)) && Два МЭС на дату выписки
          IF INLIST(Tip, '0', 'A', 'T', 'R')
           *m.recid = recid
           *rval = InsError('S', 'SMA', m.recid)
           *m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 
        
         CASE !EMPTY(Tip) And d_u < DVip AND Tip!='7'&& Разрыв
          *m.recid = recid
          *rval = InsError('S', 'DIA', m.recid)
          *m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

         CASE !EMPTY(Tip) And !INLIST(INT(Cod/1000),83,183) And d_u > DVip AND Tip!='7' && Пересечение
          *m.recid = recid
          *rval = InsError('S', 'DDA', m.recid)
          *m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

         OTHERWISE
          IF !EMPTY(Tip)
           DVip  = DVip - k_u
           DPost = d_u
          ENDIF 

        ENDCASE
        SKIP
        kMS = kMS + 1
       ENDDO 
       *Karta = c_i
      ENDIF 

     *OTHERWISE
     ELSE 
      *MESSAGEBOX(sn_pol,0+64,STR(cod,6)+" "+STR(recid,6))
     SKIP
    ENDIF  
    *ENDCASE
   ENDDO 
  ENDDO  
  *SELECT Talon
  SELECT c_talon
  ENDIF  && Отключение ошибки SMA

  IF M.MMA = .T.
   SELECT Gosp
   GO TOP 
   SET ORDER TO 
   DO WHILE !EOF()
    Karta = c_i
    CREATE CURSOR t_mma (cod n(6))
    INDEX on cod TAG cod 
    SET ORDER TO cod 
    SELECT Gosp
    DO WHILE c_i = Karta
     m.cod = cod 
     IF !(IsMES(m.cod) OR IsVMP(m.cod))
      SKIP 
      LOOP 
     ENDIF 
     IF INLIST(INT(m.cod/1000),83,183)
      SKIP 
      LOOP 
     ENDIF 
     IF SEEK(m.cod, 't_mma')
      m.recid = recid
      rval = InsError('S', 'MMA', m.recid, '',;
     	'Повтор МЭС за одну госпитализацию')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     INSERT INTO t_mma FROM MEMVAR 
     SKIP 
    ENDDO 
    USE IN t_mma 
   ENDDO  
   SELECT c_talon
  ENDIF

  IF M.DIA = .T. && два tip=0 за одну госпитализацию
   SELECT Gosp
   GO TOP 
   SET ORDER TO 
   DO WHILE !EOF()
    Karta = c_i
    CREATE CURSOR t_dia (cod n(6))
    SELECT Gosp
    DO WHILE c_i = Karta
     m.cod = cod 
     m.tip = tip
     m.d_type = d_type
     IF !(IsMES(m.cod) OR IsVMP(m.cod))
      SKIP 
      LOOP 
     ENDIF 
     IF !INLIST(m.tip,'0','A','T','R') OR m.d_type='R'
      SKIP 
      LOOP 
     ENDIF 
     IF RECCOUNT('t_dia')>0
      m.recid = recid
      rval = InsError('S', 'DIA', m.recid, '',;
     	'Повтор МЭС c tip=0 за одну госпитализацию')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     INSERT INTO t_dia FROM MEMVAR 
     SKIP 
    ENDDO 
    USE IN t_dia
   ENDDO  
   SELECT c_talon
  ENDIF

  IF m.istestmode
   INSERT INTO c_test VALUES ('SMA', SECONDS()-m.t_t)
   m.t_t = SECONDS()
  ENDIF 
  
  IF M.PPA AND m.qcod='I3' && третий вариант от 14.02.2020, здесь только стационар, остальное в ss_flk
   IF USED('nsif') && AND USED('gr_plan')
    m.ks_plan = IIF(SEEK(m.lpuid, 'nsif'), nsif.ks, 0)

    SELECT Gosp 
    GO TOP 
    DO WHILE !EOF()
     Karta = c_i
     DO WHILE c_i = Karta

      m.IsErrPPA = .F.
      
      m.recid = recid
      m.iserr = IIF(SEEK(m.recid, 'serror'), .T., .F.)
      IF m.iserr
       SKIP 
      ENDIF 

      m.cod   = cod
      m.otd   = otd
      m.usl_ok = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), profot.usl_ok, ' ')
      IF m.usl_ok<>'1'
       IF !EOF()
        SKIP 
       ENDIF 
      ENDIF 
      
      m.k_u   = k_u
      m.s_all = s_all+s_lek

      m.sn_pol = sn_pol
      m.d_u    = d_u

      IF m.ks_plan<=0
       m.IsErrPPA = .T.
      ELSE 
       m.oks_fact = nsif.ks_fact
       m.nks_fact = m.oks_fact + m.s_all
       IF !(m.nks_fact<=nsif.ks)
        m.IsErrPPA = .T.
       ENDIF 
      ENDIF 
     
      IF m.IsErrPPA = .T.
       DO WHILE c_i = Karta
        m.recid = recid
        rval    = InsError('S', 'PPA', m.recid, '', 'Превышен лимит ks (1)!')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       
        m.oks_fact = nsif.ks_fact
        m.nks_fact = m.oks_fact + m.s_all
        REPLACE ks_fact WITH m.nks_fact IN nsif
       
        IF !SEEK(Gosp.c_i, 'c_pp')
         INSERT INTO c_pp (c_i) VALUES (Gosp.c_i)
        ENDIF 
       
        SKIP
       ENDDO 

      ELSE 
      
       m.iter = 1      
       DO WHILE c_i = Karta
        m.oks_fact = nsif.ks_fact
        m.nks_fact = m.oks_fact + m.s_all
        REPLACE ks_fact WITH m.nks_fact IN nsif
       
        IF !(m.nks_fact<=nsif.ks)
         m.IsErrPPA = .T.
        ENDIF 
       
        SKIP
        m.iter = m.iter + 1
       ENDDO 
      
       IF m.IsErrPPA = .T.
        SKIP -1*m.iter
        DO WHILE c_i = Karta
         m.recid = recid
         rval    = InsError('S', 'PPA', m.recid, '', 'Превышен лимит ks (2)!')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       
         IF !SEEK(Gosp.c_i, 'c_pp')
          INSERT INTO c_pp (c_i) VALUES (Gosp.c_i)
         ENDIF 
       
         SKIP
        ENDDO 
       ENDIF 

      ENDIF 

     ENDDO 
    ENDDO  
    SELECT c_talon
   
    ENDIF 
   ENDIF 

  IF M.PPA AND m.qcod='S7' AND 1=2 && третий вариант от 14.02.2020, здесь только стационар, остальное в ss_flk
   IF USED('nsif') && AND USED('gr_plan')
    m.ks_plan = IIF(SEEK(m.lpuid, 'nsif'), nsif.ks, 0)

    SELECT Gosp 
    GO TOP 
    DO WHILE !EOF()
     Karta = c_i
     DO WHILE c_i = Karta

      m.IsErrPPA = .F.
      
      m.recid = recid
      m.iserr = IIF(SEEK(m.recid, 'serror'), .T., .F.)
      IF m.iserr
       SKIP 
      ENDIF 

      m.cod   = cod
      m.otd   = otd
      m.usl_ok = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), profot.usl_ok, ' ')
      IF m.usl_ok<>'1'
       IF !EOF()
        SKIP 
       ENDIF 
      ENDIF 
      
      m.k_u   = k_u
      m.s_all = s_all+s_lek

      m.sn_pol = sn_pol
      m.d_u    = d_u

      IF m.ks_plan<=0
       m.IsErrPPA = .T.
      ELSE 
       m.oks_fact = nsif.ks_fact
       m.nks_fact = m.oks_fact + m.s_all
       IF !(m.nks_fact<=nsif.ks)
        m.IsErrPPA = .T.
       ENDIF 
      ENDIF 
     
      IF m.IsErrPPA = .T.
       DO WHILE c_i = Karta
        m.recid = recid
        rval    = InsError('S', 'PPA', m.recid, '', 'Превышен лимит ks (1)!')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       
        m.oks_fact = nsif.ks_fact
        m.nks_fact = m.oks_fact + m.s_all
        REPLACE ks_fact WITH m.nks_fact IN nsif
       
        IF !SEEK(Gosp.c_i, 'c_pp')
         INSERT INTO c_pp (c_i) VALUES (Gosp.c_i)
        ENDIF 
       
        SKIP
       ENDDO 

      ELSE 
      
       m.iter = 1      
       DO WHILE c_i = Karta
        m.oks_fact = nsif.ks_fact
        m.nks_fact = m.oks_fact + m.s_all
        REPLACE ks_fact WITH m.nks_fact IN nsif
       
        IF !(m.nks_fact<=nsif.ks)
         m.IsErrPPA = .T.
        ENDIF 
       
        SKIP
        m.iter = m.iter + 1
       ENDDO 
      
       IF m.IsErrPPA = .T.
        SKIP -1*m.iter
        DO WHILE c_i = Karta
         m.recid = recid
         rval    = InsError('S', 'PPA', m.recid, '', 'Превышен лимит ks (2)!')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       
         IF !SEEK(Gosp.c_i, 'c_pp')
          INSERT INTO c_pp (c_i) VALUES (Gosp.c_i)
         ENDIF 
       
         SKIP
        ENDDO 
       ENDIF 

      ENDIF 

     ENDDO 
    ENDDO  
    SELECT c_talon
   
    ENDIF 
   ENDIF 

  IF m.istestmode
   INSERT INTO c_test VALUES ('PPA (step 2)', SECONDS()-m.t_t)
   m.t_t = SECONDS()
  ENDIF 


  ENDIF && Глобальное отключение ошибок счета!
  
  *IF talon.IsPr == .F. && Глобальное отключение ошибок счета!
  IF c_talon.IsPr == .F. && Глобальное отключение ошибок счета!
   IF M.DLA == .T.
    SET ORDER TO Unik
    GO TOP
    DO WHILE !EOF()
     m.c_i     = c_i
     m.add_vir = '0'
     IF OCCURS('#', m.c_i)>=3
      *m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), ;
      	SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1), '0')
      m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), ;
      	m.c_i, SPACE(30))
     ENDIF 
     *Vir = m.add_vir+sn_pol+otd+ds+Padl(cod,6,'0')+DToC(d_u)
     Vir = m.add_vir+sn_pol+otd+ds+Padl(cod,6,'0')+DToC(d_u)+pcod
     SKIP
     jjj = .T.
     m.c_i     = c_i
     m.add_vir = '0'
     IF OCCURS('#', m.c_i)>=3
      *m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), ;
      	SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1), '0')
      m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), ;
      	m.c_i, SPACE(30))
     ENDIF 
     DO WHILE m.add_vir+sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u)+pcod = vir
      IF jjj = .T.
       jjj = .F.
       SKIP -1
       *IF !INLIST(d_type,'2','9')
       IF !INLIST(d_type,'9')
        m.recid = recid
        rval = InsError('S', 'DLA', m.recid, '',;
        	'Повторная запись в файле счет '+m.add_vir+' '+ALLTRIM(sn_pol)+' '+otd+' '+ds+' '+PADL(cod,6,'0')+' '+DTOC(d_u))
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
       SKIP 1
      ENDIF
      *IF !INLIST(d_type,'2','9')
      IF !INLIST(d_type,'9')
       m.recid = recid
        rval = InsError('S', 'DLA', m.recid, '',;
        	'Повторная запись в файле счет '+m.add_vir+' '+ALLTRIM(sn_pol)+' '+otd+' '+ds+' '+PADL(cod,6,'0')+' '+DTOC(d_u))
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      SKIP 
      m.c_i     = c_i
      m.add_vir = '0'
      IF OCCURS('#', m.c_i)>=3
       *m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), ;
       	SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1), '0')
       m.add_vir = IIF(INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6'), ;
      	m.c_i, SPACE(30))
      ENDIF 
     ENDDO  
    ENDDO 
    SET ORDER TO 
   ENDIF 
  ENDIF 

  IF m.istestmode
   INSERT INTO c_test VALUES ('DLA', SECONDS()-m.t_t)
   m.t_t = SECONDS()
  ENDIF 

  CREATE CURSOR AllBad (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol 
  
  *SELECT talon 
  SELECT c_talon 
  *SET ORDER TO sn_pol && !!!
  GO TOP 

  * Самый длительный процесс!!
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
  * Самый длительный процесс!!
  
  IF m.istestmode
   INSERT INTO c_test VALUES ('allbad01', SECONDS()-m.t_t)
   m.t_t = SECONDS()
  ENDIF 
  
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
  
  IF m.istestmode
   INSERT INTO c_test VALUES ('allbad02', SECONDS()-m.t_t)
   m.t_t = SECONDS()
  ENDIF 

  * Самый длительный процесс!! решено!
  SELECT talon 
  SET RELATION TO recid INTO serror ADDITIVE 
  SUM s_all FOR !EMPTY(serror.rid) TO m.s_flk
  SUM s_lek FOR !EMPTY(serror.rid) TO m.ls_flk
  SET RELATION OFF INTO serror
  *SUM(s_all) FOR SEEK(RecId, 'sError') TO m.s_flk
  SET RELATION OFF INTO people
  * Самый длительный процесс!!
  
  IF m.istestmode
   INSERT INTO c_test VALUES ('allbad03', SECONDS()-m.t_t)
   m.t_t = SECONDS()
  ENDIF 

  =ClBase()
  
  IF USED('c_pp')
   USE IN c_pp
  ENDIF 

  SELECT AisOms
  =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
  SELECT AisOms
  m.t_1 = SECONDS()

  IF sum_flk != m.s_flk
*   MESSAGEBOX('sum_flk= '+TRANSFORM(sum_flk,'999999.99')+CHR(13)+CHR(10)+;
   	'm.s_flk= '+TRANSFORM(m.s_flk,'999999.99'),0+64,m.mcod)
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

  IF m.istestmode
   INSERT INTO c_test VALUES ('total:', SECONDS()-m.t_a)
   m.t_t = SECONDS()
  ENDIF 

  IF m.istestmode
   IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\timetest.dbf')
    fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\timetest.dbf')
   ENDIF 
   SELECT c_test
   COPY TO &pBase\&gcPeriod\&mcod\timetest
   USE IN c_test
  ENDIF 
  
  SELECT aisoms
  REPLACE sum_flk WITH m.s_flk, ls_flk WITH m.ls_flk, t_4 WITH m.t_1-m.t_0

 
 WAIT CLEAR 

 RELEASE M.PSA, M.ERA, M.ECA, M.E1A, M.E2A, M.E4A, M.E5A, M.E6A, M.E7A, M.E8A, M.H6A, M.COA, M.HCA, M.DUA, M.H8A, M.HEA, M.CSA, M.TVA, M.NLA,;
 	M.MDA, M.H3A, M.SOA, M.R1A, M.R2A, M.R3A, M.UVA, M.DVA, M.UOA, M.NOA, M.NMA, M.NUA, M.NSA, M.SMA, M.DIA, M.DDA, M.HNA, M.DLA,;
    M.DRA, M.POA, M.VDA, M.TFA, M.PPA, M.G1A, M.G2A, M.G3A, M.G4A, M.NRA, M.KEA, M.D2A, M.THA, M.TLA, M.HOA, M.SKA, M.FSA, M.PFA,;
    M.W2A, M.O0A, M.O1A, M.O2A, M.O3A, M.O4A, M.O5A, M.O6A, M.O7A, M.O8A, M.O8A, M.OAA, M.OBA, M.OCA, M.ODA, M.OEA, M.OFA, M.OGA,;
    M.OHA, M.OIA, M.OJA, M.OKA, M.OLA, M.OMA, M.ONA, M.OOA, M.OPA, M.OQA, M.ORA, M.OSA, M.OTA, M.OUA, M.OVA, M.OWA, M.OXA, M.OYA,;
    M.OZA, M.ENA, M.IPA, M.EGA, M.X1A, M.X2A, M.X3A, M.X4A, M.X5A, M.X6A, M.X7A, M.X8A, M.X9A, M.PLA, M.H7A, M.DKA, M.UMA, M.D4A,;
    M.S1A, M.MPA, M.MZA, M.NDA, M.D1A, M.PGA, M.NWA, M.HRA, M.D6A, M.DNA, M.NVA, M.TPA, M.R4A, M.PHA, M.WEA, M.MMA, M.CVA, M.CPA

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
   INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('R', 1, cError, cRecId, cDetail, cComment)
  ELSE 
  ENDIF !SEEK(cRecId, 'rError')
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(cRecId, 'sError')
   INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('S', 1, cError, cRecId, cDetail, cComment)
   RETURN .T.
  ELSE 
   IF cError != sError.c_err
    INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('S', 1, cError, cRecId, cDetail, cComment)
   ENDIF cError != sError.c_err 
  ENDIF !SEEK(cRecId, 'sError')
 ENDIF 
RETURN .F.

FUNCTION OpBase(ppath)
 tnresult = 0
 tnresult = tnresult + OpenFile(ppath+'\people', 'people', 'share', 'sn_pol')

 tnresult = tnresult + OpenFile(ppath+'\talon', 'talon', 'share')
 
 SELECT * FROM talon ORDER BY sn_pol, cod, d_u, d_type DESC INTO CURSOR c_talon READWRITE 
 SELECT c_talon 
 INDEX ON sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u)+pcod TAG unik 
 SET ORDER TO 

 tnresult = tnresult + OpenFile(ppath+'\doctor', 'doctor', 'share', 'pcod')
 IF tmonth>1
  m.deads = pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\deads'
  m.stop  = pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\stop'
  tnresult = tnresult + OpenFile(pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\'+'nsi'+'\polic_h', 'polic_h', 'share', 'sn_pol')
  IF fso.FileExists(pbase+'\'+STR(tyear,4)+PADL(tmonth-1,2,'0')+'\'+m.mcod+'\hosp.dbf')
   tnresult = tnresult + OpenFile(pbase+'\'+STR(tyear,4)+PADL(tmonth-1,2,'0')+'\'+m.mcod+'\hosp', 'hosp_p', 'share', 'c_i')
  ENDIF 
 ELSE
  m.deads = pbase +'\'+ STR(tyear-1,4)+'12'+'\deads'
  m.stop  = pbase +'\'+ STR(tyear-1,4)+'12'+'\stop'
  tnresult = tnresult + OpenFile(pbase +'\'+ STR(tyear-1,4)+'12'+'\'+'nsi'+'\polic_h', 'polic_h', 'share', 'sn_pol')
 ENDIF 
 IF fso.FileExists(m.ppath+'\hosp.dbf')
  tnresult = tnresult + OpenFile(m.ppath+'\hosp', 'hosp', 'share', 'c_i')
 ENDIF 
 IF fso.FileExists(m.deads+'.dbf')
  tnresult = tnresult + OpenFile(m.deads, 'deads', 'share', 'sn_pol')
 ENDIF 
 IF fso.FileExists(m.stop+'.dbf')
  tnresult = tnresult + OpenFile(m.stop, 'stop', 'share', 'enp')
 ENDIF 
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
 IF fso.FileExists(ppath+'\ONK_CONS'+m.qcod+'.dbf')
  tnresult = tnresult + OpenFile(ppath+'\ONK_CONS'+m.qcod, 'onk_cons', 'share', 'recid')
 ENDIF 
 IF fso.FileExists(ppath+'\CV_LS'+m.qcod+'.dbf')
  tnresult = tnresult + OpenFile(ppath+'\CV_LS'+m.qcod, 'cv_ls', 'share', 'recid_s')
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
 *tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\polic_dp', 'polic_dp', 'share', 'sn_pol')
 *tnresult = tnresult + OpenFile(pcommon+'\dsdisp', 'dsdisp', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\kpresl', 'kpresl', 'share', 'tip')
 *tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\mo_vmp', 'movmp', 'share', 'lpu_id')
 *tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spi_lpu_dd', 'spidd', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\noth', 'noth', 'share', 'cod')
 tnresult = tnresult + OpenFile(pcommon+'\lpudogs', 'lpudogs', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pcommon+'\dspcodes', 'dspcodes', 'share', 'cod')
 tnresult = tnresult + OpenFile(pcommon+'\lpuskp', 'lpuskp', 'share', 'lpuid')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\exclhors', 'exclhors', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\codprv', 'codprv', 'share', 'var')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tipnomes', 'tipnomes', 'share', 'vir')
 tnresult = tnresult + OpenFile(pcommon+'\prv002xx', 'prv002', 'share', 'profil')

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
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onoplsxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onoplsxx', 'onopls', 'share', 'unik')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\ondopkxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\ondopkxx', 'ondopk', 'share', 'COD_DKK')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\onlpshxx.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\onlpshxx', 'onlpsh', 'share', 'code_sh')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\msext.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\msext', 'msext', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\sprved.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprved', 'sprved', 'share', 'mcod')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\F003.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\F003', 'f003', 'share', 'lpu_id')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\gosp.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\gosp', 'sv_gosp', 'share', 'sn_pol')
 ENDIF 
 IF fso.FileExists(pCommon+'\dc_du.dbf')
  tnresult = tnresult + OpenFile(pCommon+'\dc_du', 'dc_du', 'share')
 ENDIF 
 IF fso.FileExists(pCommon+'\prvr2.dbf')
  tnresult = tnresult + OpenFile(pCommon+'\prvr2', 'prvr2', 'share', 'profil')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\nsi\nsio.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\nsi\nsio', 'nsio', 'share', 'unik')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\nsi\nsif.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\nsi\nsif', 'nsif', 'share', 'lpu_id')
 ENDIF 
* IF fso.FileExists(pbase+'\'+gcperiod+'\nsi\sprnco.dbf')
*  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\nsi\sprnco', 'sprnco', 'share', 'lpu_id')
* ENDIF 
 IF fso.FileExists(pCommon+'\ms_ds_prv.dbf')
  tnresult = tnresult + OpenFile(pCommon+'\ms_ds_prv', 'ms_ds_prv', 'share', 'unik')
 ENDIF 
 IF fso.FileExists(pCommon+'\gr_plan.dbf')
  tnresult = tnresult + OpenFile(pCommon+'\gr_plan', 'gr_plan', 'share', 'cod')
 ENDIF 
 IF fso.FileExists(pCommon+'\rsltishod.dbf')
  tnresult = tnresult + OpenFile(pCommon+'\rsltishod', 'rsltishod', 'share', 'unik')
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\nsi\ns36.dbf')
  tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\nsi\ns36', 'ns36', 'share', 'cod')
 ENDIF 
 **IF fso.FileExists(pCommon+'\pervpr.dbf')
 *IF fso.FileExists(pCommon+'\perv.dbf')
 * *tnresult = tnresult + OpenFile(pCommon+'\pervpr', 'ppr', 'share', 'cod')
 * tnresult = tnresult + OpenFile(pCommon+'\perv', 'ppr', 'share', 'cod')
 *ENDIF 
 
 * открываем номерники для проверки направления
 *FOR m.i_m=1 TO 3
 * m.p_period = LEFT(DTOS(GOMONTH(m.tdat1,-m.i_m)),6)
 * IF fso.FolderExists(m.pBase+'\'+m.p_period)
 *  IF fso.FileExists(m.pBase+'\'+m.p_period+'\NSI\Outs.dbf')
 *   IF OpenFile(m.pBase+'\'+m.p_period+'\NSI\Outs', 'outs&p_period', 'shar', 'enp')>0
 *    IF USED('outs&p_period')
 *     USE IN outs&p_period
 *    ENDIF 
 *   ENDIF 
 *  ENDIF 
 * ENDIF 
 *ENDFOR 
 * следующий период

  m.p_period = LEFT(DTOS(GOMONTH(m.tdat1,1)),6)
  IF fso.FolderExists(m.pBase+'\'+m.p_period)
   IF fso.FileExists(m.pBase+'\'+m.p_period+'\NSI\Outs.dbf')
    IF OpenFile(m.pBase+'\'+m.p_period+'\NSI\Outs', 'outs_n', 'shar', 'enp')>0
     IF USED('outs_n')
      USE IN outs_n
     ENDIF 
    ELSE 
     *MESSAGEBOX('НОМЕРНИК outs ' + m.p_period + ' ОТКРЫТ',0+64,'')
    ENDIF 
   ENDIF 
  ENDIF 
 * следующий период
 * открываем номерники для проверки направления

 IF !m.IsPr
 UPDATE rerror SET Tip=1, dt=DATETIME(), usr=m.gcUser WHERE SUBSTR(c_err,3,1)='A'
 DELETE FOR SUBSTR(c_err,3,1)='A' AND IIF(FIELD('et')='ET', et=1, 1=1) IN rerror

 UPDATE serror SET Tip=1, dt=DATETIME(), usr=m.gcUser WHERE SUBSTR(c_err,3,1)='A'
 DELETE FOR SUBSTR(c_err,3,1)='A' AND IIF(FIELD('et')='ET', et=1, 1=1) IN serror
 
 UPDATE rerror SET Tip=1, dt=DATETIME(), usr=m.gcUser WHERE SUBSTR(c_err,3,1)='B'
 DELETE FOR SUBSTR(c_err,3,1)='B' AND IIF(FIELD('et')='ET', et=1, 1=1) IN rerror

 UPDATE serror SET Tip=1, dt=DATETIME(), usr=m.gcUser WHERE SUBSTR(c_err,3,1)='B'
 DELETE FOR SUBSTR(c_err,3,1)='B' AND IIF(FIELD('et')='ET', et=1, 1=1) IN serror
 ENDIF 
RETURN .t.

FUNCTION ClBase()
 IF USED('talon')
  USE IN talon 
 ENDIF 
 IF USED('c_talon')
  USE IN c_talon 
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
 IF USED('gosp_d')
  USE IN Gosp_d
 ENDIF 
 IF USED('polic_h')
  USE IN polic_h
 ENDIF 
 *IF USED('polic_dp')
 * USE IN polic_dp
 *ENDIF 
 *IF USED('dsdisp')
 * USE IN dsdisp
 *ENDIF 
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
 IF USED('prv002')
  USE IN prv002
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
 IF USED('onk_cons')
  USE IN onk_cons
 ENDIF 
 IF USED('cv_ls')
  USE IN cv_ls
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
 IF USED('onopls')
  USE IN onopls
 ENDIF 
 IF USED('ondopk')
  USE IN ondopk
 ENDIF 
 IF USED('onlpsh')
  USE IN onlpsh
 ENDIF 
 IF USED('codprv')
  USE IN codprv
 ENDIF 
 IF USED('tipnomes')
  USE IN tipnomes
 ENDIF 
 IF USED('sv_gosp')
  USE IN sv_gosp
 ENDIF 
 IF USED('deads')
  USE IN deads 
 ENDIF 
 IF USED('stop')
  USE IN stop 
 ENDIF 
 IF USED('msext')
  USE IN msext
 ENDIF 
 IF USED('dc_du')
  USE IN dc_du
 ENDIF 
 *IF USED('tarion')
 * USE IN tarion
 *ENDIF 
 *IF USED('medx')
 * USE IN medx
 *ENDIF 
 *IF USED('medpack')
 * USE IN medpack
 *ENDIF 
 *IF USED('mfc')
 * USE IN mfc
 *ENDIF 
 IF USED('curs_h6')
  USE IN curs_h6
 ENDIF 
 IF USED('curs_h6p')
  USE IN curs_h6p
 ENDIF 
 IF USED('hosp_p')
  USE IN hosp_p
 ENDIF 
 IF USED('hosp')
  USE IN hosp
 ENDIF 
 IF USED('sprved')
  USE IN sprved
 ENDIF 
 IF USED('f003')
  USE IN f003
 ENDIF 
 IF USED('nsio')
  USE IN nsio
 ENDIF 
 IF USED('nsif')
  USE IN nsif
 ENDIF 
 IF USED('ms_ds_prv')
  USE IN ms_ds_prv
 ENDIF 
 IF USED('gr_plan')
  USE IN gr_plan
 ENDIF 
 IF USED('prvr2')
  USE IN prvr2
 ENDIF 
 IF USED('rsltishod')
  USE IN rsltishod
 ENDIF 
 IF USED('ns36')
  USE IN ns36
 ENDIF 
 *IF USED('ppr')
 * USE IN ppr
 *ENDIF 
* IF USED('sprnco')
*  USE IN sprnco
* ENDIF 
 *FOR m.i_m=1 TO 3
 * m.p_period = LEFT(DTOS(GOMONTH(m.tdat1,-m.i_m)),6)
 * IF USED('outs&p_period')
 *  USE IN outs&p_period
 * ENDIF 
 *ENDFOR 
 *m.p_period = LEFT(DTOS(GOMONTH(m.tdat1,+1)),6)
 IF USED('outs_n')
  USE IN outs_n
 ENDIF 
RETURN 

FUNCTION ReadErrorStatus(para1)
 LOCAL WorkIniFile
 WorkIniFile = para1

 IF EMPTY(WorkIniFile)
  M.PSA = .T.
  M.ERA = .T.
  M.ECA = .T.
  M.E1A = .T.
  M.E2A = .T.
  M.E4A = .T.
  M.E5A = .T.
  M.E6A = .T.
  M.E7A = .T.
  M.E8A = .T.

  M.H6A = .T.
  M.COA = .T.
  M.HCA = .T.
  M.DUA = .T.
  M.H8A = .T.
  M.HEA = .T.
  M.CSA = .T.
  M.TVA = .T.
  M.NLA = .T.
  M.MDA = .T.
  M.H3A = .T.
  M.SOA = .T.
  M.R1A = .T.
  M.R2A = .T.
  M.R3A = .T. 
  M.UVA = .T.
  M.DVA = .T.
  M.UOA = .T.
  M.NOA = .T.
  M.NMA = .T.
  M.NUA = .T.
  M.NSA = .T.
  M.SMA = .T.
  M.DIA = .T.
  M.DDA = .T.
  M.HNA = .T.
  M.DLA = .T.
  M.DRA = .T.
  M.POA = .T.
  M.VDA = .T.
  M.TFA = .T.
  M.PPA = .T.
  M.G1A = .F.
  M.G2A = .F.
  M.G3A = .F.
  M.G4A = .F.
  M.NRA = .T.
  M.KEA = .T.
  M.D2A = .F.
  M.THA = .T.
  M.TLA = .T.
  M.HOA = .T.
  M.S1A = .T.
  M.MPA = .T.
  M.MZA = .T.
  M.NDA = .T.
  M.D1A = .T.
  M.PGA = .T.
  M.NWA = .T.
  M.HRA = .T.
  M.D6A = .T.
  M.DNA = .T.
  M.NVA = .T.
  M.TPA = .T.
  M.R4A = .T.
  M.PHA = .T.
  M.WEA = .T.
  M.MMA = .T.
  M.CVA = .T.
  M.CPA = .T.
  M.SKA = .T.
  M.FSA = .T.
  M.PFA = .T.
  M.W2A = .T.

  M.O0A = .T.
  M.O1A = .T.
  M.O2A = .T.
  M.O3A = .T.
  M.O4A = .T.
  M.O5A = .T.
  M.O6A = .T.
  M.O7A = .T.
  M.O8A = .T.
  M.O8A = .T.
  M.OAA = .T.
  M.OBA = .T.
  M.OCA = .T.
  M.ODA = .T.
  M.OEA = .T.
  M.OFA = .T.
  M.OGA = .T.
  M.OHA = .T.
  M.OIA = .T.
  M.OJA = .T.
  M.OKA = .T.
  M.OLA = .T.
  M.OMA = .T.
  M.ONA = .T.
  M.OOA = .T.
  M.OPA = .T.
  M.OQA = .T.
  M.ORA = .T.
  M.OSA = .T.
  M.OTA = .T.
  M.OUA = .T.
  M.OVA = .T.
  M.OWA = .T.
  M.OXA = .T.
  M.OYA = .T.
  M.OZA = .T.
  M.ENA = .T.
  M.IPA = .T.
  M.EGA = .T.

  M.X1A = .T.
  M.X2A = .T.
  M.X3A = .T.
  M.X4A = .T.
  M.X5A = .T.
  M.X6A = .T.
  M.X7A = .T.
  M.X8A = .T.
  M.X9A = .T.
  M.PLA = .T.
  M.H7A  = .T.
  M.DKA = .T.
  M.UMA = .T.
  M.D4A = .T.
 ELSE 
  M.PSA = ReadFromIni(WorkIniFile, 'PSA', .T.)
  M.ERA = ReadFromIni(WorkIniFile, 'ERA', .T.)
  M.ECA = ReadFromIni(WorkIniFile, 'ECA', .T.)
  M.E1A = ReadFromIni(WorkIniFile, 'E1A', .T.)
  M.E2A = ReadFromIni(WorkIniFile, 'E2A', .T.)
  M.E4A = ReadFromIni(WorkIniFile, 'E4A', .T.)
  M.E5A = ReadFromIni(WorkIniFile, 'E5A', .T.)
  M.E6A = ReadFromIni(WorkIniFile, 'E6A', .T.)
  M.E7A = ReadFromIni(WorkIniFile, 'E7A', .T.)
  M.E8A = ReadFromIni(WorkIniFile, 'E8A', .T.)

  M.H6A = ReadFromIni(WorkIniFile, 'H6A', .T.)
  M.COA = ReadFromIni(WorkIniFile, 'COA', .T.)
  M.HCA = ReadFromIni(WorkIniFile, 'HCA', .T.)
  M.DUA = ReadFromIni(WorkIniFile, 'DUA', .T.)
  M.H8A = ReadFromIni(WorkIniFile, 'H8A', .T.)
  M.HEA = ReadFromIni(WorkIniFile, 'HEA', .T.)
  M.CSA = ReadFromIni(WorkIniFile, 'CSA', .T.)
  M.TVA = ReadFromIni(WorkIniFile, 'TVA', .T.)
  M.NLA = ReadFromIni(WorkIniFile, 'NVA', .T.)
  M.MDA = ReadFromIni(WorkIniFile, 'MDA', .T.)
  M.H3A = ReadFromIni(WorkIniFile, 'H3A', .T.)
  M.SOA = ReadFromIni(WorkIniFile, 'SOA', .T.)
  M.R1A = ReadFromIni(WorkIniFile, 'R1A', .T.)
  M.R2A = ReadFromIni(WorkIniFile, 'R2A', .T.)
  M.R3A = ReadFromIni(WorkIniFile, 'R3A', .T.)
  M.UVA = ReadFromIni(WorkIniFile, 'UVA', .T.)
  M.DVA = ReadFromIni(WorkIniFile, 'DVA', .T.)
  M.UOA = ReadFromIni(WorkIniFile, 'UOA', .T.)
  M.NOA = ReadFromIni(WorkIniFile, 'NOA', .T.)
  M.NMA = ReadFromIni(WorkIniFile, 'NMA', .T.)
  M.NUA = ReadFromIni(WorkIniFile, 'NUA', .T.)
  M.NSA = ReadFromIni(WorkIniFile, 'NSA', .T.)
  M.SMA = ReadFromIni(WorkIniFile, 'SMA', .T.)
  M.DIA = ReadFromIni(WorkIniFile, 'DIA', .T.)
  M.DDA = ReadFromIni(WorkIniFile, 'DDA', .T.)
  M.HNA = ReadFromIni(WorkIniFile, 'HNA', .T.)
  M.DLA = ReadFromIni(WorkIniFile, 'DLA', .T.)
  M.DRA = ReadFromIni(WorkIniFile, 'DRA', .T.)
  M.POA = ReadFromIni(WorkIniFile, 'POA', .T.)
  M.VDA = ReadFromIni(WorkIniFile, 'VDA', .T.)
  M.TFA = ReadFromIni(WorkIniFile, 'TFA', .T.)
  M.PPA = ReadFromIni(WorkIniFile, 'PPA', .F.)
  M.G1A = ReadFromIni(WorkIniFile, 'G1A', .F.)
  M.G2A = ReadFromIni(WorkIniFile, 'G2A', .F.)
  M.G3A = ReadFromIni(WorkIniFile, 'G3A', .F.)
  M.G4A = ReadFromIni(WorkIniFile, 'G4A', .F.)
  M.NRA = ReadFromIni(WorkIniFile, 'NRA', .T.)
  M.KEA = ReadFromIni(WorkIniFile, 'KEA', .T.)
  M.D2A = ReadFromIni(WorkIniFile, 'D2A', .F.)
  M.THA = ReadFromIni(WorkIniFile, 'THA', .T.)
  M.TLA = ReadFromIni(WorkIniFile, 'TLA', .T.)
  M.HOA = ReadFromIni(WorkIniFile, 'HOA', .T.)
  M.S1A = ReadFromIni(WorkIniFile, 'S1A', .T.)
  M.MPA = ReadFromIni(WorkIniFile, 'MPA', .T.)
  M.MZA = ReadFromIni(WorkIniFile, 'MZA', .T.)
  M.NDA = ReadFromIni(WorkIniFile, 'NDA', .T.)
  M.D1A = ReadFromIni(WorkIniFile, 'D1A', .T.)
  M.PGA = ReadFromIni(WorkIniFile, 'PGA', .T.)
  M.NWA = ReadFromIni(WorkIniFile, 'NWA', .T.)
  M.HRA = ReadFromIni(WorkIniFile, 'NWA', .T.)
  M.DNA = ReadFromIni(WorkIniFile, 'DNA', .T.)
  M.NVA = ReadFromIni(WorkIniFile, 'NVA', .T.)
  M.TPA = ReadFromIni(WorkIniFile, 'TPA', .T.)
  M.R4A = ReadFromIni(WorkIniFile, 'R4A', .T.)
  M.PHA = ReadFromIni(WorkIniFile, 'PHA', .T.)
  M.WEA = ReadFromIni(WorkIniFile, 'WEA', .T.)
  M.MMA = ReadFromIni(WorkIniFile, 'MMA', .T.)
  M.CVA = ReadFromIni(WorkIniFile, 'CVA', .T.)
  M.CPA = ReadFromIni(WorkIniFile, 'CPA', .T.)
  M.D6A = ReadFromIni(WorkIniFile, 'D6A', .T.)
  M.FSA = ReadFromIni(WorkIniFile, 'FSA', .T.)
  M.PFA = ReadFromIni(WorkIniFile, 'PFA', .T.)
  M.SKA = ReadFromIni(WorkIniFile, 'SKA', .T.)
  M.W2A = ReadFromIni(WorkIniFile, 'W2A', .T.)

  M.O0A = ReadFromIni(WorkIniFile, 'O1A', .T.)
  M.O1A = ReadFromIni(WorkIniFile, 'O1A', .T.)
  M.O2A = ReadFromIni(WorkIniFile, 'O2A', .T.)
  M.O3A = ReadFromIni(WorkIniFile, 'O3A', .T.)
  M.O4A = ReadFromIni(WorkIniFile, 'O4A', .T.)
  M.O5A = ReadFromIni(WorkIniFile, 'O5A', .T.)
  M.O6A = ReadFromIni(WorkIniFile, 'O6A', .T.)
  M.O7A = ReadFromIni(WorkIniFile, 'O7A', .T.)
  M.O8A = ReadFromIni(WorkIniFile, 'O8A', .T.)
  M.O9A = ReadFromIni(WorkIniFile, 'O9A', .T.)
  M.OAA = ReadFromIni(WorkIniFile, 'OAA', .T.)
  M.OBA = ReadFromIni(WorkIniFile, 'OBA', .T.)
  M.OCA = ReadFromIni(WorkIniFile, 'OCA', .T.)
  M.ODA = ReadFromIni(WorkIniFile, 'ODA', .T.)
  M.OEA = ReadFromIni(WorkIniFile, 'OEA', .T.)
  M.OFA = ReadFromIni(WorkIniFile, 'OFA', .T.)
  M.OGA = ReadFromIni(WorkIniFile, 'OGA', .T.)
  M.OHA = ReadFromIni(WorkIniFile, 'OHA', .T.)
  M.OIA = ReadFromIni(WorkIniFile, 'OIA', .T.)
  M.OJA = ReadFromIni(WorkIniFile, 'OJA', .T.)
  M.OKA = ReadFromIni(WorkIniFile, 'OKA', .T.)
  M.OLA = ReadFromIni(WorkIniFile, 'OLA', .T.)
  M.OMA = ReadFromIni(WorkIniFile, 'OMA', .T.)
  M.ONA = ReadFromIni(WorkIniFile, 'ONA', .T.)
  M.OOA = ReadFromIni(WorkIniFile, 'OOA', .T.)
  M.OPA = ReadFromIni(WorkIniFile, 'OPA', .T.)
  M.OQA = ReadFromIni(WorkIniFile, 'OQA', .T.)
  M.ORA = ReadFromIni(WorkIniFile, 'ORA', .T.)
  M.OSA = ReadFromIni(WorkIniFile, 'OSA', .T.)
  M.OTA = ReadFromIni(WorkIniFile, 'OTA', .T.)
  M.OUA = ReadFromIni(WorkIniFile, 'OUA', .T.)
  M.OVA = ReadFromIni(WorkIniFile, 'OVA', .T.)
  M.OWA = ReadFromIni(WorkIniFile, 'OWA', .T.)
  M.OXA = ReadFromIni(WorkIniFile, 'OXA', .T.)
  M.OYA = ReadFromIni(WorkIniFile, 'OYA', .T.)
  M.OZA = ReadFromIni(WorkIniFile, 'OZA', .T.)

  M.ENA = ReadFromIni(WorkIniFile, 'ENA', .T.)
  M.IPA = ReadFromIni(WorkIniFile, 'IPA', .T.)
  M.EGA = ReadFromIni(WorkIniFile, 'EGA', .T.)

  M.X1A = ReadFromIni(WorkIniFile, 'X1A', .T.)
  M.X2A = ReadFromIni(WorkIniFile, 'X2A', .T.)
  M.X3A = ReadFromIni(WorkIniFile, 'X3A', .T.)
  M.X4A = ReadFromIni(WorkIniFile, 'X4A', .T.)
  M.X5A = ReadFromIni(WorkIniFile, 'X5A', .T.)
  M.X6A = ReadFromIni(WorkIniFile, 'X6A', .T.)
  M.X7A = ReadFromIni(WorkIniFile, 'X7A', .T.)
  M.X8A = ReadFromIni(WorkIniFile, 'X8A', .T.)
  M.X9A = ReadFromIni(WorkIniFile, 'X9A', .T.)

  M.PLA = ReadFromIni(WorkIniFile, 'PLA', .T.)
  M.H7A = ReadFromIni(WorkIniFile, 'H7A', .T.)
  M.DKA = ReadFromIni(WorkIniFile, 'DKA', .T.)
  M.UMA = ReadFromIni(WorkIniFile, 'UMA', .T.)
  M.D4A = ReadFromIni(WorkIniFile, 'D4A', .T.)
 ENDIF 
RETURN 


