PROCEDURE FormMAG02
 IF MESSAGEBOX('СФОРМИРОВАТЬ ФОРМУ МАГ-2',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\FormMAG02.xls')
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ ШАБЛОНА FormMAG02.xls',0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  USE IN aisoms
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  USE IN aisoms
  USE IN sprlpu
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\pilot', "pilot", "shar", "lpu_id")>0
  USE IN tarif
  USE IN aisoms
  USE IN sprlpu
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\pilots', "pilots", "shar", "lpu_id")>0
  USE IN pilot
  USE IN tarif
  USE IN aisoms
  USE IN sprlpu
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\lputpn', "lputpn", "shar", "lpu_id")>0
  USE IN pilots
  USE IN pilot
  USE IN tarif
  USE IN aisoms
  USE IN sprlpu
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\horlpu', "horlpu", "shar", "lpu_id")>0
  USE IN lputpn
  USE IN pilots
  USE IN pilot
  USE IN tarif
  USE IN aisoms
  USE IN sprlpu
  IF USED('horlpu')
   USE IN horlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\horlpus', "horlpus", "shar", "lpu_id")>0
  USE IN horlpu
  USE IN lputpn
  USE IN pilots
  USE IN pilot
  USE IN tarif
  USE IN aisoms
  USE IN sprlpu
  IF USED('horlpus')
   USE IN horlpus
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdata (nrec n(7), lpuid n(4), mcod c(7), tpn c(1), tpns c(1), lpuname c(120), ;
 	col06 n(11,2), col07 n(11,2), col08 n(11,2), col09 n(11,2), col10 n(11,2), col11 n(11,2), ;
 	col12 n(11,2), col13 n(11,2), col14 n(11,2), col15 n(11,2), col16 n(11,2), col17 n(11,2),;
 	col18 n(11,2), col19 n(11,2), col20 n(11,2), col21 n(11,2), col22 n(11,2))
 SELECT curdata 
 INDEX on lpuid TAG lpuid
 INDEX on mcod TAG mcod 
 
 m.nrec = 0 

 SELECT aisoms
 SCAN 
  m.lpuid   = lpuid
  m.mcod    = mcod
  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
  m.tpn     = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.tpn), '')
  m.tpns    = IIF(SEEK(m.mcod, 'sprlpu'), IIF(!EMPTY(FIELD('tpns', 'sprlpu')), ALLTRIM(sprlpu.tpns), ''), '')
  
  WAIT m.mcod + '...' WINDOW NOWAIT 
  
  ** МЭК
  m.col10 = sum_flk
  ** МЭК

  IF USED('pr4')
   IF SEEK(m.lpuid, 'pr4')
    m.udsum = pr4.s_others
    m.koplpf=m.finval-pr4.s_others+pr4.s_guests+pr4.s_npilot+pr4.s_empty
   ENDIF 
  ENDIF 
  IF USED('pr4st')
   IF SEEK(m.lpuid, 'pr4st')
    m.udsums = pr4st.s_others
    m.koplpfs = m.finvals-pr4st.s_others+pr4st.s_guests+pr4st.s_npilot+pr4st.s_empty
   ENDIF 
  ENDIF 
 
  m.IsPilot  = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)
  m.IsPilotS = IIF(SEEK(m.lpuid, 'pilots'), .T., .F.)
  m.IsHorS   = IIF(SEEK(m.lpuid, 'horlpus'), .T., .F.)
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn', 'lpu_id'), .t., .f.)	

  m.IsStomat   = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
  m.IsIskl     = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)

  m.col06 = 0
  m.col07 = 0
  m.col08 = 0
  m.col09 = 0
  m.col11 = 0
  m.col12 = 0
  m.col13 = 0
  m.col14 = 0
  m.col15 = 0
  m.col16 = 0
  m.col17 = 0
  m.col18 = 0
  m.col19 = 0
  m.col20 = 0
  m.col21 = 0
  m.col22 = 0
  
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
*MESSAGEBOX(m.mcod,0+64,'')
  m.nrec = m.nrec + 1

  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   INSERT INTO curdata FROM MEMVAR 
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   INSERT INTO curdata FROM MEMVAR 
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   INSERT INTO curdata FROM MEMVAR 
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
   USE IN talon 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'serror', 'shar', 'rid')>0
   USE IN talon 
   USE IN people
   IF USED('serror')
    USE IN serror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\hosp.dbf')
   IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\hosp', 'hosp', 'shar', 'c_i')>0
    IF USED('hosp')
     USE IN hosp 
    ENDIF 
   ENDIF 
  ENDIF 
 
  DIMENSION dimdata(9,11)
  dimdata = 0 
 
  CREATE CURSOR paz1 (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz2 (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz3 (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz1ok (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz2ok (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz3ok (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
 
  CREATE CURSOR paz1st (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz2st (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz3st (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz1stok (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz2stok (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol
  CREATE CURSOR paz3stok (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol 
  SET ORDER TO sn_pol

  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO serror ADDITIVE 
  m.st_flk = 0
  SCAN 

  SCATTER MEMVAR 
  m.cod       = cod
  m.ds        = ds
  m.sn_pol    = sn_pol
  *m.IsErr     = IIF(!EMPTY(serror.rid), .T., .F.)
  m.IsErr     = .F.
  m.prmcod    = people.prmcod
  m.prmcods   = people.prmcods

  m.s_all     = s_all 
  m.s_lek     = s_lek
  m.s_lekok   = IIF(EMPTY(serror.rid), 0, m.s_lek)
  m.rslt      = rslt
  m.fil_id    = fil_id
  m.otd       = SUBSTR(otd,2,2)
  m.proff     = SUBSTR(otd,4,3) && профиль услуги
  m.d_type    = d_type 
  m.lpu_ord   = lpu_ord
  m.ord       = ord
  
  *m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)
  m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)

  IF m.IsLpuTpn=.t.
   m.IsUslTpn = IIF(SEEK(m.fil_id, 'lputpn', 'fil_id'), .t., .f.)
  ELSE 
   m.IsUslTpn = .f.
  ENDIF 

  m.IsUslGosp = .F.
  IF USED('hosp')
   m.IsUslGosp = IIF(IsUsl(m.cod) AND SEEK(m.c_i, 'hosp'), .T., .F.)
  ENDIF 

  m.Is02      = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
  
  *m.prmcod = IIF(m.mcod!='0344704', m.prmcod, '0344704')
  
  IF !m.IsPilotS AND !m.IsHorS AND !INLIST(m.mcod, '4344623','4344700','4344621','4344640','4344613','4134752','0343036','0244124') && С 01.02.2019 
   *MESSAGEBOX('МЫ ЗДЕСЬ (!m.IsPilotS AND !m.IsHorS)!',0+64,m.mcod)
  DO CASE 
   CASE EMPTY(m.prmcod) && неприкрепленные
    dimdata(3,2)=dimdata(3,2)+1
    dimdata(3,3)=dimdata(3,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz3')
     INSERT INTO paz3 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz3ok')
      INSERT INTO paz3ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) OR BETWEEN(m.cod,97107,97999)
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
     CASE m.IsTpnR = .T. OR (m.IsPilot AND INLIST(m.otd,'08'))
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'

     *CASE (IsMes(m.cod) AND !BETWEEN(m.cod,97107,97999)) OR IsKdS(m.cod) OR IsVMP(m.cod)
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(3,8) = dimdata(3,8) + IIF(m.IsErr,0,m.s_all)
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'
      
     CASE m.IsUslGosp
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     
     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),49,149) AND people.mcod!=people.prmcod AND people.tip_p=3 
      *dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза
     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),29,129) AND people.mcod!=people.prmcod AND people.tip_p=3 
      *dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза
     
     OTHERWISE 
       dimdata(3,5) = dimdata(3,5) + IIF(m.IsErr,0,m.s_all)
       *m.col14 = m.col14 + IIF(m.IsErr and IsPilot , m.s_all, 0)
       **m.col14 = m.col14 + IIF(m.IsErr AND Typ='p' , m.s_all, 0)
       IF m.Is02
        dimdata(3,7) = dimdata(3,7) + IIF(m.IsErr and IsPilot ,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(3,10) = dimdata(3,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(3,9) = dimdata(3,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
   
   CASE m.mcod  = m.prmcod && свои пациенты
    dimdata(1,2)=dimdata(1,2)+1
    dimdata(1,3)=dimdata(1,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz1')
     INSERT INTO paz1 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz1ok')
      INSERT INTO paz1ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 

     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) OR BETWEEN(m.cod,97107,97999)
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
     CASE m.IsTpnR = .T. OR (m.IsPilot AND INLIST(m.otd,'08'))
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'

     *CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(1,8) = dimdata(1,8) + IIF(m.IsErr,0,m.s_all)
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'
     
     CASE m.otd='93' AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'
     
     OTHERWISE 
       dimdata(1,5) = dimdata(1,5) + IIF(m.IsErr and IsPilot,0,m.s_all)
       *m.col14 = m.col14 + IIF(m.IsErr and IsPilot, m.s_all, 0)
       **m.col14 = m.col14 + IIF(m.IsErr AND Typ='p', m.s_all, 0)
       IF m.Is02
        dimdata(1,7) = dimdata(1,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(1,10) = dimdata(1,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(1,9) = dimdata(1,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
    
   CASE m.mcod != m.prmcod && чужие пациенты
    dimdata(2,2)=dimdata(2,2)+1
    dimdata(2,3)=dimdata(2,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz2')
     INSERT INTO paz2 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz2ok')
      INSERT INTO paz2ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 

     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) OR BETWEEN(m.cod,97107,97999)
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
     CASE m.IsTpnR = .T. OR (m.IsPilot AND INLIST(m.otd,'08'))
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'

     *CASE (IsMes(m.cod) AND !BETWEEN(m.cod,97107,97999)) OR IsKdS(m.cod) OR IsVMP(m.cod)
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(2,8) = dimdata(2,8) + IIF(m.IsErr,0,m.s_all)
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'

     CASE m.IsUslGosp
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!

     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),49,149) AND people.mcod!=people.prmcod AND people.tip_p=3 
      *dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза
     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),29,129) AND people.mcod!=people.prmcod AND people.tip_p=3 
      *dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза

     OTHERWISE 
	    
       dimdata(2,5) = dimdata(2,5) + IIF(m.IsErr,0,m.s_all)
       *m.col14 = m.col14 + IIF(m.IsErr and IsPilot, m.s_all, 0)
       **m.col14 = m.col14 + IIF(m.IsErr AND Typ='p', m.s_all, 0)

       IF m.Is02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
       ELSE 
        IF m.lpu_ord>0
         dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        ENDIF 
       ENDIF 

    ENDCASE 
    dimdata(2,10) = dimdata(2,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(2,9) = dimdata(2,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 

   OTHERWISE 

  ENDCASE 

  ELSE && IF !m.IsPilotS AND !m.IsHorS
  
   *MESSAGEBOX('МЫ ЗДЕСЬ !(!m.IsPilotS AND !m.IsHorS)!',0+64,m.mcod)
  m.test = 0 

  m.UslIskl      = IIF(FLOOR(m.cod/1000)=146, .T., .F.)
  m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
  m.IsStomatUsl2 = IIF(INLIST(m.cod,1101,1102,101171,101172), .T., .F.)
   
  IF ((m.IsStomat AND !m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2)) OR ;
  	 ((m.IsStomat AND m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2 OR m.UslIskl)) OR ;
  	  (!m.IsStomat AND (m.IsStomatUsl OR (m.IsStomatUsl2 AND LEFT(m.ds,2)='K0')))
  
  m.st_flk = m.st_flk + IIF(m.IsErr,m.s_all,0)
  
  DO CASE 
   CASE EMPTY(m.prmcods) && неприкрепленные
    dimdata(7,2)=dimdata(7,2)+1
    dimdata(7,3)=dimdata(7,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz3st')
     INSERT INTO paz3st FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz3stok')
      INSERT INTO paz3stok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 
     * Закомментировано 27.05.2019 под новый протокол
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR INLIST(m.otd,'08') OR BETWEEN(m.cod,97107,97999)
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * *REPLACE Mp WITH '4'

     *CASE (IsMes(m.cod) AND !BETWEEN(m.cod,97107,97999)) OR IsKdS(m.cod) OR IsVMP(m.cod)
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(7,8) = dimdata(7,8) + IIF(m.IsErr,0,m.s_all)

     * Закомментировано 27.05.2019 под новый протокол
     *CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * *REPLACE Mp WITH '4'
     
     * Закомментировано 27.05.2019 под новый протокол
     *CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * *REPLACE Mp WITH '4'
     
     * Закомментировано 27.05.2019 под новый протокол
     *CASE m.ord=7 AND m.lpu_ord=7665
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * *REPLACE Mp WITH '4'
     
     ** Добавлено 16.04.2019 по требованию Согаза
     * Закомментировано 27.05.2019 под новый протокол
     *CASE INLIST(INT(m.cod/1000),49,149) AND people.mcod!=people.prmcod AND people.tip_p=3 
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза
     ** Добавлено 16.04.2019 по требованию Согаза
     * Закомментировано 27.05.2019 под новый протокол
     *CASE INLIST(INT(m.cod/1000),29,129) AND people.mcod!=people.prmcod AND people.tip_p=3 
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза

     OTHERWISE 
       dimdata(7,5) = dimdata(7,5) + IIF(m.IsErr,0,m.s_all)
*       m.col15 = m.col15 + IIF(m.IsErr, m.s_all, 0) & !! 25.02.2019
       m.col16 = m.col16 + IIF(m.IsErr, m.s_all, 0)
       IF m.Is02
        dimdata(7,7) = dimdata(7,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(7,10) = dimdata(7,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(7,9) = dimdata(7,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
   
   CASE m.mcod  = m.prmcods && свои пациенты
    dimdata(5,2)=dimdata(5,2)+1
    dimdata(5,3)=dimdata(5,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz1st')
     INSERT INTO paz1st FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz1stok')
      INSERT INTO paz1stok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 

     *CASE m.IsTpnR = .T. OR m.d_type='s' OR INLIST(m.otd,'08') OR BETWEEN(m.cod,97107,97999)
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) && INLIST(m.otd,'08')
     CASE m.IsTpnR = .T. OR (m.IsPilot AND INLIST(m.otd,'08')) && INLIST(m.otd,'08')
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *REPLACE Mp WITH '4'

     *CASE (IsMes(m.cod) AND !BETWEEN(m.cod,97107,97999)) OR IsKdS(m.cod) OR IsVMP(m.cod)
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(5,8) = dimdata(5,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *REPLACE Mp WITH '4'
     
     CASE m.otd='93' AND IsStac(m.mcod)
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *REPLACE Mp WITH '4'
     
     OTHERWISE 
       dimdata(5,5) = dimdata(5,5) + IIF(m.IsErr,0,m.s_all)
*       m.col15 = m.col15 + IIF(m.IsErr, m.s_all, 0) & !! 25.02.2019
       IF m.Is02
        dimdata(5,7) = dimdata(5,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(5,10) = dimdata(5,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(5,9) = dimdata(5,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
    
   CASE m.mcod != m.prmcods && чужие пациенты
    dimdata(6,2)=dimdata(6,2)+1
    dimdata(6,3)=dimdata(6,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz2st')
     INSERT INTO paz2st FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz2stok')
      INSERT INTO paz2stok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 

     *CASE m.IsTpnR = .T. OR m.d_type='s' OR INLIST(m.otd,'08')
     * Закомментировано 27.05.2019 под новый протокол
     *CASE m.IsTpnR = .T. OR BETWEEN(m.cod,97107,97999)
     * dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * *REPLACE Mp WITH '4'

     *CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
     * dimdata(6,8) = dimdata(6,8) + IIF(m.IsErr,0,m.s_all)

     *CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
     * dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * REPLACE Mp WITH '4'
     
     *CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
     * dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * REPLACE Mp WITH '4'
     
     *CASE m.ord=7 AND m.lpu_ord=7665
     * dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * REPLACE Mp WITH '4'

     OTHERWISE 
	    
       dimdata(6,5) = dimdata(6,5) + IIF(m.IsErr,0,m.s_all)
*       m.col15 = m.col15 + IIF(m.IsErr, m.s_all, 0)  & !! 25.02.2019

       IF m.Is02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))
        dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
       ELSE 
        IF m.lpu_ord>0
         dimdata(6,6) = dimdata(6,6) + IIF(m.IsErr,0,m.s_all)
        ENDIF 
       ENDIF 

    ENDCASE 
    dimdata(6,10) = dimdata(6,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(6,9) = dimdata(6,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 

   OTHERWISE 

  ENDCASE 

  ELSE && IF ((m.IsStomat AND !m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2)) OR ;
  	   && ((m.IsStomat AND m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2 OR m.IsIskl)) OR ;
  	   && (!m.IsStomat AND (m.IsStomatUsl OR (m.IsStomatUsl2 AND LEFT(m.ds,2)='K0')))

  DO CASE 
   CASE EMPTY(m.prmcod) && неприкрепленные
    dimdata(3,2)=dimdata(3,2)+1
    dimdata(3,3)=dimdata(3,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz3')
     INSERT INTO paz3 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz3ok')
      INSERT INTO paz3ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR INLIST(m.otd,'08') OR BETWEEN(m.cod,97107,97999)
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) && INLIST(m.otd,'08')
     CASE m.IsTpnR = .T. OR (m.IsPilot AND INLIST(m.otd,'08')) && INLIST(m.otd,'08')
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'

     *CASE (IsMes(m.cod) AND !BETWEEN(m.cod,97107,97999)) OR IsKdS(m.cod) OR IsVMP(m.cod)
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(3,8) = dimdata(3,8) + IIF(m.IsErr,0,m.s_all)
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'
     
     CASE m.IsUslGosp
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!

     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),49,149) AND people.mcod!=people.prmcod AND people.tip_p=3 
      *dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза
     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),29,129) AND people.mcod!=people.prmcod AND people.tip_p=3 
      *dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза

     OTHERWISE 
       dimdata(3,5) = dimdata(3,5) + IIF(m.IsErr,0,m.s_all)
       *m.col14 = m.col14 + IIF(m.IsErr and m.IsPilot, m.s_all, 0)
       **m.col14 = m.col14 + IIF(m.IsErr AND Typ='p', m.s_all, 0)

       *m.col17 = m.col17 + IIF(m.IsErr and !IsPilot, m.s_all, 0)
       IF m.Is02
        dimdata(3,7) = dimdata(3,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(3,10) = dimdata(3,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(3,9) = dimdata(3,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
   
   CASE m.mcod  = m.prmcod && свои пациенты
    dimdata(1,2)=dimdata(1,2)+1
    dimdata(1,3)=dimdata(1,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz1')
     INSERT INTO paz1 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz1ok')
      INSERT INTO paz1ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 

     *CASE m.IsTpnR = .T. OR m.d_type='s' OR INLIST(m.otd,'08') OR BETWEEN(m.cod,97107,97999)
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) && INLIST(m.otd,'08')
     CASE m.IsTpnR = .T. OR (m.IsPilot AND INLIST(m.otd,'08')) && INLIST(m.otd,'08')
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'

     *CASE (IsMes(m.cod) AND !BETWEEN(m.cod,97107,97999)) OR IsKdS(m.cod) OR IsVMP(m.cod)
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(1,8) = dimdata(1,8) + IIF(m.IsErr,0,m.s_all)
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'
     
     CASE m.otd='93' AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'
     
     OTHERWISE 
       dimdata(1,5) = dimdata(1,5) + IIF(m.IsErr,0,m.s_all)
       *m.col14 = m.col14 + IIF(m.IsErr and m.IsPilot, m.s_all, 0)
       **m.col14 = m.col14 + IIF(m.IsErr AND Typ='p', m.s_all, 0)
       *m.col17 = m.col17 + IIF(m.IsErr and !IsPilot, m.s_all, 0)
*       m.col14 = m.col14 + IIF(m.IsErr, m.s_all, 0)
       IF m.Is02
        dimdata(1,7) = dimdata(1,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(1,10) = dimdata(1,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(1,9) = dimdata(1,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
    
   CASE m.mcod != m.prmcod && чужие пациенты
    dimdata(2,2)=dimdata(2,2)+1
    dimdata(2,3)=dimdata(2,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz2')
     INSERT INTO paz2 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz2ok')
      INSERT INTO paz2ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 

     *CASE m.IsTpnR = .T. OR m.d_type='s' OR INLIST(m.otd,'08') OR BETWEEN(m.cod,97107,97999)
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) && INLIST(m.otd,'08')
     CASE m.IsTpnR = .T. OR (m.IsPilot AND INLIST(m.otd,'08')) && INLIST(m.otd,'08')
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'

     *CASE (IsMes(m.cod) AND !BETWEEN(m.cod,97107,97999)) OR IsKdS(m.cod) OR IsVMP(m.cod)
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(2,8) = dimdata(2,8) + IIF(m.IsErr,0,m.s_all)
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.col17 = m.col17 + IIF(m.IsErr, m.s_all, 0)
      *REPLACE Mp WITH '8'

     CASE m.IsUslGosp
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!

     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),49,149) AND people.mcod!=people.prmcod AND people.tip_p=3 
      *dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза
     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),29,129) AND people.mcod!=people.prmcod AND people.tip_p=3 
      *dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     ** Добавлено 16.04.2019 по требованию Согаза

     OTHERWISE 
	    
       dimdata(2,5) = dimdata(2,5) + IIF(m.IsErr,0,m.s_all)
       *m.col14 = m.col14 + IIF(m.IsErr and m.IsPilot, m.s_all, 0)
       **m.col14 = m.col14 + IIF(m.IsErr AND Typ='p', m.s_all, 0)
       *m.col17 = m.col17 + IIF(m.IsErr and !IsPilot, m.s_all, 0)
*       m.col14 = m.col14 + IIF(m.IsErr, m.s_all, 0)

       IF m.Is02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
       ELSE 
        IF m.lpu_ord>0
         dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        ENDIF 
       ENDIF 

    ENDCASE 
    dimdata(2,10) = dimdata(2,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(2,9) = dimdata(2,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 

   OTHERWISE 

  ENDCASE 

  ENDIF 

  ENDIF 
  
  *IF Typ='p'
  * IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)
  *  IF m.IsPilot OR m.IsPilots
  *  ELSE IF m.IsPilot OR m.IsPilots
  *  ENDIF 
  * ELSE IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)
  *  IF m.IsPilot OR m.IsPilots
  *   m.col14 = m.col14 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Typ='p' AND !IsDental(m.cod, m.lpuid, m.mcod, m.ds), m.s_all, 0)
  *  ELSE IF m.IsPilot OR m.IsPilots
  *  ENDIF 
  * ENDIF 
  *ELSE 
  * IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)
  *  IF m.IsPilot OR m.IsPilots
  *  ELSE IF m.IsPilot OR m.IsPilots
  *  ENDIF 
  * ELSE IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)
  *  IF m.IsPilot OR m.IsPilots
  *  ELSE IF m.IsPilot OR m.IsPilots
  *  ENDIF 
  * ENDIF 
  *ENDIF 
  
  *IF m.IsPilot OR m.IsPilots
   *m.col14 = m.col14 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp='p' AND !IsDental(m.cod, m.lpuid, m.mcod, m.ds), m.s_all, 0)
   m.col14 = m.col14 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp='p', m.s_all, 0)
  *ENDIF 
  *IF m.IsPilots
   *m.col15 = m.col15 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp='p' AND IsDental(m.cod, m.lpuid, m.mcod, m.ds) , m.s_all, 0)
   m.col15 = m.col15 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp='s', m.s_all, 0)
   *m.col16 = m.col16 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp='p' AND IsDental(m.cod, m.lpuid, m.mcod, m.ds) AND EMPTY(m.prmcods), m.s_all, 0)
   m.col16 = m.col16 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp='s' AND EMPTY(m.prmcods), m.s_all, 0)
  *ENDIF 
  
  *m.col19 = m.col19 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp!='p' AND IsDental(m.cod, m.lpuid, m.mcod, m.ds) , m.s_all, 0)
  *m.col19 = m.col19 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp!='s' AND IsDental(m.cod, m.lpuid, m.mcod, m.ds) , m.s_all, 0)
  m.col19 = m.col19 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp='8', m.s_all, 0)

  *m.col17 = m.col17 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp!='p' AND !IsDental(m.cod, m.lpuid, m.mcod, m.ds), m.s_all, 0)
  *m.col17 = m.col17 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND Mp!='s' AND !IsDental(m.cod, m.lpuid, m.mcod, m.ds), m.s_all, 0)
  m.col17 = m.col17 + IIF(IIF(!EMPTY(serror.rid), .T., .F.) AND INLIST(Mp,' ','4','m'), m.s_all, 0)

  m.col21 = m.col21 + m.s_lek
  m.col22 = m.col22 + m.s_lekok

  ENDSCAN 
  SET RELATION OFF INTO serror
  SET RELATION OFF INTO people
  USE 
  USE IN people 
  USE IN serror
  IF USED('hosp')
   USE IN hosp 
  ENDIF 
  
  dimdata(1,1) = RECCOUNT('paz1')
  dimdata(2,1) = RECCOUNT('paz2')
  dimdata(3,1) = RECCOUNT('paz3')
  dimdata(1,4) = RECCOUNT('paz1ok')
  dimdata(2,4) = RECCOUNT('paz2ok')
  dimdata(3,4) = RECCOUNT('paz3ok')
 
  dimdata(5,1) = RECCOUNT('paz1st')
  dimdata(6,1) = RECCOUNT('paz2st')
  dimdata(7,1) = RECCOUNT('paz3st')
  dimdata(5,4) = RECCOUNT('paz1stok')
  dimdata(6,4) = RECCOUNT('paz2stok')
  dimdata(7,4) = RECCOUNT('paz3stok')

  dimdata(4,1) = dimdata(1,1) + dimdata(2,1)  + dimdata(3,1)
  dimdata(4,2) = dimdata(1,2) + dimdata(2,2)  + dimdata(3,2)
  dimdata(4,3) = dimdata(1,3) + dimdata(2,3)  + dimdata(3,3)
  dimdata(4,4) = dimdata(1,4) + dimdata(2,4)  + dimdata(3,4)
  dimdata(4,5) = dimdata(1,5) + dimdata(2,5)  + dimdata(3,5)
  dimdata(4,6) = dimdata(1,6) + dimdata(2,6)  + dimdata(3,6)
  dimdata(4,7) = dimdata(1,7) + dimdata(2,7)  + dimdata(3,7)
  dimdata(4,8) = dimdata(1,8) + dimdata(2,8)  + dimdata(3,8)
  dimdata(4,9) = dimdata(1,9) + dimdata(2,9)  + dimdata(3,9)
  dimdata(4,10)= dimdata(1,10)+ dimdata(2,10) + dimdata(3,10)
  dimdata(4,11)= dimdata(1,11)+ dimdata(2,11) + dimdata(3,11)
 
  dimdata(8,1) = dimdata(5,1) + dimdata(6,1)  + dimdata(7,1)
  dimdata(8,2) = dimdata(5,2) + dimdata(6,2)  + dimdata(7,2)
  dimdata(8,3) = dimdata(5,3) + dimdata(6,3)  + dimdata(7,3)
  dimdata(8,4) = dimdata(5,4) + dimdata(6,4)  + dimdata(7,4)
  dimdata(8,5) = dimdata(5,5) + dimdata(6,5)  + dimdata(7,5)
  dimdata(8,6) = dimdata(5,6) + dimdata(6,6)  + dimdata(7,6)
  dimdata(8,7) = dimdata(5,7) + dimdata(6,7)  + dimdata(7,7)
  dimdata(8,8) = dimdata(5,8) + dimdata(6,8)  + dimdata(7,8)
  dimdata(8,9) = dimdata(5,9) + dimdata(6,9)  + dimdata(7,9)
  dimdata(8,10)= dimdata(5,10)+ dimdata(6,10) + dimdata(7,10)
  dimdata(8,11)= dimdata(5,11)+ dimdata(6,11) + dimdata(7,11)
 
  dimdata(9,1) = dimdata(4,1) + dimdata(8,1)
  dimdata(9,2) = dimdata(4,2) + dimdata(8,2)
  dimdata(9,3) = dimdata(4,3) + dimdata(8,3)
  dimdata(9,4) = dimdata(4,4) + dimdata(8,4)
  dimdata(9,5) = dimdata(4,5) + dimdata(8,5)
  dimdata(9,6) = dimdata(4,6) + dimdata(8,6)
  dimdata(9,7) = dimdata(4,7) + dimdata(8,7)
  dimdata(9,8) = dimdata(4,8) + dimdata(8,8)
  dimdata(9,9) = dimdata(4,9) + dimdata(8,9)
  dimdata(9,10)= dimdata(4,10)+ dimdata(8,10)
  *dimdata(9,11)= dimdata(4,11)+ dimdata(8,11)
  dimdata(9,11)= dimdata(4,11)

  USE IN paz1
  USE IN paz2
  USE IN paz3
  USE IN paz1ok
  USE IN paz2ok
  USE IN paz3ok

  USE IN paz1st
  USE IN paz2st
  USE IN paz3st
  USE IN paz1stok
  USE IN paz2stok
  USE IN paz3stok
  
  m.col06 = dimdata(9,3) && !!!
  m.col07 = dimdata(9,8) && !!!
  m.col08 = dimdata(9,11) && !!!
  m.col18 = dimdata(8,11) && !!!
  IF m.qcod = 'I3'
   *m.col09 = dimdata(5,5)+ dimdata(6,5) + dimdata(7,5) + ;
   			 dimdata(5,11)+ dimdata(6,11) + dimdata(7,11) + ;
   			 dimdata(5,8)+ dimdata(6,8) + dimdata(7,8)
   m.col09 = dimdata(8,3) && !!!
  ELSE 
   m.col09 = dimdata(5,5)+ dimdata(6,5) + dimdata(7,5) + dimdata(5,8)+ dimdata(6,8) + dimdata(7,8) && !!!
  ENDIF 
  *m.col09 = dimdata(7,5) + dimdata(7,8) && Вернул, как было - всю стоматологию!
  m.col11 = m.col06 - m.col10
  * m.col12 = dimdata(5,5)+ dimdata(6,5) временно убрал 
  m.col13 = m.col13 + dimdata(7,3)
  
 **m.col17 = m.col10 - m.col14
  
  IF 1=2
  m.col14 = m.col14 + IIF(m.IsPilot, m.col17, 0)
  m.col17 = IIF(m.IsPilot, 0, m.col17)
  ENDIF 
  
  IF m.col09=0 AND m.col15>0 && вероятно ошибки PF
   m.col19 = m.col19 + m.col15
   m.col09 = m.col15
   m.col15 = 0
  ENDIF 
  
  IF !m.IsPilot AND m.col14>0
   m.col17 = m.col17 + m.col14
   m.col14 = 0
  ENDIF 
  
  m.col20 = dimdata(1,5)+dimdata(2,5)+dimdata(3,5)

  INSERT INTO curdata FROM MEMVAR 
  
  UPDATE aisoms SET pf_flk = m.col14 WHERE mcod = m.mcod
  
  SELECT aisoms

 ENDSCAN 
 USE IN aisoms 
 USE IN sprlpu 
 USE IN tarif 
 USE IN lputpn
 USE IN pilot
 USE IN pilots
 USE IN horlpu
 USE IN horlpus
 

 SELECT curdata 
 COPY TO &pbase\&gcperiod\FormMAG02 WITH cdx 

 m.llResult = X_Report(pTempl+'\FormMAG02.xls', pBase+'\'+m.gcperiod+'\FormMAG02.xls', .T.)

 USE 
 
 
RETURN 