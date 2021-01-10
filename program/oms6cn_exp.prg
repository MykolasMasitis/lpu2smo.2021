FUNCTION oms6cn(lcPath, IsVisible, IsQuit)
 
 tn_result = 0
 tn_result = tn_result + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx', 'sprcokr', 'shar', 'cokr')
 tn_result = tn_result + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shar', 'cod')
 IF tn_result > 0
  IF USED('sprcokr')
   USE IN sprcokr
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 
 
 m.lIsPr4 = .F.
 IF fso.FileExists(pbase+'\'+gcperiod+'\pr4.dbf')
  m.lIsPr4 = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\pr4', 'pr4', 'shar', 'lpuid')>0
   IF USED('pr4')
    USE IN pr4
   ENDIF 
  ELSE 
   m.lIsPr4 = .T.
  ENDIF 
 ENDIF 

 m.lIsPr4s = .F.
 IF fso.FileExists(pbase+'\'+gcperiod+'\pr4st.dbf')
  m.lIsPr4 = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\pr4st', 'pr4st', 'shar', 'lpuid')>0
   IF USED('pr4st')
    USE IN pr4st
   ENDIF 
  ELSE 
   m.lIsPr4st = .T.
  ENDIF 
 ENDIF 

 SELECT AisOms

 m.mcod       = mcod
 eeFile = 'e'+m.mcod

 tn_result = 0
 tn_result = tn_result + OpenFile(lcpath+'\Talon', 'talon', 'shar')
 tn_result = tn_result + OpenFile(lcpath+'\People', 'people', 'shar', 'sn_pol')
 tn_result = tn_result + OpenFile(lcpath+'\'+eeFile, 'serror', 'shar', 'rid')
 IF fso.FileExists(lcpath+'\hosp.dbf')
  IF OpenFile(lcpath+'\hosp', 'hosp', 'shar', 'c_i')>0
   IF USED('hosp')
    USE IN hosp 
   ENDIF 
  ENDIF 
 ENDIF 
 IF tn_result>0
  IF USED('sprcokr')
   USE IN sprcokr
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('serror')
   USE IN serror
  ENDIF 
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  IF USED('pr4st')
   USE IN pr4st
  ENDIF 
  IF USED('hosp')
   USE IN hosp 
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT aisoms

 m.mmy        = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 m.lpuid      = lpuid
 *m.IsStomat   = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
 *m.IsIskl     = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)
 m.lpuname    = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.cokr       = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.cokr), '')
 m.cokr_name  = IIF(SEEK(m.cokr, 'sprcokr'), ALLTRIM(sprcokr.name_okr), '')
 *m.smoname    = IIF(SEEK(m.qcod, 'smo'), ALLTRIM(smo.fullname), '')
 m.smoname    = m.qname
 m.arcfname   = 'b'+m.mcod+'.'+m.mmy
 m.datpriemki = TTOC(Recieved)
 m.finval     = finval
 m.finvals    = finvals
 m.udsum      = 0
 m.udsums     = 0
 m.koplate    = 0 
 m.koplates   = 0 
 m.koplpf     = 0
 m.koplpfs    = 0
 m.s_532      = s_532
 m.cmessage   = ALLTRIM(cmessage)

 IF USED('pr4')
  IF SEEK(m.lpuid, 'pr4')
   m.udsum = pr4.s_others
   m.koplpf=m.finval-pr4.s_others+pr4.s_guests+pr4.s_npilot+pr4.s_empty
  ENDIF 
 ENDIF 
 IF USED('pr4st')
  IF SEEK(m.lpuid, 'pr4st')
   m.udsums = pr4st.s_others
   *m.koplpfs = m.finvals-pr4st.s_others+pr4st.s_guests+pr4st.s_npilot+pr4st.s_empty
   m.koplpfs = m.finvals-pr4st.s_npilot+pr4st.s_empty && 19.11.2018!
  ENDIF 
 ENDIF 
 
 m.IsPilot    = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)
 m.IsPilotS   = IIF(SEEK(m.lpuid, 'pilots'), .T., .F.)
 m.IsHor      = IIF(SEEK(m.lpuid, 'horlpu'), .T., .F.)
 m.IsHorS     = IIF(SEEK(m.lpuid, 'horlpus'), .T., .F.)

 m.IsStPilot  = IIF(SEEK(m.lpuid, 'stpilot'), .T., .F.) && Кончаловского
 m.IsSprNCO   = IIF(SEEK(m.lpuid, 'sprnco'), .T., .F.)
 m.IsIG       = IIF(SEEK(m.lpuid, 'sprnco') AND sprnco.ig=1, .T., .F.)
 m.d_b        = IIF(SEEK(m.lpuid, 'sprnco'), sprnco.date_b, {}) && {23.03.2020}
 
 *m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn', 'lpu_id'), .t., .f.)	
 
 *m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 
 m.arcfdate = DATETIME()
 IF fso.FileExists(lcPath + '\' + arcfname)
  poi_file   = fso.GetFile(lcPath + '\' + arcfname)
  m.arcfdate = poi_file.DateLastModified
 ENDIF 
 
 ZipItemCount = 5

 m.DotName = pTempl + "\Prqqmmy.xls"
 DocName = lcPath + "\Pr" + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)

 DIMENSION dimdata(9,13)
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

 SELECT Talon 
 SET RELATION TO sn_pol INTO people
 SET RELATION TO RecId  INTO sError ADDITIVE 
 SET RELATION TO cod    INTO tarif ADDITIVE 
 REPLACE ALL Mp WITH '', Typ WITH '', vz WITH 0, dop_r WITH 0
 * Mp = '4' - допуслуги терапия
 * Mp = '8' - допуслуги стоматология
 * Mp = 'p' - подушевые терапия
 * Mp = 's' - подушевые стоматология
 * Mp = 'm' - МЭСы
 m.st_flk = 0
 
 SCAN
  SCATTER MEMVAR 
  m.cod       = cod
  m.sn_pol    = sn_pol
  m.IsErr     = IIF(!EMPTY(serror.rid), .T., .F.)
  m.prmcod    = people.prmcod
  m.pr_id     = IIF(SEEK(people.prmcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
  m.prmcods   = people.prmcods
  m.IsStPr    = IIF(USED('stpilot') AND SEEK(m.pr_id, 'stpilot'), .T., .F.)

  m.s_all     = s_all + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
  m.rslt      = rslt
  m.fil_id    = fil_id
  m.otd       = SUBSTR(otd,2,2)
  m.proff     = SUBSTR(otd,4,3) && профиль услуги
  m.d_type    = d_type 
  m.lpu_ord   = lpu_ord
  m.ord       = ord
  
  *m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)
  m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
  m.IsTpnR    = IIF(INLIST(m.cod,28211,128211) AND m.IsSprNCO AND m.d_u>=m.d_b, .T., m.IsTpnR)

  m.IsTpnR    = IIF(INLIST(m.cod,28165,128165) AND m.IsIG AND m.d_u>=m.d_b, .T., m.IsTpnR) && со счетов за май.

  m.IsTpnR    = IIF(INLIST(m.cod,37043,137043) AND ;
  	(INLIST(m.ds, 'B34.2','J02','J04','J06','J20','U07.1','U07.2') OR ;
  	 BETWEEN(LEFT(m.ds,3),'J09','J18')), .T., m.IsTpnR)
  m.IsTpnR    = IIF(INLIST(m.cod,37043,37048,137043,37044,37049,137044,137049) AND ;
  	(m.ds='C' OR BETWEEN(LEFT(m.ds,3),'D00','D09')), .T., m.IsTpnR)
  m.IsTpnR    = IIF(INLIST(m.cod,60010,160010) AND m.IsSprNCO AND m.d_u>=m.d_b, .T., m.IsTpnR)
  
  m.Is02      = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
  
  m.Mp    = ''
  m.vz    = 0
  m.dop_r = 0
  m.PrCell = ''
  
  m.test = 0 
  
  m.IsUslGosp = .F.
  IF USED('hosp')
   m.IsUslGosp = IIF(IsUsl(m.cod) AND SEEK(m.c_i, 'hosp'), .T., .F.)
  ENDIF 
  
  *IF m.IsPilotS AND IsDental(m.cod, m.lpuid, m.mcod, m.ds)
  *IF IsDental(m.cod, m.lpuid, m.mcod, m.ds) ;
  	AND m.otd<>'08' AND !(m.ord=7 AND m.lpu_ord=7665) && из-за этого проблема в Ингоссе - потерялись УМО!!!
  IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)

  m.st_flk = m.st_flk + IIF(m.IsErr,m.s_all,0)
  
  DO CASE 
   CASE EMPTY(m.prmcods) && неприкрепленные
    m.Typ = '0'
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
     CASE m.IsTpnR = .T.
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 1
      m.PrCell = IIF(!m.IsErr, '609', m.PrCell)

     CASE INLIST(m.otd,'08')
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 3
      m.PrCell = IIF(!m.IsErr, '609', m.PrCell)

     * Закомментировано 27.05.2019 под новый протокол
     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod)
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 4
      m.PrCell = IIF(!m.IsErr, '609', m.PrCell)
     
     * Закомментировано 27.05.2019 под новый протокол
     *CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
     CASE INLIST(m.otd,'01') AND IsStac(m.mcod)
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 5
      m.PrCell = IIF(!m.IsErr, '609', m.PrCell)
     
     * Закомментировано 27.05.2019 под новый протокол
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 6
      m.PrCell = IIF(!m.IsErr, '609', m.PrCell)
     
     OTHERWISE 
       m.Mp = 's'
       dimdata(7,5) = dimdata(7,5) + IIF(m.IsErr,0,m.s_all)
       m.PrCell = IIF(!m.IsErr, '606', m.PrCell)
       *IF m.Is02 OR (m.profil='100' AND INLIST(m.otd,'00','92'))
       *IF m.Is02 OR (m.profil='100' AND INLIST(m.otd,'92'))
       IF m.Is02 OR INLIST(m.otd,'92')
        dimdata(7,7) = dimdata(7,7) + IIF(m.IsErr,0,m.s_all)
        *m.PrCell = IIF(!m.IsErr, '608', m.PrCell)
       ELSE 
        *m.PrCell = IIF(!m.IsErr, '607', m.PrCell)
       ENDIF 

    ENDCASE 
    dimdata(7,10) = dimdata(7,10) + IIF(m.IsErr,0,m.s_all)
    *m.PrCell = '612'
    *dimdata(7,12) = dimdata(7,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    *IF IsVMP(m.cod)
    * dimdata(7,9) = dimdata(7,9) + IIF(m.IsErr,0,m.s_all)
    *ENDIF 
   
   CASE m.mcod  = m.prmcods && свои пациенты
    m.Typ = '1'
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
 
     CASE m.IsTpnR = .T. && tpn='r' - 3 услуги по июлю 2019, 08 - 4
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 8
      m.Mp = '8'
      m.PrCell = IIF(!m.IsErr, '409', m.PrCell)

     CASE INLIST(m.otd,'08') && tpn='r' - 3 услуги по июлю 2019, 08 - 4
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 9
      m.Mp = '8'
      m.PrCell = IIF(!m.IsErr, '409', m.PrCell)

     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod) && 23 услуги
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 10
      m.Mp = '8'
      m.PrCell = IIF(!m.IsErr, '409', m.PrCell)
     
     OTHERWISE 
       m.Mp = 's'
       dimdata(5,5) = dimdata(5,5) + IIF(m.IsErr,0,m.s_all)
       m.PrCell = IIF(!m.IsErr, '406', m.PrCell)
       *IF m.Is02 OR (m.profil='100' AND INLIST(m.otd,'00','92'))
       *IF m.Is02 OR (m.profil='100' AND INLIST(m.otd,'92'))
       IF m.Is02 OR INLIST(m.otd,'92')
        dimdata(5,7) = dimdata(5,7) + IIF(m.IsErr,0,m.s_all)
        *m.PrCell = IIF(!m.IsErr, '408', m.PrCell)
       ELSE 
        *m.PrCell = IIF(!m.IsErr, '407', m.PrCell)
       ENDIF 

    ENDCASE 
    dimdata(5,10) = dimdata(5,10) + IIF(m.IsErr,0,m.s_all)
    *m.PrCell = '612'
    *dimdata(5,12) = dimdata(5,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    *IF IsVMP(m.cod)
    * dimdata(5,9) = dimdata(5,9) + IIF(m.IsErr,0,m.s_all)
    *ENDIF 
    
   CASE m.mcod != m.prmcods && чужие пациенты
    m.Typ = '2'
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
     * Закомментировано 27.05.2019 под новый протокол
     CASE m.IsTpnR = .T.
      dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 8
      m.PrCell = IIF(!m.IsErr, '509', m.PrCell)

     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod)
      dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 8
      m.PrCell = IIF(!m.IsErr, '509', m.PrCell)
     
     *CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
     CASE INLIST(m.otd,'01') AND IsStac(m.mcod)
      dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 8
      m.PrCell = IIF(!m.IsErr, '509', m.PrCell)
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
      m.dop_r = 8
      m.PrCell = IIF(!m.IsErr, '509', m.PrCell)

     OTHERWISE 
       * сюда вставить m.vz!	    
       m.Mp = 's'

       dimdata(6,5) = dimdata(6,5) + IIF(m.IsErr,0,m.s_all)

       m.pr_id = IIF(SEEK(people.prmcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
       DO CASE 
        CASE m.ord>0 AND m.lpu_ord>0 AND IIF(m.qcod<>'I3', m.lpu_ord=m.pr_id, .T.) && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
         dimdata(6,6) = dimdata(6,6) + IIF(m.IsErr,0,m.s_all)
         *m.vz = 1
         m.PrCell = IIF(!m.IsErr, '507', m.PrCell)
        CASE m.Is02 && vz=2, неотложная помощь (по реестру медицинских услуг)
         dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
         *m.vz = 2
         m.PrCell = IIF(!m.IsErr, '508', m.PrCell)
        *CASE m.profil='100' AND INLIST(m.otd,'00','92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
        *CASE m.profil='100' AND INLIST(m.otd,'92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
        CASE INLIST(m.otd,'92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
         dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
         *m.vz = 3
         m.PrCell = IIF(!m.IsErr, '508', m.PrCell)
        CASE m.otd='08' && vz=4, услуги ЖК
         dimdata(6,6) = dimdata(6,6) + IIF(m.IsErr,0,m.s_all)
         *m.vz = 4
         m.PrCell = IIF(!m.IsErr, '508', m.PrCell)
        CASE m.otd='91' && vz=5, услуги ЦЗ
         dimdata(6,6) = dimdata(6,6) + IIF(m.IsErr,0,m.s_all)
         *m.vz = 5
         m.PrCell = IIF(!m.IsErr, '508', m.PrCell)
        OTHERWISE 
         *m.vz = 0 && то, что должно попасть в up-файл
        m.PrCell = IIF(!m.IsErr, '506', m.PrCell)
       ENDCASE 

       *F m.Is02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))
       *dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
       *LSE 
       *IF m.lpu_ord>0
       * dimdata(6,6) = dimdata(6,6) + IIF(m.IsErr,0,m.s_all)
       * ENDIF 
       *ENDIF 

    ENDCASE 
    dimdata(6,10) = dimdata(6,10) + IIF(m.IsErr,0,m.s_all)
    *m.PrCell = '512'
    *dimdata(6,12) = dimdata(6,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    *IF IsVMP(m.cod)
    * dimdata(6,9) = dimdata(6,9) + IIF(m.IsErr,0,m.s_all)
    *ENDIF 

   OTHERWISE 

  ENDCASE 

  ELSE && Терапия, здесь же МЭС!

  DO CASE 
   CASE EMPTY(m.prmcod) && неприкрепленные
    m.Is02 = IIF(SEEK(m.cod, 'pervpr') AND m.p_cel='1.1', .T., .F.)
    
    m.Typ = '0'
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
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) && Добавление условия pilot ничего не меняет
     *CASE m.IsTpnR = .T. AND (m.IsPilot OR m.IsHor) && OR m.d_type='s' OR (INLIST(m.otd,'08')) && Добавление условия pilot ничего не меняет Это пиздец как поменяло!!!
     CASE m.IsTpnR = .T.
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
       m.dop_r = 1
       m.Mp    = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '309', m.PrCell)

     *CASE INLIST(m.otd,'08') AND (m.IsPilot OR m.IsHor)
     *CASE INLIST(m.otd,'08','85') AND (m.IsPilot OR m.IsHor)
     CASE INLIST(m.otd,'08','85')
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
       m.dop_r = 3
       m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '309', m.PrCell)

     CASE INLIST(m.cod,56029,156003)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
       m.dop_r = 3
       m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '309', m.PrCell)

     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(3,8) = dimdata(3,8) + IIF(m.IsErr,0,m.s_all)
      m.Mp = 'm'
      m.PrCell = IIF(!m.IsErr, '310', m.PrCell)

     CASE INLIST(m.otd,'01') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 5
      m.Mp = '4'
      m.PrCell = IIF(!m.IsErr, '309', m.PrCell)
      *REPLACE test WITH 3

     *CASE INLIST(m.otd,'70','73','90','93') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 4
      m.Mp = '4'
      m.PrCell = IIF(!m.IsErr, '309', m.PrCell)
      *REPLACE test WITH 2
     
     CASE m.ord=7 AND m.lpu_ord=7665 AND (m.IsPilot OR m.IsHor)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 6
      m.Mp = '4'
      m.PrCell = IIF(!m.IsErr, '309', m.PrCell)
      *REPLACE test WITH 4
     
     * Эксперимент!
     ** Добавлено 16.04.2019 по требованию Согаза
     *CASE INLIST(INT(m.cod/1000),49,149) AND people.tip_p=3 
     * dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * m.Mp = '4'
      *REPLACE test WITH 5

     *CASE INLIST(INT(m.cod/1000),29,129) AND people.tip_p=3 
     * dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * m.Mp = '4'
      *REPLACE test WITH 6
     ** Добавлено 16.04.2019 по требованию Согаза

     *CASE m.IsUslGosp AND (m.IsPilot OR m.IsHor) && IsUsl(m.cod) AND people.tip_p=3 
     CASE m.IsUslGosp
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      *m.Mp = '4'
      IF m.IsPilot OR m.IsHor
      m.dop_r = 2
      m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '309', m.PrCell)
     * Эксперимент!
     
     *CASE INLIST(m.cod, 37043,37044,37048,37049,137043,137044,137049) AND LEFT(m.ds,1)='C' AND (m.IsPilot OR m.IsHor) AND INT(VAL(m.gcperiod))>=202001
     CASE INLIST(m.cod, 37043,37044,37048,37049,137043,137044,137049) AND LEFT(m.ds,1)='C' AND INT(VAL(m.gcperiod))>=202001
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
      m.Mp    = '4'
      m.dop_r = 11
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '309', m.PrCell)
     
     OTHERWISE 
       m.Mp = 'p'
       dimdata(3,5) = dimdata(3,5) + IIF(m.IsErr,0,m.s_all)
       m.PrCell = IIF(!m.IsErr, '306', m.PrCell)
       *IF m.Is02 OR (m.profil='100' AND INLIST(m.otd,'00','92'))
       *IF m.Is02 OR (m.profil='100' AND INLIST(m.otd,'92'))
       * IF m.Is02 OR INLIST(m.otd,'92')
       IF m.Is02
        dimdata(3,7) = dimdata(3,7) + IIF(m.IsErr,0,m.s_all)
        *m.PrCell = IIF(!m.IsErr, '307', m.PrCell)
       ELSE 
        *m.PrCell = IIF(!m.IsErr, '308', m.PrCell)
       ENDIF 

    ENDCASE 
    dimdata(3,10) = dimdata(3,10) + IIF(m.IsErr,0,m.s_all)
    dimdata(3,12) = dimdata(3,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    dimdata(3,13) = dimdata(3,13) + IIF(FIELD('s_lek')='S_LEK', IIF(m.IsErr,0,s_lek), 0)
    IF IsVMP(m.cod)
     dimdata(3,9) = dimdata(3,9) + IIF(m.IsErr,0,m.s_all)
     m.PrCell = IIF(!m.IsErr, '311', m.PrCell)
    ENDIF 
   
   CASE m.mcod  = m.prmcod && свои пациенты
    m.Is02 = IIF(m.p_cel='1.1', .T., .F.)

    m.Typ = '1'
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

     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
     *CASE m.IsTpnR = .T. AND (m.IsPilot OR m.IsHor) && OR m.d_type='s' OR (INLIST(m.otd,'08')) && нельзя!!
     CASE m.IsTpnR = .T.
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
      m.dop_r = 1
      m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '109', m.PrCell)

     *CASE INLIST(m.otd,'08') AND (m.IsPilot OR m.IsHor)
     *CASE INLIST(m.otd,'08','85') AND (m.IsPilot OR m.IsHor)
     CASE INLIST(m.otd,'08','85')
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
       m.dop_r = 3
       m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '109', m.PrCell)

     CASE INLIST(m.cod,56029,156003)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
       m.dop_r = 3
       m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '109', m.PrCell)

     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(1,8) = dimdata(1,8) + IIF(m.IsErr,0,m.s_all)
      IF m.IsStPilot && AND !SEEK(m.cod, 'novzms')
       m.Mp = 'p'
      ELSE 
       m.Mp = 'm'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '110', m.PrCell)

     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 4
      m.Mp = '4'
      m.PrCell = IIF(!m.IsErr, '109', m.PrCell)
     
     * Эксперимент!
     *CASE m.IsUslGosp AND (m.IsPilot OR m.IsHor) && IsUsl(m.cod) AND people.tip_p=3 
     CASE m.IsUslGosp 
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
      m.dop_r = 2
      *m.Mp = '4'
      m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '109', m.PrCell)
     * Эксперимент!

     *CASE INLIST(m.cod, 37043,37044,37048,37049,137043,137044,137049) AND (m.IsPilot OR m.IsHor) AND LEFT(m.ds,1)='C' AND INT(VAL(m.gcperiod))>=202001
     CASE INLIST(m.cod, 37043,37044,37048,37049,137043,137044,137049) AND LEFT(m.ds,1)='C' AND INT(VAL(m.gcperiod))>=202001
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
      m.Mp    = '4'
      m.dop_r = 11
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '109', m.PrCell)

     OTHERWISE 
       m.Mp = 'p'
       dimdata(1,5) = dimdata(1,5) + IIF(m.IsErr,0,m.s_all)
       m.PrCell = IIF(!m.IsErr, '106', m.PrCell)
       *IF m.Is02 OR (m.profil='100' AND INLIST(m.otd,'00','92'))
       *IF m.Is02 OR (m.profil='100' AND INLIST(m.otd,'92'))
       *IF m.Is02 OR INLIST(m.otd,'92')
       IF m.Is02
        dimdata(1,7) = dimdata(1,7) + IIF(m.IsErr,0,m.s_all)
        *m.PrCell = IIF(!m.IsErr, '108', m.PrCell)
       ELSE 
        *m.PrCell = IIF(!m.IsErr, '107', m.PrCell)
       ENDIF 

    ENDCASE 
    dimdata(1,10) = dimdata(1,10) + IIF(m.IsErr,0,m.s_all)
    dimdata(1,12) = dimdata(1,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    dimdata(1,13) = dimdata(1,13) + IIF(FIELD('s_lek')='S_LEK', IIF(m.IsErr,0,s_lek), 0)
    IF IsVMP(m.cod)
     dimdata(1,9) = dimdata(1,9) + IIF(m.IsErr,0,m.s_all)
     m.PrCell = IIF(!m.IsErr, '111', m.PrCell)
    ENDIF 
    
   CASE m.mcod != m.prmcod && чужие пациенты
    *m.Is02 = IIF((SEEK(m.cod, 'tarif') AND tarif.tpn='q') AND m.p_cel='1.1', .T., .F.)
    m.Is02 = IIF((SEEK(m.cod, 'tarif') AND tarif.tpn='q'), .T., .F.)

    m.Typ = '2'
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
    
     CASE IsStac(m.mcod) AND m.IsHor AND m.otd='00'
      m.Mp = 'p'
      dimdata(2,5) = dimdata(2,5) + IIF(m.IsErr,0,m.s_all)
      
      m.pr_id = IIF(SEEK(people.prmcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
      DO CASE 
       CASE (m.IsPilot OR m.IsHor) AND m.ord=8 AND m.lpu_ord=8888 && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
        dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        m.vz = 1
        m.PrCell = IIF(!m.IsErr, '207', m.PrCell)

       CASE (m.IsPilot OR m.IsHor) AND m.ord>0 AND m.lpu_ord>0 AND IIF(m.qcod<>'I3', m.lpu_ord=m.pr_id, .T.) && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
        dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        m.vz = 1
        m.PrCell = IIF(!m.IsErr, '207', m.PrCell)

       CASE (m.IsPilot OR m.IsHor) AND m.Is02 && vz=2, неотложная помощь (по реестру медицинских услуг)
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        m.vz = 2
        m.PrCell = IIF(!m.IsErr, '208', m.PrCell)

       OTHERWISE 
        m.dop_r = 12
        m.Mp = '4'
        m.PrCell = IIF(!m.IsErr, '209', m.PrCell)
      ENDCASE 

     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
     *CASE m.IsTpnR = .T. AND (m.IsPilot OR m.IsHor) && OR m.d_type='s' OR (INLIST(m.otd,'08')) && нельзя
     CASE m.IsTpnR = .T.
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
      m.dop_r = 1
      m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)

     *CASE INLIST(m.otd,'08') AND (m.IsPilot OR m.IsHor)
     *CASE INLIST(m.otd,'08','85') AND (m.IsPilot OR m.IsHor)
     CASE INLIST(m.otd,'08','85')
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
      m.dop_r = 3
      m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)

     CASE INLIST(m.cod,56029,156003)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
       m.dop_r = 3
       m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)

     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(2,8) = dimdata(2,8) + IIF(m.IsErr,0,m.s_all)
      IF m.IsStPr AND !SEEK(m.cod, 'novzms') AND (INLIST(m.ord,2,3) OR m.lpu_ord=1989)
       m.Mp = 'p'
       m.vz = 6
      ELSE 
       m.Mp = 'm'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '210', m.PrCell)

     CASE INLIST(m.otd,'01') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 5
      m.Mp = '4'
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)
     
     CASE INLIST(m.proff,'015','034') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor) && добавлено 14.01.2020 после сверки с ВТБ!
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 7
      m.Mp = '4'
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)
     
     *CASE INLIST(m.otd,'70','73','90','93') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 4
      m.Mp = '4'
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)
     
     CASE m.ord=7 AND m.lpu_ord=7665 AND (m.IsPilot OR m.IsHor)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.dop_r = 6
      m.Mp = '4'
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)
     
     * Эксперимент!
     ** Добавлено 16.04.2019 по требованию Согаза
     *CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=people.prmcod AND people.tip_p=3 
     * dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * m.Mp = '4'

     ** Добавлено 16.04.2019 по требованию Согаза
     *CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=people.prmcod AND people.tip_p=3 
     * dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * m.Mp = '4'
     ** Добавлено 16.04.2019 по требованию Согаза

     *CASE m.IsUslGosp AND (m.IsPilot OR m.IsHor) && IsUsl(m.cod) AND people.tip_p=3 
     CASE m.IsUslGosp
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
      *m.Mp = '4'
      m.dop_r = 2
      m.Mp = '4'
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)
     * Эксперимент!

     *CASE INLIST(m.cod, 37043,37044,37048,37049,137043,137044,137049) AND LEFT(m.ds,1)='C' AND (m.IsPilot OR m.IsHor) AND INT(VAL(m.gcperiod))>=202001
     CASE INLIST(m.cod, 37043,37044,37048,37049,137043,137044,137049) AND LEFT(m.ds,1)='C' AND INT(VAL(m.gcperiod))>=202001
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      IF m.IsPilot OR m.IsHor
      m.Mp    = '4'
      m.dop_r = 11
      ENDIF 
      m.PrCell = IIF(!m.IsErr, '209', m.PrCell)

     OTHERWISE 
	    
      m.Mp = 'p'
      dimdata(2,5) = dimdata(2,5) + IIF(m.IsErr,0,m.s_all)
      
      m.pr_id = IIF(SEEK(people.prmcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
      DO CASE 
       CASE (m.IsPilot OR m.IsHor) AND m.ord=8 AND m.lpu_ord=8888 && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
        dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        m.vz = 1
        m.PrCell = IIF(!m.IsErr, '207', m.PrCell)
       CASE (m.IsPilot OR m.IsHor) AND m.ord>0 AND m.lpu_ord>0 AND IIF(m.qcod<>'I3', m.lpu_ord=m.pr_id, .T.) && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
        dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        m.vz = 1
        m.PrCell = IIF(!m.IsErr, '207', m.PrCell)

       CASE (m.IsPilot OR m.IsHor) AND m.Is02 && vz=2, неотложная помощь (по реестру медицинских услуг)
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        m.vz = 2
        m.PrCell = IIF(!m.IsErr, '208', m.PrCell)

       *CASE m.profil='100' AND INLIST(m.otd,'00','92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
       *CASE m.profil='100' AND INLIST(m.otd,'92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)

       CASE (m.IsPilot OR m.IsHor) AND INLIST(m.otd,'92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        m.vz = 3
        m.PrCell = IIF(!m.IsErr, '208', m.PrCell)

       *CASE m.otd='08' && vz=4, услуги ЖК здесь может быть только не подушевые!! поменял 02.02.2020 по предложению Валентины
       CASE INLIST(m.otd,'08','85') && vz=4, услуги ЖК здесь может быть только не подушевые!! поменял 02.02.2020 по предложению Валентины
        *dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        m.vz = 4
        *m.PrCell = IIF(!m.IsErr, '207', m.PrCell)
        m.PrCell = IIF(!m.IsErr, '206', m.PrCell)

       CASE (m.IsPilot OR m.IsHor) AND m.otd='91' && vz=5, услуги ЦЗ
        * изменено со счетов за апрель!
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        *dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        * изменено со счетов за апрель!
        m.vz = 5
        *m.PrCell = IIF(!m.IsErr, '207', m.PrCell)
        m.PrCell = IIF(!m.IsErr, '208', m.PrCell) && почему-то к неотложке поменял 02.02.2020 по предложению Валентины

       OTHERWISE 
        m.PrCell = IIF(!m.IsErr, '206', m.PrCell)
        m.vz = 0 && то, что должно попасть в up-файл
      ENDCASE 

      *IF m.Is02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))
      * dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
      *ELSE 
      * IF m.lpu_ord>0
      *  dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
      * ENDIF 
      *ENDIF 

    ENDCASE 
    dimdata(2,10) = dimdata(2,10) + IIF(m.IsErr,0,m.s_all)
    dimdata(2,12) = dimdata(2,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    dimdata(2,13) = dimdata(2,13) + IIF(FIELD('s_lek')='S_LEK', IIF(m.IsErr,0,s_lek), 0)
    IF IsVMP(m.cod)
     dimdata(2,9) = dimdata(2,9) + IIF(m.IsErr,0,m.s_all)
     m.PrCell = IIF(!m.IsErr, '211', m.PrCell)
    ENDIF 

   OTHERWISE 

  ENDCASE 

  ENDIF IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)
  
  IF m.Mp='4' AND !m.IsPilot && у не пилотов допуслуг не должно быть!
   m.Mp    = ''
   m.vz    = 0 
   m.dop_r = 0
  ENDIF 
  IF m.Mp='8' AND !m.IsPilots && у не пилотов допуслуг не должно быть!
   m.Mp    = ''
   m.vz    = 0 
   m.dop_r = 0
  ENDIF 
  
  IF (m.IsPilot OR m.IsHor) AND INLIST(m.Mp,'4','p')
   REPLACE Mp WITH m.Mp, Typ WITH m.Typ, vz WITH m.vz, dop_r WITH m.dop_r
  ENDIF 
  IF m.IsPilotS AND INLIST(m.Mp,'8','s')
   REPLACE Mp WITH m.Mp, Typ WITH m.Typ, vz WITH m.vz, dop_r WITH m.dop_r
  ENDIF 
  IF m.IsStPr
   REPLACE Mp WITH m.Mp, Typ WITH m.Typ, vz WITH m.vz, dop_r WITH m.dop_r
  ENDIF 
  
  IF FIELD('prcell')='PRCELL'
   REPLACE PrCell WITH m.PrCell
  ENDIF 

 ENDSCAN  
 
 SET RELATION OFF INTO people
 SET RELATION OFF INTO sError
 SET RELATION OFF INTO tarif
 USE 
 USE IN sError
 USE IN people 
 IF USED('hosp')
  USE IN hosp 
 ENDIF 
 
 USE IN sprcokr
 USE IN tarif
 IF USED('pr4')
  USE IN pr4
 ENDIF 
 IF USED('pr4st')
  USE IN pr4st
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
 && Изменено 27.04.2019 под новый протокол
 dimdata(4,5) = dimdata(1,5) + dimdata(2,5)  + dimdata(3,5)
 *dimdata(4,6) = dimdata(1,6) + dimdata(2,6)  + dimdata(3,6)
 dimdata(4,6) = dimdata(2,6)
 *dimdata(4,7) = dimdata(1,7) + dimdata(2,7)  + dimdata(3,7)
 dimdata(4,7) = dimdata(2,7)
 && Изменено 27.04.2019 под новый протокол
 dimdata(4,8) = dimdata(1,8) + dimdata(2,8)  + dimdata(3,8)
 dimdata(4,9) = dimdata(1,9) + dimdata(2,9)  + dimdata(3,9)
 dimdata(4,10)= dimdata(1,10)+ dimdata(2,10) + dimdata(3,10)
 dimdata(4,12)= dimdata(1,12)+ dimdata(2,12) + dimdata(3,12)
 dimdata(4,13)= dimdata(1,13)+ dimdata(2,13) + dimdata(3,13)
 dimdata(4,11)= dimdata(1,11)+ dimdata(2,11) + dimdata(3,11)
 
 dimdata(8,1) = dimdata(5,1) + dimdata(6,1)  + dimdata(7,1)
 dimdata(8,2) = dimdata(5,2) + dimdata(6,2)  + dimdata(7,2)
 dimdata(8,3) = dimdata(5,3) + dimdata(6,3)  + dimdata(7,3)
 dimdata(8,4) = dimdata(5,4) + dimdata(6,4)  + dimdata(7,4)
 dimdata(8,5) = dimdata(5,5) + dimdata(6,5)  + dimdata(7,5)
 && Изменено 27.04.2019 под новый протокол
 dimdata(8,6) = dimdata(5,6) + dimdata(6,6)  + dimdata(7,6)
 *dimdata(8,6) = 0
 dimdata(8,7) = dimdata(5,7) + dimdata(6,7)  + dimdata(7,7)
 *dimdata(8,7) = 0
 && Изменено 27.04.2019 под новый протокол
 dimdata(8,8) = dimdata(5,8) + dimdata(6,8)  + dimdata(7,8)
 dimdata(8,9) = dimdata(5,9) + dimdata(6,9)  + dimdata(7,9)
 dimdata(8,10)= dimdata(5,10)+ dimdata(6,10) + dimdata(7,10)
 dimdata(8,12)= dimdata(5,12)+ dimdata(6,12) + dimdata(7,12)
 dimdata(8,13)= dimdata(5,13)+ dimdata(6,13) + dimdata(7,13)
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
 dimdata(9,12)= dimdata(4,12)+ dimdata(8,12)
 dimdata(9,13)= dimdata(4,13)+ dimdata(8,13)
 dimdata(9,11)= dimdata(4,11)+ dimdata(8,11)

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
 
 DO CASE && Алгоритм Лёши Маслакова 28.06.2017
  CASE !m.IsPilot AND !m.IsPilots && and !m.IsHors
   m.koplate  = 0.00
   m.koplates = 0.00
   m.koplate2 = dimdata(4,10) + dimdata(8,10)

  *CASE !m.IsPilot AND !m.IsPilots AND m.IsHors && Таких (m.IsHors) с 10.2018 нет
  * m.koplate  = 0.00
  * m.koplates = 0.00
  * m.koplate2 = dimdata(4,10) + dimdata(8,10)

  CASE m.IsPilot AND !m.IsPilots && AND !m.IsHors && Кончаловский здесь!
   m.koplate  = m.finval - m.udsum + dimdata(2,6) + dimdata(2,7) + dimdata(3,5)
   m.koplates = 0
   m.koplate2 = m.koplate + dimdata(9,11) + IIF(!m.IsStPilot, dimdata(1,8), 0) + dimdata(2,8)  + dimdata(3,8) && dimdata(9,8)

  CASE !m.IsPilot AND m.isPilots && AND !m.IsHors
   m.koplate  = 0
   *m.koplates = m.finvals - m.udsums + dimdata(6,6) + dimdata(6,7) + dimdata(7,5)
   m.koplates = m.finvals - m.udsums + dimdata(7,5) && 19.11.2018
   *m.koplate2 = m.koplates + dimdata(9,11) + dimdata(9,8)
   m.koplate2 = m.koplates + dimdata(4,5)+dimdata(9,11) + dimdata(9,8) && 20.11.2018

  CASE m.isPilot AND m.isPilots && AND !m.IsHors
   m.koplate  = m.finval - m.udsum + dimdata(2,6) + dimdata(2,7) + dimdata(3,5)
   *m.koplates = m.finvals - m.udsums + dimdata(6,6) + dimdata(6,7) + dimdata(7,5)
   m.koplates = m.finvals - m.udsums + dimdata(7,5) && изменено 18.04.2019 Согаз
   m.koplate2 = m.koplate + m.koplates + dimdata(9,11) + dimdata(9,8)

  *CASE m.isPilot and !m.IsPilots and m.isHors  && Таких (m.IsHors) с 10.2018 нет
  * m.koplate  = m.finval - m.udsum + dimdata(2,6) + dimdata(2,7) + dimdata(3,5) + dimdata(6,6) + dimdata(6,7) + dimdata(7,5)
  * m.koplates = 0
  * m.koplate2 = m.koplate + dimdata(9,11) + dimdata(9,8)
 OTHERWISE 
  MESSAGEBOX('ОШИБКА ОТНЕСЕНИЯ ЛПУ '+M.MCOD,0+64,'')
 ENDCASE 
 
 *m.koplate2 = m.koplate2 + dimdata(9,12)
 
 *m.d_created = PADL(DAY(DATE()),2,'0') + ' '+ LOWER(NameOfMonth2(MONTH(DATE())))+ ' '+STR(YEAR(DATE()),4)+' г.'
 m.d_created = PADL(goApp.d_acts,2,'0')+' '+PROPER(ALLTRIM(NameOfMonth2(MONTH(goApp.d_acts))))+' '+STR(YEAR(goApp.d_acts),4)+' г.'
 m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 CREATE CURSOR curdata (recid i)
 m.llResult = X_Report(m.dotname, m.docname+'.xls', m.IsVisible)
 USE IN curdata 

 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 
 IF fso.FileExists(m.docname+'.pdf')
  fso.DeleteFile(m.docname+'.pdf')
 ENDIF 
 oDoc = oExcel.Workbooks.Add(m.docname+'.xls')
 TRY 
  odoc.SaveAs(m.docname,57)
 CATCH 
  oWMI = GETOBJECT('winmgmts://')
  cQuery = "select * from win32_process where name='excel.exe'"
  oResult = oWMI.ExecQuery(cQuery)
  IF oResult.Count>0
   FOR EACH oProcess IN oResult
    oProcess.Terminate(1)
   NEXT
  ENDIF 

  odoc.SaveAs(m.docname,57)

 FINALLY 
  odoc.Close(0) && RPC Server
 ENDTRY 



 SELECT AisOms
 REPLACE st_flk WITH m.st_flk

RETURN  

