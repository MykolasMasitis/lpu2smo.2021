* Работающая версия на 18.10.2019!
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
 
 m.IsPilot  = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)
 m.IsPilotS = IIF(SEEK(m.lpuid, 'pilots'), .T., .F.)
 m.IsHorS   = IIF(SEEK(m.lpuid, 'horlpus'), .T., .F.)
 
 *m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn', 'lpu_id'), .t., .f.)	
 
 m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 
 m.arcfdate = DATETIME()
 IF fso.FileExists(lcPath + '\' + arcfname)
  poi_file   = fso.GetFile(lcPath + '\' + arcfname)
  m.arcfdate = poi_file.DateLastModified
 ENDIF 
 
 ZipItemCount = 5

 m.DotName = pTempl + "\Prqqmmy.xls"
 DocName = lcPath + "\Pr" + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)

 DIMENSION dimdata(9,12)
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
 REPLACE ALL Mp WITH '', Typ WITH '', vz WITH 0
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
  m.prmcods   = people.prmcods

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
  m.Is02      = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
  
  m.Mp = ''
  m.vz = 0
  
  m.test = 0 

  *IF m.IsPilotS AND IsDental(m.cod, m.lpuid, m.mcod, m.ds)
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
     *CASE m.IsTpnR = .T. OR INLIST(m.otd,'08')
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * REPLACE Mp WITH '8'

     * Закомментировано 27.05.2019 под новый протокол
     *CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * REPLACE Mp WITH '8'
     
     * Закомментировано 27.05.2019 под новый протокол
     *CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * REPLACE Mp WITH '8'
     
     * Закомментировано 27.05.2019 под новый протокол
     *CASE m.ord=7 AND m.lpu_ord=7665
     * dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * REPLACE Mp WITH '8'
     
     OTHERWISE 
       m.Mp = 's'
       dimdata(7,5) = dimdata(7,5) + IIF(m.IsErr,0,m.s_all)
       IF m.Is02
        dimdata(7,7) = dimdata(7,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(7,10) = dimdata(7,10) + IIF(m.IsErr,0,m.s_all)
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
 
     CASE m.IsTpnR = .T. OR INLIST(m.otd,'08') && tpn='r' - 3 услуги по июлю 2019, 08 - 4
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod) && 23 услуги
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
     
     CASE m.otd='93' AND IsStac(m.mcod) && ни одной!
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '8'
     
     OTHERWISE 
       m.Mp = 's'
       dimdata(5,5) = dimdata(5,5) + IIF(m.IsErr,0,m.s_all)
       IF m.Is02
        dimdata(5,7) = dimdata(5,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(5,10) = dimdata(5,10) + IIF(m.IsErr,0,m.s_all)
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
     *CASE m.IsTpnR = .T.
     * dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
     * REPLACE Mp WITH '8'

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
       * сюда вставить m.vz!	    
       m.Mp = 's'

       dimdata(6,5) = dimdata(6,5) + IIF(m.IsErr,0,m.s_all)

       DO CASE 
        CASE m.lpu_ord>0 && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
         dimdata(6,6) = dimdata(6,6) + IIF(m.IsErr,0,m.s_all)
         m.vz = 1
        CASE m.Is02 && vz=2, неотложная помощь (по реестру медицинских услуг)
         dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
         m.vz = 2
        CASE m.profil='100' AND INLIST(m.otd,'00','92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
         dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
         m.vz = 3
        CASE m.otd='08' && vz=4, услуги ЖК
         dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
         m.vz = 4
        CASE m.otd='91' && vz=5, услуги ЦЗ
         dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
         m.vz = 5
        OTHERWISE 
         m.vz = 0 && то, что должно попасть в up-файл
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
    *dimdata(6,12) = dimdata(6,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    *IF IsVMP(m.cod)
    * dimdata(6,9) = dimdata(6,9) + IIF(m.IsErr,0,m.s_all)
    *ENDIF 

   OTHERWISE 

  ENDCASE 

  ELSE && IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)

  DO CASE 
   CASE EMPTY(m.prmcod) && неприкрепленные
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
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08')) && Добавление условия pilot ничего не меняет
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (INLIST(m.otd,'08')) && Добавление условия pilot ничего не меняет
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
      *REPLACE test WITH 1

     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(3,8) = dimdata(3,8) + IIF(m.IsErr,0,m.s_all)
      m.Mp = 'm'

     CASE INLIST(m.otd,'01') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
      *REPLACE test WITH 3

     CASE INLIST(m.otd,'70','73','90','93') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
      *REPLACE test WITH 2
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
      *REPLACE test WITH 4
     
     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=people.prmcod AND people.tip_p=3 
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
      *REPLACE test WITH 5

     CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=people.prmcod AND people.tip_p=3 
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
      *REPLACE test WITH 6
     ** Добавлено 16.04.2019 по требованию Согаза

     OTHERWISE 
       m.Mp = 'p'
       dimdata(3,5) = dimdata(3,5) + IIF(m.IsErr,0,m.s_all)
       IF m.Is02
        dimdata(3,7) = dimdata(3,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(3,10) = dimdata(3,10) + IIF(m.IsErr,0,m.s_all)
    dimdata(3,12) = dimdata(3,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    IF IsVMP(m.cod)
     dimdata(3,9) = dimdata(3,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
   
   CASE m.mcod  = m.prmcod && свои пациенты
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

     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (INLIST(m.otd,'08'))
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'

     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(1,8) = dimdata(1,8) + IIF(m.IsErr,0,m.s_all)
      m.Mp = 'm'

     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
     
     OTHERWISE 
       m.Mp = 'p'
       dimdata(1,5) = dimdata(1,5) + IIF(m.IsErr,0,m.s_all)
       IF m.Is02
        dimdata(1,7) = dimdata(1,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 

    ENDCASE 
    dimdata(1,10) = dimdata(1,10) + IIF(m.IsErr,0,m.s_all)
    dimdata(1,12) = dimdata(1,12) + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
    IF IsVMP(m.cod)
     dimdata(1,9) = dimdata(1,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
    
   CASE m.mcod != m.prmcod && чужие пациенты
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

     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd,'08'))
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (INLIST(m.otd,'08'))
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'

     CASE IsMes(m.cod) OR IsVMP(m.cod)
      dimdata(2,8) = dimdata(2,8) + IIF(m.IsErr,0,m.s_all)
      m.Mp = 'm'

     CASE INLIST(m.otd,'01') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
     
     CASE INLIST(m.otd,'70','73','90','93') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'

     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=people.prmcod AND people.tip_p=3 
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'

     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=people.prmcod AND people.tip_p=3 
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      m.Mp = '4'
     ** Добавлено 16.04.2019 по требованию Согаза

     OTHERWISE 
	    
      m.Mp = 'p'
      dimdata(2,5) = dimdata(2,5) + IIF(m.IsErr,0,m.s_all)
      
      DO CASE 
       CASE m.lpu_ord>0 && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
        dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
        m.vz = 1
       CASE m.Is02 && vz=2, неотложная помощь (по реестру медицинских услуг)
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        m.vz = 2
       CASE m.profil='100' AND INLIST(m.otd,'00','92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        m.vz = 3
       CASE m.otd='08' && vz=4, услуги ЖК
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        m.vz = 4
       CASE m.otd='91' && vz=5, услуги ЦЗ
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
        m.vz = 5
       OTHERWISE 
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
    IF IsVMP(m.cod)
     dimdata(2,9) = dimdata(2,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 

   OTHERWISE 

  ENDCASE 

  ENDIF IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)

  REPLACE Mp WITH m.Mp, Typ WITH m.Typ, vz WITH m.vz

 ENDSCAN  
 
 SET RELATION OFF INTO people
 SET RELATION OFF INTO sError
 SET RELATION OFF INTO tarif
 USE 
 USE IN sError
 USE IN people 
 
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
  CASE !m.IsPilot AND !m.IsPilots and !m.IsHors
   m.koplate  = 0.00
   m.koplates = 0.00
   m.koplate2 = dimdata(4,10) + dimdata(8,10)

  CASE !m.IsPilot AND !m.IsPilots AND m.IsHors && Таких (m.IsHors) с 10.2018 нет
   m.koplate  = 0.00
   m.koplates = 0.00
   m.koplate2 = dimdata(4,10) + dimdata(8,10)

  CASE m.IsPilot AND !m.IsPilots AND !m.IsHors
   m.koplate  = m.finval - m.udsum + dimdata(2,6) + dimdata(2,7) + dimdata(3,5)
   m.koplates = 0
   m.koplate2 = m.koplate + dimdata(9,11) + dimdata(9,8)

  CASE !m.IsPilot AND m.isPilots AND !m.IsHors
   m.koplate  = 0
   *m.koplates = m.finvals - m.udsums + dimdata(6,6) + dimdata(6,7) + dimdata(7,5)
   m.koplates = m.finvals - m.udsums + dimdata(7,5) && 19.11.2018
   *m.koplate2 = m.koplates + dimdata(9,11) + dimdata(9,8)
   m.koplate2 = m.koplates + dimdata(4,5)+dimdata(9,11) + dimdata(9,8) && 20.11.2018

  CASE m.isPilot AND m.isPilots AND !m.IsHors
   m.koplate  = m.finval - m.udsum + dimdata(2,6) + dimdata(2,7) + dimdata(3,5)
   *m.koplates = m.finvals - m.udsums + dimdata(6,6) + dimdata(6,7) + dimdata(7,5)
   m.koplates = m.finvals - m.udsums + dimdata(7,5) && изменено 18.04.2019 Согаз
   m.koplate2 = m.koplate + m.koplates + dimdata(9,11) + dimdata(9,8)

  CASE m.isPilot and !m.IsPilots and m.isHors  && Таких (m.IsHors) с 10.2018 нет
   m.koplate  = m.finval - m.udsum + dimdata(2,6) + dimdata(2,7) + dimdata(3,5) + dimdata(6,6) + dimdata(6,7) + dimdata(7,5)
   m.koplates = 0
   m.koplate2 = m.koplate + dimdata(9,11) + dimdata(9,8)
 OTHERWISE 
  MESSAGEBOX('ОШИБКА ОТНЕСЕНИЯ ЛПУ '+M.MCOD,0+64,'')
 ENDCASE 
 
 *m.koplate2 = m.koplate2 + dimdata(9,12)
 
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
  odoc.Close(0)
 CATCH 
 ENDTRY 

 SELECT AisOms
 REPLACE st_flk WITH m.st_flk

RETURN  

