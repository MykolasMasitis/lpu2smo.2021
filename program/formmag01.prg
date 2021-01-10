PROCEDURE FormMag01
 IF MESSAGEBOX('СФОРМИРОВАТЬ ФОРМУ МАГ-1',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\horlpus.dbf')
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ HORLPUS.DBF!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\tarifn.dbf')
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ TARIFN.DBF!',0+16,'')
  LOOP 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpus', 'horlpu', 'shar')>0
  IF USED('horlpu')
   USE IN horlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  USE IN horlpu
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR cr (lpuid n(4), mcod c(7), str04 n(11,2), str41 n(11,2), str05 n(11,2), str51 n(11,2), str52 n(11,2), s_all n(11,2))
 INDEX on lpuid TAG lpuid
 INDEX on mcod TAG mcod

 SELECT horlpu
 SCAN 
  m.tpn = tpn
  IF !INLIST(m.tpn,'1','3')
   LOOP 
  ENDIF 
  m.mcod = mcod 
  m.lpuid = lpu_id
  m.IsStomat   = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
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
   SELECT horlpu
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   USE IN people 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT horlpu
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'serror', 'shar', 'rid')>0
   USE IN people 
   USE IN talon 
   IF USED('serror')
    USE IN serror
   ENDIF 
   SELECT horlpu
   LOOP 
  ENDIF 
  
  DIMENSION dimdata(9,11)
  dimdata = 0 
  
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO cod INTO tarif ADDITIVE 
  SET RELATION TO recid INTO serror ADDITIVE 
  SCAN 
  
  SCATTER MEMVAR 
  m.cod       = cod
  m.sn_pol    = sn_pol
  m.IsErr     = IIF(!EMPTY(serror.rid), .T., .F.)
  m.prmcod    = people.prmcod
  m.prmcods   = people.prmcods

  m.s_all     = s_all 
  m.rslt      = rslt
  m.fil_id    = fil_id
  m.otd       = SUBSTR(otd,2,2)
  m.proff     = SUBSTR(otd,4,3) && профиль услуги
  m.d_type    = d_type 
  m.lpu_ord   = lpu_ord
  m.ord       = ord
  
  m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)

*  IF m.IsLpuTpn=.t.
*   m.IsUslTpn = IIF(SEEK(m.fil_id, 'lputpn', 'fil_id'), .t., .f.)
*  ELSE 
   m.IsUslTpn = .f.
*  ENDIF 

  m.Is02      = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
  
  m.prmcod = IIF(m.mcod!='0344704', m.prmcod, '0344704')
  
  IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod) && стационар
   LOOP 
  ENDIF 
  IF m.IsTpnR OR INLIST(m.otd,'08','70','73','93') OR m.d_type='s' && допуслуги
   LOOP 
  ENDIF 
  IF (INLIST(m.otd,'01','90') AND IsStac(m.mcod)) AND people.mcod!=people.prmcods && допуслуги
   LOOP 
  ENDIF 
  IF (m.ord=7 AND m.lpu_ord=7665) AND people.mcod!=people.prmcods && допуслуги
   LOOP 
  ENDIF 

  m.UslIskl      = IIF(FLOOR(m.cod/1000)=146, .T., .F.)
  m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
  m.IsStomatUsl2 = IIF(INLIST(m.cod,1101,1102,101171,101172), .T., .F.)
   
  IF ((m.IsStomat AND !m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2)) OR ;
  	 ((m.IsStomat AND m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2 OR m.IsIskl)) OR ;
  	  (!m.IsStomat AND (m.IsStomatUsl OR (m.IsStomatUsl2 AND LEFT(m.ds,2)='K0')))
  
*  m.st_flk = m.st_flk + IIF(m.IsErr,m.s_all,0)
  
  DO CASE 
   CASE EMPTY(m.prmcods) && неприкрепленные
    dimdata(7,2)=dimdata(7,2)+1
    dimdata(7,3)=dimdata(7,3)+m.s_all
*    IF !SEEK(m.sn_pol, 'paz3st')
*     INSERT INTO paz3st FROM MEMVAR 
*    ENDIF 
    IF m.IsErr = .F.
*     IF !SEEK(m.sn_pol, 'paz3stok')
*      INSERT INTO paz3stok FROM MEMVAR 
*     ENDIF 
    ENDIF 

    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'

     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(7,8) = dimdata(7,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(7,11) = dimdata(7,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     OTHERWISE 
       dimdata(7,5) = dimdata(7,5) + IIF(m.IsErr,0,m.s_all)
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
*    IF !SEEK(m.sn_pol, 'paz1st')
*     INSERT INTO paz1st FROM MEMVAR 
*    ENDIF 
    IF m.IsErr = .F.
*     IF !SEEK(m.sn_pol, 'paz1stok')
*      INSERT INTO paz1stok FROM MEMVAR 
*     ENDIF 
    ENDIF 

    DO CASE 

     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'

     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(5,8) = dimdata(5,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE m.otd='93' AND IsStac(m.mcod)
      dimdata(5,11) = dimdata(5,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     OTHERWISE 
       dimdata(5,5) = dimdata(5,5) + IIF(m.IsErr,0,m.s_all)
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
*    IF !SEEK(m.sn_pol, 'paz2st')
*     INSERT INTO paz2st FROM MEMVAR 
*    ENDIF 
    IF m.IsErr = .F.
*     IF !SEEK(m.sn_pol, 'paz2stok')
*      INSERT INTO paz2stok FROM MEMVAR 
*     ENDIF 
    ENDIF 

    DO CASE 

     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'

     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(6,8) = dimdata(6,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(6,11) = dimdata(6,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'

     OTHERWISE 
	    
       dimdata(6,5) = dimdata(6,5) + IIF(m.IsErr,0,m.s_all)
	   IF 3=2 && Так было до 23.11.16
       IF m.Is02
        dimdata(6,7) = dimdata(6,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 
       IF (m.lpu_ord>0 AND m.Is02=.F.) OR (m.lpu_ord=0 AND INLIST(m.otd,'08','92') AND m.Is02=.F.)
        dimdata(6,6) = dimdata(6,6) + IIF(m.IsErr,0,m.s_all)
       ENDIF 
       ENDIF  && Так было до 23.11.16

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

  ELSE

  DO CASE 
   CASE EMPTY(m.prmcod) && неприкрепленные
    dimdata(3,2)=dimdata(3,2)+1
    dimdata(3,3)=dimdata(3,3)+m.s_all
*    IF !SEEK(m.sn_pol, 'paz3')
*     INSERT INTO paz3 FROM MEMVAR 
*    ENDIF 
    IF m.IsErr = .F.
*     IF !SEEK(m.sn_pol, 'paz3ok')
*      INSERT INTO paz3ok FROM MEMVAR 
*     ENDIF 
    ENDIF 

    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'

     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(3,8) = dimdata(3,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'
     
     OTHERWISE 
       dimdata(3,5) = dimdata(3,5) + IIF(m.IsErr,0,m.s_all)
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
*    IF !SEEK(m.sn_pol, 'paz1')
*     INSERT INTO paz1 FROM MEMVAR 
*    ENDIF 
    IF m.IsErr = .F.
*     IF !SEEK(m.sn_pol, 'paz1ok')
*      INSERT INTO paz1ok FROM MEMVAR 
*     ENDIF 
    ENDIF 

    DO CASE 

     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'

     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(1,8) = dimdata(1,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'
     
     CASE m.otd='93' AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'
     
     OTHERWISE 
       dimdata(1,5) = dimdata(1,5) + IIF(m.IsErr,0,m.s_all)
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
*    IF !SEEK(m.sn_pol, 'paz2')
*     INSERT INTO paz2 FROM MEMVAR 
*    ENDIF 
    IF m.IsErr = .F.
*     IF !SEEK(m.sn_pol, 'paz2ok')
*      INSERT INTO paz2ok FROM MEMVAR 
*     ENDIF 
    ENDIF 

    DO CASE 

     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'

     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(2,8) = dimdata(2,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '8'

     OTHERWISE 
	    
       dimdata(2,5) = dimdata(2,5) + IIF(m.IsErr,0,m.s_all)
*       REPLACE n_vmp WITH IIF(EMPTY(n_vmp), '12', n_vmp+' 12')

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

  ENDSCAN 
  SET RELATION OFF INTO serror
  SET RELATION OFF INTO people 
  SET RELATION OFF INTO tarif
  USE 
  USE IN people 
  USE IN serror

  dimdata(8,10)= dimdata(5,10)+ dimdata(6,10) + dimdata(7,10)
  INSERT INTO cr (lpuid, mcod, str04, str41, str05, str51, str52, s_all) VALUES ;
  	(m.lpuid, m.mcod, dimdata(2,6)+dimdata(2,7)+dimdata(6,6)+dimdata(6,7), dimdata(6,6)+dimdata(6,7), ;
  	dimdata(3,5)+dimdata(7,5), dimdata(3,5), dimdata(7,5), dimdata(8,10))
  
 ENDSCAN 
 USE IN horlpu
 USE IN tarif
 
 SELECT cr
 COPY TO &pbase\&gcperiod\mag01 CDX 

 MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64,'')
RETURN 