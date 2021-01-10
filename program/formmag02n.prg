PROCEDURE FormMAG02n
 IF MESSAGEBOX('СФОРМИРОВАТЬ ФОРМУ МАГ-2',4+32,'NEW')=7
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
 	col06 n(13,2), col21 n(13,2), col07 n(13,2), col08 n(13,2), col18 n(13,2), col09 n(13,2), col13 n(13,2),;
 	col10 n(13,2), col22 n(13,2),col11 n(13,2), col14 n(13,2), col15 n(13,2), col17 n(13,2), col19 n(13,2),;
 	col16 n(13,2), col20 n(13,2), col12 n(13,2), col23 n(13,2))

 SELECT curdata 
 INDEX on lpuid TAG lpuid
 INDEX on mcod TAG mcod 
 
 m.nrec = 0 

 SELECT aisoms
 SCAN 
  m.lpuid   = lpuid
  m.lpu_id   = lpuid
  m.mcod    = mcod
  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
  m.tpn     = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.tpn), '')
  m.tpns    = IIF(SEEK(m.mcod, 'sprlpu'), IIF(!EMPTY(FIELD('tpns', 'sprlpu')), ALLTRIM(sprlpu.tpns), ''), '')
  
  WAIT m.mcod + '...' WINDOW NOWAIT 
  
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
  m.col10 = 0
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

  m.col22 = 0 && стационар после МЭК
  m.col23 = 0 && допуслуги после МЭК
  
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
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
 
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO serror ADDITIVE 

  SCAN 

  SCATTER MEMVAR 
   m.prmcod    = people.prmcod
   m.prmcods   = people.prmcods
   
   * представлено МО
   m.col06 = m.col06 + m.s_all && представлено всего (по тарифу)
   m.col07 = m.col07 + m.s_lek

   *m.col07 = m.col07 + IIF(EMPTY(serror.c_err) AND ;
    	(IsMes(m.cod) OR IsVMP(m.cod)), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0)
   
   *m.col08 = m.col08 + IIF(IsMes(m.cod) OR IsVMP(m.cod) OR INLIST(cod,56029,156003), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0)
   m.col08 = m.col08 + IIF(m.qcod<>'I3', IIF(IsMes(m.cod) OR IsVMP(m.cod), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0), ;
   	IIF(IIF(m.mcod='0343003', people.prmcod<>'0343003', .T.) AND (IsMes(m.cod) OR IsVMP(m.cod)), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0))
   
   *m.col22 = m.col22 + IIF(EMPTY(serror.c_err) AND (IsMes(m.cod) OR IsVMP(m.cod) OR INLIST(cod,56029,156003)), m.s_all, 0)
   *m.col22 = m.col22 + IIF(EMPTY(serror.c_err) AND (IsMes(m.cod) OR IsVMP(m.cod) OR INLIST(cod,56029,156003)), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0)
   m.col22 = m.col22 + IIF(EMPTY(serror.c_err) AND (IsMes(m.cod) OR IsVMP(m.cod)), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0)
   
   *m.col08 = m.col08 + IIF(EMPTY(serror.c_err) AND INLIST(m.mp,'4','8'), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0)
   m.col09 = m.col09 + IIF(INLIST(m.mp,'4','8'), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0)
   *m.col23 = m.col23 + IIF(EMPTY(serror.c_err) AND INLIST(m.mp,'4','8'), m.s_all, 0)
   m.col23 = m.col23 + IIF(EMPTY(serror.c_err) AND INLIST(m.mp,'4','8'), m.s_all+IIF(m.qcod<>'I3', m.s_lek, 0), 0)

   m.col10 = m.col10 + IIF(INLIST(m.mp,'8'), m.s_all, 0) && стоматолгия 

   *m.col11 = m.col11 + IIF(IsDental(m.cod, m.lpu_id, m.mcod, m.ds) AND m.mp<>'8' and ;
   	(m.IsPilots OR INLIST(m.mcod, '4344623','4344700','4344621','4344640','4344613','4134752','0343036','0244124')), ;
   		m.s_all, 0)
   m.col11 = m.col11 + IIF(IsDental(m.cod, m.lpu_id, m.mcod, m.ds) AND m.mp<>'8', m.s_all, 0)

   m.col12 = m.col12 + IIF(IsDental(m.cod, m.lpu_id, m.mcod, m.ds) AND m.mp<>'8' AND m.Typ='0' and ;
   	(m.IsPilots OR INLIST(m.mcod, '4344623','4344700','4344621','4344640','4344613','4134752','0343036','0244124')), ;
   		m.s_all, 0)
   * представлено МО
   
   * МЭК
   m.col13 = m.col13 + IIF(!EMPTY(serror.c_err), m.s_all, 0) && OK!
   m.col14 = m.col14 + IIF(!EMPTY(serror.c_err), m.s_lek, 0) && OK!
   m.col15 = m.col15 + IIF(EMPTY(serror.c_err), m.s_all, 0) && OK!
   
   * терапия
   IF m.IsPilot && OR m.IsPilots
    m.col16 = m.col16 + IIF(!EMPTY(serror.c_err) AND m.mp='p', m.s_all, 0)
    m.col20 = m.col20 + IIF(!EMPTY(serror.c_err) AND !IsDental(m.cod, m.lpu_id, m.mcod, m.ds) AND m.mp<>'p', m.s_all, 0)
   ELSE 
    m.col20 = m.col20 + IIF(!EMPTY(serror.c_err) AND !IsDental(m.cod, m.lpu_id, m.mcod, m.ds), m.s_all, 0)
   ENDIF 
   * терапия
   
   * стоматология
   m.col17 = m.col17 + IIF(!EMPTY(serror.c_err) AND m.mp='s', m.s_all, 0) && OK!
   m.col18 = m.col18 + IIF(!EMPTY(serror.c_err) AND IsDental(m.cod, m.lpu_id, m.mcod, m.ds) AND m.mp<>'s', m.s_all, 0)
   m.col19 = m.col19 + IIF(!EMPTY(serror.c_err) AND m.mp='s' AND Typ='0', m.s_all, 0) && OK!
   * стоматология
   * МЭК
   
   m.col21 = m.col21 + IIF(EMPTY(serror.c_err) AND m.mp='p', m.s_all, 0) && АПП для сводной ведомости вер.3 (стомат)

  ENDSCAN 
  SET RELATION OFF INTO serror
  SET RELATION OFF INTO people
  USE 
  USE IN people 
  USE IN serror
  
  INSERT INTO curdata FROM MEMVAR 
  
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
 
 m.q_name = 'Страховая медицинская организация: '+m.qname
 m.p_period = 'Отчетный период: '+ NameOfMonth(m.tmonth)+' '+STR(m.tyear,4)+' г.'

 m.llResult = X_Report(pTempl+'\FormMAG02.xls', pBase+'\'+m.gcperiod+'\FormMAG02.xls', .T.)

 USE 
 
 
RETURN 