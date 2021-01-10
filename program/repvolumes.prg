PROCEDURE repVolumes
 IF MESSAGEBOX('СФОРМИРОВАТЬ ОТЧЕТ ПО ОБЪЕМАМ?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\repVolumes.xls')
  MESSAGEBOX('ОТСУТСТВУЕТ ШАБЛОН ОТЧЕТА repVolumes.XLS',0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN aisoms 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\profot', 'profot', 'shar', 'otd')>0
  IF USED('profot')
   USE IN profot
  ENDIF 
  USE IN sprlpu
  USE IN aisoms 
  RETURN 
 ENDIF 
 IF OpenFile(pCommon+'\gr_plan', 'gr', 'shar', 'cod')>0
  IF USED('gr')
   USE IN gr
  ENDIF 
  USE IN profot
  USE IN sprlpu
  USE IN aisoms 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curss (nrec i AUTOINC , mcod c(7), moname c(40), napp n(6), sapp n(13,2), npetkt n(6), spetkt n(13,2), nst n(6), sst n(13,2),;
 	ndst n(6), sdst n(13,2), ndstonk n(6), sdstonk n(13,2), n_x n(6), s_x n(13,2), nlt n(6), slt n(13,2), nvmp n(6), svmp n(13,2),;
 	ngem n(6), sgem n(13,2), n_eco n(6), s_eco n(13,2))
 SELECT curss
 INDEX on nrec TAG nrec
 INDEX on mcod TAG mcod 
 
 SELECT aisoms 
 SCAN 
  m.mcod = mcod
  m.moname = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.name, '')
  m.napp    = 0
  m.sapp    = 0
  m.npetkt  = 0
  m.spetkt  = 0
  m.nst     = 0
  m.sst     = 0
  m.ndst    = 0
  m.sdst    = 0
  m.ndstonk = 0
  m.sdstonk = 0
  m.n_x     = 0
  m.s_x     = 0
  m.nlt     = 0
  m.slt     = 0
  m.nvmp    = 0
  m.svmp    = 0
  m.ngem    = 0
  m.sgem    = 0
  m.n_eco   = 0
  m.s_eco   = 0

  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('err')
    USE IN err
   ENDIF 
   USE IN talon
   SELECT aisoms
   LOOP 
  ENDIF 
  
  CREATE CURSOR hosp (c_i c(30))
  SELECT hosp 
  INDEX on c_i TAG c_i 
  SET ORDER TO c_i
  
  SELECT talon 
  SET RELATION TO recid INTO err 
  SCAN 
   m.otd    = SUBSTR(otd,2,2)
   m.usl_ok = IIF(SEEK(m.otd, 'profot'), VAL(profot.usl_ok), 0)
   m.c_i    = c_i
   m.s_all  = s_all + s_lek
   m.cod    = cod
   m.ds   = ds
   m.ds_2 = ds_2
   m.ds_onk = ds_onk
   m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) OR ;
  		BETWEEN(LEFT(m.ds,3),'D00','D09') , .T., .F.)
   
   DO CASE 
    CASE m.usl_ok = 1 && Стационар 
     IF !SEEK(m.c_i, 'hosp')
      INSERT INTO hosp FROM MEMVAR 
     ENDIF 
     m.sst = m.sst + m.s_all
     
    CASE m.usl_ok = 2 && Дневной стационар
     m.ndst = m.ndst + IIF(!INLIST(m.cod,97010,197010,97041), 1, 0)
     m.sdst = m.sdst + IIF(!INLIST(m.cod,97010,197010,97041), m.s_all, 0)

     m.ndstonk = m.ndstonk + IIF(!INLIST(m.cod,97010,197010,97041) AND m.IsOnkDs, 1, 0)
     m.sdstonk = m.sdstonk + IIF(!INLIST(m.cod,97010,197010,97041) AND m.IsOnkDs, m.s_all, 0)

     m.n_eco = m.n_eco + IIF(INLIST(m.cod,97041), 1, 0)
     m.s_eco = m.s_eco + IIF(INLIST(m.cod,97041), m.s_all, 0)

     m.ngem = m.ngem + IIF(INLIST(m.cod,97010,197010), 1, 0)
     m.sgem = m.sgem + IIF(INLIST(m.cod,97010,197010), m.s_all, 0)

     m.nvmp = m.nvmp + IIF(INT(m.cod/1000)=297, 1, 0)
     m.svmp = m.svmp + IIF(INT(m.cod/1000)=297, m.s_all, 0)
     
     m.n_x = m.n_x + IIF(SEEK(m.cod, 'gr') and gr.gr_plan='on_х', 1, 0)
     m.s_x = m.s_x + IIF(SEEK(m.cod, 'gr') and gr.gr_plan='on_х', m.s_all, 0)

     m.nlt = m.nlt + IIF(SEEK(m.cod, 'gr') and gr.gr_plan='on_v', 1, 0)
     m.slt = m.slt + IIF(SEEK(m.cod, 'gr') and gr.gr_plan='on_v', m.s_all, 0)

    CASE m.usl_ok = 3 && АПП
     m.sapp = m.sapp + m.s_all
     m.npetkt = m.npetkt + IIF(INLIST(m.cod,37060,37061,37062,137060,137061), 1, 0)
     m.spetkt = m.spetkt + IIF(INLIST(m.cod,37060,37061,37062,137060,137061), m.s_all, 0)

     
    CASE m.usl_ok = 4 && внем МО, скорая помощь
    OTHERWISE         && ошибка
   ENDCASE 
  ENDSCAN 
  SET RELATION OFF INTO err 
  USE IN talon 
  USE IN err 
  
  m.nst = RECCOUNT('hosp')
  
  INSERT INTO curss FROM MEMVAR 
  
  USE IN hosp
  
  SELECT aisoms 
  
 ENDSCAN 
 USE 
 USE IN sprlpu
 USE IN profot
 USE IN gr
 
 SELECT curss 
 COPY TO &pBase\&gcPeriod\vol&qcod&gcperiod
 SET ORDER TO mcod 
 m.llResult = X_Report(pTempl+'\repVolumes.xls', pBase+'\'+m.gcPeriod+'\vol'+m.qcod+m.gcperiod+'.xls', .T.)
 USE IN curss
 
RETURN 