PROCEDURE yu_03
 IF MESSAGEBOX('СФОРМИРОВАТЬ ОТЧЕТ ГО СОГАЗ (ОНКОЛОГИЯ)',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\yu_03.xls')
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ '+UPPER(pTempl+'\yu_03.xls'),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 

 m.IsPrll  = IsRegistered("ParallelFox.Application") && Зарегистрирован ли parallelfox.exe
 *m.IsPrll  = .F.
 
 oPrm = NEWOBJECT('BaseSet', 'DataSet.prg')
 WITH oPrm
  .pBase    = m.pBase
  .pCommon  = m.pCommon
  .pTempl   = m.pTempl
  .gcPeriod = m.gcPeriod
  .qCod     = m.qCod
  .qName    = m.qname
 
  .tMonth   = m.tMonth
  .tYear    = m.tYear
 ENDWITH 

 IF m.IsPrll
  Local Parallel as Parallel of ParallelFox.vcx
  Parallel = NewObject("Parallel", "ParallelFox.vcx")
  Parallel.SetWorkerCount(1)
  
  Parallel.StartWorkers(FullPath("lpu2smo.exe"),,.t.)

  TEXT TO cScript NOSHOW
   PUBLIC fso as SCRIPTING.FileSystemObject
   SET SAFETY OFF
   fso  = CREATEOBJECT('Scripting.FileSystemObject')
   SET PROCEDURE TO Utils.prg
  ENDTEXT 
 
  Parallel.ExecScript(cScript, .T.)
  
  Parallel.Do("yu", "yu_03.prg", .F., oPrm)
  
 ELSE 
  
  DO yu IN yu_03 WITH oPrm

 ENDIF 
 RELEASE oPrm
 
RETURN 

PROCEDURE yu(para1)
 m.oPrm      = para1
 
 WITH oPrm
  m.pBase    = .pBase
  m.pCommon  = .pCommon
  m.pTempl   = .pTempl
  m.gcPeriod = .gcPeriod
  m.qCod     = .qCod
  m.qName    = .qname
 
  m.tMonth   = .tMonth
  m.tYear    = .tYear
 ENDWITH 

 m.IsAlert = IsRegistered("VFPAlert.AlertManager")   && Зарегистрирован ли AlertManager
 *m.IsAlert = .F.
 IF m.IsAlert
  oAlertMgr = CREATEOBJECT("VFPAlert.AlertManager")
  poAlert = oAlertMgr.NewAlert()
  poAlert.Alert("Вы можете продолжить работу", 8, "ЗПЗ (Таблица 5)",;
  	"Начато формирование отчета в фоновом режиме")
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\sookodxx', 'sookod', 'shar', 'er_c')>0
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'shar', 'cod')>0
  USE IN sookod
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdata (recid i AUTOINC , mcod c(7), typ c(1), typ_name c(5), ;
 	onk_sl n(8), onk_sum n(11,2), pet_sl n(9), pet_sum n(11,2), vmp_sl n(9), vmp_sum n(11,2))
 
 
 DIMENSION dimdata(60,10)

 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  m.IsPuchok = IIF(m.mcod = '0371001', .T., .F.)
  IF m.IsPuchok
   LOOP 
  ENDIF 
  
  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'er', 'shar', 'rid')>0
   IF USED('er')
    USE IN er
   ENDIF 
   USE IN talon 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  dimdata = 0

  SELECT er
  SET RELATION TO LEFT(c_err,2) INTO sookod
  SELECT talon 
  SET RELATION TO recid INTO er ADDITIVE 
  SCAN 
   m.otd   = SUBSTR(otd,2,2)
   m.cod   = cod
   m.ds    = ds
   m.ds_2  = ds_2
   m.s_all = s_all
   
   m.IsOnk = IIF(INLIST(SUBSTR(otd,4,3),'018','060'), .T., .F.)
   m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)

   IF !m.IsOnk
    *LOOP 
   ENDIF 
   	
   m.IsErr = IIF(!EMPTY(er.c_err), .T., .F.)
   m.osn230 = sookod.osn230
   
    DO CASE 
     CASE INLIST(m.otd,'00','01','08','85','90','91','92','93') && АПП
      dimdata(1,3) = dimdata(1,3) + IIF(m.IsOnk, 1, 0)
      dimdata(1,4) = dimdata(1,4) + IIF(m.IsOnk, m.s_all, 0)
      
      IF INLIST(m.cod, 37047, 137047) && ПЭТ/КТ
       dimdata(1,5) = dimdata(1,5) + IIF(m.IsOnkDs, 1, 0)
       dimdata(1,6) = dimdata(1,6) + IIF(m.IsOnkDs, m.s_all, 0)
      ENDIF 

     CASE INLIST(m.otd,'80','81') && ДСТ
      dimdata(2,3) = dimdata(2,3) + IIF(m.IsOnk, 1, 0)
      dimdata(2,4) = dimdata(2,4) + IIF(m.IsOnk, m.s_all, 0)

      IF INLIST(m.cod, 37047, 137047) && ПЭТ/КТ
       dimdata(2,5) = dimdata(2,5) + IIF(m.IsOnkDs, 1, 0)
       dimdata(2,6) = dimdata(2,6) + IIF(m.IsOnkDs, m.s_all, 0)
      ENDIF 

      IF BETWEEN(m.cod, 397001, 397056) OR INLIST(m.cod, 397058, 397059)
       dimdata(2,7) = dimdata(2,7) + 1 && IIF(m.IsOnk, 1, 0)
       dimdata(2,8) = dimdata(2,8) + m.s_all && IIF(m.IsOnk, m.s_all, 0)
      ENDIF 
      
     OTHERWISE && Стацинар
      dimdata(3,3) = dimdata(3,3) + IIF(m.IsOnk, 1, 0)
      dimdata(3,4) = dimdata(3,4) + IIF(m.IsOnk, m.s_all, 0)

      IF INLIST(m.cod, 37047, 137047) && ПЭТ/КТ
       dimdata(3,5) = dimdata(3,5) + IIF(m.IsOnkDs, 1, 0)
       dimdata(3,6) = dimdata(3,6) + IIF(m.IsOnkDs, m.s_all, 0)
      ENDIF 

      IF BETWEEN(m.cod, 200001, 200600)
       dimdata(3,7) = dimdata(3,7) + 1 && IIF(m.IsOnk, 1, 0)
       dimdata(3,8) = dimdata(3,8) + m.s_all && IIF(m.IsOnk, m.s_all, 0)
      ENDIF 
      
    ENDCASE 

  ENDSCAN 
  SELECT talon 
  SET RELATION OFF INTO er
  SELECT er
  SET RELATION OFF INTO sookod
  USE IN talon 
  USE IN er
  
  IF dimdata(1,3)=0 AND dimdata(1,5)=0 AND dimdata(1,7)=0 AND ;
  	dimdata(2,3)=0 AND dimdata(2,5)=0 AND dimdata(2,7)=0 AND ;
  	dimdata(3,3)=0 AND dimdata(3,5)=0 AND dimdata(3,7)=0
  ELSE   
  INSERT INTO curdata (mcod , typ , typ_name , onk_sl, onk_sum , pet_sl , pet_sum, vmp_sl, vmp_sum) VALUES ;
  	(m.mcod, '1', 'АПП', dimdata(1,3), dimdata(1,4), dimdata(1,5), dimdata(1,6), dimdata(1,7), dimdata(1,8))
  INSERT INTO curdata (mcod , typ , typ_name , onk_sl, onk_sum , pet_sl , pet_sum, vmp_sl, vmp_sum) VALUES ;
  	(m.mcod, '2', 'ДСТ', dimdata(2,3), dimdata(2,4), dimdata(2,5), dimdata(2,6), dimdata(2,7), dimdata(2,8))
  INSERT INTO curdata (mcod , typ , typ_name , onk_sl, onk_sum , pet_sl , pet_sum, vmp_sl, vmp_sum) VALUES ;
  	(m.mcod, '3', 'СТ', dimdata(3,3), dimdata(3,4), dimdata(3,5), dimdata(3,6), dimdata(3,7), dimdata(3,8))
  ENDIF 
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 USE IN sookod
 USE IN dspcodes

 IF m.IsAlert
  oAlertMgrr = CREATEOBJECT("VFPAlert.AlertManager")
  pooAlert = oAlertMgrr.NewAlert()
  pooAlert.Alert("Отчет находится: "+pBase+'\'+m.gcPeriod+'\vks.xls', 8, "ГО-03",;
  	"Формирование отчета закончено.")
 ENDIF 

 m.llResult = X_Report(pTempl+'\yu_03.xls', pBase+'\'+m.gcperiod+'\yu_03_'+m.gcPeriod+'.xls', .F.)

RETURN 