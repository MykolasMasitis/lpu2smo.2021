PROCEDURE MakeZPZ
 IF MESSAGEBOX('СФОРМИРОВАТЬ ОТЧЕТ ПО ФОРМЕ ЗПЗ',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\zpz.xls')
  MESSAGEBOX('ОТСУТСВУЕТ ФАЙЛ '+UPPER(pTempl+'\zpz.xls'),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 
 m.IsPrll  = IsRegistered("ParallelFox.Application") && Зарегистрирован ли parallelfox.exe
 m.IsPrll  = .F.
 
 oPrm = NEWOBJECT('BaseSet', 'DataSet.prg')
 WITH oPrm
  .pBase    = m.pBase
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
  
  Parallel.Do("mkzpz", "makezpz.prg", .F., oPrm)
  
 ELSE 
  
  DO mkzpz IN makezpz WITH oPrm

 ENDIF 
 RELEASE oPrm
 
RETURN 

PROCEDURE mkzpz(para1)
 m.oPrm      = para1
 
 WITH oPrm
  m.pBase    = .pBase
  m.pTempl   = .pTempl
  m.gcPeriod = .gcPeriod
  m.qCod     = .qCod
  m.qName    = .qname
 
  m.tMonth   = .tMonth
  m.tYear    = .tYear
 ENDWITH 

 m.IsAlert = IsRegistered("VFPAlert.AlertManager")   && Зарегистрирован ли AlertManager
 m.IsAlert = .F.
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
 
 
 DIMENSION dimdata(20,10)
 dimdata = 0

 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  m.IsPuchok = IIF(m.mcod = '0371001', .T., .F.)
  
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
  
  SELECT er
  SET RELATION TO LEFT(c_err,2) INTO sookod
  SELECT talon 
  SET RELATION TO recid INTO er ADDITIVE 
  SCAN 
   m.otd  = SUBSTR(otd,2,2)
   m.cod  = cod
   m.ds   = ds
   m.ds_2 = ds_2
   
   *m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR BETWEEN(LEFT(m.ds,3), 'D00', 'D09') OR ;
   	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
   m.IsOnk = IIF(INLIST(SUBSTR(otd,4,3),'018','060'), .T., .F.)
   	
   m.IsErr = IIF(!EMPTY(er.c_err), .T., .F.)
   m.osn230 = sookod.osn230
   
   IF m.IsPuchok
    dimdata(1,4) = dimdata(1,4) + 1
    dimdata(2,4) = dimdata(2,4) + IIF(m.IsOnk, 1, 0)
    dimdata(4,4) = dimdata(4,4) + IIF(m.IsErr, 1, 0)
    dimdata(5,4) = dimdata(5,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
    dimdata(6,4) = dimdata(6,4) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
    dimdata(7,4) = dimdata(7,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
    dimdata(8,4) = dimdata(8,4) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
    dimdata(9,4) = dimdata(9,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
    dimdata(10,4) = dimdata(10,4) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
    dimdata(11,4) = dimdata(11,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
    dimdata(12,4) = dimdata(12,4) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
    dimdata(13,4) = dimdata(13,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
    dimdata(14,4) = dimdata(14,4) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
    dimdata(17,4) = dimdata(17,4) + IIF(m.IsOnk AND m.IsErr, 1, 0)

   ELSE 
    DO CASE 
     CASE INLIST(m.otd,'00','01','08','85','90','91','92','93') && АПП
      dimdata(1,5) = dimdata(1,5) + 1
      dimdata(2,5) = dimdata(2,5) + IIF(m.IsOnk, 1, 0)
      dimdata(4,5) = dimdata(4,5) + IIF(m.IsErr, 1, 0)
      dimdata(5,5) = dimdata(5,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
      dimdata(6,5) = dimdata(6,5) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
      dimdata(7,5) = dimdata(7,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
      dimdata(8,5) = dimdata(8,5) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
      dimdata(9,5) = dimdata(9,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
      dimdata(10,5) = dimdata(10,5) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
      dimdata(11,5) = dimdata(11,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
      dimdata(12,5) = dimdata(12,5) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
      dimdata(13,5) = dimdata(13,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
      dimdata(14,5) = dimdata(14,5) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
      dimdata(17,5) = dimdata(17,5) + IIF(m.IsOnk AND m.IsErr, 1, 0)
      
     CASE INLIST(m.otd,'80','81') && ДСТ
      dimdata(1,6) = dimdata(1,6) + 1
      dimdata(2,6) = dimdata(2,6) + IIF(m.IsOnk, 1, 0)
      dimdata(4,6) = dimdata(4,6) + IIF(m.IsErr, 1, 0)
      dimdata(5,6) = dimdata(5,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
      dimdata(6,6) = dimdata(6,6) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
      dimdata(7,6) = dimdata(7,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
      dimdata(8,6) = dimdata(8,6) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
      dimdata(9,6) = dimdata(9,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
      dimdata(10,6) = dimdata(10,6) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
      dimdata(11,6) = dimdata(11,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
      dimdata(12,6) = dimdata(12,6) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
      dimdata(13,6) = dimdata(13,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
      dimdata(14,6) = dimdata(14,6) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
      dimdata(17,6) = dimdata(17,6) + IIF(m.IsOnk AND m.IsErr, 1, 0)
      
      IF BETWEEN(m.cod, 397001, 397056) OR INLIST(m.cod, 397058, 397059)
       dimdata(1,7) = dimdata(1,7) + 1
       dimdata(2,7) = dimdata(2,7) + IIF(m.IsOnk, 1, 0)
       dimdata(4,7) = dimdata(4,7) + IIF(m.IsErr, 1, 0)
       dimdata(5,7) = dimdata(5,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
       dimdata(6,7) = dimdata(6,7) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
       dimdata(7,7) = dimdata(7,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
       dimdata(8,7) = dimdata(8,7) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
       dimdata(9,7) = dimdata(9,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
       dimdata(10,7) = dimdata(10,7) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
       dimdata(11,7) = dimdata(11,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
       dimdata(12,7) = dimdata(12,7) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
       dimdata(13,7) = dimdata(13,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
       dimdata(14,7) = dimdata(14,7) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
       dimdata(17,7) = dimdata(17,7) + IIF(m.IsOnk AND m.IsErr, 1, 0)
      ENDIF 
      
     OTHERWISE && Стацинар
      dimdata(1,8) = dimdata(1,8) + 1
      dimdata(2,8) = dimdata(2,8) + IIF(m.IsOnk, 1, 0)
      dimdata(4,8) = dimdata(4,8) + IIF(m.IsErr, 1, 0)
      dimdata(5,8) = dimdata(5,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
      dimdata(6,8) = dimdata(6,8) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
      dimdata(7,8) = dimdata(7,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
      dimdata(8,8) = dimdata(8,8) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
      dimdata(9,8) = dimdata(9,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
      dimdata(10,8) = dimdata(10,8) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
      dimdata(11,8) = dimdata(11,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
      dimdata(12,8) = dimdata(12,8) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
      dimdata(13,8) = dimdata(13,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
      dimdata(14,8) = dimdata(14,8) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
      dimdata(17,8) = dimdata(17,8) + IIF(m.IsOnk AND m.IsErr, 1, 0)
      IF BETWEEN(m.cod, 200001, 200507)
       dimdata(1,9) = dimdata(1,9) + 1
       dimdata(2,9) = dimdata(2,9) + IIF(m.IsOnk, 1, 0)
       dimdata(4,9) = dimdata(4,9) + IIF(m.IsErr, 1, 0)
       dimdata(5,9) = dimdata(5,9) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
       dimdata(6,9) = dimdata(6,9) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), 1, 0)
       dimdata(7,9) = dimdata(7,9) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
       dimdata(8,9) = dimdata(8,9) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.3.1.','5.3.2.','5.3.3.'), 1, 0)
       dimdata(9,9) = dimdata(9,9) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
       dimdata(10,9) = dimdata(10,9) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.4.1.'), 1, 0)
       dimdata(11,9) = dimdata(11,9) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
       dimdata(12,9) = dimdata(12,9) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.'), 1, 0)
       dimdata(13,9) = dimdata(13,9) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
       dimdata(14,9) = dimdata(14,9) + IIF(m.IsOnk AND m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), 1, 0)
       dimdata(17,9) = dimdata(17,9) + IIF(m.IsOnk AND m.IsErr, 1, 0)
      ENDIF 
      
    ENDCASE 
   ENDIF 

  ENDSCAN 
  SELECT talon 
  SET RELATION OFF INTO er
  SELECT er
  SET RELATION OFF INTO sookod
  USE IN talon 
  USE IN er
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 USE IN sookod
 
 CREATE CURSOR curdata (recid i)
 INSERT INTO curdata (recid) VALUES (0)
 

 *dimdata(15,3)=dimdata(4,3)-(dimdata(5,3)+dimdata(7,3)+dimdata(9,3)+dimdata(11,3)+dimdata(13,3))
 dimdata(15,4)=dimdata(4,4)-(dimdata(5,4)+dimdata(7,4)+dimdata(9,4)+dimdata(11,4)+dimdata(13,4))
 dimdata(15,5)=dimdata(4,5)-(dimdata(5,5)+dimdata(7,5)+dimdata(9,5)+dimdata(11,5)+dimdata(13,5))
 dimdata(15,6)=dimdata(4,6)-(dimdata(5,6)+dimdata(7,6)+dimdata(9,6)+dimdata(11,6)+dimdata(13,6))
 dimdata(15,7)=dimdata(4,7)-(dimdata(5,7)+dimdata(7,7)+dimdata(9,7)+dimdata(11,7)+dimdata(13,7))
 dimdata(15,8)=dimdata(4,8)-(dimdata(5,8)+dimdata(7,8)+dimdata(9,8)+dimdata(11,8)+dimdata(13,8))
 dimdata(15,9)=dimdata(4,9)-(dimdata(5,9)+dimdata(7,9)+dimdata(9,9)+dimdata(11,9)+dimdata(13,9))
 
 *dimdata(16,3)=dimdata(17,3)-(dimdata(6,3)+dimdata(8,3)+dimdata(10,3)+dimdata(12,3)+dimdata(14,3))
 dimdata(16,4)=dimdata(17,4)-(dimdata(6,4)+dimdata(8,4)+dimdata(10,3)+dimdata(12,4)+dimdata(14,4))
 dimdata(16,5)=dimdata(17,5)-(dimdata(6,5)+dimdata(8,5)+dimdata(10,5)+dimdata(12,5)+dimdata(14,5))
 dimdata(16,6)=dimdata(17,6)-(dimdata(6,6)+dimdata(8,6)+dimdata(10,6)+dimdata(12,6)+dimdata(14,6))
 dimdata(16,7)=dimdata(17,7)-(dimdata(6,7)+dimdata(8,7)+dimdata(10,7)+dimdata(12,7)+dimdata(14,7))
 dimdata(16,8)=dimdata(17,8)-(dimdata(6,8)+dimdata(8,8)+dimdata(10,8)+dimdata(12,8)+dimdata(14,8))
 dimdata(16,9)=dimdata(17,9)-(dimdata(6,9)+dimdata(8,9)+dimdata(10,9)+dimdata(12,9)+dimdata(14,9))
 
 FOR n=1 TO 16 
  dimdata(n,3)=dimdata(n,4)+dimdata(n,5)+dimdata(n,6)+dimdata(n,8)
 ENDFOR 

 m.llResult = X_Report(pTempl+'\zpz.xls', pBase+'\'+m.gcperiod+'\zpz.xls', .F.)

 IF m.IsAlert
  oAlertMgr = CREATEOBJECT("VFPAlert.AlertManager")
  poAlert = oAlertMgr.NewAlert()
  poAlert.Alert("Отчет находится: "+pBase+'\'+m.gcPeriod+'\zpz.xls', 8, "ЗПЗ (Таблица 5)",;
  	"Формирование отчета закончено.")
 ENDIF 

RETURN 