PROCEDURE MakeVKS
 IF MESSAGEBOX('СФОРМИРОВАТЬ ОТЧЕТ ВКС ФФОМС',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\vks_ffoms.xls')
  MESSAGEBOX('ОТСУТСВУЕТ ФАЙЛ '+UPPER(pTempl+'\vks_ffoms.xls'),0+64,'')
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
  
  Parallel.Do("mkvks", "makevks.prg", .F., oPrm)
  
 ELSE 
  
  DO mkvks IN makevks WITH oPrm

 ENDIF 
 RELEASE oPrm
 
RETURN 

PROCEDURE mkvks(para1)
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
 
 
 DIMENSION dimdata(60,10)
 dimdata = 0

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
   
   *m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR BETWEEN(LEFT(m.ds,3), 'D00', 'D09') OR ;
   	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
   m.IsOnk = IIF(INLIST(SUBSTR(otd,4,3),'018','060'), .T., .F.)
   	
   m.IsDsp    = IIF(SEEK(m.cod, 'dspcodes') AND INLIST(dspcodes.tip,1,3), .T., .F.)
   m.IsProf   = IIF(SEEK(m.cod, 'dspcodes') AND INLIST(dspcodes.tip,2,4,5,6,7), .T., .F.)
   m.tipofcod = IIF(m.IsDsp, dspcodes.tip, 0)
   m.rslt     = rslt
   
   IF m.IsDsp  
   DO CASE 
    CASE m.tipofcod = 1 && Диспасеризация взрослых, с июля
     IF !INLIST(m.cod, 25204, 35401) AND !BETWEEN(m.rslt,316,319) AND !BETWEEN(m.rslt,352,358) 
      m.IsDsp = .F.
     ENDIF 
    CASE m.tipofcod = 2 && Профосмотры взрослых, с сентября
     IF !BETWEEN(m.rslt,343,345)
      m.IsProf = .F.
     ENDIF 
    CASE m.tipofcod = 3 && Диспасеризация детей-сирот, с июля
     IF !BETWEEN(m.rslt,321,325) AND m.rslt!=320 AND m.rslt!=390 AND !BETWEEN(m.rslt,347,351) && Это - усыновленные сироты!
      m.IsDsp = .F.
     ENDIF 
    CASE m.tipofcod = 4 && Профосмотры несовершеннолетних, с сентября
     IF !BETWEEN(m.rslt,332,336) AND m.rslt!=326
      m.IsProf = .F.
     ENDIF 
    CASE m.tipofcod = 5 && Предварительные профосмотры несовершеннолетних, с сентября
     IF !BETWEEN(m.rslt,337,341) AND !INLIST(m.rslt,326,396)
      m.IsProf = .F.
     ENDIF 
    CASE m.tipofcod = 6 && Периодические профосмотры несовершеннолетних, с сентября
     IF m.rslt!=342  AND m.rslt!=326
      m.IsProf = .F.
     ENDIF 
    CASE m.tipofcod = 7 && профосмотры, с сентября
    OTHERWISE 
     m.IsDsp = .F.
     m.IsProf
    ENDCASE 
   ENDIF 

   m.IsErr = IIF(!EMPTY(er.c_err), .T., .F.)
   m.osn230 = sookod.osn230
   
    DO CASE 
     CASE INLIST(m.otd,'00','01','08','85','90','91','92','93') && АПП
      dimdata(1,4) = dimdata(1,4) + m.s_all
      dimdata(2,4) = dimdata(2,4) + IIF(m.IsOnk, m.s_all, 0)
      dimdata(3,4) = dimdata(3,4) + IIF(m.IsDsp, m.s_all, 0)
      dimdata(4,4) = dimdata(4,4) + IIF(m.IsProf, m.s_all, 0)

      dimdata(5,4) = dimdata(5,4) + 1
      dimdata(6,4) = dimdata(6,4) + IIF(m.IsOnk, 1, 0)
      dimdata(7,4) = dimdata(7,4) + IIF(m.IsDsp, 1, 0)
      dimdata(8,4) = dimdata(8,4) + IIF(m.IsProf, 1, 0)

      dimdata(9,4)  = dimdata(9,4) + IIF(m.IsErr, m.s_all, 0)
      dimdata(10,4) = dimdata(10,4) + IIF(m.IsErr AND m.IsOnk, m.s_all, 0)
      dimdata(11,4) = dimdata(11,4) + IIF(m.IsErr AND m.IsDsp, m.s_all, 0)
      dimdata(12,4) = dimdata(12,4) + IIF(m.IsErr AND m.IsProf, m.s_all, 0)

      dimdata(13,4) = dimdata(13,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(14,4) = dimdata(14,4) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(15,4) = dimdata(15,4) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(16,4) = dimdata(16,4) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)

      dimdata(17,4) = dimdata(17,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(18,4) = dimdata(18,4) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(19,4) = dimdata(19,4) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(20,4) = dimdata(20,4) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)

      dimdata(21,4) = dimdata(21,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(22,4) = dimdata(22,4) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(23,4) = dimdata(23,4) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(24,4) = dimdata(24,4) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)

      dimdata(25,4) = dimdata(25,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(26,4) = dimdata(26,4) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(27,4) = dimdata(27,4) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(28,4) = dimdata(28,4) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)

      dimdata(29,4) = dimdata(29,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(30,4) = dimdata(30,4) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(31,4) = dimdata(31,4) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(32,4) = dimdata(32,4) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)

      dimdata(33,4) = dimdata(33,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(34,4) = dimdata(34,4) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(35,4) = dimdata(35,4) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(36,4) = dimdata(36,4) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)

      dimdata(37,4) = dimdata(37,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(38,4) = dimdata(38,4) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(39,4) = dimdata(39,4) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(40,4) = dimdata(40,4) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.6.'), m.s_all, 0)

      dimdata(41,4) = dimdata(41,4) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(42,4) = dimdata(42,4) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(43,4) = dimdata(43,4) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(44,4) = dimdata(44,4) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)

     CASE INLIST(m.otd,'80','81') && ДСТ
      dimdata(1,5) = dimdata(1,5) + m.s_all
      dimdata(2,5) = dimdata(2,5) + IIF(m.IsOnk, m.s_all, 0)
      dimdata(3,5) = dimdata(3,5) + IIF(m.IsDsp, m.s_all, 0)
      dimdata(4,5) = dimdata(4,5) + IIF(m.IsProf, m.s_all, 0)

      dimdata(5,5) = dimdata(5,5) + 1
      dimdata(6,5) = dimdata(6,5) + IIF(m.IsOnk, 1, 0)
      dimdata(7,5) = dimdata(7,5) + IIF(m.IsDsp, 1, 0)
      dimdata(8,5) = dimdata(8,5) + IIF(m.IsProf, 1, 0)

      dimdata(9,5)  = dimdata(9,5) + IIF(m.IsErr, m.s_all, 0)
      dimdata(10,5) = dimdata(10,5) + IIF(m.IsErr AND m.IsOnk, m.s_all, 0)
      dimdata(11,5) = dimdata(11,5) + IIF(m.IsErr AND m.IsDsp, m.s_all, 0)
      dimdata(12,5) = dimdata(12,5) + IIF(m.IsErr AND m.IsProf, m.s_all, 0)

      dimdata(13,5) = dimdata(13,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(14,5) = dimdata(14,5) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(15,5) = dimdata(15,5) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(16,5) = dimdata(16,5) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)

      dimdata(17,5) = dimdata(17,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(18,5) = dimdata(18,5) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(19,5) = dimdata(19,5) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(20,5) = dimdata(20,5) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)

      dimdata(21,5) = dimdata(21,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(22,5) = dimdata(22,5) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(23,5) = dimdata(23,5) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(24,5) = dimdata(24,5) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)

      dimdata(25,5) = dimdata(25,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(26,5) = dimdata(26,5) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(27,5) = dimdata(27,5) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(28,5) = dimdata(28,5) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)

      dimdata(29,5) = dimdata(29,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(30,5) = dimdata(30,5) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(31,5) = dimdata(31,5) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(32,5) = dimdata(32,5) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)

      dimdata(33,5) = dimdata(33,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(34,5) = dimdata(34,5) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(35,5) = dimdata(35,5) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(36,5) = dimdata(36,5) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)

      dimdata(37,5) = dimdata(37,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(38,5) = dimdata(38,5) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(39,5) = dimdata(39,5) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(40,5) = dimdata(40,5) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.6.'), m.s_all, 0)

      dimdata(41,5) = dimdata(41,5) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(42,5) = dimdata(42,5) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(43,5) = dimdata(43,5) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(44,5) = dimdata(44,5) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)

      
      IF BETWEEN(m.cod, 397001, 397056) OR INLIST(m.cod, 397058, 397059)
       dimdata(1,6) = dimdata(1,6) + m.s_all
       dimdata(2,6) = dimdata(2,6) + IIF(m.IsOnk, m.s_all, 0)
       dimdata(3,6) = dimdata(3,6) + IIF(m.IsDsp, m.s_all, 0)
       dimdata(4,6) = dimdata(4,6) + IIF(m.IsProf, m.s_all, 0)

       dimdata(5,6) = dimdata(5,6) + 1
       dimdata(6,6) = dimdata(6,6) + IIF(m.IsOnk, 1, 0)
       dimdata(7,6) = dimdata(7,6) + IIF(m.IsDsp, 1, 0)
       dimdata(8,6) = dimdata(8,6) + IIF(m.IsProf, 1, 0)

       dimdata(9,6)  = dimdata(9,6) + IIF(m.IsErr, m.s_all, 0)
       dimdata(10,6) = dimdata(10,6) + IIF(m.IsErr AND m.IsOnk, m.s_all, 0)
       dimdata(11,6) = dimdata(11,6) + IIF(m.IsErr AND m.IsDsp, m.s_all, 0)
       dimdata(12,6) = dimdata(12,6) + IIF(m.IsErr AND m.IsProf, m.s_all, 0)

       dimdata(13,6) = dimdata(13,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
       dimdata(14,6) = dimdata(14,6) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
       dimdata(15,6) = dimdata(15,6) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
       dimdata(16,6) = dimdata(16,6) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)

       dimdata(17,6) = dimdata(17,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
       dimdata(18,6) = dimdata(18,6) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
       dimdata(19,6) = dimdata(19,6) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
       dimdata(20,6) = dimdata(20,6) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)

      dimdata(21,6) = dimdata(21,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(22,6) = dimdata(22,6) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(23,6) = dimdata(23,6) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(24,6) = dimdata(24,6) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)

      dimdata(25,6) = dimdata(25,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(26,6) = dimdata(26,6) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(27,6) = dimdata(27,6) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(28,6) = dimdata(28,6) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)

      dimdata(29,6) = dimdata(29,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(30,6) = dimdata(30,6) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(31,6) = dimdata(31,6) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(32,6) = dimdata(32,6) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)

      dimdata(33,6) = dimdata(33,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(34,6) = dimdata(34,6) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(35,6) = dimdata(35,6) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(36,6) = dimdata(36,6) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)

      dimdata(37,6) = dimdata(37,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(38,6) = dimdata(38,6) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(39,6) = dimdata(39,6) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(40,6) = dimdata(40,6) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.6.'), m.s_all, 0)

      dimdata(41,6) = dimdata(41,6) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(42,6) = dimdata(42,6) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(43,6) = dimdata(43,6) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(44,6) = dimdata(44,6) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)

      ENDIF 
      
     OTHERWISE && Стацинар
      dimdata(1,7) = dimdata(1,7) + m.s_all
      dimdata(2,7) = dimdata(2,7) + IIF(m.IsOnk, m.s_all, 0)
      dimdata(3,7) = dimdata(3,7) + IIF(m.IsDsp, m.s_all, 0)
      dimdata(4,7) = dimdata(4,7) + IIF(m.IsProf, m.s_all, 0)

      dimdata(5,7) = dimdata(5,7) + 1
      dimdata(6,7) = dimdata(6,7) + IIF(m.IsOnk, 1, 0)
      dimdata(7,7) = dimdata(7,7) + IIF(m.IsDsp, 1, 0)
      dimdata(8,7) = dimdata(8,7) + IIF(m.IsProf, 1, 0)

      dimdata(9,7)  = dimdata(9,7) + IIF(m.IsErr, m.s_all, 0)
      dimdata(10,7) = dimdata(10,7) + IIF(m.IsErr AND m.IsOnk, m.s_all, 0)
      dimdata(11,7) = dimdata(11,7) + IIF(m.IsErr AND m.IsDsp, m.s_all, 0)
      dimdata(12,7) = dimdata(12,7) + IIF(m.IsErr AND m.IsProf, m.s_all, 0)

      dimdata(13,7) = dimdata(13,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(14,7) = dimdata(14,7) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(15,7) = dimdata(15,7) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
      dimdata(16,7) = dimdata(16,7) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)

      dimdata(17,7) = dimdata(17,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(18,7) = dimdata(18,7) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(19,7) = dimdata(19,7) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
      dimdata(20,7) = dimdata(20,7) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)

      dimdata(21,7) = dimdata(21,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(22,7) = dimdata(22,7) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(23,7) = dimdata(23,7) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(24,7) = dimdata(24,7) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)

      dimdata(25,7) = dimdata(25,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(26,7) = dimdata(26,7) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(27,7) = dimdata(27,7) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(28,7) = dimdata(28,7) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)

      dimdata(29,7) = dimdata(29,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(30,7) = dimdata(30,7) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(31,7) = dimdata(31,7) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(32,7) = dimdata(32,7) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)

      dimdata(33,7) = dimdata(33,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(34,7) = dimdata(34,7) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(35,7) = dimdata(35,7) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(36,7) = dimdata(36,7) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)

      dimdata(37,7) = dimdata(37,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(38,7) = dimdata(38,7) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(39,7) = dimdata(39,7) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(40,7) = dimdata(40,7) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.6.'), m.s_all, 0)

      dimdata(41,7) = dimdata(41,7) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(42,7) = dimdata(42,7) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(43,7) = dimdata(43,7) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(44,7) = dimdata(44,7) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)

      IF BETWEEN(m.cod, 200001, 200507)
       dimdata(1,8) = dimdata(1,8) + m.s_all
       dimdata(2,8) = dimdata(2,8) + IIF(m.IsOnk, m.s_all, 0)
       dimdata(3,8) = dimdata(3,8) + IIF(m.IsDsp, m.s_all, 0)
       dimdata(4,8) = dimdata(4,8) + IIF(m.IsProf, m.s_all, 0)

       dimdata(5,8) = dimdata(5,8) + 1
       dimdata(6,8) = dimdata(6,8) + IIF(m.IsOnk, 1, 0)
       dimdata(7,8) = dimdata(7,8) + IIF(m.IsDsp, 1, 0)
       dimdata(8,8) = dimdata(8,8) + IIF(m.IsProf, 1, 0)

       dimdata(9,8)  = dimdata(9,8) + IIF(m.IsErr, m.s_all, 0)
       dimdata(10,8) = dimdata(10,8) + IIF(m.IsErr AND m.IsOnk, m.s_all, 0)
       dimdata(11,8) = dimdata(11,8) + IIF(m.IsErr AND m.IsDsp, m.s_all, 0)
       dimdata(12,8) = dimdata(12,8) + IIF(m.IsErr AND m.IsProf, m.s_all, 0)

       dimdata(13,8) = dimdata(13,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
       dimdata(14,8) = dimdata(14,8) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
       dimdata(15,8) = dimdata(15,8) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)
       dimdata(16,8) = dimdata(16,8) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.1.  ', '5.1.1.','5.1.3.','5.1.4.','5.1.6.'), m.s_all, 0)

       dimdata(17,8) = dimdata(17,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
       dimdata(18,8) = dimdata(18,8) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
       dimdata(19,8) = dimdata(19,8) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)
       dimdata(20,8) = dimdata(20,8) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.2.1.', '5.2.2.'), m.s_all, 0)

      dimdata(21,8) = dimdata(21,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(22,8) = dimdata(22,8) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(23,8) = dimdata(23,8) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)
      dimdata(24,8) = dimdata(24,8) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.1.', '5.3.3.'), m.s_all, 0)

      dimdata(25,8) = dimdata(25,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(26,8) = dimdata(26,8) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(27,8) = dimdata(27,8) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)
      dimdata(28,8) = dimdata(28,8) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.3.2.'), m.s_all, 0)

      dimdata(29,8) = dimdata(29,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(30,8) = dimdata(30,8) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(31,8) = dimdata(31,8) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)
      dimdata(32,8) = dimdata(32,8) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.4.1.', '5.4.2.'), m.s_all, 0)

      dimdata(33,8) = dimdata(33,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(34,8) = dimdata(34,8) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(35,8) = dimdata(35,8) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)
      dimdata(36,8) = dimdata(36,8) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.5.1.', '5.5.2.', '5.5.3.'), m.s_all, 0)

      dimdata(37,8) = dimdata(37,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(38,8) = dimdata(38,8) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(39,8) = dimdata(39,8) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.6.'), m.s_all, 0)
      dimdata(40,8) = dimdata(40,8) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.6.'), m.s_all, 0)

      dimdata(41,8) = dimdata(41,8) + IIF(m.IsErr AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(42,8) = dimdata(42,8) + IIF(m.IsErr AND m.IsOnk AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(43,8) = dimdata(43,8) + IIF(m.IsErr AND m.IsDsp AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)
      dimdata(44,8) = dimdata(44,8) + IIF(m.IsErr AND m.IsProf AND INLIST(m.osn230, '5.7.1.','5.7.2.','5.7.3.','5.7.5.','5.7.6.'), m.s_all, 0)

      ENDIF 
      
    ENDCASE 

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
 USE IN dspcodes
 
 CREATE CURSOR curdata (recid i)
 INSERT INTO curdata (recid) VALUES (0)
 

 **dimdata(15,3)=dimdata(4,3)-(dimdata(5,3)+dimdata(7,3)+dimdata(9,3)+dimdata(11,3)+dimdata(13,3))
 *dimdata(15,4)=dimdata(4,4)-(dimdata(5,4)+dimdata(7,4)+dimdata(9,4)+dimdata(11,4)+dimdata(13,4))
 *dimdata(15,5)=dimdata(4,5)-(dimdata(5,5)+dimdata(7,5)+dimdata(9,5)+dimdata(11,5)+dimdata(13,5))
 *dimdata(15,6)=dimdata(4,6)-(dimdata(5,6)+dimdata(7,6)+dimdata(9,6)+dimdata(11,6)+dimdata(13,6))
 *dimdata(15,7)=dimdata(4,7)-(dimdata(5,7)+dimdata(7,7)+dimdata(9,7)+dimdata(11,7)+dimdata(13,7))
 *dimdata(15,8)=dimdata(4,8)-(dimdata(5,8)+dimdata(7,8)+dimdata(9,8)+dimdata(11,8)+dimdata(13,8))
 *dimdata(15,9)=dimdata(4,9)-(dimdata(5,9)+dimdata(7,9)+dimdata(9,9)+dimdata(11,9)+dimdata(13,9))
 
 **dimdata(16,3)=dimdata(17,3)-(dimdata(6,3)+dimdata(8,3)+dimdata(10,3)+dimdata(12,3)+dimdata(14,3))
 *dimdata(16,4)=dimdata(17,4)-(dimdata(6,4)+dimdata(8,4)+dimdata(10,3)+dimdata(12,4)+dimdata(14,4))
 *dimdata(16,5)=dimdata(17,5)-(dimdata(6,5)+dimdata(8,5)+dimdata(10,5)+dimdata(12,5)+dimdata(14,5))
 *dimdata(16,6)=dimdata(17,6)-(dimdata(6,6)+dimdata(8,6)+dimdata(10,6)+dimdata(12,6)+dimdata(14,6))
 *dimdata(16,7)=dimdata(17,7)-(dimdata(6,7)+dimdata(8,7)+dimdata(10,7)+dimdata(12,7)+dimdata(14,7))
 *dimdata(16,8)=dimdata(17,8)-(dimdata(6,8)+dimdata(8,8)+dimdata(10,8)+dimdata(12,8)+dimdata(14,8))
 *dimdata(16,9)=dimdata(17,9)-(dimdata(6,9)+dimdata(8,9)+dimdata(10,9)+dimdata(12,9)+dimdata(14,9))
 
 FOR n=1 TO 50
  dimdata(n,3)=dimdata(n,4)+dimdata(n,5)+dimdata(n,7)
 ENDFOR 

 IF m.IsAlert
  oAlertMgrr = CREATEOBJECT("VFPAlert.AlertManager")
  pooAlert = oAlertMgrr.NewAlert()
  pooAlert.Alert("Отчет находится: "+pBase+'\'+m.gcPeriod+'\vks.xls', 8, "ВКС ФФОМС",;
  	"Формирование отчета закончено.")
 ENDIF 

 m.llResult = X_Report(pTempl+'\vks_ffoms.xls', pBase+'\'+m.gcperiod+'\vks_'+m.gcPeriod+'.xls', .F.)

RETURN 