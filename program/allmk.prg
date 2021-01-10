PROCEDURE AllMK
 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 m.mmy = PADL(m.tMonth,2,'0') + SUBSTR(STR(m.tYear,4),4,1)
 
 m.t_beg = SECONDS()

 Local Parallel as Parallel of ParallelFox.vcx

 m.IsPparallel = .T.
 TRY 
  Parallel = NewObject("Parallel", "ParallelFox.vcx")
  Parallel.SetWorkerCount(parallel.CPUCount, parallel.CPUCount)

 CATCH 
  m.IsPparallel = .F.
 ENDTRY 

 IF m.IsPparallel = .T.
  Parallel.StartWorkers(FullPath("lpu2smo.exe"),,.t.)
  Parallel.Do("SetEnv", "AllMk.prg", .T., ;
  	m.qcod, m.pBase+'\'+m.gcPeriod+'\nsi', m.gcPeriod, m.qname, m.pTempl, m.tMonth, m.pBase)
 ENDIF 

* TRY 
*  oExcel = GETOBJECT(,"Excel.Application")
* CATCH 
*  oExcel = CREATEOBJECT("Excel.Application")
* ENDTRY 

 SCAN FOR !DELETED()
  MailView.refresh

  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\people.dbf') OR ;
   !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\talon.dbf')
   LOOP 
  ENDIF 

  m.mcod  = mcod
  m.lpuid = lpuid
  m.l_path  = m.pbase+'\'+m.gcperiod+'\'+m.mcod
  m.mk_file = "Mk" + STR(m.lpuid,4) + m.qcod + m.mmy

  IF fso.FileExists(m.l_path+'\'+m.mk_file+'.pdf')
   LOOP 
  ENDIF 

  m.t_1 = SECONDS()

  IF m.IsPparallel = .T.
   =MkPrn2(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)
   *Parallel.Do("MkPrn2","",,pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)
  ELSE 
   =MkPrn2(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)
  ENDIF 
  
  SELECT aisoms

  m.t_2 = SECONDS()
  m.t_mk = m.t_2 - m.t_1
  
  UPDATE aisoms SET t_mk=m.t_mk WHERE mcod=m.mcod

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 


 SET ESCAPE &OldEscStatus

 IF m.IsPparallel = .T.
  Parallel.Wait()
  Parallel.StopWorkers()
  Parallel.Wait()
 ENDIF 

 *oExcel.Quit
 
 m.t_end = SECONDS()
 *m.t_last = ROUND((m.t_end - m.t_beg)/60,2)
 m.t_last = m.t_end - m.t_beg
 
 SELECT aisoms 

 MESSAGEBOX(TRANSFORM(m.t_last,'9999999.99'),0+64,'')

RETURN 

FUNCTION SetEnv(para1, para2, para3, para4, para5, para6, para7)
 PUBLIC  m.qcod, m.fso, m.gcPeriod, m.qname, m.pTempl, m.tMonth, oExcel as Excel.Application , m.pBase
 LOCAL m.lPath

 m.qcod     = para1
 m.lPath    = para2
 m.gcPeriod = para3
 m.qname    = para4
 m.pTempl   = para5
 m.tMonth   = para6
 m.pBase    = para7
 
 SET SAFETY OFF
 fso  = CREATEOBJECT('Scripting.FileSystemObject')
 SET PROCEDURE TO Utils.prg

 *TRY 
 * oExcel = GETOBJECT(,"Excel.Application")
 *CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 *ENDTRY 

 =OpenFile(m.lPath+'\sprlpuxx', "sprlpu", "shar", "lpu_id")
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'shar', 'er_c')
 
RETURN 
