PROCEDURE pdf_test_002
 IF MESSAGEBOX('PDFCREATOR?',64,'')=7
  RETURN 
 ENDIF 
 
 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 
 IF OpenFile(pBase+'\'+gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aioms
  ENDIF 
  RETURN 
 ENDIF 
 
 oQueue = CreateObject('PDFCreator.JobQueue')
 oQueue.Initialize
 
 t_beg = SECONDS()
 n_lpu = 0

 SELECT aisoms
 SCAN 
  m.mcod  = mcod
  m.lpuid = lpuid
  IF !fso.FolderExists(pbase+'\'+gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  MTFile    = 'MT'+STR(m.lpuid,4)+m.qcod+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,4,1)+'.xls'
  MTFilePdf = 'MT'+STR(m.lpuid,4)+m.qcod+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,4,1)+'.pdf'
  IF !fso.FileExists(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+MTFile)
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 

  n_lpu = n_lpu + 1

  oBook = oExcel.Workbooks.Add(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+MTFile)

  obook.PrintOut()
  
  DO WHILE oQueue.Count=0
  ENDDO 
  
  myJob = oQueue.NextJob
  
  IF fso.FileExists(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+MTFilePdf)
   fso.DeleteFile(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+MTFilePdf)
  ENDIF 
  myjob.ConvertTo(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+MTFilePdf)

  DO WHILE !myJob.IsFinished
  ENDDO 
  
  oBook.Close
  
  WAIT CLEAR 

 ENDSCAN 
 USE 

 oQueue.ReleaseCom()
 
 t_end = SECONDS()
 
 MESSAGEBOX('Общее время выполнения: '+TRANSFORM(t_end-t_beg,'9999.99')+CHR(13)+CHR(10)+;
 	'Общее количество МО: '+TRANSFORM(m.n_lpu,'999')+CHR(13)+CHR(10)+;
 	'Среднее время выполнения: '+TRANSFORM((t_end-t_beg)/n_lpu,'9999.99'), 0+64, '')

RETURN 

