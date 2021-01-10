PROCEDURE pdf_test_001
 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 

  *Excel.ActivePrinter='PDFCreator (Ne05:)'
  oBook = oexcel.Workbooks.Add('d:\lpu2smo\001.xlsx')
  
  oQueue = CreateObject('PDFCreator.JobQueue')
  oQueue.Initialize
  
  obook.PrintOut()
  
  DO WHILE oQueue.Count=0
  ENDDO 
  
  myJob = oQueue.NextJob
  *myJob.SetProfileByName("DefaultGuid")
  myjob.ConvertTo('d:\lpu2smo\001.pdf')
  
  DO WHILE !myJob.IsFinished
  ENDDO 
  
  oBook.Close
  
  MESSAGEBOX('OK!',0+64,'')
  
  oQueue.ReleaseCom()

 
RETURN 