FUNCTION PrintReportPDFCreator(strReportName, strFileName)

 LOCAL oExcel as Excel.Application
 oExcel = CREATEOBJECT('Excel.Application')

 LOCAL OldPrinterName as Character, PDFPrintername As Character
 
 LOCAL PDF As PDFCreator.PdfCreatorObj
 *LOCAL PDFQueue As PDFCreator.JobQueue
 *LOCAL MyJob As PDFCreator.PrintJob
 *LOCAL PDFDevices As PDFCreator.Printers
 *LOCAL objPrn As Printer
  
  * Hhold old to the default printer for reset at end of procedure
  OldPrinterName = oExcel.ActivePrinter

  * fire up PDFCreator
  PDF = CREATEOBJECT('PDFCreator.PdfCreatorObj')

  * get handle on PDFCreator printer
  PDFDevices     = PDF.GetPDFCreatorPrinters
  PDFPrintername = PDFDevices.GetPrinterByIndex(0)

  * set PDFCreator printer as current printer
  * you could just as easily grab with Application.Printers(“PDFCreator”) 
  *SET PRINTER TO NAME PDFPrintername
  oExcel.ActivePrinter = PDFPrintername
  
  PDFQueue = CreateObject('PDFCreator.JobQueue')
  *dimmed as New above so no need for this line above
  PDFQueue.Initialize

  * it’s important that the report is set to use the default printer, not a specific printer
  *  DoCmd.OpenReport strReportName, acViewNormal

  * wait for output
  *  Do Until PDFQueue.Count > 0
  *      DoEvents
  *  Loop

  * get a handle on the job and save to filename
  MyJob = PDFQueue.NextJob

  MyJob.SetProfileSetting([OpenViewer],Iif(mes=6,[true],[false]))
  *MyJob.SetProfileByGuid(this._SaveAs) && _SaveAs is a property that can be set as "DefaultGuid", "PdfaGuid", etc, described in help file
  MyJob.SetProfileSetting([ShowOnlyErrorNotifications],[true])
  MyJob.SetProfileSetting([ShowAllNotifications],[false])
  MyJob.SetProfileSetting([ShowProgress],[false])
  MyJob.SetProfileSetting([ShowQuickActions],[false])
  
  MyJob.ConvertTo(strFileName)

  *waiting 
   DO WHILE !MyJob.IsFinished
   ENDDO 

  *reset default printer
  *  For Each objPrn In Application.Printers
  *      If objPrn.DeviceName = OldPrinterName Then
  *          Set Application.Printer = objPrn
  *          Exit For
  *      End If
  *  Next

  * close all objects
  *  Set objPrn = Nothing
  *  Set MyJob = Nothing

 oPDF.ReleaseCom

 RELEASE PDFQueue, PDFDevices, PDF
 
ENDFUNC 

oPDF=CreateObject([PDFCreator.JobQueue])
oPDF.Initialize()
oPDF.WaitForJob(10)

oJob=oPDF.NextJob
oJob.SetProfileByGuid(this._SaveAs) && _SaveAs is a property that can be set as "DefaultGuid", "PdfaGuid", etc, described in help file
oJob.SetProfileSetting([OpenViewer],Iif(mes=6,[true],[false]))
oJob.SetProfileSetting([ShowOnlyErrorNotifications],[true])
oJob.SetProfileSetting([ShowAllNotifications],[false])
oJob.SetProfileSetting([ShowProgress],[false])
oJob.SetProfileSetting([ShowQuickActions],[false])
oJob.ConvertTo(lcFile)
oPDF.ReleaseCom()