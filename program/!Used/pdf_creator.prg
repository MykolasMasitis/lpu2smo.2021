Public Function PrintReportPDFCreator(strReportName As String, strFileName As String) As Boolean
    
    Dim OldPrinterName As String, PDFPrintername As String

    Dim PDF As PDFCreator.PdfCreatorObj
    Dim PDFQueue As New PDFCreator.Queue
    Dim MyJob As PDFCreator.PrintJob
    Dim PDFDevices As PDFCreator.Printers

    Dim objPrn As Printer

    'hold old to the default printer for reset at end of procedure
    OldPrinterName = Application.Printer.DeviceName

    'fire up PDFCreator
    Set PDF = New PDFCreator.PdfCreatorObj

    'get handle on PDFCreator printer
    Set PDFDevices = PDF.GetPDFCreatorPrinters
    PDFPrintername = PDFDevices.GetPrinterByIndex(0)

    'set PDFCreator printer as current printer
    'you could just as easily grab with Application.Printers(“PDFCreator”) 
    For Each objPrn In Application.Printers
        If objPrn.DeviceName = PDFPrintername Then
            Set Application.Printer = objPrn
            Exit For
        End If
    Next

    'Set PDFQueue = New PDFCreator.Queue
    'dimmed as New above so no need for this line above
    PDFQueue.Initialize

    'it’s important that the report is set to use the default printer, not a specific printer
    DoCmd.OpenReport strReportName, acViewNormal

    'wait for output
    Do Until PDFQueue.Count > 0
        DoEvents
    Loop

    'get a handle on the job and save to filename
    Set MyJob = PDFQueue.NextJob
    MyJob.SetProfileSetting “OpenViewer”, “False”
    MyJob.ConvertTo strFileName

    'waiting 
    Do Until MyJob.IsFinished
        DoEvents
    Loop

ProcExit:
    'reset default printer
    For Each objPrn In Application.Printers
        If objPrn.DeviceName = OldPrinterName Then
            Set Application.Printer = objPrn
            Exit For
        End If
    Next

    'close all objects
    Set objPrn = Nothing
    Set MyJob = Nothing

    PDFQueue.ReleaseCom
    Set PDFQueue = Nothing
    Set PDFDevices = Nothing
    Set PDF = Nothing


End Function


PROCEDURE PDFCreator

DECLARE Sleep IN kernel32 INTEGER
LOCAL lcRepName,lcFileName,lcFolder,oPDFC,lcOldDefaPrint,DefaultPrinter,laf[1],lnCount

CREATE CURSOR cc (ii I)
APPEND BLANK
GO top
lcFolder = ADDBS(JUSTPATH(SYS(16)))
lcRepName = m.lcFolder  + "cc.frx"
CREATE REPORT (m.lcRepName) FROM cc
lcFileName = "cc.pdf"

IF ADIR(laf,m.lcFolder+m.lcFileName) > 0
	ERASE (m.lcFolder+m.lcFileName)
	IF ADIR(laf,m.lcFolder+m.lcFileName) > 0
		MESSAGEBOX("Can't create " + m.lcFolder+m.lcFileName,16,"Erase the old PDF")
		RETURN
	ENDIF
ENDIF
lcOldDefaPrint = Alltrim(Set('PRINTER', 2))

oPDFC  = CREATEOBJECT("PDFCreator.clsPDFCreator","pdfcreator")
oPDFC.cStart("/NoProcessingAtStartup")
oPDFC.cOption("UseAutosave") = 1
oPDFC.cOption("UseAutosaveDirectory") = 1
oPDFC.cOption("AutosaveFormat") = 0
* 0 = PDF format
* 1 = PNG
* 2 = JPEG
* 3 = BMP
* 4 = PCX
* 5 = TIFF
DefaultPrinter = oPDFC.cDefaultprinter
oPDFC.cDefaultprinter = "pdfcreator"
oPDFC.cClearCache
ReadyState = 0
oPDFC.cOption("AutosaveFilename") = m.lcFileName
oPDFC.cOption("AutosaveDirectory") = m.lcFolder
oPDFC.cprinterstop=.F.

SET PRINTER TO NAME (oPDFC.cDefaultprinter) && Fix this 
REPORT FORM (m.lcRepName) NOCONSOLE TO PRINTER 

lnCount = 0
DO WHILE ADIR(laf,m.lcFolder+m.lcFileName) = 0 AND m.lnCount <= 40
    sleep(50)
    lnCount = m.lnCount + 1
ENDDO

oPDFC.cDefaultprinter = DefaultPrinter
****************************
* This is the the new line *
****************************
oPDFC.cOption("UseAutosave") = 0
****************************
****************************
oPDFC.cClearCache
RELEASE m.oPDFC
Set Printer To Name (m.lcOldDefaPrint) 

**********************************************
* Small procedure to reset the autosave mode *
**********************************************
oPDFC  = CREATEOBJECT("PDFCreator.clsPDFCreator","pdfcreator")
oPDFC.cStart("/NoProcessingAtStartup")
oPDFC.cOption("UseAutosave") = 0
oPDFC.cClearCache
RELEASE m.oPDFC 