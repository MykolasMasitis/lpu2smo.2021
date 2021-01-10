*****************************************************************************
*!* ����� FoxPDFText. �������� PDF ����� �� 100% ������ VFP               *!*
*!* ������ 1.0  25/07/2005                                                *!*
*!* �����: Loic Carrere - loic.carrere@asi-concept.fr                     *!*
*!* ������� � ������������ � ��������  05/12/2005                         *!*
*!* ������� �������: ������� ����� ���������� - rvc44@rambler.ru          *!*
*!* ������������ �� www.foxclub.ru � ������� ���������� ������            *!*
*!* �������������� �������� ����������:                                   *!*
*!* PDF Reference. Adobe Portable Document Format (pdfreference.pdf)      *!*
*!* http://partners.adobe.com/public/developer/pdf/index_reference.html#5 *!*
*!* ������ �������������:                                                 *!*
*!* 	If AT("FoxPDFText", SET("PROCEDURE")) = 0                         *!*
*!* 		SET PROCEDURE TO FoxPDFText ADDITIVE                      *!*
*!* 	EndIf                                                             *!*
*!* 	oPDF = CreateObject("FoxPDFText")                                 *!*
*!* 	oPDF.NewPDF("C:\������ � ������.pdf")                             *!*
*!* 	oPDF.NewPage()                                                    *!*
*!* 	oPDF.WriteText("Any text!", .F.)   			          *!*
*!* 	oPDF.ClosePDF()                                                   *!*
*****************************************************************************

&& ���� �������
#DEFINE pdfRegular    1
#DEFINE pdfItalic     2
#DEFINE pdfBold       3
#DEFINE pdfBoldItalic 4

&& ������� ���������
#DEFINE pdfA3         1
#DEFINE pdfA4         2
#DEFINE pdfA5         3
#DEFINE pdfTABLOID    4
#DEFINE pdfLEDGER     5
#DEFINE pdfLEGAL      6
#DEFINE pdfSTATEMENT  7
#DEFINE pdfEXECUTIVE  8
#DEFINE pdfCUSTOMSIZE 9

#DEFINE LF  CHR(10)
#DEFINE CR  CHR(13)

DEFINE CLASS FoxPDFText as Custom

   sProducer = ""          && ��� Producer ����� PDF
   sTitle = ""             && ��� Title ����� PDF
   sSubject = ""           && ��� Subject ����� PDF
   sAuthor = ""            && ��� Author ����� PDF
   nPaperSize = pdfA4      && ������ ���������
   bLandScape = .F.        && ����� �������
   nTopMargin = 50         && ������� ���� ��������
   nLeftMargin = 50        && ����� ���� ��������
   nFontSize = 10          && ������ ������
   nFontType = pdfRegular  && ��� ������
   nVertSpace = 12         && ������ �������
   nDegree = 0             && ���� ��������
   nCustomHeight = 0       && ��������� ������ ���������
   nCustomWidth = 0        && ��������� ������ ���������

   PROTECTED nStartStream,;
             nObject,;               && ����� ��������
             tbPages(1),;            && ������� �������
             tbObjectsRefs(1),;      && ������� ������ �� ������� (�������� ������� ������� ������� ����� PDF)
             nPaperWidth,;           && ������ ������
             nPaperHeight,;          && ������ ������
             nPDfHandle,;            && ���������� ����� PDF
             nPages,;                && ����� �������
             sObj,;                  && ��������� ����������, ���������� ������ �� ������� ������
             sFilePath               && ���� � ����� PDF


   && �������� ������ ����� PDF (���������� ���������� �����)
   FUNCTION NewPDF(sPath) as long
      PRIVATE nLen, nBuffer

      This.sFilePath = sPath
      DO CASE
         CASE This.nPaperSize = pdfA3
              This.nPaperWidth = 842
              This.nPaperHeight = 1190
         CASE This.nPaperSize = pdfA4
              This.nPaperWidth = 595
              This.nPaperHeight = 842
         CASE This.nPaperSize = pdfA5
              This.nPaperWidth = 421
              This.nPaperHeight = 595
         CASE This.nPaperSize = pdfTABLOID
              This.nPaperWidth = 792
              This.nPaperHeight = 1224
         CASE This.nPaperSize = pdfLEDGER
              This.nPaperWidth = 1224
              This.nPaperHeight = 792
         CASE This.nPaperSize = pdfLETTER
              This.nPaperWidth = 612
              This.nPaperHeight = 1008
         CASE This.nPaperSize = pdfSTATEMENT
              This.nPaperWidth = 396
              This.nPaperHeight = 612
         CASE This.nPaperSize = pdfEXECUTIVE
              This.nPaperWidth = 540
              This.nPaperHeight = 720
         CASE This.nPaperSize = pdfCUSTOMSIZE
              This.nPaperWidth = nCustomHeight
              This.nPaperHeight = nCustomWidth
      ENDCASE

      If This.bLandScape
         nBuffer = This.nPaperWidth
         This.nPaperWidth = This.nPaperHeight
         This.nPaperHeight = nBuffer
      EndIf

      This.nPDfHandle = FCREATE(This.sFilePath)
      If This.nPDfHandle > 0
         FPUTS(This.nPDfHandle, "%PDF-1.2" + LF + "%����" + LF)
         This.nObject = 1
         This.sObj = "/CreationDate (D:" + AllT(Str(YEAR(DATE()))) +;
                                      PADL(AllT(Str(MONTH(DATE()))), 2, '0') +;
                                      PADL(AllT(Str(DAY(DATE()))), 2, '0') +;
                                      PADL(AllT(Str(HOUR(DATETIME()))), 2, '0') +;
                                      PADL(AllT(Str(MINUTE(DATETIME()))), 2, '0') +;
                                      PADL(AllT(Str(SEC(DATETIME()))), 2, '0') + ")"

         This.sObj = This.sObj + LF + "/Creator (FoxPDF 1.0)"
         This.sObj = This.sObj + LF + "/Producer (" + This.sProducer + ")"
         This.sObj = This.sObj + LF + "/Title ("    + This.sTitle    + ")"
         This.sObj = This.sObj + LF + "/Subject ("  + This.sSubject  + ")"
         This.sObj = This.sObj + LF + "/Author ("   + This.sAuthor   + ")"

         && ������ �������
         This.WriteObject()

         && ����� ������� (regular)
         This.nObject = 4

		 && (Required) The type of PDF object that this dictionary describes; must be
		 && Font for a font dictionary
         This.sObj = "/Type /Font"  					 && �������� ����

		 && (Required) The type of font; must be Type1 for a Type 1 font
         This.sObj = This.sObj + LF + "/Subtype /Type1"  && �������� �������

		 && (Required in PDF 1.0; optional otherwise) The name by which this font is referenced
		 && in the Font subdictionary of the current resource dictionary.
         This.sObj = This.sObj + LF + "/Name /F1"  		 && F1 ��� ������� ���

		 && (Optional) A specification of the font�s character encoding, if different from
		 && dictionary its built-in encoding. The value of Encoding may be either the name of a predefined
		 && encoding (MacRomanEncoding, MacExpertEncoding, or WinAnsi-
		 && Encoding, as described in Appendix D) or an encoding dictionary that
		 && specifies differences from the font�s built-in encoding or from a specified predefined
		 && encoding (see Section 5.5.5, �Character Encoding�).
         This.sObj = This.sObj + LF + "/Encoding 8 0 R"

		 && (Required) The PostScript name of the font. For Type 1 fonts, this is usually
		 && the value of the FontName entry in the font program; for more information,
		 && see Section 5.2 of the PostScript Language Reference, Third Edition. The Post-
		 && Script name of the font can be used to find the font�s definition in the viewer
		 && application or its environment. It is also the name that will be used when
		 && printing to a PostScript output device.
         This.sObj = This.sObj + LF + "/BaseFont /Courier"
         This.WriteObject()

         && ����� ������ (italic)
         This.nObject = This.nObject + 1
         This.sObj = "/Type /Font"
         This.sObj = This.sObj + LF + "/Subtype /Type1"
         This.sObj = This.sObj + LF + "/Name /F2"
         This.sObj = This.sObj + LF + "/Encoding 8 0 R"
         This.sObj = This.sObj + LF + "/BaseFont /Courier-Oblique"
         This.WriteObject()

         && ����� ���������� (bold)
         This.nObject = This.nObject + 1
         This.sObj = "/Type /Font"
         This.sObj = This.sObj + LF + "/Subtype /Type1"
         This.sObj = This.sObj + LF + "/Name /F3"
         This.sObj = This.sObj + LF + "/Encoding 8 0 R"
         This.sObj = This.sObj + LF + "/BaseFont /Courier-Bold"
         This.WriteObject()

         && ����� ���������� ������ (bold-italic)
         This.nObject = This.nObject + 1
         This.sObj = "/Type /Font"
         This.sObj = This.sObj + LF + "/Subtype /Type1"
         This.sObj = This.sObj + LF + "/Name /F4"
         This.sObj = This.sObj + LF + "/Encoding 8 0 R"
         This.sObj = This.sObj + LF + "/BaseFont /Courier-BoldOblique"
         This.WriteObject()

         && ��������� ������
         This.nObject = This.nObject + 1
         This.sObj = "/Type /Encoding"
         This.sObj = This.sObj + LF + "/BaseEncoding /WinAnsiEncoding"
         This.WriteObject()

         && ������ ������
         This.nObject = This.nObject + 1
         This.sObj = " /Font << /F1 4 0 R /F2 5 0 R /F3 6 0 R /F4 7 0 R >>"
         This.sObj = This.sObj + LF + " /ProcSet [ /PDF /Text ]"
         This.WriteObject()
      EndIf

      RETURN This.nPDfHandle
   ENDFUNC


   && �������� � ������ ����� � ������� native PDF
   PROCEDURE ClosePDF()
      PRIVATE nCpt, nOffset

      If This.nPDfHandle > 0
          This.EndPage()

          && ������ ��������
          This.sObj = "2 0 obj"
          This.sObj = This.sObj + LF + "<<"
          This.sObj = This.sObj + LF + "/Type /Catalog"
          This.sObj = This.sObj + LF + "/Pages 3 0 R"
          This.sObj = This.sObj + LF + "/PageLayout /OneColumn"
          This.sObj = This.sObj + LF + ">>"
          This.sObj = This.sObj + LF + "endobj"
          nOffset = This.WriteTextObjet()
          This.tbObjectsRefs(2) = AllT(Str(nOffset))

          && ���������� ����������
          This.sObj = "3 0 obj"
          This.sObj = This.sObj + LF + "<<"
          This.sObj = This.sObj + LF + "/Type /Pages"
          This.sObj = This.sObj + LF + "/Count " + AllT(Str(This.nPages))
          This.sObj = This.sObj + LF + "/MediaBox [ 0 0 " + AllT(Str(This.nPaperWidth)) + " " + AllT(Str(This.nPaperHeight)) + " ]"
          This.sObj = This.sObj + LF + "/Kids [ "

          FOR nCpt = 1 To This.nPages
              This.sObj = This.sObj + AllT(Str(This.tbPages(nCpt))) + " 0 R "
          ENDFOR
          This.sObj = This.sObj + "]"
          This.sObj = This.sObj + LF + ">>"
          This.sObj = This.sObj + LF + "endobj"
          nOffset = This.WriteTextObjet()
          This.tbObjectsRefs(3) = AllT(Str(nOffset))

          && ������������ ������
          This.nObject = This.nObject + 1
          This.sObj = "xref"
          This.sObj = This.sObj + LF + "0 " + AllT(Str(This.nObject))
          This.sObj = This.sObj + LF + "0000000000 65535 f "
  
          FOR nCpt = 1 To This.nObject - 1
              This.sObj = This.sObj + CR + padl(This.tbObjectsRefs(nCpt), 10, '0') + " 00000 n "
          ENDFOR
          This.sObj = This.sObj + CR + "trailer"
          nOffset = This.WriteTextObjet()

          && Trailer
          This.sObj = "<<"
          This.sObj = This.sObj + LF + "/Size " + AllT(Str(This.nObject))
          This.sObj = This.sObj + LF + "/Root 2 0 R"
          This.sObj = This.sObj + LF + "/Info 1 0 R"
          This.sObj = This.sObj + LF + ">>"
          This.sObj = This.sObj + LF + "startxref"
          This.sObj = This.sObj + LF + AllT(Str(nOffset))
          This.sObj = This.sObj + LF + "%%EOF"
          This.WriteTextObjet()
          FCLOSE(This.nPDfHandle)
          This.nPDfHandle = -1
      EndIf
   ENDPROC

   && ��������� ����� ��������
   PROCEDURE NewPage()
      PRIVATE nLen, nOffset

      && �������� ���������� ��������
      If This.nPages > 0
          This.EndPage()
      EndIf

      && ������� ��������
      This.nObject = This.nObject + 1
      This.sObj = "/Type /Page"
      This.sObj = This.sObj + LF + "/Parent 3 0 R"
      This.sObj = This.sObj + LF + "/Resources 9 0 R"
      This.sObj = This.sObj + LF + "/Contents " + AllT(Str(This.nObject + 1)) + " 0 R"
      This.WriteObject()

      && ������ �� ������ ��������
      This.nPages = This.nPages + 1
      DIMENSION This.tbPages(This.nPages)
      This.tbPages(This.nPages) = This.nObject

      && ����������� �������, ����������� ����� ��������
      This.nObject = This.nObject + 1
      This.sObj = AllT(Str(This.nObject)) + " 0 obj"
      This.sObj = This.sObj + LF + "<<"
      This.sObj = This.sObj + LF + "/Length " + AllT(Str(This.nObject + 1)) + " 0 R"
      This.sObj = This.sObj + LF + ">>"
      This.sObj = This.sObj + LF + "stream"
      This.sObj = This.sObj + LF + "BT"
      nOffset = This.WriteTextObjet()
      DIMENSION This.tbObjectsRefs(This.nObject)
      This.tbObjectsRefs(This.nObject) = AllT(Str(nOffset))

      && ����������� ������� ������ ������
      This.nStartStream = nOffset + Len(This.sObj) - 5

      This.sObj = "/F" + AllT(Str(This.nFontType)) + " " + AllT(Str(This.nFontSize)) + " Tf"
      This.WriteTextObjet()

      This.SetOrigin(This.nLeftMargin, This.nPaperHeight - This.nTopMargin, This.nDegree)
   ENDPROC


   && ������ ������ sText � ���� PDF
   PROCEDURE WriteText(sText As String, bNewRow As Boolean)
      PRIVATE sRestoreFont

      This.sObj = "/F" + AllT(Str(This.nFontType)) + " " + AllT(Str(This.nFontSize)) + " Tf" + LF
      sRestoreFont = LF + "/F" + AllT(Str(This.nFontType)) + " " + AllT(Str(This.nFontSize)) + " Tf"

      If bNewRow Then
          This.sObj = This.sObj + "T* "
      EndIf

      This.sObj = This.sObj + "(" + sText + ") Tj"
      This.sObj = This.sObj + sRestoreFont
      This.WriteTextObjet()

   ENDPROC

   && ������ �������������� ������ sText � ���� PDF
   PROCEDURE WriteMultiLineText(sText As String, bNewRow As Boolean)
      PRIVATE sStartObj,;
              sRestoreFont,;
              nCpt,;
              sLine,;
              nBegin,;
              nlen

      nLineCount = OCCURS(CR + LF, sText)
      sStartObj =  "/F" + AllT(Str(This.nFontType)) + " " + AllT(Str(This.nFontSize)) + " Tf" + LF
      sRestoreFont = LF + "/F" + AllT(Str(This.nFontType)) + " " + AllT(Str(This.nFontSize)) + " Tf"
      If bNewRow Then
         sStartObj = sStartObj + "T* "
      EndIf
      FOR nCpt = 1 TO nLineCount
          DO CASE
             Case nCpt = 1
                  nBegin = 1
                  nLen = AT(CR + LF, sText, nCpt)
             Case nCpt = nLineCount
                  nBegin = AT(CR + LF, sText, nCpt) + 2
                  nLen = LEN(sText) - nBegin + 2
             Otherwise
                  nBegin = AT(CR + LF, sText, nCpt - 1) + 2
                  nLen = AT(CR + LF, sText, nCpt) - nBegin
          ENDCASE
          sLine = SubStr(sText, nBegin, nLen)
          This.sObj = sStartObj + "(" + sLine + ") Tj"
          This.sObj = This.sObj + sRestoreFont
          This.WriteTextObjet()
      ENDFOR
   ENDPROC

   && ��������� ��������� �������, � ��� �� ������������� ���� ��������
   PROTECTED PROCEDURE SetOrigin(nStartX As Long, nStartY As Long, nDegree As Long)
      PRIVATE a, b, c, d, nPi

      nPi = 3.141592654
      a = Cos(nPI * nDegree / 180)
      b = Sin(nPI * nDegree / 180)
      c = -b
      d = a

      This.sObj = AllT(Str(a, 3))    + " " + ;
                  AllT(Str(b, 3))    + " " + ;
                  AllT(Str(c, 3))    + " " + ;
                  AllT(Str(d, 3))    + " " + ;
                  AllT(Str(nStartX)) + " " + ;
                  AllT(Str(nStartY)) + " Tm"
  
      This.WriteTextObjet()
      This.sObj = AllT(Str(This.nVertSpace)) + " TL"

      This.WriteTextObjet()
   ENDPROC

   && ���������� ������� � ���� PDF
   PROTECTED PROCEDURE WriteObject()
      PRIVATE nOffset, sObjectBuf

      sObjectBuf = This.sObj
      This.sObj = AllT(Str(This.nObject)) + " 0 obj"
      This.sObj = This.sObj + LF + "<<"
      This.sObj = This.sObj + LF + sObjectBuf
      This.sObj = This.sObj + LF + ">>"
      This.sObj = This.sObj + LF + "endobj"
      nOffset = This.WriteTextObjet()
      DIMENSION This.tbObjectsRefs(This.nObject)
      This.tbObjectsRefs(This.nObject) = AllT(Str(nOffset))
   ENDPROC

   && ������ �������, � ����� ������ � ���� pdf � ������� ��� �������
   PROTECTED FUNCTION WriteTextObjet() As Long
      PRIVATE nSeek
   
      nSeek = FSEEK(This.nPDfHandle, 0, 2)
      FWRITE(This.nPDfHandle, This.sObj + LF)
      RETURN nSeek
   ENDFUNC

   && ����������� ��������� ��������
   PROTECTED PROCEDURE EndPage()
      PRIVATE nLen, nOffset

      This.sObj = "ET"
      This.sObj = This.sObj + LF + "endstream"
      This.sObj = This.sObj + LF + "endobj"

      && ������ ����� ��������
      nLen = This.WriteTextObjet() - This.nStartStream
      This.nObject = This.nObject + 1
      This.sObj = AllT(Str(This.nObject)) + " 0 obj"
      This.sObj = This.sObj + LF + AllT(Str(nLen))
      This.sObj = This.sObj + LF + "endobj"
      nOffset = This.WriteTextObjet()
      DIMENSION This.tbObjectsRefs(This.nObject)
      This.tbObjectsRefs(This.nObject) = AllT(Str(nOffset))
   ENDPROC

   && ��� ����� init
   PROTECTED PROCEDURE Init()
      This.nObject = 0
      This.nPages = 0
      This.nStartStream = 0
      This.nPaperWidth = 595
      This.nPaperHeight = 842
   ENDPROC

ENDDEFINE
