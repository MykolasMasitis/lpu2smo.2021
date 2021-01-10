PROCEDURE medpack_sax
 IF MESSAGEBOX(CHR(13)+CHR(10)+'КОНВЕРТИРОВАТЬ medicament_man_pack.xml?'+CHR(13)+CHR(10),;
 	4+32,'medicament_man_pack.xml')=7
  RETURN 
 ENDIF 
 
 pUpdDir = fso.GetParentFolderName(pbin)+'\UPDATE'
 IF !fso.FolderExists(pUpdDir)
  fso.CreateFolder(pUpdDir)
 ENDIF 

 SET DEFAULT TO (pUpdDir)
 csprfile = ''
 csprfile = GETFILE('xml')
 IF EMPTY(csprfile)
  MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ НИЧЕГО НЕ ВЫБРАЛИ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 CREATE CURSOR curss(R_UP C(10), NAME c(100), NNAME C(100), PMP_ID C(10), TYPE C(50), ;
 	QTY_VALUE n(11,2), QTY_UNIT c(10), MASS_VALUE n(6,2), MASS_UNIT C(10), VOL_VALUE n(6,2), VOL_UNIT C(10))
 SELECT curss

 *LOCAL loWriter as MSXML2.MXXMLWriter40, loAtrs as MSXML2.SAXAttributes40
 LOCAL loWriter as MSXML2.MXXMLWriter30, loAtrs as MSXML2.SAXAttributes30
 *loWriter = CREATEOBJECT("Msxml2.MXXMLWriter.4.0")
 loWriter = CREATEOBJECT("Msxml2.MXXMLWriter.3.0")
 IF VARTYPE(loWriter) # 'O'
  RETURN .F.
 ENDIF
 loAtrs = CREATEOBJECT("Msxml2.SAXAttributes.3.0")
 IF VARTYPE(loAtrs) # 'O'
  RETURN .F.
 ENDIF
 
 loWorker = NEWOBJECT('custom')

 loWorker.AddProperty('oWriter', loWriter)
 loWorker.AddProperty('oAtrs', loAtrs)
 loWorker.AddObject('oContentHandlerImpl', 'ContentHandlerImpl')
 IF TYPE('loWorker.oContentHandlerImpl') # 'O' OR ISNULL(loWorker.oContentHandlerImpl)
  RETURN .F.
 ENDIF 
 
 LOCAL loException as Exception, loRdr As SAXXMLReader40

 TRY
	*-- Create SAXXMLReader
	*loRdr = CREATEOBJECT("Msxml2.SAXXMLReader.4.0")
	loRdr = CREATEOBJECT("Msxml2.SAXXMLReader.3.0")
	loRdr.contentHandler = loWorker.oContentHandlerImpl && Set the content handler for the reader.
	loRdr.errorHandler = loWorker.oContentHandlerImpl && Set the error handler for the reader.
	loWorker.oContentHandlerImpl.oContentHandler = loWorker.oWriter && Set the writer for the content handler
	loWorker.oContentHandlerImpl.oErrorHandler = loWorker.oWriter && Set the error handler for the writer
	*
	*-- Configure output for the writer.
	loWorker.oWriter.indent = .T.
	loWorker.oWriter.standalone = .T.
	loWorker.oWriter.output = ""
	loWorker.oWriter.omitXMLDeclaration = .T.
	*
	*-- Set FilterCriteria
	loWorker.oContentHandlerImpl.FilterCriteria = ALLTRIM('ROW') && Важный параметр!
	*
	*-- parseURL
	loRdr.parseURL(csprfile)
	*
	*-- Set result
	*loWorker.edit1.Value = loWorker.oWriter.output  && сейчас мне это не надо!
 CATCH TO loException
		MESSAGEBOX("**** Error **** ";
			+ LTRIM(STR(loException.ErrorNo)) ;
			+ " : " + loException.Message;
			,16, _SCREEN.Caption)
 FINALLY
	STORE NULL TO ;
		loException, loRdr
 ENDTRY

 
 SELECT curss 
 REPLACE ALL name WITH nname
 ALTER table curss drop COLUMN nname
 INDEX on r_up TAG r_up
 SET ORDER TO r_up

 *SET SAFETY OFF
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medpack.dbf')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medpack.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medpack.cdx')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medpack.cdx')
 ENDIF 

 COPY TO &pBase/&gcPeriod/nsi/medpack WITH cdx 
 *SET SAFETY ON
 USE 
 
 IF TYPE('loWorker.oContentHandlerImpl') = 'O'
	loWorker.RemoveObject('oContentHandlerImpl') 
 ENDIF
 IF TYPE('loWorker.oWriter') = 'O'
	loWorker.oWriter = NULL
 ENDIF
 IF TYPE('loWorker.oAtrs') = 'O'
	loWorker.oAtrs = NULL
 ENDIF
 RELEASE loWorker

 *MESSAGEBOX('ФАЙЛ СФОРМИРОВАН!',0+64,'medpack')
 MESSAGEBOX('Надеюсь, Вы оценили'+CHR(13)+CHR(10)+;
 	'разницу в скорости загрузки файла...'+CHR(13)+CHR(10)+;
 	'Поверьте, это было не просто...'+CHR(13)+CHR(10);
 	,0+64,'SAX')
 

RETURN 

DEFINE CLASS ContentHandlerImpl AS session
	oContentHandler = NULL	&& the content handler object
	oErrorHandler = NULL	&& the error handler object
	errorHappen = .F.		&& Flag to indicate if the error handler has thrown a fatal error.
	FilterTrue = .F.		&& Flag to indicate if the element is in scope.
	FilterCriteria = ""		&& String to hold the element name
	
	CurrentField = ""

    PMP_ID     = ""
    R_UP       = ""
    NNAME      = ""
    TYPE       = ""
    QTY_VALUE  = 0
    QTY_UNIT   = ""
    MASS_VALUE = 0
    MASS_UNIT  = ""
    VOL_VALUE  = 0
    VOL_UNIT   = ""
    nrecs      = 0 

	*IMPLEMENTS IVBSAXContentHandler EXCLUDE IN "msxml4.dll"
	*IMPLEMENTS IVBSAXContentHandler EXCLUDE IN "msxml3.dll"
	IMPLEMENTS IVBSAXContentHandler EXCLUDE IN "msxml6.dll"
	*IMPLEMENTS IVBSAXErrorHandler EXCLUDE IN "msxml4.dll"
	*IMPLEMENTS IVBSAXErrorHandler EXCLUDE IN "msxml3.dll"
	IMPLEMENTS IVBSAXErrorHandler EXCLUDE IN "msxml6.dll"

	PROCEDURE Destroy
		STORE NULL TO ;
			this.oContentHandler, this.oErrorHandler
	ENDPROC

	PROTECTED PROCEDURE Initialize
		this.errorHappen = .F.
		this.FilterTrue = .F.
	ENDPROC

	*////////////////////////////////////
	*// IMPLEMENTS IVBSAXContentHandler
	*//
	PROCEDURE IVBSAXContentHandler_put_documentLocator(RHS As VARIANT) AS VARIANT;
			HELPSTRING "Receive an object for locating the origin of SAX document events."
		this.Initialize()
	ENDPROC

	PROCEDURE IVBSAXContentHandler_startDocument() AS VOID;
			HELPSTRING "Receive notification of the beginning of a document."
			*MESSAGEBOX('The beginning of a document',0+64,'sax')
		* add user code here
	ENDPROC

	PROCEDURE IVBSAXContentHandler_endDocument() AS VOID;
			HELPSTRING "Receive notification of the end of a document."
		* add user code here
			*MESSAGEBOX('The end of a document',0+64,'sax')
	ENDPROC

	PROCEDURE IVBSAXContentHandler_startPrefixMapping(strPrefix AS STRING @, strURI AS STRING @) AS VOID;
			HELPSTRING "Begin the scope of a prefix-URI Namespace mapping."
		* add user code here
	ENDPROC

	PROCEDURE IVBSAXContentHandler_endPrefixMapping(strPrefix AS STRING @) AS VOID;
			HELPSTRING "End the scope of a prefix-URI mapping."
		* add user code here
	ENDPROC

	PROCEDURE IVBSAXContentHandler_startElement(strNamespaceURI AS STRING @, strLocalName AS STRING @, strQName AS STRING @, oAttributes AS VARIANT) AS VOID;
			HELPSTRING "Receive notification of the beginning of an element."
		IF strLocalName == this.FilterCriteria
        	this.FilterTrue = .T.
		ENDIF
		IF this.FilterTrue
			this.oContentHandler.startElement(@strNamespaceURI, @strLocalName, @strQName, oAttributes)

			DO CASE 
			 CASE strLocalName = 'ROW'
			  this.PMP_ID     = ""
			  this.R_UP       = ""
			  this.NNAME      = ""
			  this.TYPE       = ""
			  this.QTY_VALUE  = 0
			  this.QTY_UNIT   = ""
			  this.MASS_VALUE = 0
			  this.MASS_UNIT  = ""
			  this.VOL_VALUE  = 0
			  this.VOL_UNIT   = ""

			 CASE strLocalName = 'COLUMN'

			  m.attrName = UPPER(ALLTRIM(oAttributes.getValue(0)))
			  THIS.CurrentField = m.attrName
			 OTHERWISE 
			ENDCASE 
			*MESSAGEBOX('oAttributes.getValue='+oAttributes.getValue(0),0+64,'startElement')
		ENDIF
*				'
	ENDPROC

	PROCEDURE IVBSAXContentHandler_endElement(strNamespaceURI AS STRING @, strLocalName AS STRING @, strQName AS STRING @) AS VOID;
			HELPSTRING "Receive notification of the end of an element."
		IF this.FilterTrue
         	this.oContentHandler.endElement(@strNamespaceURI;
         		,@strLocalName, @strQName)
    	ENDIF
		IF strLocalName == this.FilterCriteria
			INSERT INTO curss FROM NAME this 
			this.FilterTrue = .F.
			this.nrecs = this.nrecs + 1
			IF this.nrecs/100 = INT(this.nrecs/100)
				WAIT 'Обработано '+STR(this.nrecs,6) + ' записей...' WINDOW NOWAIT 
			ENDIF 
		ENDIF
	ENDPROC

	PROCEDURE IVBSAXContentHandler_characters(strChars AS STRING @) AS VOID;
			HELPSTRING "Receive notification of character data."
		IF this.FilterTrue
			LOCAL lcVal
			lcVal = ALLTRIM(strChars)
			this.oContentHandler.characters(@strChars)
			IF !EMPTY(lcVal)
			  DO CASE
			   CASE THIS.CurrentField = 'PMP_MEDICAMENT_MAN_ID'
  				this.PMP_ID = lcVal
			   CASE THIS.CurrentField = 'CODE'
				this.R_UP = lcVal
			   CASE THIS.CurrentField = 'NAME'
				this.NNAME = lcVal
			   CASE THIS.CurrentField = 'PRIM_TYPE'
				this.TYPE = lcVal
			   CASE THIS.CurrentField = 'PRIM_QTY_VALUE'
				this.QTY_VALUE = VAL(ALLTRIM(STRTRAN(lcVal,',','.')))
			   CASE THIS.CurrentField = 'PRIM_QTY_UNIT'
				this.QTY_UNIT = lcVal 
			   CASE THIS.CurrentField = 'PRIM_MASS_VALUE'
				this.MASS_VALUE = VAL(ALLTRIM(STRTRAN(lcVal,',','.')))
			   CASE THIS.CurrentField = 'PRIM_MASS_UNIT'
  				this.MASS_UNIT = lcVal
			   CASE THIS.CurrentField = 'PRIM_VOL_VALUE'
  				this.VOL_VALUE = VAL(ALLTRIM(STRTRAN(lcVal,',','.')))
			   CASE THIS.CurrentField = 'PRIM_VOL_UNIT'
  				this.VOL_UNIT = lcVal
  			   OTHERWISE 
			  ENDCASE 
			ENDIF 
		ENDIF
	ENDPROC

	PROCEDURE IVBSAXContentHandler_ignorableWhitespace(strChars AS STRING @) AS VOID;
			HELPSTRING "Receive notification of ignorable whitespace in element content."
		* add user code here
	ENDPROC

	PROCEDURE IVBSAXContentHandler_processingInstruction(strTarget AS STRING @, strData AS STRING @) AS VOID;
			HELPSTRING "Receive notification of a processing instruction."
		* add user code here
	ENDPROC

	PROCEDURE IVBSAXContentHandler_skippedEntity(strName AS STRING @) AS VOID;
			HELPSTRING "Receive notification of a skipped entity."
		* add user code here
	ENDPROC

	*////////////////////////////////////
	*// IMPLEMENTS IVBSAXErrorHandler
	*//
	PROCEDURE IVBSAXErrorHandler_error(oLocator AS VARIANT, strErrorMessage AS STRING @, nErrorCode AS Number) AS VOID;
			HELPSTRING "Receive notification of a recoverable error."
		* add user code here
	ENDPROC

	PROCEDURE IVBSAXErrorHandler_fatalError(oLocator AS VARIANT, strErrorMessage AS STRING @, nErrorCode AS Number) AS VOID;
			HELPSTRING "Receive notification of a non-recoverable error."
		MESSAGEBOX('A non-recoverable error occured!',0+64,'sax')
		*IF TYPE('thisForm.edit1') = 'O' AND !ISNULL(thisForm.edit1)
		*	thisForm.edit1.Value = strErrorMessage + " [" + TRANSFORM(nErrorCode, '@0')+"]"
		*ENDIF	 	    
		*this.errorHappen = .T.
	ENDPROC

	PROCEDURE IVBSAXErrorHandler_ignorableWarning(oLocator AS VARIANT, strErrorMessage AS STRING @, nErrorCode AS Number) AS VOID;
			HELPSTRING "Receive notification of an ignorable warning."
		* add user code here
	ENDPROC

ENDDEFINE
