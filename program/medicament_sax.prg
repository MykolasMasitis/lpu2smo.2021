PROCEDURE medicament_sax
 IF MESSAGEBOX(CHR(13)+CHR(10)+' ŒÕ¬≈–“»–Œ¬¿“‹ medicament.xml?'+CHR(13)+CHR(10),;
 	4+32,'SAX (XML->DBF)')=7
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
  MESSAGEBOX(CHR(13)+CHR(10)+'¬€ Õ»◊≈√Œ Õ≈ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 CREATE CURSOR curss (DD_SID C(10), DD_NAME C(100), IS_OMS L, ;
 	MASS_VALUE n(6,2), MASS_UNIT C(10), VOL_VALUE n(6,2), VOL_UNIT C(10), GD_SID C(8), GD_NAME C(100), IS_TARGET n(1), MAX_DOSE N(11,4))
 SELECT curss
 INDEX on dd_sid TAG dd_sid

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
	loWorker.oContentHandlerImpl.FilterCriteria = ALLTRIM('ROW') && ¬‡ÊÌ˚È Ô‡‡ÏÂÚ!
	*
	*-- parseURL
	loRdr.parseURL(csprfile)
	*
	*-- Set result
	*loWorker.edit1.Value = loWorker.oWriter.output  && ÒÂÈ˜‡Ò ÏÌÂ ˝ÚÓ ÌÂ Ì‡‰Ó!
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

 *SET SAFETY OFF
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medicament.dbf')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medicament.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medicament.cdx')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medicament.cdx')
 ENDIF 

 COPY TO &pBase/&gcPeriod/nsi/medicament WITH cdx 
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

 MESSAGEBOX('‘¿…À —‘Œ–Ã»–Œ¬¿Õ!',0+64,'medicament')
 

RETURN 

DEFINE CLASS ContentHandlerImpl AS session
	oContentHandler = NULL	&& the content handler object
	oErrorHandler = NULL	&& the error handler object
	errorHappen = .F.		&& Flag to indicate if the error handler has thrown a fatal error.
	FilterTrue = .F.		&& Flag to indicate if the element is in scope.
	FilterCriteria = ""		&& String to hold the element name
	
	CurrentField = ""

	is_oms     = .f.
  	gd_sid     = ''
	dd_sid     = ''
	dd_name    = ''
	gd_name    = ''
	is_target  = .f.
	mass_value = 0
	mass_unit  = ''
	vol_value  = 0
  	vol_unit   = ''
  	max_dose   = 0
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
			  this.is_oms     = .f.
  			  this.gd_sid     = ''
			  this.dd_sid     = ''
			  this.dd_name    = ''
			  this.gd_name    = ''
			  this.is_target  = 0
			  this.mass_value = 0
			  this.mass_unit  = ''
			  this.vol_value  = 0
  			  this.vol_unit   = ''
  			  this.max_dose   = 0

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
				WAIT 'Œ·‡·ÓÚ‡ÌÓ '+STR(this.nrecs,6) + ' Á‡ÔËÒÂÈ...' WINDOW NOWAIT 
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
			   CASE THIS.CurrentField = 'IS_OMS_TARIF'
			   	this.is_oms = IIF(lcVal='1', .T., .F.)
			   CASE THIS.CurrentField = 'GD_SID'
  				this.gd_sid = ALLTRIM(lcVal)
			   CASE THIS.CurrentField = 'CODE'
				this.dd_sid = ALLTRIM(lcVal)
			   CASE THIS.CurrentField = 'NAME'
				this.dd_name = ALLTRIM(lcVal)
			   CASE THIS.CurrentField = 'GD_NAME'
				this.gd_name = ALLTRIM(lcVal)
			   CASE THIS.CurrentField = 'IS_TARGET'
				this.is_target = VAL(ALLTRIM(STRTRAN(lcVal,',','.')))
			   CASE THIS.CurrentField = 'STRENGTH_MASS_VALUE'
				this.mass_value = VAL(ALLTRIM(STRTRAN(lcVal,',','.')))
			   CASE THIS.CurrentField = 'STRENGTH_MASS_UNIT'
				this.mass_unit = ALLTRIM(lcVal)
			   CASE THIS.CurrentField = 'STRENGTH_VOL_VALUE'
				this.vol_value = VAL(ALLTRIM(STRTRAN(lcVal,',','.')))
			   CASE THIS.CurrentField = 'STRENGTH_VOL_UNIT'
  				this.vol_unit = ALLTRIM(lcVal)
			   CASE THIS.CurrentField = 'MAX_SINGLE_DOSE'
  				this.max_dose = VAL(ALLTRIM(STRTRAN(lcVal,',','.')))
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
