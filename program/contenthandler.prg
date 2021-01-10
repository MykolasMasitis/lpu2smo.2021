DEFINE CLASS ContentHandlerImpl AS session
	oContentHandler = NULL	&& the content handler object
	oErrorHandler = NULL	&& the error handler object
	errorHappen = .F.		&& Flag to indicate if the error handler has thrown a fatal error.
	*PROTECTED simpStack
	*simpStack = NULL
	FilterTrue = .F.		&& Flag to indicate if the element is in scope.
	FilterCriteria = ""		&& String to hold the element name
	nrecs         = 0
	
	isperson = .f.
	ispolicy = .f.
	isattach = .f.
	isdudl   = .f.
	is_mo    = .f.
	
	CurrentField = ""
	
	n_lpu  = 0
	n_rec  = 0
	recid  = ""
    n_pol  = ""
    d_rq   = {}
    d_u    = {}
    fam    = ""
    im     = ""
    ot     = ""
    dr     = ""
    w      = 0
    tip_d  = ""
    ans_r  = '0*0'
    tip_d  = ""
    q      = ""
    d_beg  = {}
    d_h    = {}
    d_end  = {}
    moCode = 0
	
	lpu_tip  = 0
    
    lpu_id = 0
    st_id  = 0
    pd_id  = 0
    
    dtr_off = {}
    dst_off = {}
    dpd_off = {}
    ttr_off = {}
    tst_off = {}
    tpd_off = {}
    
    st_code = ""
    st_name = ""
    
    *dAttachE = {}
    *tAttachE = {}
    *dstAttachE = {}
    *dsttAttachE = {}

	IMPLEMENTS IVBSAXContentHandler EXCLUDE IN "msxml6.dll"
	IMPLEMENTS IVBSAXErrorHandler EXCLUDE IN "msxml6.dll"

	PROCEDURE Destroy
		*this.simpStack = NULL
		STORE NULL TO ;
			this.oContentHandler, this.oErrorHandler
	ENDPROC

	PROTECTED PROCEDURE Initialize
		this.errorHappen = .F.
		this.FilterTrue = .F.
		*this.simpStack = CREATEOBJECT("simpStack")
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
		IF 'person' = strLocalName
			*MESSAGEBOX('Start person!',0+64,'')
			this.isperson = .t.
		ENDIF 
		IF 'policy' = strLocalName
			this.ispolicy = .t.
		ENDIF 
		IF 'attach' = strLocalName
			this.isattach = .t.
		    this.ttr_off = {}
		    this.tst_off = {}
		    this.n_lpu = this.n_lpu + 1
		ENDIF 
		IF 'dudl' = strLocalName
			this.isdudl   = .t.
		ENDIF 
		IF 'mo' = strLocalName
			this.is_mo   = .t.
		ENDIF 

		*this.simpStack.Push(strLocalName)

		IF strLocalName == this.FilterCriteria

			this.n_rec = this.n_rec + 1
			this.recid = PADL(this.n_rec,6,'0')
        	this.FilterTrue = .T.
			this.n_pol  = ""
			this.d_rq   = {}
			this.d_u    = {}
		    this.fam    = ""
		    this.im     = ""
		    this.ot     = ""
		    this.dr     = ""
		    this.w      = 0
		    this.tip_d  = ""
		    this.ans_r  = '0*0'
		    this.tip_d  = ""
		    this.q      = ""
		    this.d_beg  = {}
		    this.d_h    = {}
		    this.d_end  = {}

			this.lpu_tip = 0
		    this.lpu_id  = 0
		    this.st_id   = 0
		    this.pd_id   = 0
		    this.moCode  = 0

		    this.dtr_off = {}
		    this.dpd_off = {}
		    this.dst_off = {}
		    this.ttr_off = {}
		    this.tpd_off = {}
		    this.tst_off = {}
		    
		    this.st_code = ""
		    this.st_name = ""

		    *this.dAttachE   = {}
		    *this.dstAttachE = {}
		    this.n_lpu      = 0
		ENDIF

		IF this.FilterTrue
			this.oContentHandler.startElement(@strNamespaceURI, @strLocalName, @strQName, oAttributes)

			DO CASE 
			 CASE strLocalName = 'policySerNum'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'rqDate'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'surname'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'namep'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'patronymic'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'sexId'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'dateBirth'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'policyTCode'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'insuranceQQ'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'plDateE'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'plDateB'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'plDateH'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'policyStatusCode'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'policyStatus'
			  THIS.CurrentField = strLocalName
			 
			 CASE strLocalName = 'areaTId'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'moCode'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'dateAttachB'
			  THIS.CurrentField = strLocalName
			 CASE strLocalName = 'dateAttachE'
			  THIS.CurrentField = strLocalName
			 OTHERWISE 
			  THIS.CurrentField = NULL
			ENDCASE 
			*MESSAGEBOX('oAttributes.getValue='+oAttributes.getValue(0),0+64,'startElement')
		ENDIF
*				'
	ENDPROC

	PROCEDURE IVBSAXContentHandler_endElement(strNamespaceURI AS STRING @, strLocalName AS STRING @, strQName AS STRING @) AS VOID;
			HELPSTRING "Receive notification of the end of an element."
		IF 'person' = strLocalName
			this.isperson = .f.
		ENDIF 
		IF 'policy' = strLocalName
			this.ispolicy = .f.
		ENDIF 
		IF 'attach' = strLocalName
			this.isattach = .f.
			DO CASE 
			 CASE INLIST(this.lpu_tip,1) && Взрослый
				IF EMPTY(this.ttr_off )
			    	this.ttr_off  = {31.12.2099}
				ENDIF 
		        IF this.ttr_off  >= this.dtr_off 
		         this.dtr_off  = this.ttr_off 
			 	 this.lpu_id = this.moCode
		        ENDIF 

			 CASE INLIST(this.lpu_tip,2) && Детский
				IF EMPTY(this.tpd_off )
			    	this.tpd_off  = {31.12.2099}
				ENDIF 
		        IF this.tpd_off  >= this.dpd_off 
		         this.dpd_off  = this.tpd_off 
			 	 this.pd_id = this.moCode
		        ENDIF 

			 CASE this.lpu_tip = 5 && Стоматология
				IF EMPTY(this.tst_off )
			    	this.tst_off  = {31.12.2099}
				ENDIF 
		        IF this.tst_off  >= this.dst_off 
		         this.dst_off  = this.tst_off 
			 	 this.st_id = this.moCode
		        ENDIF 
			 OTHERWISE 
			 	&& ничего не делаем!
			ENDCASE 
		ENDIF 
		IF 'dudl' = strLocalName
			this.isdudl   = .f.
		ENDIF 
		IF 'mo' = strLocalName
			this.is_mo   = .f.
		ENDIF 
		*this.simpStack.Pop()
		IF this.FilterTrue
         	this.oContentHandler.endElement(@strNamespaceURI;
         		,@strLocalName, @strQName)
    	ENDIF
		IF strLocalName == this.FilterCriteria
		    IF this.d_end >= this.d_rq
		     m.priz = IIF(this.n_lpu>2, .t., .f.)
			 IF SEEK(this.n_pol, 'answer')
			 DO CASE 
			  CASE this.q=m.qcod AND this.q != answer.q
			   DELETE IN answer 
			   INSERT INTO answer FROM NAME this 
			  CASE answer.q=m.qcod AND this.q != answer.q
			   * не добавляем!
			  OTHERWISE 
			   INSERT INTO answer FROM NAME this 
			 ENDCASE 
			  *IF this.d_beg > answer.d_beg
			  * DELETE IN answer 
			  * INSERT INTO answer FROM NAME this 
			  *ELSE 
			  * && не добавляем ничего!
			  *ENDIF 
			 ELSE 
			  INSERT INTO answer FROM NAME this 
			 ENDIF 
		    ENDIF 
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
			   CASE THIS.CurrentField = 'policySerNum' AND !this.ispolicy
  				this.n_pol = lcVal
			   CASE THIS.CurrentField = 'rqDate'
				m.d_rq = STRTRAN(lcVal,'-','')
				m.d_u  = m.d_rq
				m.d_rq = CTOD(SUBSTR(m.d_rq,7,2)+'.'+SUBSTR(m.d_rq,5,2)+'.'+SUBSTR(m.d_rq,1,4))
  				this.d_rq = m.d_rq
  				this.d_u  = m.d_u
			   CASE THIS.CurrentField = 'surname'
			    this.fam = lcVal
			   CASE THIS.CurrentField = 'namep'
			    this.im = lcVal
			   CASE THIS.CurrentField = 'patronymic'
			    this.ot = lcVal
			   CASE THIS.CurrentField = 'sexId' AND this.isperson=.t.
			    this.w = INT(VAL(lcVal))
			   CASE THIS.CurrentField = 'dateBirth'
			    this.dr = STRTRAN(lcVal,'-','')
			   CASE THIS.CurrentField = 'policyTCode'
			    this.tip_d = lcVal
			   CASE THIS.CurrentField = 'insuranceQQ'
			    this.q = lcVal
			   CASE THIS.CurrentField = 'plDateE'
			    m.d_end = lcVal
				m.d_end = STRTRAN(m.d_end,'-','')
				m.d_end = CTOD(SUBSTR(m.d_end,7,2)+'.'+SUBSTR(m.d_end,5,2)+'.'+SUBSTR(m.d_end,1,4))
			    this.d_end = m.d_end
			   CASE THIS.CurrentField = 'plDateB'
			    m.d_beg = lcVal
				m.d_beg = STRTRAN(m.d_beg,'-','')
				m.d_beg = CTOD(SUBSTR(m.d_beg,7,2)+'.'+SUBSTR(m.d_beg,5,2)+'.'+SUBSTR(m.d_beg,1,4))
			    this.d_beg = m.d_beg
			   CASE THIS.CurrentField = 'plDateH'
			    m.d_h = lcVal
				m.d_h = STRTRAN(m.d_h,'-','')
				m.d_h = CTOD(SUBSTR(m.d_h,7,2)+'.'+SUBSTR(m.d_h,5,2)+'.'+SUBSTR(m.d_h,1,4))
			    this.d_h = m.d_h

			   CASE THIS.CurrentField = 'policyStatusCode'
			    this.st_code = lcVal
			    IF this.st_code<>'2'
			     this.ans_r = '211'
			    ENDIF 
			   CASE THIS.CurrentField = 'policyStatus'
			    this.st_name = lcVal

			   CASE THIS.CurrentField = 'areaTId'
       			this.lpu_tip = INT(VAL(lcVal))
			   CASE THIS.CurrentField = 'moCode' AND this.is_mo
       			this.moCode = INT(VAL(lcVal))
			   CASE THIS.CurrentField = 'dateAttachE'
			    DO CASE 
			     CASE this.lpu_tip = 1
			     CASE this.lpu_tip = 2
			     CASE this.lpu_tip = 5
			    ENDCASE 
			    
			    IF INLIST(this.lpu_tip,1,2)
		         m.ttr_off  = STRTRAN(lcVal,'-','')
		         this.ttr_off  = CTOD(SUBSTR(m.ttr_off ,7,2)+'.'+SUBSTR(m.ttr_off ,5,2)+'.'+SUBSTR(m.ttr_off ,1,4))
		        ELSE 
		         m.tst_off  = STRTRAN(lcVal,'-','')
		         this.tst_off  = CTOD(SUBSTR(m.tst_off ,7,2)+'.'+SUBSTR(m.tst_off ,5,2)+'.'+SUBSTR(m.tst_off ,1,4))
		        ENDIF 
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

DEFINE CLASS simpStack AS session
	PROTECTED stackValue
	PROTECTED stackSep

	* Init props
	PROCEDURE Init()
		this.stackValue = ""
		this.stackSep = ","
	ENDFUNC

	* Tests if the stack is empty.
	FUNCTION Empty() AS Boolean
		RETURN (LENC(this.stackValue) = 0)
	ENDFUNC

	* Removes the element from the top of the stack.
	FUNCTION Pop() AS VOID
		IF !this.Empty()
			LOCAL lnPos as Integer
			lnPos = ATC(this.stackSep, this.stackValue)
			IF lnPos > 0
				this.stackValue = SUBSTRC(this.stackValue, lnPos + 1)
			ELSE
				this.stackValue = ""
			ENDIF
		ENDIF
	ENDFUNC

	* Adds an element to the top of the stack.
	FUNCTION Push(tcValue as String) AS VOID
		IF !this.Empty()
			this.stackValue = tcValue + this.stackSep + this.stackValue
		ELSE
			this.stackValue = tcValue
		ENDIF
	ENDFUNC

	* Returns the number of elements in the stack.
	FUNCTION Size() AS Integer
		IF !this.Empty()
			RETURN GETWORDCOUNT(this.stackValue, this.stackSep)
		ENDIF
		RETURN 0
	ENDFUNC

	* Returns a reference to an element at the top of the stack.
	FUNCTION Top() AS String
		IF !this.Empty()
			RETURN GETWORDNUM(this.stackValue, 1, this.stackSep)
		ENDIF
		RETURN ""
	ENDFUNC

ENDDEFINE
