PROCEDURE uploadMail
 PARAMETERS para0, para1, para2, para3, para4, para5 
 
 *m.Mmy = SUBSTR(m.gcPeriod,5,2) + SUBSTR(m.gcPeriod,4,1)
 *m.attName = 'E' + STR(m.lpu_id,4) + m.qcod + '.' + m.mmy

 m.attName   = para0 && m.attName = 'OT'+STR(m.lpu_id,4)+m.qcod+'.'+m.mmy
 m.lpu_id    = para1
 m.mcod      = para2
 m.GWlogId   = para3
 m.lIsSilent = para4
 m.loForm    = para5

 *m.attName = 'OT'+STR(m.lpu_id,4)+m.qcod+'.'+m.mmy
 *m.checkSum = '123456789'

 *Описание SOAP метода
  *<wsdl:operation name="uploadMail">
 *	<soap:operation soapAction="" style="document"/>
 *	<wsdl:input name="uploadMail">
 *		<soap:body use="literal"/>
 *	</wsdl:input>
 *	<wsdl:output name="uploadMailResponse">
 *		<soap:body use="literal"/>
 *	</wsdl:output>
 *</wsdl:operation>
 
 * Элементы
 *<xs:element name="uploadMail" type="tns:uploadMail"/>
 *<xs:element name="uploadMailResponse" type="tns:uploadMailResponse"/>

 * Комплексный тип uploadMail
 *<xs:complexType name="uploadMail">
 *	<xs:sequence>
 *		<xs:element name="authInfo" type="tns:wsAuthInfo"/>
 *		<xs:element name="request" type="tns:uploadMailRequest"/>
 *	</xs:sequence>
 *</xs:complexType>

 * Комплексный тип uploadMailRequest
 *<xs:complexType name="uploadMailRequest">
 *	<xs:sequence>
 *		<xs:element minOccurs="1" name="attachment" nillable="true" type="tns:mailAttachment"/>
 *	</xs:sequence>
 *</xs:complexType>

 * Комплексный тип mailAttachment
 *<xs:complexType name="mailAttachment">
 *	<xs:sequence>
 *		<xs:element minOccurs="1" name="parentMailGWlogId" type="xs:long"/>
 *		<xs:element minOccurs="1" name="attachmentName" type="xs:string"/>
 *		<xs:element minOccurs="1" name="checkSum" type="xs:long"/>
 *		<xs:element minOccurs="1" name="attachmentData" type="xs:base64Binary" xmime:expectedContentTypes="application/zip"/>
 *	</xs:sequence>
 *</xs:complexType>

* тестовые адреса
* m.address = 'http://192.168.192.111:8080/module-pmp/ws/smoIOWs'
* m.host = '192.168.192.111:8080'

* промышленные адреса
 m.address = 'http://192.168.192.119:8080/module-pmp/ws/smoIOWs'
 m.host = '192.168.192.119:8080'

 IF EMPTY(m.address)
  RETURN
 ENDIF 
 
 * Проверяем наличие директорий
 m.curSoapDir = pSoap+'\'+DTOS(DATE())
 =CheckSOAPDirs(m.curSoapDir)
 * Проверяем наличие директорий

 * Генерируем уникальные имена 
 m.un_id    = SYS(3)
 m.httpFile = m.un_id + '.txt'
 m.xmlFile  = m.un_id + '.xml'
 * Генерируем уникальные имена 

 LOCAL oXML  AS MsXml2.DOMDocument
 LOCAL oNode as MSXML2.IXMLDOMNode
 LOCAL oRoot as MSXML2.IXMLDOMElement
 LOCAL oElem as MSXML2.IXMLDOMElement
 LOCAL oBody as MSXML2.IXMLDOMNode
 LOCAL oRequest as MSXML2.IXMLDOMNode
 LOCAL oClient as MSXML2.IXMLDOMNode
 
 LOCAL oHttp AS MsXml2.XMLHTTP
 
 oHttp = CREATEOBJECT("MsXml2.XMLHTTP")

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")

 * Create a procesing instruction.  
 oXML.appendChild(oXML.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'"))
 oXML.resolveExternals = .T.

 oRoot = CreateEnvelopePump(oXML)

 oXML.appendChild(oRoot)
 
 oNode = oXML.createElement("soapenv:Header")
 oRoot.appendChild(oNode)

 oBody = oXML.createElement("soapenv:Body")
 oRoot.appendChild(oBody)

 oRequest = oXML.createElement("ws:uploadMail")
 oBody.appendChild(oRequest)
 
 oClient = CreateClientPump(oXml, '')
 oRequest.appendChild(oClient)
 
 oXXX = oXML.createElement("request")
 oRequest.appendChild(oXXX)

* m.zip = FILETOSTR('&pBase\&gcPeriod\&mcod\&attachmentName')
 m.zip = FILETOSTR('&pBase\&gcPeriod\&mcod\&attName')

 *m.Base64zip = Base64(m.zip)

 *Dim objBase64
 *Const otSafeArray = 0
 *Const otString = 2
 *Set objBase64 = Server.CreateObject("XStandard.Base64")
 *Response.Write objBase64.Encode("Hello World!", otString)
 *Set objBase64 = Nothing
 
 *oBase64 = CREATEOBJECT("XStandard.Base64")
 *m.Base64zip = oBase64.Encode(m.zip,2)
 *m.Base64zip = EncodeStr64(m.zip)

 m.Base64zip = IIF(!INLIST(m.qcod,'R2','S7'), Base64(m.zip), ToBase64(m.zip))
 m.checkSum = crc32(m.zip)

 oYYY = oXML.createElement("attachment")
 oYYY.appendChild(oXML.createElement("parentMailGWlogId")).text = m.GWlogId
 oYYY.appendChild(oXML.createElement("attachmentName")).text    = m.attName
 oYYY.appendChild(oXML.createElement("checkSum")).text          = m.checkSum
 
 oYYY.appendChild(oXML.createElement("attachmentData")).text = m.Base64zip

 oXXX.appendChild(oYYY)

 oXML.save('&curSoapDir\OUTPUT\&xmlFile')
 length = fso.GetFile('&curSoapDir\OUTPUT\&xmlFile').Size
 
 LOCAL oEx as Exception

 m.err = .f. 
 TRY 
*  ohttp.open('post', m.address, .f.) && .f. - синхронное соединение, .t. - асинхронное соединение!
  ohttp.open('post', m.address, .t.) && .f. - синхронное соединение, .t. - асинхронное соединение!
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
 IF m.err = .t. 
  RELEASE oXML, oHttp
  IF  !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  ENDIF 
  RETURN 
 ENDIF 

 ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 ohttp.setRequestHeader("Content-Type", "application/xml")
* ohttp.setRequestHeader("Content-Type", ;
 	'multipart/related; type="application/xop+xml"; boundary="uuid:c9697379-9b14-4294-988b-233bf40f0f30"; start="<root.message@cxf.apache.org>"; start-info="text/xml"; charset=UTF-8')
 ohttp.setRequestHeader("Content-Length", m.length)
 ohttp.setRequestHeader("Host", m.host)
 ohttp.setRequestHeader("User-Agent", "Mike Ruby Software")
 
 poi = fso.CreateTextFile('&curSoapDir\OUTPUT\&httpFile')
 poi.WriteLine('Accept-Encoding: gzip,deflate')
 poi.WriteLine('Content-Type: "application/xml"')
* poi.WriteLine('Content-Type: "multipart/related; type="application/xop+xml"; boundary="uuid:c9697379-9b14-4294-988b-233bf40f0f30"; start="<root.message@cxf.apache.org>"; start-info="text/xml"; charset=UTF-8"')
 poi.WriteLine('Content-Length: '+ALLTRIM(STR(m.length)))
 poi.WriteLine('Host: ' + m.host)
 poi.WriteLine('User-Agent: "Mike Ruby Software"')
 poi.Close
 
 TRY 
  ohttp.send(oXml.xml) && Для get-запросов тела нет, был бы null, для post - есть, поэтому передаем парметр
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  IF  !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  ENDIF 
  RETURN 
 ENDIF 

 m.IsCancelled = .f.
 DO WHILE ohttp.readyState<4
  WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 

  IF CHRSAW(0) 
   IF INKEY() == 27
    WAIT CLEAR 
    IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
     m.IsCancelled = .t.
     EXIT 
    ENDIF 
    WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 
   ENDIF 
  ENDIF 

 ENDDO 
 WAIT CLEAR 
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

 m.httpStatus = ohttp.status
 IF  !INLIST(m.httpStatus, 200, 500)
  MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText), 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'uploadMail'))
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 
 
 * Сохраняем http-заголовок
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * Сохраняем http-заголовок
 
 *Попытка без сохранения файла выделить из него xml
 m.XmlFromFile = ExtractEnvelope(ohttp)
 IF EMPTY(m.XmlFromFile)
  IF !m.lIsSilent
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'uploadMail'))
  ENDIF 
  RETURN .T.
 ENDIF 

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 IF !oxml.loadXML(m.XmlFromFile)
  RELEASE oXML
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'uploadMail'))
  ENDIF 
  RETURN .T.
 ENDIF 
 oXml.save('&curSoapDir\INPUT\&xmlFile')
 *Попытка без сохранения файла выделить из него xml

 IF m.httpStatus=500
  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs=0
   RELEASE oXml
   IF !m.lIsSilent
    MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'')
   ENDIF 
   RETURN .T.
  ELSE 
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text
   
   IF !m.lIsSilent
    MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'uploadMail'))
   ENDIF 
  
   RELEASE oXml
   RETURN .T.
  ENDIF 
 ENDIF 

*<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
*<soap:Body>
*<ns2:uploadMailResponse xmlns:ns2="http://ws.smo.pmp.ibs.ru/">
*<mailGWlogid>3409301</mailGWlogid>
*</ns2:uploadMailResponse>
*</soap:Body>
*</soap:Envelope>
 
 *m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/ns2:uploadMailResponse/list').length
 m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/ns2:uploadMailResponse').length
 IF m.n_recs=0
  RELEASE oXml
  IF  !m.lIsSilent
   MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'')
  ENDIF 
  RETURN 
 ELSE 
  IF  !m.lIsSilent
   *MESSAGEBOX('ОБНАРУЖЕНО '+STR(m.n_recs)+' ЗАПИСЕЙ!',0+64,'')
  ENDIF 
 ENDIF 

 *m.orec = oxml.selectNodes('soap:Envelope/soap:Body/ns2:uploadMailResponse/list').item(0)
 m.orec = oxml.selectNodes('soap:Envelope/soap:Body/ns2:uploadMailResponse').item(0)
  
 m.mailGWlogid = ALLTRIM(orec.selectNodes('mailGWlogid').item(0).text)
 
 IF UPPER(m.attName) = 'E'
  oXML.save(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\flk_'+m.GWlogId+'.xml')
 ELSE 
  oXML.save(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\mek_'+m.GWlogId+'.xml')
 ENDIF 

 m.bansfile = IIF(UPPER(m.attName) = 'E', 'b_flk_'  + m.mcod, 'b_mek_'  + m.mcod)

 poi = fso.CreateTextFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.bansfile)
 poi.WriteLine('Date&Time  : ' + TTOC(DATETIME()))
 poi.WriteLine('Версия счета: '+m.GWlogId)
 poi.Close

 IF !ISNULL(m.loForm)
  loForm.resplogid = m.mailGWlogid
 ENDIF 

 IF !m.lIsSilent
  MESSAGEBOX('mailGWlogid = '+m.mailGWlogid, 0+64, '')
 ENDIF 
RETURN 
  
FUNCTION ToBase64(cSrc)
	LOCAL nFlags, nBufsize, cDst
	nFlags=1  && base64

	nBufsize=0
	= CryptBinaryToString(@cSrc, LEN(cSrc),;
		m.nFlags, NULL, @nBufsize)

	cDst = REPLICATE(CHR(0), m.nBufsize)
	IF CryptBinaryToString(@cSrc, LEN(cSrc), m.nFlags,;
		@cDst, @nBufsize) = 0
		RETURN ""
	ENDIF
RETURN cDst
ENDFUNC 

FUNCTION FromBase64(cSrc)
	LOCAL nFlags, nBufsize, cDst
	nFlags=1  && base64

	nBufsize=0
	= CryptStringToBinary(@cSrc, LEN(m.cSrc),;
		nFlags, NULL, @nBufsize, 0,0)

	cDst = REPLICATE(CHR(0), m.nBufsize)
	IF CryptStringToBinary(@cSrc, LEN(m.cSrc),;
		nFlags, @cDst, @nBufsize, 0,0) = 0
		RETURN ""
	ENDIF
RETURN m.cDst
ENDFUNC 