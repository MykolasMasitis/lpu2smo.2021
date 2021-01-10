FUNCTION getAttachment2(para1, para2, para3) && mailGWlogid!
 IF PARAMETERS() > 0
  m.mailGWlogid = para1 && char(7)
 ELSE 
  m.mailGWlogid = ''
 ENDIF 
 IF PARAMETERS() > 1
  m.lIsSilent = para2
 ELSE 
  m.lIsSilent = .F.
 ENDIF 
 IF PARAMETERS() > 2
  m.loForm = para3
 ELSE 
  m.loForm = NULL
 ENDIF 
 
 IF EMPTY(m.mailGWlogid)
  RETURN .F.
 ENDIF 

 *Описание SOAP метода
 *<wsdl:operation name="getAttachment">
 *	<soap:operation soapAction="" style="document"/>
 *	<wsdl:input name="getAttachment">
 *		<soap:body use="literal"/>
 *	</wsdl:input>
 *	<wsdl:output name="getAttachmentResponse">
 *		<soap:body use="literal"/>
 *	</wsdl:output>
 *</wsdl:operation>

 *Элементы
 *<xs:element name="getAttachment" type="tns:getAttachment"/>
 *<xs:element name="getAttachmentResponse" type="tns:getAttachmentResponse"/>

 *Комплексный тип getAttachment
 *<xs:complexType name="getAttachment">
 *	<xs:sequence>
 *		<xs:element name="authInfo" type="tns:wsAuthInfo"/>
 *		<xs:element name="request" type="tns:getAttachmentRequest"/>
 *	</xs:sequence>
 *</xs:complexType>
 
* тестовые адреса
* m.address = 'http://192.168.192.111:8080/module-pmp/ws/smoIOWs'
* m.host = '192.168.192.111:8080'

* промышленные адреса
 m.address = 'http://192.168.192.119:8080/module-pmp/ws/smoIOWs'
 m.host = '192.168.192.119:8080'

 IF EMPTY(m.address)
  RETURN .F.
 ENDIF 

 * Проверяем наличие директорий
 m.curSoapDir = pSoap+'\'+DTOS(DATE())
 =CheckSOAPDirs(m.curSoapDir)
 * Проверяем наличие директорий

 * Генерируем уникальные имена 
 m.un_id    = SYS(3)
 m.httpFile = m.un_id + '.txt'
 m.xmlFile  = m.un_id + '.xml'
 m.zipFile  = m.un_id + '.zip'
 * Генерируем уникальные имена 

 *LOCAL oXML  AS MsXml2.DOMDocument
 *LOCAL oNode as MSXML2.IXMLDOMNode
 *LOCAL oRoot as MSXML2.IXMLDOMElement
 *LOCAL oElem as MSXML2.IXMLDOMElement
 *LOCAL oBody as MSXML2.IXMLDOMNode
 *LOCAL oRequest as MSXML2.IXMLDOMNode
 *LOCAL oClient as MSXML2.IXMLDOMNode
 
 *LOCAL oHttp AS MsXml2.XMLHTTP
 *LOCAL oHttp AS MsXml2.XMLHTTP.3.0 && эта версия рабочая!
 *LOCAL oHttp AS MsXml2.XMLHTTP.4.0
 *LOCAL oHttp AS MsXml2.XMLHTTP.6.0
 *LOCAL oHttp AS MSXML2.ServerXMLHTTP 
 
 *oHttp = CREATEOBJECT("MsXml2.XMLHTTP")
 *oHttp = CREATEOBJECT("MsXml2.XMLHTTP.3.0") && эта версия рабочая!
 *oHttp = CREATEOBJECT("MsXml2.XMLHTTP.4.0")
 *oHttp = CREATEOBJECT("MsXml2.XMLHTTP.6.0")
 *oHttp = CREATEOBJECT("MSXML2.ServerXMLHTTP")
 
 * попробовать!
 LOCAL oHttp as WinHttp.WinHttpRequest.5.1
 oHttp = CREATEOBJECT("WinHttp.WinHttpRequest.5.1")
 *https://docs.microsoft.com/en-us/windows/win32/winhttp/winhttprequest
 * попробовать!

 *oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 oXML  = CREATEOBJECT("MsXml2.DOMDocument.6.0")

 * Create a procesing instruction.  
 oXML.appendChild(oXML.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'"))
 oXML.resolveExternals = .T.

 oRoot = CreateEnvelopePump(oXML)

 oXML.appendChild(oRoot)
 
 oNode = oXML.createElement("soapenv:Header")
 oRoot.appendChild(oNode)

 oBody = oXML.createElement("soapenv:Body")
 oRoot.appendChild(oBody)

 oRequest = oXML.createElement("ws:getAttachment")
 oBody.appendChild(oRequest)
 
 oClient = CreateClientPump(oXml, '')
 oRequest.appendChild(oClient)
 
 oXXX = oXML.createElement("request")
 oXXX.appendChild(oXML.createElement("mailGWlogid")).text = m.mailGWlogid
 oRequest.appendChild(oXXX)
 
 IF oXML.parseError.errorCode != 0 
  MESSAGEBOX(oXML.parseError.reason,0+64,'')
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 
 
 oXML.save('&curSoapDir\OUTPUT\&xmlFile')
 length = fso.GetFile('&curSoapDir\OUTPUT\&xmlFile').Size
 
 LOCAL oEx as Exception

 m.err = .f. 
 TRY 
  *ohttp.open('post', m.address, .f.) && .f. - синхронное соединение, .t. - асинхронное соединение!
  ohttp.open('post', m.address, .t.) && .f. - синхронное соединение, .t. - асинхронное соединение!
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
 * readyState
 * 0 UNINITIALIZED	The object has been created, but not initialized (the open method has not been called).
 * 1 LOADING	    The object has been created, but the send method has not been called.
 * 2 LOADED	    The send method has been called, but the status and headers are not yet available.
 * 3 INTERACTIVE	Some data has been received. Calling the responseBody and responseText properties at this state to obtain partial results will return an error, because status and response headers are not fully available.
 * 4 COMPLETED	    All the data has been received, and the complete data is available in the responseBody and responseText properties.
 * This property returns a 4-byte integer.

 *m.IsCancelled = .f.
 *DO WHILE ohttp.readyState<1
 * *WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 

 * IF CHRSAW(0) 
 *  IF INKEY() == 27
 *   *WAIT CLEAR 
 *   IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОЖИДАНИЕ ОТВЕТА?',4+32,'') == 6
 *    KEYBOARD '{ESC}'
 *    m.IsCancelled = .t.
 *    EXIT 
 *   ENDIF 
 *   *WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 
 *  ENDIF 
 * ENDIF 

 *ENDDO 
 
 *IF  m.IsCancelled = .t.
 * RELEASE oXML, oHttp
 * RETURN .F.
 *ENDIF 
 
 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN .F.
 ENDIF 

 ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 ohttp.setRequestHeader("Content-Type", "application/xml")
 ohttp.setRequestHeader("Content-Length", m.length)
 ohttp.setRequestHeader("Host", m.host)
 ohttp.setRequestHeader("User-Agent", "Mike Ruby Software (9950825@mail.ru; +79637820825)")
 
 poi = fso.CreateTextFile('&curSoapDir\OUTPUT\&httpFile')
 poi.WriteLine('Accept-Encoding: gzip,deflate')
 poi.WriteLine('Content-Type: "application/xml"')
 poi.WriteLine('Content-Length: '+ALLTRIM(STR(m.length)))
 poi.WriteLine('Host: ' + m.host)
 poi.WriteLine('User-Agent: "Mike Ruby Software (9950825@mail.ru; +79637820825)"')
 poi.Close
 
 TRY 
  ohttp.send(oXml.xml) && Для get-запросов тела нет, был бы null, для post - есть, поэтому передаем парметр
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN .F.
 ENDIF 

 *m.IsCancelled = .f.
 *DO WHILE ohttp.readyState<4
 * WAIT "checkattachments ("+mcod+", ОЖИДАНИЕ ОТВЕТА)..." WINDOW NOWAIT 

 * IF CHRSAW(0) 
 *  IF INKEY() == 27
 *   WAIT CLEAR 
 *   IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
 *    m.IsCancelled = .t.
 *    EXIT 
 *   ENDIF 
 *   WAIT "checkattachments ("+mcod+", ОЖИДАНИЕ ОТВЕТА)..." WINDOW NOWAIT 
 *  ENDIF 
 * ENDIF 

 *ENDDO 
 
 *IF  m.IsCancelled = .t.
 * RELEASE oXML, oHttp
 * RETURN .F. 
 *ENDIF 

 DO WHILE !oHttp.WaitForResponse()
 ENDDO 

 m.s_tatus = 0
 TRY 
  m.httpStatus = ohttp.status
  *m.s_tatusText = ohttp.statusText
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('ОШИБКА ohttp.status!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN .F.
 ENDIF 

 *IF  ohttp.status<>200
 IF  !INLIST(m.httpStatus, 200, 500) && m.s_tatus<>200
  *MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
  *MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(m.s_tatus)+CHR(13)+CHR(10)+ALLTRIM(m.s_tatus),0+64,'')
  *MESSAGEBOX('ОШИБКА HTTP:  '+STR(m.httpStatus)+CHR(13), 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  MESSAGEBOX('ОШИБКА HTTP:  '+STR(m.httpStatus)+CHR(13), 0+64, 'getAttachment')
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 


 *m.httpStatus = ohttp.status
 *IF  !INLIST(m.httpStatus, 200, 500)
 * IF !m.lIsSilent
 *  MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
 * ENDIF 
 * RELEASE oXML, oHttp
 * RETURN .F.
 *ENDIF 
 
 * Сохраняем http-заголовок
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * Сохраняем http-заголовок

 poi = fso.OpenTextFile('&curSoapDir\INPUT\&httpFile')
 m.boundary = Element('boundary', ReadTheHead(poi, 'Content-Type'))
 poi.Close

 TRY 
  m.respText = ohttp.responseText
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  IF !m.lIsSilent
   MESSAGEBOX('ОШИБКА ohttp.responseText!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  ENDIF 
  RETURN .F.
 ENDIF 

 m.realBound = ALLTRIM(MLINE(ohttp.responseText, ATLINE(m.boundary, ohttp.responseText))) && вот здесь проблема!!!
 m.finalBound = m.realBound + '--'

* poi = fso.OpenTextFile('&curSoapDir\INPUT\&httpFile')
* m.start      = Element('start', ReadTheHead(poi, 'Content-Type'))
* poi.Close
* MESSAGEBOX(m.start,0+64,'')
 
* poi = fso.OpenTextFile('&curSoapDir\INPUT\&httpFile')
* m.start_info = Element('start-info', ReadTheHead(poi, 'Content-Type'))
* poi.Close
* MESSAGEBOX(m.start_info,0+64,'')
  
 m.XmlFromFile = ExtractEnvelope(ohttp)
 IF EMPTY(m.XmlFromFile)
  IF !m.lIsSilent
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!', 0+64, 'getAttachment')
  ENDIF 
  RETURN .T.
 ENDIF 

 *oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 oXML  = CREATEOBJECT("MsXml2.DOMDocument.6.0")
 oXML.async = .F.
 oXML.setProperty("SelectionNamespaces", ;
 	"xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:ns2='http://ws.smo.pmp.ibs.ru/'")

 IF !oxml.loadXML(m.XmlFromFile)
  RELEASE oXML
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!', 0+64, 'getAttachment')
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
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, 'getAttachment')
   ENDIF 
  
   RELEASE oXml
   RETURN .T.
  ENDIF 
 ENDIF 
 
 m.n_len = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getAttachmentResponse/attachment/attachmentName').length
 IF m.n_len <> 1
  MESSAGEBOX('ЗНАЧЕНИЕ m.n_len='+STR(m.n_len,3), 0+64, 'getAttachment')
  RELEASE oXml
  RETURN .T.
 ENDIF 
  
 m.fname    = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getAttachmentResponse/attachment/attachmentName').item(0).text
 m.checkSum = INT(VAL(oxml.selectNodes('soap:Envelope/soap:Body/ns2:getAttachmentResponse/attachment/checkSum').item(0).text))

 IF EMPTY(m.fname)
  MESSAGEBOX('ЗНАЧЕНИЕ attachmentName='+m.fname, 0+64, 'getAttachment')
  RELEASE oXml
  RETURN .T.
 ENDIF 

 IF !INLIST(LEN(m.fname), 12, 9)
  IF !m.lIsSilent
   MESSAGEBOX('В ПОЛУЧЕННОМ НЕКОРРЕКТНЫЙ name: '+m.fname, 0+64, 'name')
  ENDIF 
  RETURN .F.
 ENDIF 

 m.lIsBadLpu = .F.
 IF USED('sprlpu')
  IF LEN(m.fname) = 12 && mcod
   m.mcod = SUBSTR(m.fname,2,7)
   IF !SEEK(m.mcod, 'sprlpu', 'mcod')
   ELSE 
    m.lIsBadLpu = .T.
   ENDIF 
  ENDIF 
  IF LEN(m.fname) = 9 && lpuid
   m.lpu_id = SUBSTR(m.fname,2,4)
   IF !SEEK(INT(VAL(m.lpu_id)), 'sprlpu', 'lpu_id')
   ELSE 
    m.lIsBadLpu = .T.
    m.mcod = sprlpu.mcod
   ENDIF 
  ENDIF 
 ENDIF 
 
 IF !m.lIsBadLpu
  m.zip = ohttp.responseBody
   
  IF fso.FileExists(m.curSoapDir+'\INPUT\'+m.fname)
   fso.DeleteFile(m.curSoapDir+'\INPUT\'+m.fname)
  ENDIF 
  m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip)), m.curSoapDir+'\INPUT\'+m.fname)

  IF !m.lIsSilent
   MESSAGEBOX('НЕКОРРЕКТОЕ НАИМЕНОВАНИЕ ФАЙЛА: '+m.fname+CHR(13)+CHR(10)+'ФАЙЛ СОХРАНЕН В '+m.curSoapDir+'\INPUT\', 0+64, '')
  ENDIF 
  RETURN .F.
 ENDIF 

 m.zip = ohttp.responseBody
 IF AT('PK', m.zip)=0
  RELEASE m.zip
  IF !m.lIsSilent
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ ZIP НЕ ОБНАРУЖЕН!',0+64,'%PDF')
  ENDIF 
  RETURN .F.
 ENDIF 

 IF fso.FolderExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod)

  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)

   *m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
   *FOR tParam = 1 TO 999
   * m.fname000 = STRTRAN(m.fname, m.mmy, PADL(tParam,3,'0'))
   * IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname000)
   *  EXIT 
   * ENDIF 
   *ENDFOR 
   
   *fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname, ;
   *	m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname000)
   *fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)

  ENDIF 

  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)
  ENDIF 
  m.crc    = INT(VAL(crc32(SUBSTR(m.zip, AT('PK',m.zip), AT(m.finalBound,m.zip)-AT('PK',m.zip)-2))))
  m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip), AT(m.finalBound,m.zip)-AT('PK',m.zip)-2), ;
 	 m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)
  m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\b'+mailGWlogid+'.'+m.mmy)
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\b'+mailGWlogid+'.'+m.mmy)
  ENDIF 
  m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip), AT(m.finalBound,m.zip)-AT('PK',m.zip)-2), ;
 	 m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\b'+mailGWlogid+'.'+m.mmy)

  IF !m.lIsSilent
   IF m.crc<>m.checkSum
    MESSAGEBOX('КОНТРОЛЬНАЯ СУММА       ' + STR(m.crc,10)+CHR(13)+CHR(10)+;
    	       'НЕ СОВПАДАЕТ С ИСХОДНОЙ ' + STR(m.checkSum,10), 0+64, 'getAttachment')
   ELSE
*    MESSAGEBOX('КОНТРОЛЬНАЯ СУММА      ' + STR(m.crc,10)+CHR(13)+CHR(10)+;
    	       'СООТВЕТСТВУЕТ ИСХОДНОЙ ' + STR(m.checkSum,10), 0+64, 'getAttachment')
   ENDIF 
  ENDIF 
  
  RETURN .T.

 ELSE 

  IF fso.FileExists(m.curSoapDir+'\INPUT\'+m.fname)
   fso.DeleteFile(m.curSoapDir+'\INPUT\'+m.fname)
  ENDIF 
  *m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip)), ;
 	 m.curSoapDir+'\INPUT\'+m.fname)
  m.crc    = INT(VAL(crc32(SUBSTR(m.zip, AT('PK',m.zip), AT(m.finalBound,m.zip)-AT('PK',m.zip)-2))))
  m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip), AT(m.finalBound,m.zip)-AT('PK',m.zip)-2), ;
 	 m.curSoapDir+'\INPUT\'+m.fname)
 	 
*  m.nBytes = STRTOFILE(m.zip, m.curSoapDir+'\INPUT\qwerty.zip')

  IF !m.lIsSilent
   IF c.crc<>m.checkSum
    MESSAGEBOX('КОНТРОЛЬНАЯ СУММА       ' + STR(m.crc,10)+CHR(13)+CHR(10)+;
    	       'НЕ СОВПАДАЕТ С ИСХОДНОЙ ' + STR(m.checkSum,10), 0+64, 'getAttachment')
   ELSE
*    MESSAGEBOX('КОНТРОЛЬНАЯ СУММА      ' + STR(m.crc,10)+CHR(13)+CHR(10)+;
    	       'СООТВЕТСТВУЕТ ИСХОДНОЙ ' + STR(m.checkSum,10), 0+64, 'getAttachment')
   ENDIF 
  ENDIF 

  RETURN .F.

 ENDIF 
