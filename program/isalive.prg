FUNCTION IsAlive (para1, para2) 
 IF PARAMETERS() > 0
  m.service = para1 && 0 - ПУМП, 1 - ЕРЗЛ (не работает!)
 ELSE 
  m.service = 0 && по умолчанию - ПУМП
 ENDIF 
 IF PARAMETERS() > 1
  m.lIsSilent = para2
 ELSE 
  m.lIsSilent = .F.
 ENDIF 
 
 * Описание SOAP-метода
 *wsdl:operation name="isAlive">
 *	<soap:operation soapAction="" style="document"/>
 *	<wsdl:input name="isAlive">
 *		<soap:body use="literal"/>
 *	</wsdl:input>
 *	<wsdl:output name="isAliveResponse">
 *		<soap:body use="literal"/>
 *	</wsdl:output>
 *</wsdl:operation>

 * Элементы
 *<xs:element name="isAlive" type="tns:isAlive"/>
 *<xs:element name="isAliveResponse" type="xsd:boolean"/>
 
 * Комплексный тип isalive
 *<xs:complexType name="isAlive">
 *	<xs:sequence>
 *		<xs:element name="authInfo" type="tns:wsAuthInfo"/>
 *	</xs:sequence>
 *</xs:complexType>


 IF m.service = 0
  *адрес ПУМП
  m.address = 'http://192.168.192.119:8080/module-pmp/ws/smoIOWs'
  m.host = '192.168.192.119:8080'
 ELSE 
  * адрес ЕРЗЛ
  m.address = 'http://192.168.192.118:8080/erzl-for-smo/ws/'
  m.host = '192.168.192.118:8080'
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
 
 oHttp = CREATEOBJECT("MsXml2.XMLHTTP.3.0")

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

 oRequest = oXML.createElement("ws:isAlive")
 oBody.appendChild(oRequest)
 
 oClient = CreateClientPump(oXml, '')
 oRequest.appendChild(oClient)
 
 IF oXML.parseError.errorCode != 0 
  IF !m.lIsSilent
   MESSAGEBOX(oXML.parseError.reason,0+64,'')
  ENDIF 
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 
 
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
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  ENDIF 
  RELEASE oXML, oHttp
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
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  ENDIF 
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 

 m.IsCancelled = .f.
 DO WHILE ohttp.readyState<4
  *WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 

  IF CHRSAW(0) 
   IF INKEY() == 27
    *WAIT CLEAR 
    IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
     m.IsCancelled = .t.
     EXIT 
    ENDIF 
    *WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 
   ENDIF 
  ENDIF 

 ENDDO 
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 

 m.httpStatus = ohttp.status
 IF  !INLIST(m.httpStatus, 200, 500)
  IF !m.lIsSilent
   MESSAGEBOX('ОШИБКА HTTP: '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
  ENDIF 
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
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!', 0+64, 'IsAlive')
  ENDIF 
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 

 oXML  = CREATEOBJECT("MsXml2.DOMDocument.6.0")
 oXML.async = .F.
 oXML.setProperty("SelectionNamespaces", ;
 	"xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:ns2='http://ws.smo.pmp.ibs.ru/'")
 
 IF !oxml.loadXML(m.XmlFromFile)
  RELEASE oXML
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!', 0+64, 'IsAlive')
  ENDIF 
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 
 oXml.save('&curSoapDir\INPUT\&xmlFile')
 
 IF m.httpStatus=500
  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs=0
   IF !m.lIsSilent
    MESSAGEBOX('HTTP Error 500, ПУСТОЙ ОТВЕТ!', 0+64, 'IsAvlive')
   ENDIF 
   RELEASE oXml, oHttp
   RETURN .F.
  ELSE 
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text
   
   IF !m.lIsSilent
    MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, 'IsAlive')
   ENDIF 
   RELEASE oXML, oHttp
   RETURN .F.
  ENDIF 
 ENDIF 
 
 * ответ
 *<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 * <soap:Body>
 *  <ns2:isAliveResponse xmlns:ns2="http://ws.smo.pmp.ibs.ru/">true</ns2:isAliveResponse>
 * </soap:Body>
 *</soap:Envelope>
 * ответ
 
 m.orec    = oxml.selectNodes('soap:Envelope/soap:Body/ns2:isAliveResponse')
 m.success = LOWER(ALLTRIM(m.orec.item(0).text))

 IF !m.lIsSilent
  MESSAGEBOX('success= '+m.success+CHR(13)+CHR(10),0+64,'IsAlive')
 ENDIF 
 RELEASE oXML, oHttp

RETURN IIF(m.success = 'true', .T., .F.)