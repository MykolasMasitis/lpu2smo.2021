PROCEDURE changeBillStatus(para1, para2, para3)
 IF PARAMETERS() > 0
  m.billid    = para1
 ELSE 
  m.billid    = ''
 ENDIF 
 IF PARAMETERS() > 1
  m.status    = para2 && RECIEVED or RECIEVED_E
 ELSE 
  m.status    = ''
 ENDIF 
 IF PARAMETERS() > 2
  m.lIsSilent = para3
 ELSE 
  m.lIsSilent = .F.
 ENDIF 
 
 IF EMPTY(m.billid)
  RETURN .F.
 ENDIF 
 IF !INLIST(m.status, 'RECIEVED', 'RECIEVED_E')
  RETURN .F.
 ENDIF 
 
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

 oRequest = oXML.createElement("ws:changeBillStatus")
 oBody.appendChild(oRequest)
 
 oClient = CreateClientPump(oXml, '')
 oRequest.appendChild(oClient)
 
 oXXX = oXML.createElement("changeBillStatusRequest")
 oXXX.appendChild(oXML.createElement("billId")).text = m.billid
 oRequest.appendChild(oXXX)
 oXXX.appendChild(oXML.createElement("billStatus")).text = m.status
 oRequest.appendChild(oXXX)

 IF oXML.parseError.errorCode != 0 
  MESSAGEBOX(oXML.parseError.reason,0+64,'changeBillStatus(1)')
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
  RELEASE oXML, oHttp
  MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'changeBillStatus(2)')
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
  MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'changeBillStatus(3)')
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
  RETURN .T.
 ENDIF 

 m.httpStatus = ohttp.status
 IF  !INLIST(m.httpStatus, 200, 500)
  IF !m.lIsSilent
   MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'changeBillStatus(4)')
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
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'changeBillStatus(5)'))
  ENDIF 
  RETURN .T.
 ENDIF 

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 IF !oxml.loadXML(m.XmlFromFile)
  RELEASE oXML
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'changeBillStatus(6)'))
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
    MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'changeBillStatus(7)')
   ENDIF 
   RETURN .T.
  ELSE 
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text
   
   IF !m.lIsSilent
    MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, STR(m.lpu_id,4))
   ENDIF 
  
   RELEASE oXml
   RETURN .T.
  ENDIF 
 ENDIF 

 m.orec = oxml.selectNodes('soap:Envelope/soap:Body/ns2:changeBillStatusResponse/return').item(0)
  
 m.billid     = orec.selectNodes('billId').item(0).text
 m.billStatus = orec.selectNodes('billStatus').item(0).text
 m.success    = LOWER(ALLTRIM(orec.selectNodes('success').item(0).text))

 IF !m.lIsSilent
  MESSAGEBOX('billid= '+m.billid+CHR(13)+CHR(10)+;
			'billStatus= '+m.billStatus+CHR(13)+CHR(10)+;
			'success= '+m.success+CHR(13)+CHR(10),0+64,'changeBillStatus')
 ENDIF 

RETURN IIF(m.success = 'true', .T., .F.)