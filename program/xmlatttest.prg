FUNCTION XmlAttTest
 IF MESSAGEBOX('ТЕСТИРОВАТЬ getXMLAttachment?'+CHR(13)+CHR(10)+'',4+32,'')=7
  RETURN 
 ENDIF 
 
 CREATE CURSOR c_curs (cod c(6))
 
 DO CASE 
  CASE m.qcod='S7'
   INSERT INTO c_curs (cod) VALUES ('443356')
   INSERT INTO c_curs (cod) VALUES ('443306')
   INSERT INTO c_curs (cod) VALUES ('443256')
   INSERT INTO c_curs (cod) VALUES ('444308')
   INSERT INTO c_curs (cod) VALUES ('444258')
   INSERT INTO c_curs (cod) VALUES ('444006')
   INSERT INTO c_curs (cod) VALUES ('442956')
   INSERT INTO c_curs (cod) VALUES ('443951')
   INSERT INTO c_curs (cod) VALUES ('444106')
   INSERT INTO c_curs (cod) VALUES ('443156')
   INSERT INTO c_curs (cod) VALUES ('443106')
   INSERT INTO c_curs (cod) VALUES ('443206')
   INSERT INTO c_curs (cod) VALUES ('444156')
   
  CASE m.qcod='I3'
   INSERT INTO c_curs (cod) VALUES ('444107')
   INSERT INTO c_curs (cod) VALUES ('443357')
   INSERT INTO c_curs (cod) VALUES ('443207')
   INSERT INTO c_curs (cod) VALUES ('444007')
   INSERT INTO c_curs (cod) VALUES ('443007')
   INSERT INTO c_curs (cod) VALUES ('443307')
   INSERT INTO c_curs (cod) VALUES ('442957')
   INSERT INTO c_curs (cod) VALUES ('443107')
   INSERT INTO c_curs (cod) VALUES ('443952')
   INSERT INTO c_curs (cod) VALUES ('443157')
   INSERT INTO c_curs (cod) VALUES ('444306')
   INSERT INTO c_curs (cod) VALUES ('443057')
   INSERT INTO c_curs (cod) VALUES ('444250')
   INSERT INTO c_curs (cod) VALUES ('443257')
   INSERT INTO c_curs (cod) VALUES ('444157')
   
  OTHERWISE 
 ENDCASE 
 
 SELECT c_curs
 SCAN 
  m.mailGWlogid = cod
  WAIT m.mailGWlogid+'...' WINDOW NOWAIT 
  =getXMLAttachment(m.mailGWlogid, .F., NULL)
  WAIT CLEAR 
  *EXIT 
 ENDSCAN 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 

FUNCTION getXMLAttachment(para1, para2, para3) && mailGWlogid!
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
 
* тестовые адреса
 m.address = 'http://192.168.192.111:8080/module-pmp/ws/smoIOWs'
 m.host = '192.168.192.111:8080'

* промышленные адреса
* m.address = 'http://192.168.192.119:8080/module-pmp/ws/smoIOWs'
* m.host = '192.168.192.119:8080'

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

 oRequest = oXML.createElement("ws:getXMLAttachment")
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
*  ohttp.open('post', m.address, .f.) && .f. - синхронное соединение, .t. - асинхронное соединение!
  ohttp.open('post', m.address, .t.) && .f. - синхронное соединение, .t. - асинхронное соединение!
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
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
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN .F. 
 ENDIF 

 m.httpStatus = ohttp.status
 IF  !INLIST(m.httpStatus, 200, 500)
  IF !m.lIsSilent
   MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
  ENDIF 
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 
 
 * Сохраняем http-заголовок
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * Сохраняем http-заголовок

 poi = fso.OpenTextFile('&curSoapDir\INPUT\&httpFile')
 m.boundary = Element('boundary', ReadTheHead(poi, 'Content-Type'))
 poi.Close
* MESSAGEBOX(m.boundary,0+64,'')
 
 m.realBound = ALLTRIM(MLINE(ohttp.responseText, ATLINE(m.boundary, ohttp.responseText)))
 m.finalBound = m.realBound + '--'
* MESSAGEBOX(realBound,0+64,'')

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
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!', 0+64, 'getXMLAttachment')
  ENDIF 
  RETURN .T.
 ENDIF 

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 IF !oxml.loadXML(m.XmlFromFile)
  RELEASE oXML
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!', 0+64, 'getXMLAttachment')
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
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, 'getXMLAttachment')
   ENDIF 
  
   RELEASE oXml
   RETURN .T.
  ENDIF 
 ENDIF 
 
 m.n_len = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getXMLAttachmentResponse/xmlAttachment/xmlAttachmentName').length
 IF m.n_len <> 1
  MESSAGEBOX('ЗНАЧЕНИЕ m.n_len='+STR(m.n_len,3), 0+64, 'getXMLAttachment')
  RELEASE oXml
  RETURN .T.
 ENDIF 
  
 m.fname    = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getXMLAttachmentResponse/xmlAttachment/xmlAttachmentName').item(0).text
 m.checkSum = INT(VAL(oxml.selectNodes('soap:Envelope/soap:Body/ns2:getXMLAttachmentResponse/xmlAttachment/checkSum').item(0).text))

 IF EMPTY(m.fname)
  MESSAGEBOX('ЗНАЧЕНИЕ xmlAttachmentName='+m.fname, 0+64, 'getXMLAttachment')
  RELEASE oXml
  RETURN .T.
 ENDIF 

 *IF !INLIST(LEN(m.fname), 12, 9)
 * IF !m.lIsSilent
 *  MESSAGEBOX('В ПОЛУЧЕННОМ НЕКОРРЕКТНЫЙ name: '+m.fname, 0+64, 'name')
 * ENDIF 
 * RETURN .F.
 *ENDIF 

 *m.lIsBadLpu = .F.
 *IF USED('sprlpu')
 * IF LEN(m.fname) = 12 && mcod
 *  m.mcod = SUBSTR(m.fname,2,7)
 *  IF !SEEK(m.mcod, 'sprlpu', 'mcod')
 *  ELSE 
 *   m.lIsBadLpu = .T.
 *  ENDIF 
 * ENDIF 
 * IF LEN(m.fname) = 9 && lpuid
 *  m.lpu_id = SUBSTR(m.fname,2,4)
 *  IF !SEEK(INT(VAL(m.lpu_id)), 'sprlpu', 'lpu_id')
 *  ELSE 
 *   m.lIsBadLpu = .T.
 *   m.mcod = sprlpu.mcod
 *  ENDIF 
 * ENDIF 
 *ENDIF 
 
* IF !m.lIsBadLpu
*  m.zip = ohttp.responseBody
   
*  IF fso.FileExists(m.curSoapDir+'\INPUT\'+m.fname)
*   fso.DeleteFile(m.curSoapDir+'\INPUT\'+m.fname)
*  ENDIF 
*  m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip)), m.curSoapDir+'\INPUT\'+m.fname)

*  IF !m.lIsSilent
*   *MESSAGEBOX('НЕКОРРЕКТОЕ НАИМЕНОВАНИЕ ФАЙЛА: '+m.fname+CHR(13)+CHR(10)+'ФАЙЛ СОХРАНЕН В '+m.curSoapDir+'\INPUT\', 0+64, '')
*  ENDIF 
*  *RETURN .F.
* ENDIF 

 m.zip = ohttp.responseBody
 IF AT('PK', m.zip)=0
  RELEASE m.zip
  IF !m.lIsSilent
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ ZIP НЕ ОБНАРУЖЕН!',0+64,'%PDF')
  ENDIF 
  RETURN .F.
 ENDIF 

 IF fso.FileExists(m.curSoapDir+'\INPUT\'+m.fname)
  fso.DeleteFile(m.curSoapDir+'\INPUT\'+m.fname)
 ENDIF 

 m.crc    = INT(VAL(crc32(SUBSTR(m.zip, AT('PK',m.zip), AT(m.finalBound,m.zip)-AT('PK',m.zip)-2))))
 m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip), AT(m.finalBound,m.zip)-AT('PK',m.zip)-2), ;
 	 m.curSoapDir+'\INPUT\'+m.fname)
 	 
 IF !m.lIsSilent
  IF m.crc <> m.checkSum
   MESSAGEBOX('КОНТРОЛЬНАЯ СУММА       ' + STR(m.crc,10)+CHR(13)+CHR(10)+;
    	       'НЕ СОВПАДАЕТ С ИСХОДНОЙ ' + STR(m.checkSum,10), 0+64, 'getAttachment')
  ENDIF 
 ENDIF 

 RETURN .T.

