FUNCTION getDictionary
 PARAMETERS para1
 m.lIsSilent  = .f.
 m.dictionaryName = para1
 
* промышленные адреса
 m.address = 'http://192.168.192.119:8080/module-nsi/ws/nsiWs'
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

 oRoot = CreateEnvelopeNSI(oXML)

 oXML.appendChild(oRoot)
 
 oNode = oXML.createElement("soapenv:Header")
 oRoot.appendChild(oNode)

 oBody = oXML.createElement("soapenv:Body")
 oRoot.appendChild(oBody)

 oClient = CreateClientPump(oXml, '')

 oRequest = oXML.createElement("ws:getDictionary")
 oRequest.appendChild(oClient)
 oFilter = oXML.createElement("ws:filter")
 oFilter.appendChild(oXML.createElement("dictionaryName")).text = m.dictionaryName
 oRequest.appendChild(oFilter)
 oBody.appendChild(oRequest)
 
 IF oXML.parseError.errorCode != 0 
  MESSAGEBOX(oXML.parseError.reason,0+64,'')
  RELEASE oXML, oHttp
  RETURN 
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
  RETURN 
 ENDIF 

 ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 ohttp.setRequestHeader("Content-Type", "application/xml")
 ohttp.setRequestHeader("Content-Length", m.length)
 ohttp.setRequestHeader("Host", m.host)
 ohttp.setRequestHeader("User-Agent", "Mike Ruby Software")
 
 poi = fso.CreateTextFile('&curSoapDir\OUTPUT\&httpFile')
 poi.WriteLine('Accept-Encoding: gzip,deflate')
 poi.WriteLine('Content-Type: "application/xml"')
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
  MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
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
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

* IF  ohttp.status<>200
 m.httpStatus = ohttp.status
 IF  !INLIST(m.httpStatus, 200, 500)
  MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
 * Сохраняем http-заголовок
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * Сохраняем http-заголовок
 
 poi   = FCREATE('&curSoapDir\INPUT\&xmlFile')
 nSize = FWRITE(poi, ohttp.responseBody)
 =FCLOSE(poi)

 IF m.httpStatus=500
  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs=0
   RELEASE oXml
   IF !m.lIsSilent
    MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'soap:Fault')
   ENDIF 
   RETURN 
  ELSE 
   IF !m.lIsSilent
    MESSAGEBOX('ОБНАРУЖЕНО '+STR(m.n_recs)+' ЗАПИСЕЙ!',0+64,'')
   ENDIF 
  ENDIF 

  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs=0
   RELEASE oXml
   IF !m.lIsSilent
    MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'')
   ENDIF 
   RETURN 
  ELSE 
   IF !m.lIsSilent
    MESSAGEBOX('ОБНАРУЖЕНО '+STR(m.n_recs)+' ЗАПИСЕЙ!',0+64,'')
   ENDIF 
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text
   
   IF !m.lIsSilent
    MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10),0+64,'changeBillStatus')
   ENDIF 
  
   RELEASE oXml
   RETURN 
  ENDIF 
 ENDIF 

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 IF !oxml.load('&curSoapDir\INPUT\&xmlFile')
  RELEASE oXML
  MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!',0+64,'')
  RETURN 
 ENDIF 

 m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getDictionariesResponse/return/response/nsiDictionaryList/list').length
 IF m.n_recs=0
  RELEASE oXml
  MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'')
  RETURN 
 ELSE 
  MESSAGEBOX('ОБНАРУЖЕНО '+STR(m.n_recs)+' ЗАПИСЕЙ!',0+64,'')
 ENDIF 

 CREATE CURSOR answer (name_eta c(8), intr_data d, full_name c(100), cur_ver c(10))
 FOR m.n_rec = 0 TO m.n_recs-1
  m.orec = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getDictionariesResponse/return/response/nsiDictionaryList/list').item(m.n_rec)
   
  m.name_eta  = orec.selectNodes('code').item(0).text
  m.intr_data = LEFT(orec.selectNodes('dateVersion').item(0).text,10)
  m.intr_data = CTOD(SUBSTR(m.intr_data,9,2)+'.'+SUBSTR(m.intr_data,6,2)+'.'+SUBSTR(m.intr_data,1,4))
  m.full_name = orec.selectNodes('name').item(0).text
  m.cur_ver   = orec.selectNodes('version').item(0).text
  
  INSERT INTO answer FROM MEMVAR 

 ENDFOR 
  
 COPY TO &pbase\&gcperiod\sprspr
* BROWSE 
 USE IN answer 

 MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64,'')

RETURN 