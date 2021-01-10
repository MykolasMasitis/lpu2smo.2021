FUNCTION getMailGw(para1, para2, para3, para4)
 IF PARAMETERS() > 0
  m.lpu_id    = para1
 ELSE 
  m.lpu_id    = 0
 ENDIF 
 IF PARAMETERS() > 1
  m.loForm    = para2
 ELSE 
  m.loForm    = NULL
 ENDIF 
 IF PARAMETERS() > 2
  m.billId    = para3
 ELSE 
  m.billId    = ''
 ENDIF 
 IF PARAMETERS() > 3
  m.lIsSilent = para4
 ELSE 
  m.lIsSilent = .F.
 ENDIF 
 
 IF m.lpu_id>0 AND EMPTY(m.billId)
  MESSAGEBOX('НЕ ЗАДАН BILLID!', 0+64, STR(m.lpu_id,4))
  RETURN .F.
 ENDIF 

 * Описание SOAP-метода
 *<wsdl:operation name="getMailGw">
 *	<soap:operation soapAction="" style="document"/>
 *	<wsdl:input name="getMailGw">
 *		<soap:body use="literal"/>
 *	</wsdl:input>
 *	<wsdl:output name="getMailGwResponse">
 *		<soap:body use="literal"/>
 *	</wsdl:output>
 *</wsdl:operation>

 * Элементы
 *<xs:element name="getMailGw" type="tns:getMailGw"/>
 *<xs:element name="getMailGwResponse" type="tns:getMailGwResponse"/>

 *Комплексный тип getMailGw
 *<xs:complexType name="getMailGw">
 *	<xs:sequence>
 *		<xs:element name="authInfo" type="tns:wsAuthInfo"/>
 *		<xs:element name="getMailGwRequest" type="tns:getMailGwRequest"/>
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
 * Генерируем уникальные имена 

 *LOCAL oXML  AS MsXml2.DOMDocument
 *LOCAL oNode as MSXML2.IXMLDOMNode
 *LOCAL oRoot as MSXML2.IXMLDOMElement
 *LOCAL oElem as MSXML2.IXMLDOMElement
 *LOCAL oBody as MSXML2.IXMLDOMNode
 *LOCAL oRequest as MSXML2.IXMLDOMNode
 *LOCAL oClient as MSXML2.IXMLDOMNode
 
 *LOCAL oHttp AS MsXml2.XMLHTTP
 *LOCAL oHttp AS MsXml2.XMLHTTP.6.0
 LOCAL oHttp AS MsXml2.XMLHTTP.3.0
 
 *oHttp = CREATEOBJECT("MsXml2.XMLHTTP")
 *oHttp = CREATEOBJECT("MsXml2.XMLHTTP.6.0")
 oHttp = CREATEOBJECT("MsXml2.XMLHTTP.3.0")

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

 oRequest = oXML.createElement("ws:getMailGw")
 oBody.appendChild(oRequest)
 
 *oClient = CreateClientPump(oXml, '')
 oClient = CreateClientPump(oXml, m.un_id)
 oRequest.appendChild(oClient)
 
 oXXX = oXML.createElement("getMailGwRequest")
 m.lcperiod = LEFT(m.gcperiod,4)+'-'+SUBSTR(m.gcperiod,5,2)
 oXXX.appendChild(oXML.createElement("period")).text = m.lcperiod
 oRequest.appendChild(oXXX)
 m.mailDirection = 'OUT'
 oXXX.appendChild(oXML.createElement("mailDirection")).text = m.mailDirection  && был *
 oRequest.appendChild(oXXX)  && был *
 m.parcelType = 'REPORT_IN_SMO'
 oXXX.appendChild(oXML.createElement("parcelType")).text = m.parcelType  && был *
 oRequest.appendChild(oXXX)  && был *
 IF m.lpu_id>0
  m.moId = STR(m.lpu_id,4)
  oXXX.appendChild(oXML.createElement("moId")).text = m.moId
  oRequest.appendChild(oXXX)
 ENDIF 
 m.smoId = m.qobjid
 oXXX.appendChild(oXML.createElement("smoId")).text = m.smoId
 oRequest.appendChild(oXXX)
 IF m.lpu_id>0
  oXXX.appendChild(oXML.createElement("billId")).text = m.billId
  oRequest.appendChild(oXXX)
 ENDIF 
 *oXXX.appendChild(oXML.createElement("status")).text = 'SENT'
 *oRequest.appendChild(oXXX)
 
 IF oXML.parseError.errorCode != 0 
   MESSAGEBOX(oXML.parseError.reason, 0+64, STR(m.lpu_id,4))
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
 
 * readyState
 * 0 UNINITIALIZED	The object has been created, but not initialized (the open method has not been called).
 * 1 LOADING	    The object has been created, but the send method has not been called.
 * 2 LOADED	    The send method has been called, but the status and headers are not yet available.
 * 3 INTERACTIVE	Some data has been received. Calling the responseBody and responseText properties at this state to obtain partial results will return an error, because status and response headers are not fully available.
 * 4 COMPLETED	    All the data has been received, and the complete data is available in the responseBody and responseText properties.
 * This property returns a 4-byte integer.

 m.IsCancelled = .f.
 DO WHILE ohttp.readyState<1
  *WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 

  IF CHRSAW(0) 
   IF INKEY() == 27
    *WAIT CLEAR 
    IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОЖИДАНИЕ ОТВЕТА?',4+32,'') == 6
     KEYBOARD '{ESC}'
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

 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message, 0+64, STR(m.lpu_id,4))
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
  MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message, 0+64, STR(m.lpu_id,4))
  RETURN .F.
 ENDIF 

 m.IsCancelled = .f.
 WAIT "CheckMailGw (ОЖИДАНИЕ ОТВЕТА)..." WINDOW NOWAIT 
 DO WHILE ohttp.readyState<4

  IF CHRSAW(0) 
   IF INKEY() == 27
    WAIT CLEAR 
    IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
     m.IsCancelled = .t.
     EXIT 
    ENDIF 
    WAIT "CheckMailGw (ОЖИДАНИЕ ОТВЕТА)..." WINDOW NOWAIT 
   ENDIF 
  ENDIF 

 ENDDO 
 WAIT "CheckMailGw..." WINDOW NOWAIT 
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN IIF(m.lpu_id>0, .T., .F.)
 ENDIF 

 m.httpStatus = ohttp.status
 IF  !INLIST(m.httpStatus, 200, 500)
  MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText), 0+64, STR(m.lpu_id,4))
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
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  ENDIF 
  RETURN IIF(m.lpu_id>0, .T., .F.)
 ENDIF 

 *oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 oXML  = CREATEOBJECT("MsXml2.DOMDocument.6.0")
 oXML.async = .F.
 oXML.setProperty("SelectionNamespaces", ;
 	"xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:ns2='http://ws.smo.pmp.ibs.ru/'")

 IF !oxml.loadXML(m.XmlFromFile)
  RELEASE oXML
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  ENDIF 
  RETURN IIF(m.lpu_id>0, .T., .F.)
 ENDIF 
 oXml.save('&curSoapDir\INPUT\&xmlFile')
 oXml.save('&pBase\&gcPeriod\getMailGw.xml')
 *Попытка без сохранения файла выделить из него xml

 *IF m.httpStatus=500
  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs>0
   *RELEASE oXml
   *IF !m.lIsSilent
   * MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'')
   *ENDIF 
   *RETURN IIF(m.lpu_id>0, .T., .F.)
  *ELSE 
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text
   
   IF !m.lIsSilent
    *MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, STR(m.lpu_id,4))
   ENDIF 
  
   RELEASE oXml
   RETURN IIF(m.lpu_id>0, .T., .F.)
  ENDIF 
 *ENDIF 

 m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getMailGwResponse/return/list').length
 m.n_size = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getMailGwResponse/return/size').length

 CREATE CURSOR answer (recid c(6), period c(6), mGwLogId c(7), mDirection c(3), parcelType c(100), moId n(4), ;
 	smoId n(4), messageId c(75), sendDate t, billId c(9), ver c(6), msgStatus c(10), ok l)
 INDEX on moId tag moId
 SET ORDER TO moId
 *INDEX on sendDate TAG sendDate DESCENDING 
 *SET ORDER TO sendDate

 CREATE CURSOR FullAnswer (recid c(6), period c(6), mGwLogId c(7), mDirection c(3), parcelType c(100), moId n(4), ;
 	smoId n(4), messageId c(75), sendDate t, billId c(9), ver c(6), msgStatus c(10), ok l)
 INDEX on moId tag moId
 SET ORDER TO moId

 FOR m.n_rec = 0 TO m.n_recs-1
  m.orec = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getMailGwResponse/return/list').item(m.n_rec)
  
  m.period = STRTRAN(orec.selectNodes('period').item(0).text,'-','')
  m.mDirection = orec.selectNodes('mailDirection').item(0).text
  m.parcelType = orec.selectNodes('parcelType').item(0).text
   
  IF m.mDirection!='OUT'
   *LOOP 
  ENDIF 
  IF m.parcelType != 'REPORT_IN_SMO'
   *LOOP 
  ENDIF 

  m.mGwLogId  = orec.selectNodes('mailGwLogId').item(0).text
  m.moId      = INT(VAL(orec.selectNodes('moId').item(0).text))
  m.smoId     = INT(VAL(orec.selectNodes('smoId').item(0).text))
  m.messageId = orec.selectNodes('messageId').item(0).text
  m.sendDate  = LEFT(orec.selectNodes('sendDate').item(0).text,19)
  m.sendDate  = CTOT(SUBSTR(m.sendDate,9,2)+'.'+SUBSTR(m.sendDate,6,2)+'.'+SUBSTR(m.sendDate,1,4)+' '+SUBSTR(m.sendDate,12,8))
  m.billid    = orec.selectNodes('billId').item(0).text
  m.ver       = orec.selectNodes('versionNumber').item(0).text
  m.msgStatus = orec.selectNodes('messageStatus').item(0).text
   
  INSERT INTO FullAnswer FROM MEMVAR 

  IF m.mDirection!='OUT' OR m.parcelType != 'REPORT_IN_SMO'
  ELSE 
  IF !SEEK(m.moId, 'answer')
   INSERT INTO answer FROM MEMVAR 
  ELSE 
   IF m.sendDate > answer.sendDate
    UPDATE answer SET messageId=m.messageId, sendDate=m.sendDate, ver=m.ver, msgStatus=m.msgStatus, mGwLogId=m.mGwLogId ;
    	WHERE moId=m.moId
   ENDIF 
  ENDIF 
  ENDIF 

 ENDFOR 
 
 SELECT answer
 IF fso.FileExists(pbase+'\'+gcperiod+'\getMailGw.dbf')
  fso.DeleteFile(pbase+'\'+gcperiod+'\getMailGw.dbf')
 ENDIF 
 COPY TO &pbase\&gcperiod\getMailGw
 SELECT aisoms

 IF !ISNULL(m.loForm)

 ELSE 

  IF !fso.FileExists('&pbase\&gcperiod\aisoms.dbf')
   USE IN answer
   MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ '+pbase+'\'+gcperiod+'\aisoms.dbf!', 0+64, 'getBillStatuses')
   RETURN .F.
  ENDIF 
  IF OpenFile('&pbase\&gcperiod\aisoms', 'aisoms', 'shar', 'lpuid')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   USE IN amswer
   RETURN .F.
  ENDIF 

 ENDIF !ISNULL(m.loForm)

 WAIT "CheckMailGw (ОБРАБОТКА)..." WINDOW NOWAIT 
 SELECT aisoms
 orec = RECNO('aisoms')
 SET RELATION TO lpuid INTO answer ADDITIVE 
 SCAN 
  m.lpuid = lpuid
  m.mcod  = mcod
  IF m.lpu_id>0 AND m.lpu_id<>m.lpuid
   LOOP 
  ENDIF 
  IF EMPTY(answer.mGwLogId)
   LOOP 
  ENDIF 
  

  *SELECT * FROM FullAnswer WHERE moId=m.lpuid ;
  	INTO TABLE &pBase\&gcPeriod\&mcod\bill_vers 
  SELECT * FROM FullAnswer WHERE moId=m.lpuid INTO CURSOR b_vers
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\bill_vers.dbf')
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\bill_vers.dbf')
  ENDIF   
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\bill_vers.cdx')
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\bill_vers.cdx')
  ENDIF   
  SELECT b_vers
  INDEX on mGwLogId TAG mGwLogId
  COPY TO &pBase\&gcPeriod\&mcod\bill_vers WITH cdx 
  USE 

  *SELECT period, billId, sendDate, mGwLogId, ver, messageId FROM FullAnswer WHERE moId=m.lpuid ;
  	INTO TABLE &pBase\&gcPeriod\&mcod\bill_vers
  *SELECT bill_vers
  *INDEX on mGwLogId TAG mGwLogId
  *USE IN bill_vers
  SELECT aisoms
  
  IF !EMPTY(aisoms.gwlogid) AND aisoms.gwlogid != answer.mgwlogid AND INLIST(soapsts, 'ACCEPTED')
   oAlertMgr = CREATEOBJECT("VFPAlert.AlertManager")
   poAlert = oAlertMgr.NewAlert()
   poAlert.Alert(aisoms.gwlogid+' !=: '+answer.mgwlogid, 8, "Расхождение в версиях счетов!",;
  	IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, ''))
*   MESSAGEBOX('ВЕРСИЯ ПРИНЯТОГО СЧЕТА       : '+aisoms.gwlogid+CHR(13)+CHR(10)+;
 	          'НЕ СОТВЕТСТВУЕТ ВЫСТАВЛЕННОЙ : '+answer.mgwlogid, 0+64, STR(m.lpuid,4))
  ENDIF 
  * Сюда добавить проверку gwlogid - если поменялся, то взвести триггер IsNewAcc
  IF gwlogid<>answer.mGwLogId
   IF LOCK()
    REPLACE IsNewAcc WITH .T.
    UNLOCK RECORD RECNO()
   ENDIF 
  ENDIF 
  IF LOCK()
   REPLACE gwlogid WITH answer.mGwLogId, cmessage WITH answer.messageid, ;
 	sent WITH answer.sendDate, ver WITH answer.ver
   UNLOCK RECORD RECNO()
  ENDIF 

  IF !ISNULL(m.loForm)
   loForm.refresh
  ENDIF 
   
 ENDSCAN 
 SET RELATION OFF INTO answer
 USE IN answer
 USE IN FullAnswer
 IF BETWEEN(m.orec, 1, RECCOUNT('aisoms'))
  GO (orec)
 ENDIF 
 
 WAIT CLEAR 
 
 IF !ISNULL(m.loForm)
  loForm.refresh
 ENDIF 

 IF !m.lIsSilent
  MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64,'')
 ENDIF 

RETURN .T.