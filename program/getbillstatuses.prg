FUNCTION getBillStatuses(para1, para2, para3, para4) && Метод получения billid и статуса счета
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
  m.lIsSilent = para3
 ELSE 
  m.lIsSilent = .F.
 ENDIF 
 IF PARAMETERS() > 3
  m.TypeOfBill = para4
 ELSE 
  m.TypeOfBill = ''
 ENDIF 
 
 SET REPROCESS TO 1 SECONDS 
 
 * Описание SOAP-метода
 *<wsdl:operation name="getBillStatuses">
 *	<soap:operation soapAction="" style="document"/>
 *	<wsdl:input name="getBillStatuses">
 *		<soap:body use="literal"/>
 *	</wsdl:input>
 *	<wsdl:output name="getBillStatusesResponse">
 *		<soap:body use="literal"/>
 *	</wsdl:output>
 *</wsdl:operation>

 * Элементы
 *<xs:element name="getBillStatuses" type="tns:getBillStatuses"/>
 *<xs:element name="getBillStatusesResponse" type="tns:getBillStatusesResponse"/>

 *Комплексный тип getBillStatuses
 *<xs:complexType name="getBillStatuses">
 *	<xs:sequence>
 *		<xs:element name="authInfo" type="tns:wsAuthInfo"/>
 *		<xs:element name="request" type="tns:getBillStatusesRequest"/>
 *	</xs:sequence>
 *</xs:complexType>

*Комплексный тип getBillStatusesRequest
*<xs:complexType name="getBillStatusesRequest">
*<xs:sequence>
*<xs:element minOccurs="1" name="period" type="xs:string"/>
*<xs:element minOccurs="0" name="moId" type="xs:long"/>
*<xs:element minOccurs="0" name="status" type="xs:string"/>
*</xs:sequence>
*</xs:complexType>

*Комплексный тип billItem
*<xs:complexType name="billItem">
*	<xs:sequence>
*		<xs:element minOccurs="1" name="period" type="xs:string"/>
*		<xs:element minOccurs="1" name="billId" type="xs:long"/>
*		<xs:element minOccurs="1" name="moId" type="xs:long"/>
*		<xs:element minOccurs="1" name="smoId" type="xs:long"/>
*		<xs:element minOccurs="1" name="patients_count" type="xs:int"/>
*		<xs:element minOccurs="1" name="invoice_count" type="xs:int"/>
*		<xs:element minOccurs="1" name="amount" type="xs:int"/>
*		<xs:element minOccurs="1" name="status" type="tns:billStatus"/>
*		<xs:element minOccurs="1" name="statusChangeDate" type="xs:date"/>
*	</xs:sequence>
*</xs:complexType>

*Простой тип billStatus
*<xs:simpleType name="billStatus">
*	<xs:restriction base="xs:string">
*		<xs:enumeration value="GENERATION"/>
*		<xs:enumeration value="GENERATED"/>
*		<xs:enumeration value="P_CREATING"/>
*		<xs:enumeration value="P_CREATING_E"/>
*		<xs:enumeration value="P_SENDING"/>
*		<xs:enumeration value="P_SENDING_E"/>
*		<xs:enumeration value="SENT"/>
*		<xs:enumeration value="ACCEPTED"/>
*		<xs:enumeration value="RECIEVED"/>
*		<xs:enumeration value="RECIEVED_E"/>
*		<xs:enumeration value="DRAFT"/>
*		<xs:enumeration value="RECREATE_QUEUE"/>
*		<xs:enumeration value="SEND_QUEUE"/>
*		<xs:enumeration value="RECREATE_QUEUE_WFLK"/>
*		<xs:enumeration value="FLK_CHECKING"/>
*		<xs:enumeration value="GENERATED_WFLK"/>
*	</xs:restriction>
*</xs:simpleType>

* тестовые адреса
* m.address = 'http://192.168.192.111:8080/module-pmp/ws/smoIOWs'
* m.host = '192.168.192.111:8080'

* промышленные адреса
 m.address = 'http://192.168.192.119:8080/module-pmp/ws/smoIOWs'
 m.host = '192.168.192.119:8080'

 IF EMPTY(m.address)
  MESSAGEBOX('ПУСТОЙ АДРЕС m.address!', 0+64, 'getBillStatuses')
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
 
 * Закомментировано 05.12.2019
 *LOCAL oXML  AS MsXml2.DOMDocument
 *LOCAL oNode as MSXML2.IXMLDOMNode
 *LOCAL oRoot as MSXML2.IXMLDOMElement
 *LOCAL oElem as MSXML2.IXMLDOMElement
 *LOCAL oBody as MSXML2.IXMLDOMNode
 *LOCAL oRequest as MSXML2.IXMLDOMNode
 *LOCAL oClient as MSXML2.IXMLDOMNode
 * Закомментировано 05.12.2019
 
 LOCAL oHttp AS MsXml2.XMLHTTP
 
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

 oRequest = oXML.createElement("ws:getBillStatuses")
 oBody.appendChild(oRequest)
 
 oClient = CreateClientPump(oXml, m.un_id)
 oRequest.appendChild(oClient)
 
 oXXX = oXML.createElement("request")
 m.lcperiod = LEFT(m.gcperiod,4)+'-'+SUBSTR(m.gcperiod,5,2)
 oXXX.appendChild(oXML.createElement("period")).text = m.lcperiod
 IF m.lpu_id>0
  oXXX.appendChild(oXML.createElement("moId")).text = STR(m.lpu_id,4)
 ENDIF 
 oRequest.appendChild(oXXX)

 IF oXML.parseError.errorCode != 0 
   MESSAGEBOX(oXML.parseError.reason, 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses')) 
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
  MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message, 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
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
  MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  RETURN .F.
 ENDIF 

 m.IsCancelled = .f.
 WAIT "CheckStatus (ОЖИДАНИЕ ОТВЕТА)..." WINDOW NOWAIT 
 DO WHILE ohttp.readyState<4
  IF CHRSAW(0) 
   IF INKEY() == 27
    WAIT CLEAR 
    IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
     m.IsCancelled = .t.
     EXIT 
    ENDIF 
    WAIT "CheckStatus (ОЖИДАНИЕ ОТВЕТА)..." WINDOW NOWAIT 
   ENDIF 
  ENDIF 
 ENDDO 
 WAIT "CheckStatus..." WINDOW NOWAIT 
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN IIF(m.lpu_id>0, .T., .F.)
 ENDIF 

 TRY 
  m.s_tatus     = ohttp.status
  m.s_tatusText = ohttp.statusText
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('ОШИБКА ohttp.status!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN .F.
 ENDIF 

 *m.httpStatus = ohttp.status
 *IF  !INLIST(m.httpStatus, 200, 500)
 IF  !INLIST(m.s_tatus, 200, 500)
  *MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText), 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  MESSAGEBOX('ОШИБКА HTTP:  '+STR(m.s_tatus)+CHR(13)+CHR(10)+ALLTRIM(m.s_tatusText ), 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
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
 oXML.save('&pBase\&gcPeriod\BillStatuses.xml')

 *Попытка без сохранения файла выделить из него xml

* IF m.httpStatus=500
  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs>0
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text
   
   IF !m.lIsSilent
    MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
   ENDIF 
  
   RELEASE oXml
   RETURN IIF(m.lpu_id>0, .T., .F.)
  ENDIF 
  
*  IF m.n_recs=0
*   *RELEASE oXml
*   *IF !m.lIsSilent
*   * MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'')
*   *ENDIF 
*   *RETURN IIF(m.lpu_id>0, .T., .F.)
*  ELSE 
*   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

*   m.faultcode = orec.selectNodes('faultcode').item(0).text
*  m.faultstring = orec.selectNodes('faultstring').item(0).text
*   
*   *IF !m.lIsSilent
*    MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
*			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
*   *ENDIF 
  
*   RELEASE oXml
*   RETURN IIF(m.lpu_id>0, .T., .F.)
*  ENDIF 
** ENDIF 

 m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getBillStatusesResponse/list').length
* IF m.n_recs <> 1
*  IF !m.lIsSilent
*   MESSAGEBOX('ОБНАРУЖЕНО '+STR(m.n_recs)+' ЗАПИСЕЙ!',0+64,'')
*  ENDIF 
*  IF m.lpu_id > 0
*   RELEASE oXml
*   RETURN .T.
*  ENDIF 
* ENDIF 
 
 CREATE CURSOR answer (recid c(6), lpuid n(4), billid c(9), d_u d, paz n(6), nsch n(6), s_pred n(13,2), "status" c(25), ok l)
 SELECT answer
 INDEX on lpuid TAG lpuid 
 SET ORDER TO lpuid
 FOR m.n_rec = 0 TO m.n_recs-1
  m.orec = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getBillStatusesResponse/list').item(m.n_rec)
  m.billType = orec.selectNodes('billType').item(0).text
  IF m.billType != m.TypeOfBill
   LOOP 
  ENDIF 
  
  m.recid  = PADL(m.n_rec+1,6,'0')
  m.billid = orec.selectNodes('billId').item(0).text
  m.lpuid  = INT(VAL(orec.selectNodes('moId').item(0).text))
  m.paz    = INT(VAL(orec.selectNodes('patients_count').item(0).text))
  m.nsch   = INT(VAL(orec.selectNodes('invoice_count').item(0).text))
  m.s_pred = VAL(LEFT(ALLTRIM(orec.selectNodes('amount').item(0).text), LEN(ALLTRIM(orec.selectNodes('amount').item(0).text))-2) + '.' + ;
 	RIGHT(ALLTRIM(orec.selectNodes('amount').item(0).text),2))
  m.status = orec.selectNodes('status').item(0).text
  m.d_u    = LEFT(orec.selectNodes('statusChangeDate').item(0).text,10)
  m.d_u    = CTOD(SUBSTR(m.d_u,9,2)+'.'+SUBSTR(m.d_u,6,2)+'.'+SUBSTR(m.d_u,1,4))
  
  INSERT INTO answer FROM MEMVAR 

 ENDFOR 
 
 IF fso.FileExists(pbase+'\'+gcperiod+'\BillStatuses.dbf')
  fso.DeleteFile(pbase+'\'+gcperiod+'\BillStatuses.dbf')
 ENDIF 
 COPY TO &pbase\&gcperiod\BillStatuses

 IF !ISNULL(m.loForm)
  *WITH m.loForm
  * .paz     = m.paz
  * .nsch    = m.nsch
  * .s_pred  = m.s_pred
  * .d_u     = m.d_u
  * .billid  = m.billid
  * .soapsts = m.soapsts
  *ENDWITH 
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
   USE IN answer
   RETURN .F.
  ENDIF 
 ENDIF !ISNULL(m.loForm)
 
 WAIT "CheckStatus (ОБРАБОТКА)..." WINDOW NOWAIT 
 SELECT aisoms 
 m.flt = FILTER('aisoms')
 IF !EMPTY(m.flt)
  SET FILTER TO IN aisoms 
 ENDIF 
 orec = RECNO('aisoms')
 SET RELATION TO lpuid INTO answer 
 SCAN 
  m.lpuid = lpuid 
  IF !EMPTY(bname)
   *LOOP 
  ENDIF  
  IF m.lpu_id>0 AND m.lpu_id<>m.lpuid
   LOOP 
  ENDIF
  
  IF LOCK()
   REPLACE billid WITH answer.billid, soapsts WITH answer.status
   UNLOCK RECORD RECNO()
  ENDIF 
  
  IF !EMPTY(bname)
   IF INLIST(soapsts, 'SENT', 'RECIEVED', 'RECIEVED_E', 'ACCEPTED')
    IF LOCK()
     REPLACE bname WITH '', dname WITH ''
     UNLOCK RECORD RECNO()
    ENDIF 
   ENDIF 
  ENDIF  
  
  IF aisoms.s_pred>0 AND aisoms.s_pred+aisoms.s_lek != answer.s_pred AND INLIST(answer.status, 'ACCEPTED')
   oAlertMgr = CREATEOBJECT("VFPAlert.AlertManager")
   poAlert = oAlertMgr.NewAlert()
   poAlert.Alert(TRANSFORM(aisoms.s_pred+aisoms.s_lek,'9999999999.99')+' !=: '+TRANSFORM(answer.s_pred,'9999999999.99'), 8, "Расхождение в суммах!",;
  	IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, ''))
*   MESSAGEBOX('РАСЧИТАННАЯ МО СУММА ПРЕДСТАВЛЕННОГО СЧЕТА: '+TRANSFORM(aisoms.s_pred,'9999999.99')+CHR(13)+CHR(10)+;
 	'НЕ СОВПАДАТЕТ С СУММОЙ, РАСЧИТАННОЙ СМО   : '+TRANSFORM(answer.s_pred,'9999999.99'),0+64,STR(m.lpuid,4))
  ENDIF 
  IF !ISNULL(m.loForm)
   loForm.refresh
  ENDIF 
 ENDSCAN 

 IF !EMPTY(m.flt)
  SET FILTER TO &flt IN aisoms 
 ENDIF 

 SET RELATION OFF INTO answer
 USE IN answer 
 
 IF BETWEEN(m.orec, 1, RECCOUNT('aisoms'))
  GO (orec)
 ENDIF 
 
 WAIT CLEAR 

 IF !ISNULL(m.loForm)
  loForm.refresh
 ENDIF 

 IF !m.lIsSilent
  IF m.lpu_id>0
   MESSAGEBOX('СТАТУС: ' + m.status, 0+64, STR(m.lpu_id,4))
  ELSE 
   MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64,'')
  ENDIF 
 ENDIF 
 
RETURN .T.