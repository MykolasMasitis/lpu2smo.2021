FUNCTION getAttachment_t(para1, para2, para3) && mailGWlogid!
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
  IF !m.lIsSilent
   MESSAGEBOX('ПУСТОЙ mailGWlogid!',0+64,'getAttachment')
  ENDIF 
  RETURN .F.
 ENDIF 
* m.mailGWlogid = para1 && char(7)
* m.lIsSilent   = para2
* m.loForm      = para3
* m.mailGWlogid = IIF(!EMPTY(m.mailGWlogid), m.mailGWlogid, '1586256')

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

 oRequest = oXML.createElement("ws:getAttachment")
 oBody.appendChild(oRequest)
 
 oClient = CreateClientPump(oXml, '')
 oRequest.appendChild(oClient)
 
 oXXX = oXML.createElement("request")
 oXXX.appendChild(oXML.createElement("mailGWlogid")).text = m.mailGWlogid
 oRequest.appendChild(oXXX)
 
 IF oXML.parseError.errorCode != 0 
*  MESSAGEBOX(oXML.parseError.reason,0+64,'')
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 
 
 oXML.save('&curSoapDir\OUTPUT\&xmlFile')
 length = fso.GetFile('&curSoapDir\OUTPUT\&xmlFile').Size
 
 LOCAL oEx as Exception

 m.err = .f. 
IF 3=2
 TRY 
  *ohttp.open('post', m.address, .f.) && .f. - синхронное соединение, .t. - асинхронное соединение!
  ohttp.open('post', m.address, .t.) && .f. - синхронное соединение, .t. - асинхронное соединение!
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
ENDIF  
 IF m.err = .t. 
  RELEASE oXML, oHttp
 *MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN .F.
 ENDIF 
IF 3=2 
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
ENDIF 
 IF m.err = .t. 
  RELEASE oXML, oHttp
*  MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN .F.
 ENDIF 

 m.IsCancelled = .f.
IF 3=2 
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
* IF  ohttp.status<>200
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
 
 poi   = FCREATE('&curSoapDir\INPUT\&xmlFile')
 nSize = FWRITE(poi, ohttp.responseBody)
 =FCLOSE(poi)
ENDIF   
 
 m.httpFile = '001.txt'
 m.xmlFile  = '001.xml'
 
 poi = fso.OpenTextFile('&curSoapDir\INPUT\&httpFile')
 m.FlInMem = poi.ReadAll
 poi.Close
 poi = fso.OpenTextFile('&curSoapDir\INPUT\&httpFile')
 m.boundary   = Element('boundary', ReadTheHead(poi, 'Content-Type'))
 MESSAGEBOX(m.boundary,0+64,'')
 poi.Close

 m.realBound = ALLTRIM(MLINE(m.FlInMem, ATLINE(m.boundary, m.FlInMem)))
 m.finalBound = m.realBound + '--'
 poi = fso.OpenTextFile('&curSoapDir\INPUT\&httpFile')
 m.start      = Element('start', ReadTheHead(poi, 'Content-Type'))
 MESSAGEBOX(m.start,0+64,'')
 poi.Close

 poi = fso.OpenTextFile('&curSoapDir\INPUT\&httpFile')
 m.start_info = Element('start-info', ReadTheHead(poi, 'Content-Type'))
 MESSAGEBOX(m.start_info,0+64,'')
 poi.Close

 poi = fso.OpenTextFile('&curSoapDir\INPUT\&xmlFile')
 ohttp = poi.ReadAll

 *Попытка без сохранения файла выделить из него xml
 m.XmlFromFile = ExtractEnvelope(ohttp)
 IF EMPTY(m.XmlFromFile)
  IF !m.lIsSilent
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  ENDIF 
  RETURN .T.
 ENDIF 

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 IF !oxml.loadXML(m.XmlFromFile)
  RELEASE oXML
  IF !m.lIsSilent
   MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  ENDIF 
  RETURN .T.
 ENDIF 
* oXml.save('&curSoapDir\INPUT\&xmlFile')
 oXml.save('&curSoapDir\INPUT\002.xml')
 *Попытка без сохранения файла выделить из него xml

* IF m.httpStatus=500
 IF 1=2
  poi = fso.OpenTextFile('&curSoapDir\INPUT\&xmlFile')
  m.bodypart = poi.ReadLine
  DO WHILE m.bodypart != '<soap:Envelope'
   m.bodypart = poi.ReadLine
  ENDDO 
  DO WHILE RIGHT(RTRIM(m.bodypart),16) != '</soap:Envelope'
  m.bodypart = m.bodypart + poi.ReadLine
  ENDDO 
  poi.close	
  
  oXML  = CREATEOBJECT("MsXml2.DOMDocument")
  IF m.bodypart='<soap:Envelope'
   IF !oxml.loadXML(m.bodypart)
    RELEASE oXML
    IF !m.lIsSilent
     MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ XML ФАЙЛ!', 0+64, STR(m.lpu_id,4))
    ENDIF 
    RETURN .F.
   ENDIF 
  ELSE 
   RELEASE oXML
   IF !m.lIsSilent
    MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!', 0+64, STR(m.lpu_id,4))
   ENDIF 
   RETURN .F.
  ENDIF 

  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs=0
   RELEASE oXml
   IF !m.lIsSilent
    MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'')
   ENDIF 
   RETURN .F.
  ELSE 
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text
   
   IF !m.lIsSilent
    MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, STR(m.lpu_id,4))
   ENDIF 
  
   RELEASE oXml
   RETURN .F.
  ENDIF 
 ENDIF 

* poi = fso.OpenTextFile('&curSoapDir\INPUT\&xmlFile')
* m.bodypart = poi.ReadLine
* DO WHILE m.bodypart != 'Content-Disposition'
*  m.bodypart = poi.ReadLine
* ENDDO 
* poi.close	
 
* IF m.bodypart != 'Content-Disposition'
*  IF !m.lIsSilent
*   MESSAGEBOX('В ПОЛУЧЕННОМ ФАЙЛЕ ОТСУТСТВУЕТ Content-Disposition!',0+64,'')
*  ENDIF 
*  RETURN .F.
* ENDIF 
 
* IF OCCURS('"',m.bodypart)!=2
*  IF !m.lIsSilent
*   MESSAGEBOX('В ПОЛУЧЕННОМ НЕКОРРЕКТНЫЙ Content-Disposition!',0+64,'2')
*  ENDIF 
*  RETURN .F.
* ENDIF 
 
* m.fname = STRTRAN(ALLTRIM(SUBSTR(m.bodypart, AT('"',m.bodypart))), '"', '')
* IF !INLIST(LEN(m.fname), 12, 9)
*  IF !m.lIsSilent
*   MESSAGEBOX('В ПОЛУЧЕННОМ НЕКОРРЕКТНЫЙ name: '+m.fname, 0+64, 'name')
*  ENDIF 
*  RETURN .F.
* ENDIF 

 m.n_len = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getAttachmentResponse/attachment/attachmentName').length
 IF m.n_len <> 1
  MESSAGEBOX('ЗНАЧЕНИЕ m.n_len='+STR(m.n_len,3), 0+64, 'getAttachment')
  RELEASE oXml
  RETURN .T.
 ELSE 
  MESSAGEBOX('ЗНАЧЕНИЕ m.n_len='+STR(m.n_len,3), 0+64, 'getAttachment')
 ENDIF 
  
 m.fname = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getAttachmentResponse/attachment/attachmentName').item(0).text
 m.f_crc32 = oxml.selectNodes('soap:Envelope/soap:Body/ns2:getAttachmentResponse/attachment/checkSum').item(0).text

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
  m.FileInMem = ohttp
  m.zip = m.FileInMem
   
  IF fso.FileExists(m.curSoapDir+'\INPUT\'+m.fname)
   fso.DeleteFile(m.curSoapDir+'\INPUT\'+m.fname)
  ENDIF 
  m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip)), m.curSoapDir+'\INPUT\'+m.fname)

  IF !m.lIsSilent
   MESSAGEBOX('НЕКОРРЕКТОЕ НАИМЕНОВАНИЕ ФАЙЛА: '+m.fname+CHR(13)+CHR(10)+'ФАЙЛ СОХРАНЕН В '+m.curSoapDir+'\INPUT\', 0+64, '')
  ENDIF 
  RETURN .F.
 ENDIF 

 m.FileInMem = ohttp
 m.zip = m.FileInMem
 IF AT('PK', m.zip)=0
  RELEASE m.zip
  IF !m.lIsSilent
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ ZIP НЕ ОБНАРУЖЕН!',0+64,'%PDF')
  ENDIF 
  RETURN .F.
 ENDIF 

 IF fso.FolderExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod)

  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)
  ENDIF 
  m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip)), ;
 	 m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)

  IF !m.lIsSilent
   MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64,'')
  ENDIF 
  
  RETURN .T.

 ELSE 

  IF fso.FileExists(m.curSoapDir+'\INPUT\'+m.fname)
   fso.DeleteFile(m.curSoapDir+'\INPUT\'+m.fname)
  ENDIF 
  m.nBytes = STRTOFILE(SUBSTR(m.zip, AT('PK',m.zip)), ;
 	 m.curSoapDir+'\INPUT\'+m.fname)
 	 
  IF !m.lIsSilent
   MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64,'')
  ENDIF 

  RETURN .F.

 ENDIF 
