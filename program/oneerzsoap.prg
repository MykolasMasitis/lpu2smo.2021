FUNCTION OneERZSoap(lcDir, IsOK)

 m.t_beg = SECONDS()
 IF !fso.FolderExists(lcDir)
  RETURN .F.
 ENDIF 

 IF !fso.FileExists(lcDir + '\People.dbf')
  RETURN .F.
 ENDIF 

 IF OpenFile("&lcDir\People", "People", "SHARE")>0
  RETURN .F.
 ENDIF 
 
 IF RECCOUNT('people')<=0
  USE IN people 
  RETURN .F.
 ENDIF 

 * В конфигурационном файле ищет сервис erzlsmowebsvc.wsdl
 *m.address = WS2Address('erzlsmowebsvc.wsdl')
 *m.host = SUBSTR(m.address, AT('//',m.address)+2, RAT('/',m.address) - (AT('//',m.address) + 2)) && 192.168.192.106:9090
  m.address = 'http://192.168.192.118:8080/erzl-for-smo/ws/'
  m.host = '192.168.192.118:8080'
 * Должны получить вот такой адрес: 'http://192.168.192.106:9090/ws', полный адрес сервиса: http://192.168.192.106:9090/ws/erzlsmowebsvc.wsdl
 * Адрес боевого сервиса: http://192.168.192.106:8080/erzl-for-smo/ws/erzlsmowebsvc.wsdl 
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

   m.bpCode = 101 && Поиск и просмотр записи о ЗЛ ОМС Москвы
 * m.bpCode = 102 && Поиск и просмотр записи об иногородних
 * m.bpCode = 103 && Поиск и просмотр записи о новорожденных
 * m.bpCode = 104 && Поиск и просмотр сведений о прикреплениях ЗЛ
 * m.bpCode = 105 && Утверждение прикрепления по смене места жительства
 * m.bpCode = 106 && Аннулирование прикрепления
 * m.bpCode = 107 && Поиск и просмотр сведений о результатах запросов к ЦС ЕРЗЛ
 * m.bpCode = 108 && Просмотр рекомендаций пользователям по дальнейшим действиям (директив)

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

 oRoot = CreateEnvelope(oXML)

 oXML.appendChild(oRoot)
 
 oNode = oXML.createElement("soapenv:Header")
 oRoot.appendChild(oNode)

 oBody = oXML.createElement("soapenv:Body")
 oRoot.appendChild(oBody)

 oRequest = oXML.createElement("ser:getPersonInsuranceDataAsyncRequest")
 oBody.appendChild(oRequest)
 
 oClient = CreateClient(oXml, m.bpcode)
 oRequest.appendChild(oClient)
 
 SELECT people
 SCAN 
  SCATTER MEMVAR 
  
  m.date_in  = IIF(!EMPTY(m.d_end), m.d_end, m.tdat2) && С 07.03.2018
  m.date_out = IIF(!EMPTY(m.d_end), m.d_end, m.tdat2)
  m.d_out    = STR(YEAR(m.date_out),4) + '-' + PADL(MONTH(m.date_out),2,'0') + '-' + PADL(DAY(m.date_out),2,'0')

  m.recid    = PADL(m.recid,6,'0')
  m.fam      = m.fam
  m.im       = m.im
  m.ot       = m.ot
  m.q        = m.qcod
  
  opRq = Create_pRq(oXml, ALLTRIM(m.sn_pol), m.d_out)
  oRequest.appendChild(opRq)
    
 ENDSCAN 
 USE IN people 
 
 IF oXML.parseError.errorCode != 0 
  MESSAGEBOX(oXML.parseError.reason,0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
 oXML.save('&curSoapDir\OUTPUT\&xmlFile')
 length = fso.GetFile('&curSoapDir\OUTPUT\&xmlFile').Size
 
 LOCAL oEx as Exception
 
 m.t_rqst = SECONDS()

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
 
 CreateHeader(ohttp, m.length, m.host, .T., curSoapDir+'\OUTPUT\'+httpFile)

 *ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 *ohttp.setRequestHeader("Content-Type", "text/xml; charset=UTF-8")
 *ohttp.setRequestHeader("SOAPAction", "")
 *ohttp.setRequestHeader("Content-Length", m.length)
 *ohttp.setRequestHeader("Host", m.host)
 *ohttp.setRequestHeader("Connection", "Keep-Alive")
 *ohttp.setRequestHeader("User-Agent", "Apache-HttpClient/4.1.1 (java 1.5)")
 
 *poi = fso.CreateTextFile('&curSoapDir\OUTPUT\&httpFile')
 *poi.WriteLine('Accept-Encoding: gzip,deflate')
 *poi.WriteLine('Content-Type: "text/xml; charset=UTF-8"')
 *poi.WriteLine('SOAPAction: ""')
 *poi.WriteLine('Content-Length: '+ALLTRIM(STR(m.length)))
 *poi.WriteLine('Host: ' + m.host)
 *poi.WriteLine('Connection: Keep-Alive')
 *poi.WriteLine('User-Agent: "Apache-HttpClient/4.1.1 (java 1.5)"')
 *poi.Close
 
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
 
 m.t_send_rq = SECONDS()

* MESSAGEBOX('ОТПРАВКА НА МОДУЛЬ'+CHR(13)+CHR(10)+m.address+'/erzlsmowebsvc.wsdl'+CHR(13)+CHR(10)+'ПРОШЛА УСПЕШНО!',0+64,'')
 
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

 IF  ohttp.status<>200
  MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
 * Сохраняем http-заголовок
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * Сохраняем http-заголовок
 
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&xmlFile')
 poi.Write(ohttp.responseText)
 poi.Close

 poi = fso.OpenTextFile('&curSoapDir\INPUT\&xmlFile')
 m.bodypart = poi.ReadLine
 DO WHILE m.bodypart != '<SOAP-ENV:Envelope'
  m.bodypart = poi.ReadLine
 ENDDO 
 poi.close	
  
 IF m.bodypart='<SOAP-ENV:Envelope'
  oxml.loadXML(m.bodypart)
  
  m.pollTag = ''
  IF oxml.selectNodes('//SOAP-ENV:Envelope/SOAP-ENV:Body/ns2:getPersonInsuranceDataAsyncResponse/ns2:pollTag').length>0
   m.pollTag = oxml.selectNodes('//SOAP-ENV:Envelope/SOAP-ENV:Body/ns2:getPersonInsuranceDataAsyncResponse/ns2:pollTag').item(0).text

*   RELEASE oXML, oHttp
*   MESSAGEBOX('pollTag: '+m.pollTag,0+64,'')

  ELSE 

   RELEASE oXML, oHttp
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ pollTag НЕ ОБНАРУЖЕН!',0+64,'')

  ENDIF 

 ELSE 
 
  RELEASE oXML, oHttp
  MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!',0+64,'')

 ENDIF 
 
 m.t_poll_tag = SECONDS()
 
 m.status = '?'
 DO WHILE m.status!='F'

  oHttp = CREATEOBJECT("MsXml2.XMLHTTP")
  oXML  = CREATEOBJECT("MsXml2.DOMDocument")
  oXML.resolveExternals = .T.
  oHttp = CREATEOBJECT("MsXml2.XMLHTTP")

  m.un_id    = SYS(3)
  m.httpFile = m.un_id + '.txt'
  m.xmlFile  = m.un_id + '.xml'
  m.zipFile  = m.un_id + '.zip'

  * Create a procesing instruction.  
  oNode = oXML.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
 
  oRoot = CreateEnvelope(oXML)

  oXML.appendChild(oRoot)
 
  oNode = oXML.createElement("soapenv:Header")
  oRoot.appendChild(oNode)

  oBody = oXML.createElement("soapenv:Body")
  oRoot.appendChild(oBody)

  oRequest = oXML.createElement("ser:pollPersonInsuranceDataRequest")
  oBody.appendChild(oRequest)
 
  oClient = CreateClient(oXml, m.bpcode)
  oRequest.appendChild(oClient)
  oRequest.appendChild(oXML.createElement("ser:pollTag")).text = m.pollTag
 
  IF oXML.parseError.errorCode != 0 
   RELEASE oXML, oHttp
   MESSAGEBOX(oXML.parseError.reason,0+64,'')
   LOOP 
  ENDIF 
 
  oXML.save('&curSoapDir\OUTPUT\&xmlFile')
  length = fso.GetFile('&curSoapDir\OUTPUT\&xmlFile').Size
  fso.DeleteFile('&curSoapDir\OUTPUT\&xmlFile')

  LOCAL oEx as Exception

  m.err = .f. 
  TRY 
   *ohttp.open('post', m.address, .f.) && .f. - синхронное соединение, .t. - асинхронное соединение!
   ohttp.open('post', m.address, .t.) && .f. - синхронное соединение, .t. - асинхронное соединение!
  CATCH TO oEx
   m.err = .t. 
  ENDTRY 
 
  IF m.err = .t. 
   RELEASE oXML, oHttp
   MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
   RETURN 
  ENDIF 

  CreateHeader(ohttp, m.length, m.host, .F.)
  *ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
  *ohttp.setRequestHeader("Content-Type", "text/xml; charset=UTF-8")
  *ohttp.setRequestHeader("SOAPAction", "")
  *ohttp.setRequestHeader("Content-Length", m.length)
  *ohttp.setRequestHeader("Host", m.host)
  *ohttp.setRequestHeader("Connection", "Keep-Alive")
  *ohttp.setRequestHeader("User-Agent", "Apache-HttpClient/4.1.1 (java 1.5)")
 
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

  IF  ohttp.status<>200
   MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
   RELEASE oXML, oHttp
   RETURN 
  ENDIF 
 
  * Сохраняем http-заголовок
*  poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
*  poi.Write(ohttp.getAllResponseHeaders())
*  poi.Close
  * Сохраняем http-заголовок
 
  poi = fso.CreateTextFile('&curSoapDir\INPUT\&xmlFile')
  poi.Write(ohttp.responseText)
  poi.Close

  poi = fso.OpenTextFile('&curSoapDir\INPUT\&xmlFile')
  m.bodypart = poi.ReadLine
  DO WHILE m.bodypart != '<SOAP-ENV:Envelope'
   m.bodypart = poi.ReadLine
  ENDDO 
  poi.close	
  fso.DeleteFile('&curSoapDir\INPUT\&xmlFile')
  
  IF m.bodypart='<SOAP-ENV:Envelope'
   oxml.loadXML(m.bodypart)
  
   IF oxml.selectNodes('//SOAP-ENV:Envelope/SOAP-ENV:Body/ns2:pollPersonInsuranceDataResponse/ns2:status').length>0
    m.status = oxml.selectNodes('//SOAP-ENV:Envelope/SOAP-ENV:Body/ns2:pollPersonInsuranceDataResponse/ns2:status').item(0).text
   ELSE 
    RELEASE oXML, oHttp
    MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ status НЕ ОБНАРУЖЕН!',0+64,'status')
    RETURN 
   ENDIF 
  ELSE 
   RELEASE oXML, oHttp
   MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ <SOAP-ENV:Envelope> НЕ ОБНАРУЖЕН!',0+64,'status!')
   RETURN 
  ENDIF 
 
 ENDDO 
 
 m.t_ans = SECONDS()
  
 * Сохраняем http-заголовок
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * Сохраняем http-заголовок
 
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&xmlFile')
 poi.Write(ohttp.responseText)
 poi.Close

 poi   = FCREATE('&curSoapDir\INPUT\&zipFile')
 nSize = FWRITE(poi, ohttp.responseBody)
 
 IF !FCLOSE(poi)
  MESSAGEBOX('НЕ УДАЛОСЬ ЗАКРЫТЬ ФАЙЛ'+CHR(13)+CHR(10)+m.curSoapDir+'\INPUT\'+zipFile,0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

 IF !UnzipOpen('&curSoapDir\INPUT\&zipFile')
  MESSAGEBOX('ПОЛУЧЕННЫЙ ФАЙЛ НЕ ZIP-АРХИВ!', 0+64, '')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

 ZipDir = curSoapDir + '\INPUT\'
 IF !UnzipGotoFileByName('data.xml')
  UnzipClose()
  MESSAGEBOX('В ПОЛУЧЕННОМ ZIP-АРХИВЕ НЕ СОДЕРЖИТСЯ DATA.XML!', 0+64, '')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
 m.ZipDir = curSoapDir + '\INPUT\'
 UnzipFile(ZipDir)
 UnzipClose()

 m.un_id    = SYS(3)
 m.xmlFile  = m.un_id + '.xml'
 fso.CopyFile(curSoapDir+'\INPUT\data.xml', curSoapDir+'\INPUT\'+m.xmlFile, .t.)
  
 RELEASE oXML, oHttp
 
 m.t_save_ans = SECONDS()
 
 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 WAIT "ЗАГРУЗКА XML..." WINDOW NOWAIT 
* IF !oxml.load('&curSoapDir\INPUT\m.xmlFile')
 IF !oxml.load('&curSoapDir\INPUT\data.xml')
  RELEASE oXml
*  MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ '+m.xmlFile+' ФАЙЛ!',0+64,'oxml.load()')
  MESSAGEBOX('НЕ УДАЛОСЬ ЗАГРУЗИТЬ data.xml ФАЙЛ!',0+64,'oxml.load()')
  RETURN 
 ENDIF 
 WAIT CLEAR 
 
 m.n_recs = oxml.selectNodes('PersonInsuranceDataSet/data').length
 IF m.n_recs=0
  RELEASE oXml
  MESSAGEBOX('В ОТВЕТЕ НИ ОДНОЙ ЗАПИСИ!',0+64,'')
  RETURN 
 ENDIF 
 
 CREATE CURSOR answer (recid c(6), s_pol c(6), n_pol c(16), d_u c(8), q c(2), fam c(25), im c(20), ot c(20), ;
  dr c(8), w n(1), ans_r c(3), tip_d c(1), lpu_id n(6), st_id n(6), d_rq d, d_end d)
 INDEX on n_pol TAG n_pol 
 
 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 WAIT "XML->DBF..."  WINDOW NOWAIT 
 FOR m.n_rec = 0 TO m.n_recs-1

  m.orec = oxml.selectNodes('PersonInsuranceDataSet/data').item(m.n_rec)

  m.d_rq = orec.selectNodes('rqDate').item(0).text
  m.d_rq = STRTRAN(m.d_rq,'-','')
  m.d_rq = CTOD(SUBSTR(m.d_rq,7,2)+'.'+SUBSTR(m.d_rq,5,2)+'.'+SUBSTR(m.d_rq,1,4))

  m.recid = PADL(m.n_rec+1,6,'0')
  m.n_pol = orec.selectNodes('policySerNum').item(0).text
  m.d_u   = orec.selectNodes('rqDate').item(0).text
  m.d_u   = STRTRAN(m.d_u, '-', '')

  m.fam    = ""
  m.im     = ""
  m.ot     = ""
  m.dr     = ""
  m.w      = 0
  m.tip_d  = ""
  m.lpu_id = 0
  m.st_id  = 0
  m.ans_r  = '0*0'

  IF orec.selectNodes('person/surname').length<=0
   INSERT INTO answer FROM MEMVAR 
   LOOP 
  ENDIF 
  
  m.fam   = orec.selectNodes('person/surname').item(0).text
  m.ans_r = '211'
  m.im    = orec.selectNodes('person/namep').item(0).text
  IF orec.selectNodes('person/patronymic').length>0
   m.ot    = orec.selectNodes('person/patronymic').item(0).text
  ENDIF 
  m.w     = orec.selectNodes('person/sexId').item(0).text
  m.w     = INT(VAL(m.w))
  m.dr     = orec.selectNodes('person/dateBirth').item(0).text
  m.dr    = STRTRAN(m.dr,'-','')

  IF orec.selectNodes('policy/policyTCode').length<=0
   INSERT INTO answer FROM MEMVAR 
   LOOP 
  ENDIF 
  
  m.tip_d = orec.selectNodes('policy/policyTCode').item(0).text
  m.q     = orec.selectNodes('policy/insuranceQQ').item(0).text

  m.d_end = orec.selectNodes('policy/plDateE').item(0).text
  m.d_end = STRTRAN(m.d_end,'-','')
  m.d_end = CTOD(SUBSTR(m.d_end,7,2)+'.'+SUBSTR(m.d_end,5,2)+'.'+SUBSTR(m.d_end,1,4))
  
  m.n_atts = orec.selectNodes('attach').length
  m.dAttachB = {}
  m.dAttachE = {}

  IF m.n_atts>0
   FOR m.n_att = 0 TO m.n_atts-1
    m.o_att = orec.selectNodes('attach').item(m.n_att)
    m.lpu_tip = INT(VAL(o_att.selectNodes('areaTId').item(0).text))

    IF m.lpu_tip != 5
     m.lpu_id = INT(VAL(o_att.selectNodes('mo/moCode').item(0).text))
    ELSE 
     IF o_att.selectNodes('dateAttachB').length>0
      m.dateAttachB = o_att.selectNodes('dateAttachB').item(0).text
      m.dAttachB = STRTRAN(m.dateAttachB,'-','')
      m.dAttachB = CTOD(SUBSTR(m.dAttachB,7,2)+'.'+SUBSTR(m.dAttachB,5,2)+'.'+SUBSTR(m.dAttachB,1,4))
     ENDIF 
     IF o_att.selectNodes('dateAttachE').length > 0
      m.dateAttachE = o_att.selectNodes('dateAttachE').item(0).text
      m.tAttachE = STRTRAN(m.dateAttachE,'-','')
      m.tAttachE = CTOD(SUBSTR(m.tAttachE,7,2)+'.'+SUBSTR(m.tAttachE,5,2)+'.'+SUBSTR(m.tAttachE,1,4))
      m.dAttachE = MAX(m.tAttachE, m.dAttachE)
     ELSE 
      m.tAttachE = {31.12.2099}
      m.dAttachE = {31.12.2099}
     ENDIF 
     IF m.tAttachE >= m.dAttachE
      m.dAttachE = m.tAttachE
      m.st_id = INT(VAL(o_att.selectNodes('mo/moCode').item(0).text))
    *  MESSAGEBOX(DTOC(m.dAttachE)+' '+STR(m.st_id,4), 0+64, m.n_pol)
     ENDIF 
    ENDIF 
   ENDFOR 
  ENDIF 

  IF m.d_end < m.d_rq
   LOOP 
  ENDIF 

  INSERT INTO answer FROM MEMVAR 

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
     EXIT 
    ELSE 
     WAIT "XML->DBF..."  WINDOW NOWAIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDFOR 
 WAIT CLEAR 
 
 SELECT answer
 IF fso.FileExists(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\soapans.dbf')
  fso.DeleteFile(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\soapans.dbf')
 ENDIF 
 COPY TO &pbase\&gcPeriod\&mcod\soapans
 USE IN answer 
 
 m.t_ans_proc = SECONDS()

   IF OpenFile(pbase+'\'+m.gcperiod+'\'+mcod+'\soapans', 'answer', 'excl')>0
    IF USED('answer')
     USE IN answer
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
 
   WAIT m.mcod WINDOW NOWAIT 
 
   lcDir = m.pBase + '\' + m.gcperiod + '\' + m.mcod
 
   tn_result = 0
   tn_result = tn_result + OpenFile(lcDir+'\People', 'People', 'Share')
   tn_result = tn_result + OpenFile(lcDir+'\Talon', 'Talon', 'Share', 'sn_pol')
   tn_result = tn_result + OpenFile(lcDir+'\e'+mcod, 'sError', 'Share', 'rid')
   tn_result = tn_result + OpenFile(lcDir+'\e'+mcod, 'rError', 'Share', 'rrid', 'again')
   tn_result = tn_result + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoerzxx', 'OsoERZ', 'Shar', 'ans_r')
 
   IF tn_result>0
    IF USED('people')
     USE IN people
    ENDIF 
    IF USED('talon')
     USE IN talon 
    ENDIF 
    IF USED('serror')
     USE IN serror
    ENDIF 
    IF USED('rerror')
     USE IN rerror
    ENDIF 
    IF USED('answer')
     USE IN answer
    ENDIF 
    IF USED('osoerz')
     USE IN osoerz
    ENDIF 
    LOOP 
   ENDIF 
 
   m.p_recs = RECCOUNT('people')
   IF m.n_recs != m.p_recs
    * MESSAGEBOX('КОЛ-ВО ЗАПИСЕЙ В ЗАПРОСЕ: '+STR(m.p_recs,6)+CHR(13)+CHR(10)+'НЕ СООТВЕТСТВУЕТ'+CHR(13)+CHR(10)+;
  	'КОЛ-ВУ ЗАПИСЕЙ В ОТВЕТЕ: '+STR(m.n_recs,6), 0+64, m.mcod)
   ENDIF 
 
   SELECT Answer 
   DELETE TAG ALL 
   SET SAFETY OFF
   INDEX ON RecId TAG RecId
   INDEX on n_pol TAG n_pol
   SET SAFETY ON 
   * SET ORDER TO RecId
   SET ORDER TO n_pol
 
   SELECT People
   * SET RELATION TO PADL(RecId,6,'0') INTO Answer
   SET RELATION TO LEFT(sn_pol,16) INTO Answer
 
 SCAN 
  m.recid  = recid
  m.sn_pol = sn_pol
  m.IsVs = IsVs(m.sn_pol)
  m.d_type = d_type
  
  IF EMPTY(Answer.RecId)
*   MESSAGEBOX('ВНИМАНИЕ!'+CHR(13)+CHR(10)+;
   	'ФАЙЛ ЗАПРОСА НЕ ПОЛНОСТЬЮ СВЯЗАЛСЯ С ФАЙЛОМ ОТВЕТА!'+CHR(13)+CHR(10)+m.sn_pol, 0+64, m.mcod)
    REPLACE prmcod WITH '', prmcods WITH ''
   LOOP 
  ENDIF 
  
  m.llpu = Answer.lpu_id
  m.slpu = Answer.st_id
  m.prmcod  = ''
  m.prmcods = ''

  IF m.llpu>0
   IF SEEK(m.llpu, 'pilot')
    m.prmcod = IIF(SEEK(m.llpu, 'sprlpu'), sprlpu.mcod, '')
   ELSE 
    IF m.slpu>0 AND SEEK(m.slpu, 'pilot')
     m.prmcod = IIF(SEEK(m.slpu, 'sprlpu'), sprlpu.mcod, '')
    ELSE 
     m.prmcod  = ''
*     MESSAGEBOX('ВНИМАНИЕ!'+CHR(13)+CHR(10)+;
     'LPU_ID '+STR(m.llpu,4)+' ПРИКРЕПЛЕНИЯ'+CHR(13)+CHR(10)+;
     'ОТСУТСТВУЕТ В СПРАВОЧНИКЕ PILOT.DBF!',0+48,'1')
    ENDIF 
   ENDIF 
  ENDIF 

  IF m.slpu>0
   IF SEEK(m.slpu, 'pilots')
    m.prmcods = IIF(SEEK(m.slpu, 'sprlpu'), sprlpu.mcod, '')
   ELSE 
    IF m.llpu>0 AND SEEK(m.llpu, 'pilots')
     m.prmcods = IIF(SEEK(m.llpu, 'sprlpu'), sprlpu.mcod, '')
    ELSE 
     m.prmcods = ''
*     MESSAGEBOX('ВНИМАНИЕ!'+CHR(13)+CHR(10)+;
     'LPU_ID '+STR(m.llpu,4)+' ПРИКРЕПЛЕНИЯ'+CHR(13)+CHR(10)+;
     'ОТСУТСТВУЕТ В СПРАВОЧНИКЕ PILOT.DBF!',0+48,'2')
    ENDIF 
   ENDIF 
  ENDIF 
  
  REPLACE qq WITH Answer.Q, sv WITH Answer.ans_r, prmcod WITH IIF(m.d_type!='9', m.prmcod, ''), ;
  	prmcods WITH m.prmcods

    m.sv = sv
  
    m.IsGood = IIF(SEEK(m.sv, 'osoerz') AND osoerz.kl == 'y', .T., .F.)
    IF m.IsGood=.f. AND m.IsVs AND USED('kms')
     m.vvs = SUBSTR(m.sn_pol,7,9)
     IF SEEK(m.vvs, 'kms')
      m.IsGood = .t.
      REPLACE qq WITH m.qcod, sv WITH '110'
     ENDIF 
    ENDIF 
  
    IF IsGood == .f.
     IF !SEEK(m.RecId, 'rError')
      INSERT INTO rError (f, c_err, rid) VALUES ('R', 'ERA', m.RecId)
      = SEEK(m.sn_pol, 'Talon')
      DO WHILE talon.sn_pol = m.sn_pol
       m.t_recid = Talon.RecId
       IF !SEEK(m.t_recid, 'sError')
        INSERT INTO sError (f, c_err, rid) VALUES ('S', 'PKA', m.t_recid)
       ENDIF 
       SKIP IN Talon
      ENDDO 
     ENDIF 
     LOOP 
    ENDIF 

    IF qq != m.qcod
     IF !SEEK(m.RecId, 'rError')
      INSERT INTO rError (f, c_err, rid) VALUES ('R', 'ECA', m.RecId)
      = SEEK(m.sn_pol, 'Talon')
      DO WHILE talon.sn_pol = m.sn_pol
       m.t_recid = Talon.RecId
       IF !SEEK(m.t_recid, 'sError')
        INSERT INTO sError (f, c_err, rid) VALUES ('S', 'PKA', m.t_recid)
       ENDIF 
       SKIP IN Talon
      ENDDO 
     ENDIF 
     LOOP 
    ENDIF 

   ENDSCAN 
 
   SET RELATION OFF INTO Answer
   USE 
   USE IN rError 
   USE IN sError
   SELECT Answer
   SET ORDER TO 
   DELETE TAG ALL 
   USE
   USE IN talon 
   USE IN OsoERZ

   SELECT aisoms

   REPLACE soapstatus WITH m.status, soaprcv WITH m.soaprcv, erz_status WITH 2

  RELEASE oXML, oHttp

 MESSAGEBOX('ПОДГОТОВКА ЗАПРОСА: ' + SecToHrs(m.t_rqst - m.t_beg) + CHR(13)+CHR(10)+;
 	'ОТПРАВКА ЗАПРОСА: ' + SecToHrs(m.t_send_rq - m.t_rqst) + CHR(13)+CHR(10)+;
 	'ОЖИДАНИЕ POLLTAG: ' + SecToHrs(m.t_poll_tag - m.t_send_rq) + CHR(13)+CHR(10)+;
 	'ОЖИДАНИЕ ОТВЕТА: ' + SecToHrs(m.t_ans - m.t_poll_tag) + CHR(13)+CHR(10)+;
 	'СОХРАНЕНИЕ ОТВЕТА: ' + SecToHrs(m.t_save_ans - m.t_ans) + CHR(13)+CHR(10)+;
 	'ОБРАБОТКА ОТВЕТА: ' + SecToHrs(m.t_ans_proc - m.t_save_ans) + CHR(13)+CHR(10), 0+64, '')
