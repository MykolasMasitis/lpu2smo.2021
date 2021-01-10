PROCEDURE erzsend_p
 LPARAMETERS para1

 m.oPrm      = para1

 WITH oPrm
  m.pBase    = .pBase
  m.pTempl   = .pTempl
  m.gcPeriod = .gcPeriod
  m.qCod     = .qCod
  m.qName    = .qname
 
  m.tMonth   = .tMonth
  m.tYear    = .tYear

  m.mcod     = .mcod
  m.lpuid    = .lpuid
   
  m.orgCode    = .orgCode
  m.soapSystem = .soapSystem
  m.erzlUser   = .erzlUser
  m.erzlPass   = .erzlPass
 ENDWITH 

 IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar')>0
  IF USED('people')
   USE IN people
  ENDIF 
  *SELECT aisoms
  *LOOP 
  RETURN 
 ENDIF 

  * В конфигурационном файле ищет сервис erzlsmowebsvc.wsdl
  * m.address = WS2Address('erzlsmowebsvc.wsdl')
  * m.host = SUBSTR(m.address, AT('//',m.address)+2, RAT('/',m.address) - (AT('//',m.address) + 2)) && 192.168.192.106:9090
  m.address = 'http://192.168.192.118:8080/erzl-for-smo/ws/'
  m.host = '192.168.192.118:8080'
  * Должны получить вот такой адрес: 'http://192.168.192.106:9090/ws', полный адрес сервиса: http://192.168.192.106:9090/ws/erzlsmowebsvc.wsdl
  * Адрес боевого сервиса: http://192.168.192.106:8080/erzl-for-smo/ws/erzlsmowebsvc.wsdl 
  IF EMPTY(m.address)
   MESSAGEBOX('СЕРВИС НЕДОСТУПЕН!',0+16,'m.address?')
   RETURN
  ENDIF 

  * Генерируем уникальные имена 
  *m.un_id    = SYS(3)
  *m.rqHTTP = m.un_id + '.txt'
  *m.rqXML  = m.un_id + '.xml'
  * Генерируем уникальные имена 

  m.bpCode = 101 && Поиск и просмотр записи о ЗЛ ОМС Москвы

  m.rqHTTP  = 'request.http'
  m.rqXML   = 'request.xml'
  m.ansHTTP = 'polltag.http'
  m.ansXML  = 'polltag.xml'

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
  *SELECT aisoms
 
  IF oXML.parseError.errorCode != 0 
   MESSAGEBOX(oXML.parseError.reason,0+64,'')
   RELEASE oXML, oHttp
   RETURN 
   *EXIT 
  ENDIF 
 
  oXML.save(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rqXML)
  length = fso.GetFile(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rqXML).Size
 
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
   *EXIT 
  ENDIF 
 
  CreateHeader(ohttp, m.length, m.host, .T., m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+rqHTTP)
 
  TRY 
   ohttp.send(oXml.xml) && Для get-запросов тела нет, был бы null, для post - есть, поэтому передаем парметр
  CATCH TO oEx
   m.err = .t. 
  ENDTRY 

  IF m.err = .t. 
   RELEASE oXML, oHttp
   MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
   RETURN 
   *EXIT 
  ENDIF 
 
  m.IsCancelled = .f.
  DO WHILE ohttp.readyState<4
   *WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 

   IF CHRSAW(0) 
    IF INKEY() == 27
     WAIT CLEAR 
     IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОЖИДАНИЕ ОТВЕТА?',4+32,'') == 6
      KEYBOARD '{ESC}'
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
   *EXIT 
  ENDIF 

  IF  ohttp.status<>200
   MESSAGEBOX('ОШИБКА ЗАПРОСА СТАТУС '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
   RELEASE oXML, oHttp
   RETURN 
   *EXIT 
  ENDIF 
 
  *Сохраняем http-заголовок ответа
  poi = fso.CreateTextFile(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ansHTTP)
  poi.Write(ohttp.getAllResponseHeaders())
  poi.Close
  m.cdate = ""

  CFG = FOPEN(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ansHTTP)
  =ReadHTTPHead()
  =FCLOSE(CFG)
  m.polltagdt = RFC2date(m.cdate) + 3*60*60
  * Сохраняем http-заголовок ответа
 
  poi = fso.CreateTextFile(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+ansXML)
  poi.Write(ohttp.responseText)
  poi.Close

  poi = fso.OpenTextFile(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+ansXML)
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

    *MESSAGEBOX('pollTag: '+m.pollTag,0+64,'')

   ELSE 

    *MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ pollTag НЕ ОБНАРУЖЕН!',0+64,'')

   ENDIF 

  ELSE 
 
   *MESSAGEBOX('В ПОЛУЧЕННОМ ОТВЕТЕ XML НЕ ОБНАРУЖЕН!',0+64,'')

  ENDIF 
  RELEASE oXML, oHttp
 
  *MESSAGEBOX('pollTag: '+m.pollTag,0+64,'Дима!')
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
   IF USED('aisoms')
    USE IN aisoms 
   ENDIF 
   RETURN 
  ENDIF 
  *REPLACE polltag WITH m.polltag, polltagdt WITH m.polltagdt, erz_status WITH 1 IN aisoms 
  UPDATE aisoms SET polltag=m.polltag, polltagdt=m.polltagdt, erz_status=1, soapstatus='' WHERE mcod = m.mcod
  *SELECT aisoms 
  USE IN aisoms 
  RETURN m.mcod 
