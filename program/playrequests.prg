PROCEDURE PlayRequests
 IF MESSAGEBOX('��������� �������?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curss (mcod c(7), pollag c(36), polltagdt t)

 SELECT aisoms 
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  
  WAIT m.mcod+"..." WINDOW NOWAIT 
  
  m.polltag = SendRequest(m.mcod)
  
  IF !EMPTY(m.polltag)
   INSERT INTO curss FROM MEMVAR 
  ENDIF 
  
  SELECT aisoms 
  
  WAIT CLEAR 
  
 ENDSCAN
 USE IN aisoms
 
 SELECT curss
 COPY TO &pBase\&gcPeriod\allans WITH cdx 
 
 MESSAGEBOX('OK!',0+64,'')


RETURN 


FUNCTION SendRequest(para1)

  m.mcod    = para1

  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   *LOOP 
   RETURN ""
  ENDIF 
  
  IF RECCOUNT('people')<=0
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   *LOOP 
   RETURN ""
  ENDIF 

  IF !BETWEEN(RECCOUNT('people'),10,20)
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   *LOOP 
   RETURN ""
  ENDIF 

  * � ���������������� ����� ���� ������ erzlsmowebsvc.wsdl
  * m.address = WS2Address('erzlsmowebsvc.wsdl')
  * m.host = SUBSTR(m.address, AT('//',m.address)+2, RAT('/',m.address) - (AT('//',m.address) + 2)) && 192.168.192.106:9090
  m.address = 'http://192.168.192.118:8080/erzl-for-smo/ws/'
  m.host = '192.168.192.118:8080'
  * ������ �������� ��� ����� �����: 'http://192.168.192.106:9090/ws', ������ ����� �������: http://192.168.192.106:9090/ws/erzlsmowebsvc.wsdl
  * ����� ������� �������: http://192.168.192.106:8080/erzl-for-smo/ws/erzlsmowebsvc.wsdl 
  IF EMPTY(m.address)
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   MESSAGEBOX('������ ����������!',0+16,'m.address?')
   RETURN ""
  ENDIF 

  * ���������� ���������� ����� 
  *m.un_id    = SYS(3)
  *m.rqHTTP = m.un_id + '.txt'
  *m.rqXML  = m.un_id + '.xml'
  * ���������� ���������� ����� 

  m.bpCode = 101 && ����� � �������� ������ � �� ��� ������

  m.rqHTTP  = 'request.http'
  m.rqXML   = 'request.xml'
  m.ansHTTP = 'polltag.http'
  m.ansXML  = 'polltag.xml'

  *oHttp = CREATEOBJECT("MsXml2.XMLHTTP")
  oHttp = CREATEOBJECT("MsXml2.XMLHTTP.3.0")
  *oHttp = CREATEOBJECT("MsXml2.XMLHTTP.6.0")

  oXML  = CREATEOBJECT("MsXml2.DOMDocument")
  *oXML  = CREATEOBJECT("MsXml2.DOMDocument.5.0")
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
  
   m.date_in  = IIF(!EMPTY(m.d_end), m.d_end, m.tdat2) && � 07.03.2018
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
  SELECT aisoms
 
  IF oXML.parseError.errorCode != 0 
   MESSAGEBOX(oXML.parseError.reason,0+64,'')
   RELEASE oXML, oHttp
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN ""
   *EXIT 
  ENDIF 
 
  oXML.save(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rqXML)
  length = fso.GetFile(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.rqXML).Size
 
  m.err = .f. 
  TRY 
   *ohttp.open('post', m.address, .f.) && .f. - ���������� ����������, .t. - ����������� ����������!
   ohttp.open('post', m.address, .t.) && .f. - ���������� ����������, .t. - ����������� ����������!
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
   *WAIT "�������� ������..." WINDOW NOWAIT 

   IF CHRSAW(0) 
    IF INKEY() == 27
     *WAIT CLEAR 
     IF MESSAGEBOX('�� ������ �������� �������� ������?',4+32,'') == 6
      KEYBOARD '{ESC}'
      m.IsCancelled = .t.
      EXIT 
     ENDIF 
     *WAIT "�������� ������..." WINDOW NOWAIT 
    ENDIF 
   ENDIF 

  ENDDO 
 
  IF  m.IsCancelled = .t.
   RELEASE oXML, oHttp
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN ""
   *EXIT 
  ENDIF 

  IF m.err = .t. 
   RELEASE oXML, oHttp
   MESSAGEBOX('�� ������� ���������� ����������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN ""
   *EXIT 
  ENDIF 
 
  CreateHeader(ohttp, m.length, m.host, .T., m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+rqHTTP)
 
  TRY 
   ohttp.send(oXml.xml) && ��� get-�������� ���� ���, ��� �� null, ��� post - ����, ������� �������� �������
  CATCH TO oEx
   m.err = .t. 
  ENDTRY 

  IF m.err = .t. 
   RELEASE oXML, oHttp
   MESSAGEBOX('�� ������� ���������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN ""
   * EXIT 
  ENDIF 
 
  m.IsCancelled = .f.
  DO WHILE ohttp.readyState<4
   *WAIT "�������� ������..." WINDOW NOWAIT 

   IF CHRSAW(0) 
    IF INKEY() == 27
     *WAIT CLEAR 
     IF MESSAGEBOX('�� ������ �������� �������� ������?',4+32,'') == 6
      KEYBOARD '{ESC}'
      m.IsCancelled = .t.
      EXIT 
     ENDIF 
     *WAIT "�������� ������..." WINDOW NOWAIT 
    ENDIF 
   ENDIF 

  ENDDO 
 
  IF  m.IsCancelled = .t.
   RELEASE oXML, oHttp
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN ""
   * EXIT 
  ENDIF 

  TRY 
   m.s_tatus = ohttp.status
  CATCH TO oEx
   m.err = .t. 
  ENDTRY 

  IF m.err = .t. 
   RELEASE oXML, oHttp
   MESSAGEBOX('������ ohttp.status!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN ""
   *EXIT 
  ENDIF 

  *IF  ohttp.status<>200
  IF  m.s_tatus<>200
   *MESSAGEBOX('������ ������� ������ '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
   MESSAGEBOX('������ ������� ������ '+STR(m.s_tatus)+CHR(13)+CHR(10)+ALLTRIM(STR(m.s_tatus)),0+64,'')
   RELEASE oXML, oHttp
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN ""
   *EXIT 
  ENDIF 
 
  *��������� http-��������� ������
  poi = fso.CreateTextFile(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ansHTTP)
  poi.Write(ohttp.getAllResponseHeaders())
  poi.Close
  m.cdate = ""

  CFG = FOPEN(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.ansHTTP)
  =ReadHTTPHead()
  =FCLOSE(CFG)
  m.polltagdt = RFC2date(m.cdate) + 3*60*60
  * ��������� http-��������� ������
 
  poi = fso.CreateTextFile(m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\'+ansXML)
  TRY 
   m.respText = ohttp.responseText
  CATCH TO oEx
   m.err = .t. 
  ENDTRY 
  IF m.err = .t. 
   poi.Write('��������� ������ ohttp.responseText! ������ � ���������� msxml3.dll!')
   poi.Close
   RELEASE oXML, oHttp
   *MESSAGEBOX('������ ohttp.responseText!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   RETURN ""
   *EXIT 
  ENDIF 
  *poi.Write(ohttp.responseText)
  poi.Write(m.respText)
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

    *MESSAGEBOX('� ���������� ������ pollTag �� ���������!',0+64,'')

   ENDIF 

  ELSE 
 
   *MESSAGEBOX('� ���������� ������ XML �� ���������!',0+64,'')

  ENDIF 
  RELEASE oXML, oHttp
 
RETURN m.polltag