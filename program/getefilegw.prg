FUNCTION getEFileGw(para1, para2, para3, para4)
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
  MESSAGEBOX('�� ����� BILLID!', 0+64, STR(m.lpu_id,4))
  RETURN .F.
 ENDIF 

 * �������� SOAP-������
 *<wsdl:operation name="getMailGw">
 *	<soap:operation soapAction="" style="document"/>
 *	<wsdl:input name="getMailGw">
 *		<soap:body use="literal"/>
 *	</wsdl:input>
 *	<wsdl:output name="getMailGwResponse">
 *		<soap:body use="literal"/>
 *	</wsdl:output>
 *</wsdl:operation>

 * ��������
 *<xs:element name="getMailGw" type="tns:getMailGw"/>
 *<xs:element name="getMailGwResponse" type="tns:getMailGwResponse"/>

 *����������� ��� getMailGw
 *<xs:complexType name="getMailGw">
 *	<xs:sequence>
 *		<xs:element name="authInfo" type="tns:wsAuthInfo"/>
 *		<xs:element name="getMailGwRequest" type="tns:getMailGwRequest"/>
 *	</xs:sequence>
 *</xs:complexType>

* �������� ������
* m.address = 'http://192.168.192.111:8080/module-pmp/ws/smoIOWs'
* m.host = '192.168.192.111:8080'

* ������������ ������
 m.address = 'http://192.168.192.119:8080/module-pmp/ws/smoIOWs'
 m.host = '192.168.192.119:8080'

 IF EMPTY(m.address)
  RETURN .F.
 ENDIF 
 
 * ��������� ������� ����������
 m.curSoapDir = pSoap+'\'+DTOS(DATE())
 =CheckSOAPDirs(m.curSoapDir)
 * ��������� ������� ����������

 * ���������� ���������� ����� 
 m.un_id    = SYS(3)
 m.httpFile = m.un_id + '.txt'
 m.xmlFile  = m.un_id + '.xml'
 * ���������� ���������� ����� 

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

 oRequest = oXML.createElement("ws:getMailGw")
 oBody.appendChild(oRequest)
 
 oClient = CreateClientPump(oXml, '')
 oRequest.appendChild(oClient)
 
 oXXX = oXML.createElement("getMailGwRequest")
 m.lcperiod = LEFT(m.gcperiod,4)+'-'+SUBSTR(m.gcperiod,5,2)
 oXXX.appendChild(oXML.createElement("period")).text = m.lcperiod
 oRequest.appendChild(oXXX)
 m.mailDirection = 'IN'
 *oXXX.appendChild(oXML.createElement("mailDirection")).text = m.mailDirection
 *oRequest.appendChild(oXXX)
 m.parcelType = 'REPORT_SMO_IN_MO'
 *oXXX.appendChild(oXML.createElement("parcelType")).text = m.parcelType
 *oRequest.appendChild(oXXX)
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
*  ohttp.open('post', m.address, .f.) && .f. - ���������� ����������, .t. - ����������� ����������!
  ohttp.open('post', m.address, .t.) && .f. - ���������� ����������, .t. - ����������� ����������!
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('�� ������� ���������� ����������!'+CHR(13)+CHR(10)+oEx.Message, 0+64, STR(m.lpu_id,4))
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
  ohttp.send(oXml.xml) && ��� get-�������� ���� ���, ��� �� null, ��� post - ����, ������� �������� �������
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('�� ������� ���������!'+CHR(13)+CHR(10)+oEx.Message, 0+64, STR(m.lpu_id,4))
  RETURN .F.
 ENDIF 

 m.IsCancelled = .f.
 WAIT "CheckMailGw (�������� ������)..." WINDOW NOWAIT 
 DO WHILE ohttp.readyState<4

  IF CHRSAW(0) 
   IF INKEY() == 27
    WAIT CLEAR 
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
     m.IsCancelled = .t.
     EXIT 
    ENDIF 
    WAIT "CheckMailGw (�������� ������)..." WINDOW NOWAIT 
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
  MESSAGEBOX('������ ������� ������ '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText), 0+64, STR(m.lpu_id,4))
  RELEASE oXML, oHttp
  RETURN .F.
 ENDIF 
 
 * ��������� http-���������
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * ��������� http-���������
 
 *������� ��� ���������� ����� �������� �� ���� xml
 m.XmlFromFile = ExtractEnvelope(ohttp)

 IF EMPTY(m.XmlFromFile)
  IF !m.lIsSilent
   MESSAGEBOX('� ���������� ������ XML �� ���������!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  ENDIF 
  RETURN IIF(m.lpu_id>0, .T., .F.)
 ENDIF 

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 IF !oxml.loadXML(m.XmlFromFile)
  RELEASE oXML
  IF !m.lIsSilent
   MESSAGEBOX('�� ������� ��������� XML ����!', 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  ENDIF 
  RETURN IIF(m.lpu_id>0, .T., .F.)
 ENDIF 
 oXml.save('&curSoapDir\INPUT\&xmlFile')
 oXml.save('&pBase\&gcPeriod\getMailGw.xml')
 *������� ��� ���������� ����� �������� �� ���� xml

 IF m.httpStatus=500
  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs=0
   RELEASE oXml
   IF !m.lIsSilent
    MESSAGEBOX('� ������ �� ����� ������!',0+64,'')
   ENDIF 
   RETURN IIF(m.lpu_id>0, .T., .F.)
  ELSE 
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text
   
   IF !m.lIsSilent
    MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10), 0+64, STR(m.lpu_id,4))
   ENDIF 
  
   RELEASE oXml
   RETURN IIF(m.lpu_id>0, .T., .F.)
  ENDIF 
 ENDIF 

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
   
  IF m.mDirection!='IN'
   *LOOP 
  ENDIF 
  IF m.parcelType != 'REPORT_SMO_IN_MO'
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

  IF m.mDirection!='IN' OR m.parcelType != 'REPORT_SMO_IN_MO'
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
 IF fso.FileExists(pbase+'\'+gcperiod+'\getEFileGw.dbf')
  fso.DeleteFile(pbase+'\'+gcperiod+'\getEFileGw.dbf')
 ENDIF 
 COPY TO &pbase\&gcperiod\getEFileGw
 SELECT aisoms

 IF !ISNULL(m.loForm)

 ELSE 

  IF !fso.FileExists('&pbase\&gcperiod\aisoms.dbf')
   USE IN answer
   MESSAGEBOX('����������� ���� '+pbase+'\'+gcperiod+'\aisoms.dbf!', 0+64, 'getBillStatuses')
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

 WAIT "CheckMailGw (���������)..." WINDOW NOWAIT 
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
  
  SELECT aisoms

  REPLACE RespLogId WITH answer.mGwLogId

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
  MESSAGEBOX('��������� ���������!',0+64,'')
 ENDIF 

RETURN .T.