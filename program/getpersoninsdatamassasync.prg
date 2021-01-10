PROCEDURE GetPersonInsDataMassAsync && GetPersonInsuranceDataMassAsync
 IF MESSAGEBOX('GetPersonInsuranceDataMassAsync?',4+32,'SOAP') = 7
  RETURN 
 ENDIF 

* m.address = 'http://192.168.192.106:8080/erzl-for-smo/ws/' && �������� �����
* m.host = '192.168.192.106:8080'
 m.address = 'http://192.168.192.118:8080/erzl-for-smo/ws/' && ������������ �����
 m.host = '192.168.192.118:8080'

 IF EMPTY(m.address)
  MESSAGEBOX('������ ����������!',0+16,'m.address?')
  RETURN
 ENDIF 

 * ��������� ������� ����������
 m.curSoapDir = pSoap+'\'+DTOS(DATE())
 =CheckSOAPDirs(m.curSoapDir)
 * ��������� ������� ����������

 * ���������� ���������� ����� 
 m.un_id    = SYS(3)
 m.httpFile = m.un_id + '.txt'
 m.xmlFile  = m.un_id + '.xml'
 m.xxmlFile = m.un_id + 'x.xml'
 * ���������� ���������� ����� 

   m.bpCode = 101 && ����� � �������� ������ � �� ��� ������
 * m.bpCode = 102 && ����� � �������� ������ �� �����������
 * m.bpCode = 103 && ����� � �������� ������ � �������������
 * m.bpCode = 104 && ����� � �������� �������� � ������������� ��
 * m.bpCode = 105 && ����������� ������������ �� ����� ����� ����������
 * m.bpCode = 106 && ������������� ������������
 * m.bpCode = 107 && ����� � �������� �������� � ����������� �������� � �� ����
 * m.bpCode = 108 && �������� ������������ ������������� �� ���������� ��������� (��������)

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

 oRequest = oXML.createElement("ser:getPersonInsuranceDataMassAsyncRequest")
 oBody.appendChild(oRequest)
 
 oClient = CreateClient(oXml, m.bpcode)
 oRequest.appendChild(oClient)
 
 oContent = oXML.createElement("ser:content")
 oInclude = oXML.createElement("xop:Include")
 oInclude.SetAttribute("xmlns:xop", "http://www.w3.org/2004/08/xop/include")
 oInclude.SetAttribute("href", "cid:data.zip")
 oContent.appendChild(oInclude)
* oRequest.appendChild(oXML.createElement("ser:content")).text = "cid:data.zip"
 oRequest.appendChild(oContent)
 
 IF oXML.parseError.errorCode != 0 
  MESSAGEBOX(oXML.parseError.reason,0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
 oXML.save('&curSoapDir\OUTPUT\&xmlFile')
 
 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 oXML.appendChild(oXML.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'"))
 oXML.resolveExternals = .T.
 oRoot = oXML.createElement("PersonInsuranceDataRequestSet")
 oRoot.SetAttribute("xmlns", "http://erzl.org/services")
 oXML.appendChild(oRoot)
 
 opRq = oXML.createElement("pRq")
 opRq.appendChild(oXML.createElement("sn")).text = "770000 5077024107"
 opRq.appendChild(oXML.createElement("dt")).text = "2018-04-30"
 oRoot.appendChild(opRq)

 opRq = oXML.createElement("pRq")
 opRq.appendChild(oXML.createElement("sn")).text = "770000 4046700674"
 opRq.appendChild(oXML.createElement("dt")).text = "2018-04-30"
 oRoot.appendChild(opRq)

 oXML.save('&curSoapDir\OUTPUT\&xxmlFile')
 ZipOpen('&curSoapDir\OUTPUT\data.zip')
 ZipFile('&curSoapDir\OUTPUT\&xxmlFile')
 fso.DeleteFile('&curSoapDir\OUTPUT\&xxmlFile')
 ZipClose()

 poi = FOPEN('&curSoapDir\OUTPUT\data.zip', 0)
 pkSize = FSEEK(poi, 0, 2)
 = FSEEK(poi, 0, 0)
 pkString = FREAD(poi, pkSize)
 FCLOSE(poi)
 
 poi = FOPEN('&curSoapDir\OUTPUT\&xmlFile',12)
 =FSEEK(poi, 0, 2)
 =FPUTS(poi, '----=_Part_01_' + m.un_id)
 =FPUTS(poi, 'Content-Type: application/octet-stream')
 =FPUTS(poi, 'Content-ID: <data.zip>')
 =FPUTS(poi, 'Content-Transfer-Encoding: binary')
 =FPUTS(poi, '')
 =FPUTS(poi, pkString)
* =FPUTS(poi, 'Content-Type: application/xop+xml; charset=utf-8; type="text/xml"')
* =FPUTS(poi, ZipString(oXml.xml))
 =FPUTS(poi, '----=_Part_01_' + m.un_id+'--')
 FCLOSE(poi)

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
  MESSAGEBOX('�� ������� ���������� ����������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

 ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 ohttp.setRequestHeader("Content-Type", 'Multipart/Related; boundary="----=_Part_'+m.un_id+'"; type="application/xop+xml"; start-info="text/xml"')
 ohttp.setRequestHeader("SOAPAction", "")
 ohttp.setRequestHeader("Content-Length", m.length)
 ohttp.setRequestHeader("Host", m.host)
 ohttp.setRequestHeader("Connection", "Keep-Alive")
 ohttp.setRequestHeader("Transfer-Encoding", "chunked")
 ohttp.setRequestHeader("User-Agent", "Visual FoxPro 9.0 (MsXml2.XMLHTTP/MsXml2.DOMDocument)")
 
 poi = fso.CreateTextFile('&curSoapDir\OUTPUT\&httpFile')
 poi.WriteLine('Accept-Encoding: gzip,deflate')
 poi.WriteLine('Content-Type: Multipart/Related; boundary="----=_Part_'+m.un_id+'"; type="application/xop+xml"; start-info="text/xml"')
 poi.WriteLine('SOAPAction: ""')
 poi.WriteLine('Content-Length: '+ALLTRIM(STR(m.length)))
 poi.WriteLine('Host: ' + m.host)
 poi.WriteLine('Connection: Keep-Alive')
 poi.WriteLine('Transfer-Encoding: chunked')
 poi.WriteLine('User-Agent: "Visual FoxPro 9.0 (MsXml2.XMLHTTP/MsXml2.DOMDocument)"')
 poi.Close
 
 TRY 
  ohttp.send(oXml.xml) && ��� get-�������� ���� ���, ��� �� null, ��� post - ����, ������� �������� �������
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('�� ������� ���������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

* MESSAGEBOX('�������� �� ������'+CHR(13)+CHR(10)+m.address+'/erzlsmowebsvc.wsdl'+CHR(13)+CHR(10)+'������ �������!',0+64,'')
 
 m.IsCancelled = .f.
 DO WHILE ohttp.readyState<4
  WAIT "�������� ������..." WINDOW NOWAIT 

  IF CHRSAW(0) 
   IF INKEY() == 27
    WAIT CLEAR 
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
     m.IsCancelled = .t.
     EXIT 
    ENDIF 
    WAIT "�������� ������..." WINDOW NOWAIT 
   ENDIF 
  ENDIF 

 ENDDO 
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

 *IF  ohttp.status<>200
 m.httpStatus = ohttp.status
 IF  !INLIST(m.httpStatus, 200, 500)
  MESSAGEBOX('������ ������� ������ '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
* MESSAGEBOX('������� �����: '+STR(ohttp.status,3),0+64,'')
 
 * ��������� http-���������
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * ��������� http-���������
 
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&xmlFile')
 poi.Write(ohttp.responseText)
 poi.Close

 poi = fso.OpenTextFile('&curSoapDir\INPUT\&xmlFile')
 m.bodypart = poi.ReadLine
 DO WHILE m.bodypart != '<SOAP-ENV:Envelope'
  m.bodypart = poi.ReadLine
 ENDDO 
 poi.close	
  
 IF m.httpStatus=500
  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs=0
   RELEASE oXml
   MESSAGEBOX('� ������ �� ����� ������!',0+64,'soap:Fault')
   RETURN 
  ELSE 
   MESSAGEBOX('���������� '+STR(m.n_recs)+' �������!',0+64,'')
  ENDIF 

  m.n_recs = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').length
  IF m.n_recs=0
   RELEASE oXml
   MESSAGEBOX('� ������ �� ����� ������!',0+64,'')
   RETURN 
  ELSE 
   *MESSAGEBOX('���������� '+STR(m.n_recs)+' �������!',0+64,'')
   m.orec = oxml.selectNodes('soap:Envelope/soap:Body/soap:Fault').item(0)

   m.faultcode = orec.selectNodes('faultcode').item(0).text
   m.faultstring = orec.selectNodes('faultstring').item(0).text

   MESSAGEBOX('faultcode= '+m.faultcode+CHR(13)+CHR(10)+;
			'faultstring= '+m.faultstring+CHR(13)+CHR(10),0+64,'changeBillStatus')
  
   RELEASE oXml
   RETURN 
  ENDIF 
 ENDIF 

 IF m.bodypart='<SOAP-ENV:Envelope'
  oxml.loadXML(m.bodypart)
  
  m.pollTag = ''
  IF oxml.selectNodes('//SOAP-ENV:Envelope/SOAP-ENV:Body/ns2:getPersonInsuranceDataAsyncResponse/ns2:pollTag').length>0
   m.pollTag = oxml.selectNodes('//SOAP-ENV:Envelope/SOAP-ENV:Body/ns2:getPersonInsuranceDataAsyncResponse/ns2:pollTag').item(0).text

*   RELEASE oXML, oHttp
   MESSAGEBOX('pollTag: '+m.pollTag,0+64,'')

  ELSE 

   RELEASE oXML, oHttp
   MESSAGEBOX('� ���������� ������ pollTag �� ���������!',0+64,'')

  ENDIF 

 ELSE 
 
  RELEASE oXML, oHttp
  MESSAGEBOX('� ���������� ������ XML �� ���������!',0+64,'')

 ENDIF 

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
 
* oRoot = oXML.createElement("soapenv:Envelope")
* oRoot.SetAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
* oRoot.SetAttribute("xmlns:ser", "http://erzl.org/services")
* oXML.appendChild(oRoot)
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
  MESSAGEBOX(oXML.parseError.reason,0+64,'')
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
  MESSAGEBOX('�� ������� ���������� ����������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

* MESSAGEBOX('���������� � �������'+CHR(13)+CHR(10)+m.address+'/erzlsmowebsvc.wsdl'+CHR(13)+CHR(10)+'�����������!',0+64,'')

* MESSAGEBOX('������� ���������� ���������...',0+64, '')
 
 ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 ohttp.setRequestHeader("Content-Type", "text/xml; charset=UTF-8")
 ohttp.setRequestHeader("SOAPAction", "")
 ohttp.setRequestHeader("Content-Length", m.length)
 ohttp.setRequestHeader("Host", m.host)
 ohttp.setRequestHeader("Connection", "Keep-Alive")
 ohttp.setRequestHeader("User-Agent", "Visual FoxPro 9.0 (MsXml2.XMLHTTP/MsXml2.DOMDocument)")
 
* MESSAGEBOX('��������� ����������!',0+64,'')
 
 poi = fso.CreateTextFile('&curSoapDir\OUTPUT\&httpFile')
 poi.Close
 
 TRY 
  ohttp.send(oXml.xml) && ��� get-�������� ���� ���, ��� �� null, ��� post - ����, ������� �������� �������
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('�� ������� ���������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

* MESSAGEBOX('�������� �� ������'+CHR(13)+CHR(10)+m.address+'/erzlsmowebsvc.wsdl'+CHR(13)+CHR(10)+'������ �������!',0+64,'')
 
 m.IsCancelled = .f.
 DO WHILE ohttp.readyState<4
  WAIT "�������� ������..." WINDOW NOWAIT 

  IF CHRSAW(0) 
   IF INKEY() == 27
    WAIT CLEAR 
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
     m.IsCancelled = .t.
     EXIT 
    ENDIF 
    WAIT "�������� ������..." WINDOW NOWAIT 
   ENDIF 
  ENDIF 

 ENDDO 
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

 IF  ohttp.status<>200
  MESSAGEBOX('������ ������� ������ '+STR(ohttp.status)+CHR(13)+CHR(10)+ALLTRIM(ohttp.statusText),0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
* MESSAGEBOX('������� �����: '+STR(ohttp.status,3),0+64,'')
 
 * ��������� http-���������
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * ��������� http-���������
 
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&xmlFile')
 poi.Write(ohttp.responseText)
 poi.Close
 
 poi   = FCREATE('&curSoapDir\INPUT\&zipFile')
 nSize = FWRITE(poi, ohttp.responseBody)
 
 IF !FCLOSE(poi)
  MESSAGEBOX('�� ������� ������� ����'+CHR(13)+CHR(10)+m.curSoapDir+'\INPUT\'+zipFile,0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

 IF !UnzipOpen('&curSoapDir\INPUT\&zipFile')
  MESSAGEBOX('���������� ���� �� ZIP-�����!', 0+64, '')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

 ZipDir = curSoapDir + '\INPUT\'
 IF !UnzipGotoFileByName('data.xml')
  UnzipClose()
  MESSAGEBOX('� ���������� ZIP-������ �� ���������� DATA.XML!', 0+64, '')
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

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 m.t_save_ans = SECONDS()

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
* IF !oxml.load('&curSoapDir\INPUT\m.xmlFile')
 WAIT "�������� XML..." WINDOW NOWAIT 
 IF !oxml.load('&curSoapDir\INPUT\data.xml')
  RELEASE oXml
*  MESSAGEBOX('�� ������� ��������� '+m.xmlFile+' ����!',0+64,'oxml.load()')
  MESSAGEBOX('�� ������� ��������� data.xml ����!',0+64,'oxml.load()')
  RETURN 
 ENDIF 
 WAIT CLEAR 
 
 m.n_recs = oxml.selectNodes('PersonInsuranceDataSet/data').length
 IF m.n_recs=0
  RELEASE oXml
  MESSAGEBOX('� ������ �� ����� ������!',0+64,'')
  RETURN 
 ENDIF 
 
 CREATE CURSOR answer (recid c(6), s_pol c(6), n_pol c(16), d_u c(8), q c(2), fam c(25), im c(20), ot c(20), ;
	dr c(8), w n(1), ans_r c(3), tip_d c(1), lpu_id n(6), st_id n(6))
 WAIT "XML->DBF..."  WINDOW NOWAIT 
 FOR m.n_rec = 0 TO m.n_recs-1
  m.orec = oxml.selectNodes('PersonInsuranceDataSet/data').item(m.n_rec)
  
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
  
*  IF orec.selectNodes('person/surname').length>0
   m.fam   = orec.selectNodes('person/surname').item(0).text
   m.ans_r = '211'
*  ENDIF 
*  IF orec.selectNodes('person/namep').length>0
   m.im    = orec.selectNodes('person/namep').item(0).text
*  ENDIF 
  IF orec.selectNodes('person/patronymic').length>0
   m.ot    = orec.selectNodes('person/patronymic').item(0).text
  ENDIF 
*  IF orec.selectNodes('person/sexId').length>0
   m.w     = orec.selectNodes('person/sexId').item(0).text
   m.w     = INT(VAL(m.w))
*  ENDIF 
*  IF orec.selectNodes('person/dateBirth').length>0
   m.dr     = orec.selectNodes('person/dateBirth').item(0).text
   m.dr    = STRTRAN(m.dr,'-','')
*  ENDIF 

*  IF orec.selectNodes('policy/policyTCode').length>0
   m.tip_d     = orec.selectNodes('policy/policyTCode').item(0).text
*  ENDIF 

  IF orec.selectNodes('attach').length>0
   m.lpu_id     = orec.selectNodes('attach').item(0).selectNodes('mo/moCode').item(0).text
   m.lpu_id = INT(VAL(m.lpu_id))
  ENDIF 

  IF orec.selectNodes('attach').length>1
   m.st_id     = orec.selectNodes('attach').item(1).selectNodes('mo/moCode').item(0).text
   m.st_id = INT(VAL(m.st_id))
  ENDIF 

  INSERT INTO answer FROM MEMVAR 

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
     EXIT 
    ELSE 
     WAIT "XML->DBF..."  WINDOW NOWAIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDFOR 
 m.t_ans_proc = SECONDS()
 WAIT CLEAR 
 SELECT answer
 COPY TO &curSoapDir\INPUT\soapans
 BROWSE 
* COPY TO &lcDir/soapans
 USE IN answer 
 
 
* MESSAGEBOX('���������� �������: ' + SecToHrs(m.t_rqst - m.t_beg), 0+64, '')

* MESSAGEBOX('�������� �������: ' + SecToHrs(m.t_send_rq - m.t_rqst), 0+64, '')

* MESSAGEBOX('�������� POLLTAG: ' + SecToHrs(m.t_poll_tag - m.t_send_rq), 0+64, '')

* MESSAGEBOX('�������� ������: ' + SecToHrs(m.t_ans - m.t_poll_tag), 0+64, '')

* MESSAGEBOX('���������� ������: ' + SecToHrs(m.t_save_ans - m.t_ans), 0+64, '')

 SET ESCAPE &OldEscStatus

 MESSAGEBOX('��������� ������: ' + SecToHrs(m.t_ans_proc - m.t_save_ans), 0+64, '')

* MESSAGEBOX('���������� �������: ' + SecToHrs(m.t_rqst - m.t_beg) + CHR(13)+CHR(10)+;
 	'�������� �������: ' + SecToHrs(m.t_send_rq - m.t_rqst) + CHR(13)+CHR(10)+;
 	'�������� POLLTAG: ' + SecToHrs(m.t_poll_tag - m.t_send_rq) + CHR(13)+CHR(10)+;
 	'�������� ������: ' + SecToHrs(m.t_ans - m.t_poll_tag) + CHR(13)+CHR(10)+;
 	'���������� ������: ' + SecToHrs(m.t_save_ans - m.t_ans) + CHR(13)+CHR(10)+;
 	'��������� ������: ' + SecToHrs(m.t_ans_proc - m.t_save_ans) + CHR(13)+CHR(10), 0+64, '')
 	

*ENDIF 

* MESSAGEBOX('��������� ���������'+CHR(13)+CHR(10)+'���� ������: '+m.xmlFile,0+64,'')

RETURN 

