PROCEDURE getPersonPolicy

m.curSoapDir = pSoap+'\'+DTOS(DATE())
IF !fso.FolderExists(m.curSoapDir)
 fso.CreateFolder(m.curSoapDir)
ENDIF 

* bpCodes:
* 101 - ����� � �������� ������ � �� ��� ������
* 102 - ����� � �������� ������ �� �����������
* 103 - ����� � �������� ������ � �������������
* 104 - ����� � �������� �������� � ������������� ��
* 105 - ����������� ������������ �� ����� ����� ����������
* 106 - ������������� ������������
* 107 - ����� � �������� �������� � ����������� �������� � �� ����
* 108 - �������� ������������ ������������� �� ���������� ��������� (��������)

 PUBLIC oXML  AS MsXml2.DOMDocument
 PUBLIC oNode as MSXML2.IXMLDOMNode
 PUBLIC oRoot as MSXML2.IXMLDOMElement
 PUBLIC oElem as MSXML2.IXMLDOMElement
 PUBLIC oBody as MSXML2.IXMLDOMNode
 PUBLIC oRequest as MSXML2.IXMLDOMNode
 PUBLIC oClient as MSXML2.IXMLDOMNode
 
 PUBLIC oHttp AS MsXml2.XMLHTTP
 
 oHttp = CREATEOBJECT("MsXml2.XMLHTTP")
 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 oXML.resolveExternals = .T.

 * Create a procesing instruction.  
 oNode = oXML.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
 oXML.appendChild(oNode)
 
 oRoot = oXML.createElement("soapenv:Envelope")
 oRoot.SetAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
 oRoot.SetAttribute("xmlns:ser", "http://erzl.org/services")
 oXML.appendChild(oRoot)
 
 oNode = oXML.createElement("soapenv:Header")
 oRoot.appendChild(oNode)

 oBody = oXML.createElement("soapenv:Body")
 oRoot.appendChild(oBody)

 oRequest = oXML.createElement("ser:findPersonByPolicyRequest")
 oBody.appendChild(oRequest)
 
 m.bpCode = '101'
 
 oClient = oXML.createElement("ser:client")
 oClient.appendChild(oXML.createElement("ser:orgCode")).text  = m.orgCode
 oClient.appendChild(oXML.createElement("ser:bpCode")).text   = m.bpCode
 oClient.appendChild(oXML.createElement("ser:system")).text   = m.soapSystem
 oClient.appendChild(oXML.createElement("ser:user")).text     = m.erzlUser
 oClient.appendChild(oXML.createElement("ser:password")).text = m.erzlPass
 oClient.appendChild(oXML.createElement("ser:comment")).text  = "����� ��� �� ���� ��� �����������" && ��� �����
 oRequest.appendChild(oClient)
 
 oRequest.appendChild(oXML.createElement("ser:ukl")).text="4501591777"
 oRequest.appendChild(oXML.createElement("ser:pageSize")).text="20" && Optional
 oRequest.appendChild(oXML.createElement("ser:offset")).text="0"    && Optional
 oRequest.appendChild(oXML.createElement("ser:date")).text=""       && Optional
 oRequest.appendChild(oXML.createElement("ser:dateTo")).text=""     && Optional
 
  IF oXML.parseError.errorCode != 0 
   MESSAGEBOX(oXML.parseError.reason,0+64,'')
  ENDIF 
 
 oXML.save('&curSoapDir\getPersonPolicyRequest.xml')
 length = fso.GetFile("&curSoapDir\getPersonPolicyRequest.xml").Size
 
* ������� ������:
* ����: 192.168.192.106:8080
* ����: 192.168.192.111:8080
 
 LOCAL oEx as Exception

 m.err = .f. 
 TRY 
*  ohttp.open('post', 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl', .f.)
  ohttp.open('post', 'http://192.168.192.106:8080/erzlws/policyService', .f.)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('�� ������� ���������� ����������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

* MESSAGEBOX('���������� � �������'+CHR(13)+CHR(10)+;
* 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl'+CHR(13)+CHR(10)+;
* 	'�����������!',0+64,'')

 MESSAGEBOX('���������� � �������'+CHR(13)+CHR(10)+;
 'http://192.168.192.106:8080/erzlws/policyService'+CHR(13)+CHR(10)+;
 	'�����������!',0+64,'')

 MESSAGEBOX('������� ���������� ���������...',0+64, '')
 
 ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 ohttp.setRequestHeader("Content-Type", "text/xml; charset=UTF-8")
 ohttp.setRequestHeader("SOAPAction", "")
 ohttp.setRequestHeader("Content-Length", m.length)
 ohttp.setRequestHeader("Host", "192.168.192.106:8080")
 ohttp.setRequestHeader("Connection", "Keep-Alive")
 ohttp.setRequestHeader("User-Agent", "Apache-HttpClient/4.1.1 (java 1.5)")
 
 MESSAGEBOX('��������� ����������!',0+64,'')
 
 TRY 
  ohttp.send(oXml.xml)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('�� ������� ���������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

* MESSAGEBOX('�������� �� ������'+CHR(13)+CHR(10)+;
* 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl'+CHR(13)+CHR(10)+;
* 	'������ �������!',0+64,'')

 MESSAGEBOX('�������� �� ������'+CHR(13)+CHR(10)+;
 'http://192.168.192.106:8080/erzlws/policyService'+CHR(13)+CHR(10)+;
 	'������ �������!',0+64,'')
 
 m.IsCancelled = .f.
 DO WHILE ohttp.readyState<4
  WAIT "�������� ������..." WINDOW NOWAIT 

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
     m.IsCancelled = .t.
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDDO 
 
 IF  m.IsCancelled = .t.
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 

 IF  ohttp.status<>200
  MESSAGEBOX('������ ������� ������ '+STR(ohttp.status,3),0+64,'')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
 MESSAGEBOX('������� �����: '+STR(ohttp.status,3),0+64,'')
 
 oXml.loadXML(ohttp.responseText)
 IF oXml.childNodes.length<=0
  MESSAGEBOX('������� ������ �����!', 0+64, '')
  RELEASE oXML, oHttp
  RETURN 
 ENDIF 
 
 oXML.save('&curSoapDir\getPersonPolicyAnswer.xml')
 
 RELEASE oXML, oHttp
 MESSAGEBOX('OK!',0+64,'')

RETURN 