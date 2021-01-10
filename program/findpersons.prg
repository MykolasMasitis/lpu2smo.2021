PROCEDURE findPersons && soap03

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
 lcTarget = "xml"  && oNewPI.Target  
 lcPItext = "version='1.0' encoding='UTF-8'"  && oNewPI.Data  
 oNode = oXML.createProcessingInstruction(lcTarget, lcPItext)  
    
* Add the processing instruction node to the document.  
 oXML.appendChild(oNode)
 
 oRoot = oXML.createElement("soapenv:Envelope")
 oRoot.SetAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
 oRoot.SetAttribute("xmlns:ser", "http://erzl.org/services")
 oXML.appendChild(oRoot)
 oNode = oXML.createElement("soapenv:Header")
 oRoot.appendChild(oNode)

 oBody = oXML.createElement("soapenv:Body")
 oRoot.appendChild(oBody)

 oRequest = oXML.createElement("ser:findPersonsRequest")
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
 
 oSearchQuery = oXML.createElement("ser:searchQuery")
 oSearchQuery.appendChild(oXML.createElement("op")).text    = 'AND'
 oSearchQuery.appendChild(oXML.createElement("code")).text  = 'surname'
 oSearchQuery.appendChild(oXML.createElement("cmp")).text   = 'EQ'
 oSearchQuery.appendChild(oXML.createElement("value")).text = '�����'
 oRequest.appendChild(oSearchQuery)
 
 oSearchQuery = oXML.createElement("ser:searchQuery")
 oSearchQuery.appendChild(oXML.createElement("op")).text    = 'AND'
 oSearchQuery.appendChild(oXML.createElement("code")).text  = 'namep'
 oSearchQuery.appendChild(oXML.createElement("cmp")).text   = 'EQ'
 oSearchQuery.appendChild(oXML.createElement("value")).text = '������'
 oRequest.appendChild(oSearchQuery)

 oSearchQuery = oXML.createElement("ser:searchQuery")
 oSearchQuery.appendChild(oXML.createElement("op")).text    = 'AND'
 oSearchQuery.appendChild(oXML.createElement("code")).text  = 'patronymic'
 oSearchQuery.appendChild(oXML.createElement("cmp")).text   = 'EQ'
 oSearchQuery.appendChild(oXML.createElement("value")).text = '�������������'
 oRequest.appendChild(oSearchQuery)

 oSearchQuery = oXML.createElement("ser:searchQuery")
 oSearchQuery.appendChild(oXML.createElement("op")).text    = 'AND'
 oSearchQuery.appendChild(oXML.createElement("code")).text  = 'sexId'
 oSearchQuery.appendChild(oXML.createElement("cmp")).text   = 'EQ'
 oSearchQuery.appendChild(oXML.createElement("value")).text = '1'
 oRequest.appendChild(oSearchQuery)

 oSearchQuery = oXML.createElement("ser:searchQuery")
 oSearchQuery.appendChild(oXML.createElement("op")).text    = 'AND'
 oSearchQuery.appendChild(oXML.createElement("code")).text  = 'dateBirth'
 oSearchQuery.appendChild(oXML.createElement("cmp")).text   = 'EQ'
 oSearchQuery.appendChild(oXML.createElement("value")).text = '1974-06-20'
 oRequest.appendChild(oSearchQuery)

 oSort = oXML.createElement("sort")
* oSearchQuery.appendChild(oXML.createElement("code")).text  = 'surname'
* oSearchQuery.appendChild(oXML.createElement("order")).text   = 'A'
 oRequest.appendChild(oSort)


 oRequest.appendChild(oXML.createElement("ser:pageSize")).text="20" && Optional
 oRequest.appendChild(oXML.createElement("ser:offset")).text="0"    && Optional
 oRequest.appendChild(oXML.createElement("ser:date")).text=""       && Optional
 oRequest.appendChild(oXML.createElement("ser:dateTo")).text=""     && Optional
 
 IF oXML.parseError.errorCode != 0 
  MESSAGEBOX(oXML.parseError.reason,0+64,'')
 ENDIF 
 
 oXML.save('&curSoapDir\findPersonsRequest.xml')
 length = fso.GetFile("&curSoapDir\findPersonsRequest.xml").Size
 
* ������� ������:
* ����: 192.168.192.106:8080
* ����: 192.168.192.111:8080
 
 LOCAL oEx as Exception

 m.err = .f. 
 TRY 
  ohttp.open('post', 'http://192.168.192.106:8080/erzlws/policyService', .f.)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('�� ������� ���������� ����������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

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
 
 oXML.save('&curSoapDir\findPersonsAnswer.xml')
 
 RELEASE oXML, oHttp
 MESSAGEBOX('OK!',0+64,'')

RETURN 