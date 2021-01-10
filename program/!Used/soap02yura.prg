PROCEDURE soap02

* bpCodes:
* 101 - ����� � �������� ������ � �� ��� ������
* 102 - ����� � �������� ������ �� �����������
* 103 - ����� � �������� ������ � �������������
* 104 - ����� � �������� �������� � ������������� ��
* 105 - ����������� ������������ �� ����� ����� ����������
* 106 - ������������� ������������
* 107 - ����� � �������� �������� � ����������� �������� � �� ����
* 108 - �������� ������������ ������������� �� ���������� ��������� (��������)

*Dim Doc As MSXML2.DOMDocument
*Dim Node As MSXML2.IXMLDOMNode
*Dim Root As MSXML2.IXMLDOMElement
*Dim Elem As MSXML2.IXMLDOMElement

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
 
 * ����� ��������: <ds:Transform Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"></ds:Transform>  
*  oTransform = oXML.CreateElement("ds:Transform")  
*  oTransform.SetAttribute("Algorithm", "http://www.w3.org/2001/10/xml-exc-c14n#")  
*  oXML.AppendChild(oTransform)   
*  xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ser="http://erzl.org/services
 oRoot = oXML.createElement("soapenv:Envelope")
 oRoot.SetAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
 oRoot.SetAttribute("xmlns:ser", "http://erzl.org/services")
 oXML.appendChild(oRoot)
 oNode = oXML.createElement("soapenv:Header")
 oRoot.appendChild(oNode)

 oBody = oXML.createElement("soapenv:Body")
 oRoot.appendChild(oBody)

* oRequest = oXML.createElement("getPersonPolicyRequest")
 oRequest = oXML.createElement("ser:findPersonByPolicyRequest")
* oRequest.setAttribute("xmlns", "http://erzl.org/services")
 oBody.appendChild(oRequest)
 
 m.bpCode = '101'
 DO CASE 
  CASE m.qcod='S7'
   m.orgId    = '1234' && ���, ���� ����� ���� ������!
   m.orgCode  =  '9876' && && ���, ���� ����� ���� ������!
   m.system   = 'lpu2smo'
   m.user     = 'yura_smagin_erzl_in' && ���, ���� ����� ���� ������!
   m.password = '12w#er' && ���, ���� ����� ���� ������!
  OTHERWISE 
   m.orgId    = '0000'
   m.orgCode  = '9999'
   m.system   = 'lpu2smo'
   m.user     = 'user'
   m.password = 'password'
 ENDCASE 
 
 oClient = oXML.createElement("ser:client")
* oClient = oXML.createElement("authInfo")
 oClient.appendChild(oXML.createElement("ser:orgCode")).text=m.orgCode
 oClient.appendChild(oXML.createElement("ser:bpCode")).text=m.bpCode
* oClient.appendChild(oXML.createElement("orgId")).text = m.orgId
 oClient.appendChild(oXML.createElement("ser:system")).text = m.system
 oClient.appendChild(oXML.createElement("ser:user")).text = m.user
 oClient.appendChild(oXML.createElement("ser:password")).text = m.password
 oClient.appendChild(oXML.createElement("ser:comment")).text = "����� ��� �� ���� ��� �����������" && ��� �����
 oRequest.appendChild(oClient)
 
 oRequest.appendChild(oXML.createElement("ser:policySerNum")).text="7758720874002365"
 oRequest.appendChild(oXML.createElement("ser:pageSize")).text="20" && Optional
 oRequest.appendChild(oXML.createElement("ser:offset")).text="0"    && Optional
 oRequest.appendChild(oXML.createElement("ser:date")).text=""       && Optional
 oRequest.appendChild(oXML.createElement("ser:dateTo")).text=""     && Optional
 
* oEnvelope = oXML.getElementsByTagName("soapenv:Envelope").item.
* oHeader = oEnvelope.createElement("soapenv:Header")
* oEnvelope.appendChild(oHeader)

* IF 3=2
* m.MethodName = 'getParcelMoPage'
 
* m.Envelope = '<?xml version="1.0" ?>'
* m.Envelope = m.Envelope + '<soapenv:Envelope ' && ������ ������ �����
* m.Envelope = m.Envelope + 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
* m.Envelope = m.Envelope + 'xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
* m.Envelope = m.Envelope + 'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'

* m.Envelope = m.Envelope + '<soapenv:Header/>' && ������� ��������� ���������
* m.Envelope = m.Envelope + '<soapenv:Body>' && ������ �������� �������������� ����� ���������
* m.Envelope = m.Envelope + '<'+m.MethodName+'>' && ������ ����������� ������, ��� ��� - ������������ ������
* m.Envelope = m.Envelope + '<'+m.MethodName+'Request>'

 *m.Envelope = m.Envelope + '<>' && ���� ���������
 *m.Envelope = m.Envelope + '<>' && ���� ���������
 *m.Envelope = m.Envelope + '<>' && ���� ���������
 *m.Envelope = m.Envelope + '<>' && ���� ���������

* m.Envelope = m.Envelope + '</'+m.MethodName+'Request>'
* m.Envelope = m.Envelope + '<'+m.MethodName+'>'
* m.Envelope = m.Envelope + '</soapenv:Envelope>' && ��������� ������ �����
 
 *oXML.appendChild(m.Envelope)
* oXML.async = .f.
* oXML.appendChild(m.Envelope)
* IF !oXML.loadXML(m.Envelope)
  IF oXML.parseError.errorCode != 0 
   MESSAGEBOX(oXML.parseError.reason,0+64,'')
  ENDIF 
 *  MESSAGEBOX('�� ������� ��������� loadXML!',0+64,'')
 *ENDIF 
* ENDIF 
 
 oXML.save('&pBase\myEnvelope.xml')
 
* ������� ������:
* ����: 192.168.192.106:8080
* ����: 192.168.192.111:8080
 
 LOCAL oEx as Exception

 m.err = .f. 
 TRY 
  ohttp.open('post', 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl', .f.)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('�� ������� ���������� ����������!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

 MESSAGEBOX('���������� � �������'+CHR(13)+CHR(10)+;
 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl'+CHR(13)+CHR(10)+;
 	'�����������!',0+64,'')

 MESSAGEBOX('������� ���������� ���������...',0+64,'���1')
 ohttp.setRequestHeader("Content-Type", "application/soap+xml; charset=utf-8")
 MESSAGEBOX('������� ���������� ���������...',0+64,'���2')
 
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
 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl'+CHR(13)+CHR(10)+;
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
 
 MESSAGEBOX('������� �����: '+STR(ohttp.status,3),0+64,'')
 
 oXml.loadXML(ohttp.responseText)
 oXML.save('&pBase\myAnswer.xml')
 
 RELEASE oXML, oHttp
 MESSAGEBOX('OK!',0+64,'')

*VBSTART

*Function StripToNumeric(sNumber)
*    Dim oHttp
*    Dim oXML
*    Dim tEnvelope
*    Dim tResult

*    ' Preparation of SOAP header.
*    tEnvelope = "<?xml version=""1.0"" ?>"
*    tEnvelope = tEnvelope & "<soap:Envelope "
*    tEnvelope = tEnvelope & "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" "
*    tEnvelope = tEnvelope & "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" "
*    tEnvelope = tEnvelope & "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
*    tEnvelope = tEnvelope & "<soap:Body>"
*    ' Here I define the operation to call
*    tEnvelope = tEnvelope & "<StripToNumeric xmlns=""http://webservices.DataAccess.Net/ElevenTest"">"
*    tEnvelope = tEnvelope & "<sNumber>" & sNumber & "</sNumber>"
*    tEnvelope = tEnvelope & "</StripToNumeric>"
*    tEnvelope = tEnvelope & "</soap:Body>"
*    tEnvelope = tEnvelope & "</soap:Envelope>"

*    ' Create Object of MtEnvelope2.XMLHTTP
*    Set oHttp = CreateObject("MsXml2.XMLHTTP")
*    Set oXML = CreateObject("MsXml2.DOMDocument")

*    ' Load the header as XML
*    oXML.loadXML tEnvelope
*    ' Open the web service location
*    oHttp.open "POST","http://webservices.daehosting.com/services/eleventest.wso", False
*    ' Add the SOAPAction header
*    oHttp.setRequestHeader "SOAPAction", "StripToNumeric"
*    ' We are working with XML
*    oHttp.setRequestHeader "Content-Type", "text/xml"
*    ' Send the SOAP Message
*    oHttp.send oXML.xml

*    ' responseText property contains the full answer received from the server
*    ' wscript.echo(oHttp.responseText)

*    ' Treat the response as XML
*    oXML.LoadXml oHttp.responseText
*    ' What I want is the returning value, contained in the tag value. I get it and use it
*    Set objNodeList = oXML.selectNodes("//soap:Envelope/soap:Body/m:StripToNumericResponse/m:StripToNumericResult")
*    If objNodeList.length > 0 Then
*        ' If there is at least one node named "valid", get the answer.
*        ' Otherwise there are errors, and the web service should have sent a fault message.
*      tResult = oXML.selectSingleNode("//soap:Envelope/soap:Body/m:StripToNumericResponse/m:StripToNumericResult").text
*    End If

*    ' Clean up objects
*    Set oXML = Nothing
*    Set oHttp = Nothing

*    ' Return the result
*    StripToNumeric = tResult
*End Function

*VBEND

*VBEval>StripToNumeric("1a2b3c4d5e6"),response
*MessageModal>response
RETURN 