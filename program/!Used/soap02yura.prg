PROCEDURE soap02

* bpCodes:
* 101 - Поиск и просмотр записи о ЗЛ ОМС Москвы
* 102 - Поиск и просмотр записи об иногородних
* 103 - Поиск и просмотр записи о новорожденных
* 104 - Поиск и просмотр сведений о прикреплениях ЗЛ
* 105 - Утверждение прикрепления по смене места жительства
* 106 - Аннулирование прикрепления
* 107 - Поиск и просмотр сведений о результатах запросов к ЦС ЕРЗЛ
* 108 - Просмотр рекомендаций пользователям по дальнейшим действиям (директив)

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
 
 * Хотим получить: <ds:Transform Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"></ds:Transform>  
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
   m.orgId    = '1234' && Юра, сюда забей свои данные!
   m.orgCode  =  '9876' && && Юра, сюда забей свои данные!
   m.system   = 'lpu2smo'
   m.user     = 'yura_smagin_erzl_in' && Юра, сюда забей свои данные!
   m.password = '12w#er' && Юра, сюда забей свои данные!
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
 oClient.appendChild(oXML.createElement("ser:comment")).text = "Здесь мог бы быть ваш комментарий" && Это опция
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
* m.Envelope = m.Envelope + '<soapenv:Envelope ' && Первая строка файла
* m.Envelope = m.Envelope + 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
* m.Envelope = m.Envelope + 'xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
* m.Envelope = m.Envelope + 'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'

* m.Envelope = m.Envelope + '<soapenv:Header/>' && Сегмент заголовка сообщения
* m.Envelope = m.Envelope + '<soapenv:Body>' && Начало сегмента содержательной части сообщения
* m.Envelope = m.Envelope + '<'+m.MethodName+'>' && Начало вызываемого метода, где ХХХ - наименование метода
* m.Envelope = m.Envelope + '<'+m.MethodName+'Request>'

 *m.Envelope = m.Envelope + '<>' && Поля сообщения
 *m.Envelope = m.Envelope + '<>' && Поля сообщения
 *m.Envelope = m.Envelope + '<>' && Поля сообщения
 *m.Envelope = m.Envelope + '<>' && Поля сообщения

* m.Envelope = m.Envelope + '</'+m.MethodName+'Request>'
* m.Envelope = m.Envelope + '<'+m.MethodName+'>'
* m.Envelope = m.Envelope + '</soapenv:Envelope>' && Последняя строка файла
 
 *oXML.appendChild(m.Envelope)
* oXML.async = .f.
* oXML.appendChild(m.Envelope)
* IF !oXML.loadXML(m.Envelope)
  IF oXML.parseError.errorCode != 0 
   MESSAGEBOX(oXML.parseError.reason,0+64,'')
  ENDIF 
 *  MESSAGEBOX('НЕ УДАЛОСЬ ВЫПОЛНИТЬ loadXML!',0+64,'')
 *ENDIF 
* ENDIF 
 
 oXML.save('&pBase\myEnvelope.xml')
 
* Игровые адреса:
* ЕРЗЛ: 192.168.192.106:8080
* ПУМП: 192.168.192.111:8080
 
 LOCAL oEx as Exception

 m.err = .f. 
 TRY 
  ohttp.open('post', 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl', .f.)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 
 
 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('НЕ УДАЛОСЬ УСТАНОВИТЬ СОЕДИНЕНИЕ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

 MESSAGEBOX('СОЕДИНЕНИЕ С МОДУЛЕМ'+CHR(13)+CHR(10)+;
 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl'+CHR(13)+CHR(10)+;
 	'УСТАНОВЛЕНО!',0+64,'')

 MESSAGEBOX('ПОПЫТКА УСТАНОВИТЬ ЗАГОЛОВОК...',0+64,'ШАГ1')
 ohttp.setRequestHeader("Content-Type", "application/soap+xml; charset=utf-8")
 MESSAGEBOX('ПОПЫТКА УСТАНОВИТЬ ЗАГОЛОВОК...',0+64,'ШАГ2')
 
 TRY 
  ohttp.send(oXml.xml)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  RELEASE oXML, oHttp
  MESSAGEBOX('НЕ УДАЛОСЬ ОТПРАВИТЬ!'+CHR(13)+CHR(10)+oEx.Message,0+64,'')
  RETURN 
 ENDIF 

 MESSAGEBOX('ОТПРАВКА НА МОДУЛЬ'+CHR(13)+CHR(10)+;
 'http://192.168.192.106:8080/erzlws/policyService/policies.wsdl'+CHR(13)+CHR(10)+;
 	'ПРОШЛА УСПЕШНО!',0+64,'')
 
 m.IsCancelled = .f.
 DO WHILE ohttp.readyState<4
  WAIT "ОЖИДАНИЕ ОТВЕТА..." WINDOW NOWAIT 

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
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
 
 MESSAGEBOX('ПОЛУЧЕН ОТВЕТ: '+STR(ohttp.status,3),0+64,'')
 
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