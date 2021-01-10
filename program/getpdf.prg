PROCEDURE getPdf
* �������� ������
* m.address = 'http://192.168.192.111:8080/module-pmp/ws/smoIOWs'
* m.host = '192.168.192.111:8080'

* ������������ ������
 m.address = 'http://192.168.192.119:8080/module-pmp/ws/smoIOWs'
 m.host = '192.168.192.119:8080'

 IF EMPTY(m.address)
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

 oRequest = oXML.createElement("ws:getPdf")
 oBody.appendChild(oRequest)
 
 oClient = CreateClientPump(oXml, '')
 oRequest.appendChild(oClient)
 
* oXXX = oXML.createElement("request")
 oXXX = oXML.createElement("request")
 oXXX.appendChild(oXML.createElement("mailGWlogid")).text = '1586256'
 oRequest.appendChild(oXXX)
 
 IF oXML.parseError.errorCode != 0 
  MESSAGEBOX(oXML.parseError.reason,0+64,'')
  RELEASE oXML, oHttp
  RETURN 
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

 ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 ohttp.setRequestHeader("Content-Type", "application/xml")
 ohttp.setRequestHeader("Content-Length", m.length)
 ohttp.setRequestHeader("Host", m.host)
 ohttp.setRequestHeader("User-Agent", "Mike Ruby Software")
 
 poi = fso.CreateTextFile('&curSoapDir\OUTPUT\&httpFile')
 poi.WriteLine('Accept-Encoding: gzip,deflate')
 poi.WriteLine('Content-Type: "application/xml"')
 poi.WriteLine('Content-Length: '+ALLTRIM(STR(m.length)))
 poi.WriteLine('Host: ' + m.host)
 poi.WriteLine('User-Agent: "Mike Ruby Software"')
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
 
 * ��������� http-���������
 poi = fso.CreateTextFile('&curSoapDir\INPUT\&httpFile')
 poi.Write(ohttp.getAllResponseHeaders())
 poi.Close
 * ��������� http-���������
 
 poi   = FCREATE('&curSoapDir\INPUT\&xmlFile')
 nSize = FWRITE(poi, ohttp.responseBody)
 =FCLOSE(poi)

 poi = fso.OpenTextFile('&curSoapDir\INPUT\&xmlFile')
 m.bodypart = poi.ReadLine
 DO WHILE m.bodypart != 'Content-Disposition'
  m.bodypart = poi.ReadLine
 ENDDO 
 poi.close	
 
 IF m.bodypart != 'Content-Disposition'
  MESSAGEBOX('� ���������� ����� ����������� Content-Disposition!',0+64,'')
  RETURN 
 ENDIF 
 
 IF OCCURS('"',m.bodypart)!=2
  MESSAGEBOX('� ���������� ������������ Content-Disposition!',0+64,'2')
  RETURN 
 ENDIF 
 
 m.fname = STRTRAN(ALLTRIM(SUBSTR(m.bodypart, AT('"',m.bodypart))), '"', '')
 IF LEN(m.fname)<=0
  MESSAGEBOX('� ���������� ������������ name: '+m.fname, 0+64, 'name')
  RETURN 
 ENDIF 

* poi = FOPEN('&curSoapDir\INPUT\&xmlFile')
 m.pdf = FILETOSTR('&curSoapDir\INPUT\&xmlFile')
 IF AT('%PDF', m.pdf)=0
  RELEASE m.pdf
  MESSAGEBOX('� ���������� ������ PDF �� ���������!',0+64,'%PDF')
  RETURN 
 ENDIF 
 IF AT('%%EOF', m.pdf)=0
  RELEASE m.pdf
  MESSAGEBOX('� ���������� ������ ��������� ������������ PDF!',0+64,'%%EOF')
  RETURN 
 ENDIF 
 m.nBytes = STRTOFILE(SUBSTR(m.pdf, AT('%PDF',m.pdf), ;
 	AT('%%EOF',m.pdf)-AT('%PDF',m.pdf)+5), ;
 	m.curSoapDir+'\INPUT\'+m.fname)

 MESSAGEBOX('��������� ���������!',0+64,'')
RETURN 