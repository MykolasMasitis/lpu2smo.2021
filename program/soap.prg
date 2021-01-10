FUNCTION ReadTheHead(para1, para2) && para1 - указатель на открытый файл, para2 - что ищем
 IF VARTYPE(para1)!='O'
  RETURN ''
 ENDIF 
 m.poi        = para1
 m.WhatToSeek = para2
 IF poi.AtEndOfStream
  RETURN ''
 ENDIF 
 
 DO WHILE !poi.AtEndOfStream
  m.lcString = poi.ReadLine
  IF m.lcString = m.WhatToSeek
   EXIT 
  ENDIF 
 ENDDO 
 
RETURN IIF(m.lcString != m.WhatToSeek, '', ALLTRIM(STRTRAN(m.lcString, m.WhatToSeek+':', '')))

FUNCTION Element(para0, para1)
 m.lcParam  = para0 && boundary, start, start-info etc
 m.lcString = para1 && строка
 IF OCCURS(m.lcParam, m.lcString)=0
  RETURN ''
 ENDIF 
 
 m.tmp = SUBSTR(m.lcString, AT(m.lcParam, m.lcString)+LEN(m.lcParam)+1)
 m.tmp = STRTRAN(SUBSTR(m.tmp, 1, AT(';', m.tmp)-1), '"', '')
 
RETURN m.tmp

FUNCTION ExtractEnvelope(ohttp)
 m.FileInMem     = ohttp.responseText
 *m.FileInMem     = ohttp
 m.BegOfEnvelope = AT('<soap:Envelope', m.FileInMem)
 m.EndOfEnvelope = AT('</soap:Envelope', m.FileInMem)
 IF m.BegOfEnvelope<=0 OR m.EndOfEnvelope<=0 OR m.EndOfEnvelope < m.BegOfEnvelope
  RETURN ''
 ENDIF 

 m.t_temp = SUBSTR(m.FileInMem, m.BegOfEnvelope, ;
 	m.EndOfEnvelope - m.BegOfEnvelope + LEN('</soap:Envelope') + 1) 
RETURN m.t_temp

FUNCTION CreateEnvelope(oXml)
 oRoot = oXML.createElement("soapenv:Envelope")
 oRoot.SetAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
 oRoot.SetAttribute("xmlns:ser", "http://erzl.org/services")
RETURN oRoot

FUNCTION CreateEnvelopePump(oXml)
 oRoot = oXML.createElement("soapenv:Envelope")
 oRoot.SetAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
 oRoot.SetAttribute("xmlns:ws", "http://ws.smo.pmp.ibs.ru/")
RETURN oRoot

FUNCTION CreateEnvelopeNSI(oXml)
 oRoot = oXML.createElement("soapenv:Envelope")
 oRoot.SetAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
 oRoot.SetAttribute("xmlns:ws", "http://ws.pmp.ibs.ru/")
RETURN oRoot

FUNCTION CreateClient(oXml, para1)
 PRIVATE m.bpCode
 m.bpCode  = para1
 oClient = oXML.createElement("ser:client")
 oClient.appendChild(oXML.createElement("ser:orgCode")).text  = m.orgCode && глобальная переменная
 oClient.appendChild(oXML.createElement("ser:bpCode")).text   = m.bpCode
 oClient.appendChild(oXML.createElement("ser:system")).text   = m.soapSystem && глобальная переменная
 oClient.appendChild(oXML.createElement("ser:user")).text     = m.erzlUser && глобальная переменная
 oClient.appendChild(oXML.createElement("ser:password")).text = m.erzlPass && глобальная переменная
* oClient.appendChild(oXML.createElement("ser:comment")).text  = "This soap-message was formed and sent by Lpu2SMO software. The author is Mike Ruby, 9950825@mail.ru" 
 IF !m.IsTestMode
  oClient.appendChild(oXML.createElement("ser:comment")).text  = "@ Mike Ruby Software 2018 (Lpu2SMO); 9950825@mail.ru, +79637820825" 
 ENDIF 
RETURN oClient

FUNCTION CreateClientPump(oXml, para1)
 PRIVATE m.bpCode, m.requestId
 m.requestId = para1
* m.bpCode  = para1
 oClient = oXML.createElement("authInfo")
* oClient.appendChild(oXML.createElement("orgCode")).text  = m.orgCode && глобальная переменная
 oClient.appendChild(oXML.createElement("orgId")).text  = STR(m.qobjid,4) && глобальная переменная
* oClient.appendChild(oXML.createElement("orgId")).text  = '5400' && глобальная переменная
* oClient.appendChild(oXML.createElement("bpCode")).text   = m.bpCode
 oClient.appendChild(oXML.createElement("system")).text   = m.soapSystem && глобальная переменная
 oClient.appendChild(oXML.createElement("user")).text     = m.pumpUser && глобальная переменная
 oClient.appendChild(oXML.createElement("password")).text = m.pumpPass && глобальная переменная
 oClient.appendChild(oXML.createElement("requestId")).text = m.requestId 
* oClient.appendChild(oXML.createElement("comment")).text  = "@ Mike Ruby Software 2018 (Lpu2SMO); 9950825@mail.ru, +79637820825" 
RETURN oClient

FUNCTION WS2Address(para1)
 m.addr = ALLTRIM(para1)
 IF EMPTY(m.addr)
  RETURN ""
 ENDIF 
 IF !fso.FileExists(pBin+'\soap.cfg')
  MESSAGEBOX('ОТСУТСТВУЕТ КОНФИГУРАЦИОННЫЙ ФАЙЛ SOAP.CFG!',0+16,'')
  RETURN ""
 ENDIF 
 IF OpenFile(pBin+'\soap.cfg', 'soap', 'shar')>0
  IF USED('soap')
   USE IN soap
  ENDIF 
  RETURN ""
 ENDIF 
 SELECT soap
 SELECT * FROM soap WHERE ws = m.addr INTO CURSOR curSoap
 IF RECCOUNT('curSoap')=0 
  USE IN soap
  USE IN curSoap
  MESSAGEBOX('В КОНФИГУРАЦИОННОМ ФАЙЛЕ SOAP.CFG!'+CHR(13)+CHR(10)+;
  	'ОТСУТСТВУЮТ НАСТРОЙКИ ДЛЯ СЕРВИСА erzlsmowebsvc.wsdl!',0+16,'')
  RETURN ""
 ENDIF 
 IF RECCOUNT('curSoap')>1
  USE IN soap
  USE IN curSoap
  MESSAGEBOX('В КОНФИГУРАЦИОННОМ ФАЙЛЕ SOAP.CFG!'+CHR(13)+CHR(10)+;
  	'БОЛЕЕ ОДНОЙ НАСТРОЙКИ ДЛЯ СЕРВИСА erzlsmowebsvc.wsdl!',0+16,'')
  RETURN ""
 ENDIF 
 USE IN soap
 SELECT curSoap
 m.address = ALLTRIM(address)
 * m.address = 'http://192.168.192.106:9090/ws' && http://192.168.192.106:9090/ws/erzlsmowebsvc.wsdl
 USE IN curSoap
RETURN m.address

FUNCTION CheckSOAPDirs(para1)
 IF !fso.FolderExists(m.curSoapDir)
  fso.CreateFolder(m.curSoapDir)
  fso.CreateFolder(m.curSoapDir+'\INPUT')
  fso.CreateFolder(m.curSoapDir+'\OUTPUT')
 ELSE 
  IF !fso.FolderExists(m.curSoapDir+'\INPUT')
   fso.CreateFolder(m.curSoapDir+'\INPUT')
  ENDIF 
  IF !fso.FolderExists(m.curSoapDir+'\OUTPUT')
   fso.CreateFolder(m.curSoapDir+'\OUTPUT')
  ENDIF 
 ENDIF 
RETURN 

FUNCTION Create_pRq(oXml, para1, para2)
 PRIVATE m.policy, m.dt
 m.policy = para1
 m.dt     = para2
 opRq = oXML.createElement("ser:pRq")
 opRq.appendChild(oXML.createElement("ser:sn")).text = m.policy
 opRq.appendChild(oXML.createElement("ser:dt")).text = m.dt
RETURN opRq

FUNCTION CreateHeader(ohttp, para1, para2, para3, para4) && para3 - надо сохранять в текстовом файле? para4 - куда?
 PRIVATE m.length, m.host, m.NeedToSave, m.FileName
 m.length     = para1
 m.host       =  para2
 m.NeedToSave = para3
 IF m.NeedToSave
  m.FileName     = para4
 ENDIF 

 ohttp.setRequestHeader("Accept-Encoding", "gzip,deflate")
 ohttp.setRequestHeader("Content-Type", "text/xml; charset=UTF-8")
 ohttp.setRequestHeader("SOAPAction", "")
 ohttp.setRequestHeader("Content-Length", m.length)
 ohttp.setRequestHeader("Host", m.host)
 ohttp.setRequestHeader("Connection", "Keep-Alive")
* ohttp.setRequestHeader("User-Agent", "Apache-HttpClient/4.1.1 (java 1.5)")
 ohttp.setRequestHeader("User-Agent", "Visual FoxPro 9.0 (MsXml2.XMLHTTP/MsXml2.DOMDocument)")

 IF m.NeedToSave
  poi = fso.CreateTextFile('&FileName')
  poi.WriteLine('Accept-Encoding: gzip,deflate')
  poi.WriteLine('Content-Type: "text/xml; charset=UTF-8"')
  poi.WriteLine('SOAPAction: ""')
  poi.WriteLine('Content-Length: '+ALLTRIM(STR(m.length)))
  poi.WriteLine('Host: ' + m.host)
  poi.WriteLine('Connection: Keep-Alive')
*  poi.WriteLine('User-Agent: "Apache-HttpClient/4.1.1 (java 1.5)"')
  poi.WriteLine('User-Agent: "Visual FoxPro 9.0 (MsXml2.XMLHTTP/MsXml2.DOMDocument)"')
  poi.Close
 ENDIF 

RETURN 

FUNCTION ReadHTTPHead
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  DO CASE
   CASE UPPER(READCFG) = 'DATE'
    m.cdate = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
  ENDCASE
 ENDDO
RETURN 

FUNCTION RFC2date(lcData) && Конверитруем из [Thu,  2 Feb 2012 08:42:29 +0300] (RFC822) в datetime-формат
 IF SET("Hours")!=24
  SET HOURS TO 24
 ENDIF 

* startpos = RAT(' ',lcData,5)
 startpos = 6
 lcDay    = PADL(ALLTRIM(SUBSTR(lcData,startpos,2)),2,'0')
 lcMonthT = SUBSTR(lcData,startpos+3,3)
 DO CASE
  CASE lcMonthT = 'Jan'
   lcMonth = '01'
  CASE lcMonthT = 'Feb'
   lcMonth = '02'
  CASE lcMonthT = 'Mar'
   lcMonth = '03'
  CASE lcMonthT = 'Apr'
   lcMonth = '04'
  CASE lcMonthT = 'May'
   lcMonth = '05'
  CASE lcMonthT = 'Jun'
   lcMonth = '06'
  CASE lcMonthT = 'Jul'
   lcMonth = '07'
  CASE lcMonthT = 'Aug'
   lcMonth = '08'
  CASE lcMonthT = 'Sep'
   lcMonth = '09'
  CASE lcMonthT = 'Oct'
   lcMonth = '10'
  CASE lcMonthT = 'Nov'
   lcMonth = '11'
  CASE lcMonthT = 'Dec'
   lcMonth = '12'
  OTHERWISE 
   lcMonth = '00'
 ENDCASE 

 lcYear  =  SUBSTR(lcData, startpos+7, 4)
 lcDate  = lcDay +'.' + lcMonth + '.' + lcYear
 lcTime = SUBSTR(lcData, startpos+12, 8)
 lcRealData = CTOT(lcDate + ' ' + lcTime)

RETURN lcRealData
