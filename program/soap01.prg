PROCEDURE soap01

*VBSTART
*Function SoapExample(input)
* dim SOAPClient
* set SOAPClient = createobject("MSSOAP.SOAPClient")
* on error resume next
* SOAPClient.mssoapinit("http://webservices.daehosting.com/services/eleventest.wso?wsdl")
* if err then
*    MsgBox SOAPClient.faultString
*    MsgBox SOAPClient.detail
* end if
* SoapExample = SOAPClient.StripToNumeric(input)
* if err then
*    MsgBox SOAPClient.faultString
*    MsgBox SOAPClient.detail
* end if
*End Function
*VBEND

*VBEval>SoapExample("1a2b3c4d5e6"),OUT
*MessageModal>OUT
RETURN 