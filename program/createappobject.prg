Procedure CreateAppObject
	If Type("_Screen.oApp") <> "O"
		_Screen.AddProperty("oApp")
	EndIf 
	_Screen.oApp = CreateObject("AppObject")
EndProc 

DEFINE CLASS AppObject AS Custom

Procedure SimulateWork
	Local i

	For i = 1 to 1000000
		* Peg CPU
	EndFor
EndProc 

Procedure Test
	Lparameters lnUnits
	Local i

	? Program(), lnUnits
	For i = 1 to lnUnits
		This.SimulateWork()
	EndFor 	

EndProc 

ENDDEFINE
