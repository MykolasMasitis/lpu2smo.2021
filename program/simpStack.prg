DEFINE CLASS simpStack AS session
	PROTECTED stackValue
	PROTECTED stackSep

	* Init props
	PROCEDURE Init()
		this.stackValue = ""
		this.stackSep = ","
	ENDFUNC

	* Tests if the stack is empty.
	FUNCTION Empty() AS Boolean
		RETURN (LENC(this.stackValue) = 0)
	ENDFUNC

	* Removes the element from the top of the stack.
	FUNCTION Pop() AS VOID
		IF !this.Empty()
			LOCAL lnPos as Integer
			lnPos = ATC(this.stackSep, this.stackValue)
			IF lnPos > 0
				this.stackValue = SUBSTRC(this.stackValue, lnPos + 1)
			ELSE
				this.stackValue = ""
			ENDIF
		ENDIF
	ENDFUNC

	* Adds an element to the top of the stack.
	FUNCTION Push(tcValue as String) AS VOID
		IF !this.Empty()
			this.stackValue = tcValue + this.stackSep + this.stackValue
		ELSE
			this.stackValue = tcValue
		ENDIF
	ENDFUNC

	* Returns the number of elements in the stack.
	FUNCTION Size() AS Integer
		IF !this.Empty()
			RETURN GETWORDCOUNT(this.stackValue, this.stackSep)
		ENDIF
		RETURN 0
	ENDFUNC

	* Returns a reference to an element at the top of the stack.
	FUNCTION Top() AS String
		IF !this.Empty()
			RETURN GETWORDNUM(this.stackValue, 1, this.stackSep)
		ENDIF
		RETURN ""
	ENDFUNC

ENDDEFINE
