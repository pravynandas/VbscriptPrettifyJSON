' ==============================================================================================
' Adaptation of JSONToXML() function for VbscriptPrettifyJSON().
' Author: Praveen Nandagiri (pravynandas@gmail.com)
'
' JSONToXML() Credits:
' Visit: https://stackoverflow.com/a/12171836/1751166
' Author: https://stackoverflow.com/users/881441/stephen-quan
' ==============================================================================================
Const stateRoot = 0
Const stateNameQuoted = 1
Const stateNameFinished = 2
Const stateValue = 3
Const stateValueQuoted = 4
Const stateValueQuotedEscaped = 5
Const stateValueQuotedEscapedHex = 6
Const stateValueUnquoted = 7
Const stateValueUnquotedEscaped = 8
Dim iLevel
Dim bLastLevel

Function VbscriptPrettifyJSON(json)
  Dim out
  Dim i, ch, state, name, value, sHex
  out = ""
  bLastLevel = ""
  state = stateRoot
  For i = 1 to Len(json)
    ch = Mid(json, i, 1)
    Select Case state
    Case stateRoot
      Select Case ch
      Case "["
		out = Append(out, ch, "", 1)
      Case "{"
		out = Append(out, ch, "", 1)
      Case """"
        state = stateNameQuoted 
        name = ""
		out = Append(out, ch, "/r/n", 0)
      Case ","
        state = stateRoot
		out = Append(out, ch, "", 0)		
      Case "}"
		out = Append(out, ch, "/r/n", -1)
      Case "]"
		out = Append(out, ch, "/r/n", -1)
      End Select
    Case stateNameQuoted 
      Select Case ch
      Case """"
        state = stateNameFinished
		out = Append(out, ch, "", 0)
      Case Else
        name = name + ch
		out = Append(out, ch, "", 0)
      End Select
    Case stateNameFinished
      Select Case ch
      Case ":"
        value = ""
        State = stateValue
		out = Append(out, ch, " ", 0)
      Case Else						
        State = stateRoot
		out = Append(out, ch, "", 0)		
      End Select
    Case stateValue
      Select Case ch
      Case """"
        State = stateValueQuoted
		out = Append(out, ch, " ", 0)
      Case "{"
        State = stateRoot
		out = Append(out, ch, "/r/n", 1)
      Case "["
        State = stateRoot
		out = Append(out, ch, "/r/n", 1)
      Case " "
      Case Chr(9)
      Case vbCr
      Case vbLF
      Case Else
        value = ch
        State = stateValueUnquoted
      End Select
    Case stateValueQuoted
      Select Case ch
      Case """"
        state = stateRoot
		out = Append(out, value + ch, "", 0)
      Case "\"
        state = stateValueQuotedEscaped
      Case Else
        value = value + ch
      End Select
    Case stateValueQuotedEscaped 
	  If ch = "u" Then	'Four digit hex. Ex: o = 00f8
	  	sHex = ""
	  	state = stateValueQuotedEscapedHex
	  Else
	  	Select Case ch
	  	Case """"
	  		value = value + """"
	  	Case "\"
	  		value = value + "\"
	  	Case "/"
	  		value = value + "/"
	  	Case "b"	'Backspace
	  		value = value + chr(08)
	  	Case "f"	'Form-Feed
	  		value = value + chr(12)
	  	Case "n"	'New-line (LineFeed(10))
	  		value = value + vbLF
	  	Case "r"	'New-line (CarriageReturn/CRLF(13))
	  		value = value + vbCR
	  	Case "t"	'Horizontal-Tab (09)
	  		value = value + vbTab
	  	Case Else
	  		'do not accept any other escape sequence
	  	End Select
	  	state = stateValueQuoted
	  End If
	Case stateValueQuotedEscapedHex
	  sHex = sHex + ch
	  If len(sHex) = 4 Then
	  	on error resume next
	  	value = value + Chr("&H" & sHex)	'Hex to String conversion
	  	on error goto 0
	  	state = stateValueQuoted
	  End If
    Case stateValueUnquoted
      Select Case ch
      Case "}"
        state = stateRoot
		out = Append(out, ch, "/r/n", 1)
      Case "]"
        state = stateRoot
		out = Append(out, ch, "/r/n", 1)
      Case ","
        state = stateRoot
		out = Append(out, ch, "", 0)
      Case "\"
         state = stateValueUnquotedEscaped
      Case Else
        value = value + ch
      End Select
    Case stateValueUnquotedEscaped ' @@TODO: Handle escape sequences
      value = value + ch
      state = stateValueUnquoted
    End Select
  Next
  VbscriptPrettifyJSON = out
End Function

Function Append(out, sChar, sPrepend, iLvl)
	Select Case sPrepend
	Case vbCr, "/r"
			out = out + vbCr + sChar
	Case vbLF, "/n"
			out = out + vbLF + sChar
	Case vbCrLF, "/r/n"
			If iLvl = 1 Then 
				iLevel = iLevel + 1
			ElseIf iLvl = -1 Then 
				iLevel = iLevel - 1	
			End IF
			
			If sChar <> "{" And sChar <> "[" And sChar <> "}" And sChar <> "]" Then sChar = "  " + sChar
			If iLvl = -1 Then	'Fixing an issue in levelling down
				out = out + vbCrLF + Joiner(iLevel + 1, vbTab) + sChar
			Else
				out = out + vbCrLF + Joiner(iLevel, vbTab) + sChar
			End If
			bLastLevel = iLvl
	Case vbTab, "/t"
			out = out + vbTab + sChar
	Case Else
		If Left(sPrepend, 2) = "/s" Then
			on error resume next
			iSpaces = cInt(Right(sPrepend, len(sPrepend)-2))
			out = out + Space(iSpaces) + sChar
			on error goto 0
		Else
			sPrepend = Replace(sPrepend, "/r/n", vbCrLF)
			sPrepend = Replace(sPrepend, "/r", vbLF)
			sPrepend = Replace(sPrepend, "/n", vbLF)
			sPrepend = Replace(sPrepend, "/t", vbTab)
			out = out + sPrepend + sChar
		End If
	End Select
	Append = out
End Function

Function Joiner(iLvl, sChar)
	Dim lOut
	lOut = ""
	For i = 1 To iLvl
		lOut = lOut + sChar
	Next
	Joiner = lOut
End Function