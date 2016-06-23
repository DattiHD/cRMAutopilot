Option Explicit

' eMail Autopilot Objekte
Dim oMail
Set oMail = HostApp.Mail ' Gibt die aktuelle eMail als Mail-Objekt zurück. S193 

' eMail Autopilot Variablen
Dim sMailSender
sMailSender = oMail.SenderAddressResolved 'Name des Absenders inklusive eMail-Adresse. S196

Dim sMailBody
sMailBody = oMail.BodyMessage 'Liefert den kompletten Nachrichtentext zurück. S195

Dim sMailSubject
sMailSubject = oMail.Subject 'Gibt den Betreff der eMail zurück. S197

Dim sMailSendDateTime
sMailSendDateTime = oMail.SendDateTime ' Gibt das Sendedatum inkl. Uhrzeit der eMail zurück. Das Format wird gemäß den Windows Systemeinstellungen zurückgegeben, d.h. bei einer Regionseinstellung "Deutsch" erfolgt die Ausgabe umgerechnet in die lokale Zeit im Format: 16.07.2012 11:09:04 S196

'--------------------------------------------------------------------------------------------------------------
' Konfiguration:
'--------------------------------------------------------------------------------------------------------------
Const sMailView = "Bewerber" ' in dieser Ansicht befinden sich die eMail-Adressen
' Const sRMContactContainer = "Kontakte" ' Ansicht in der die Newsletterbestellung protokolliert wird EVTL RELEVANT FÜR PROTOKOLLIERUNG IN DEN AKTIVITÄTEN

' Protokolldatei:
Const sLogFile = "cRM-kontakt_u_regform.log"

' cRM Projektvariable
Dim ocRMListViewConfigs 'Liefert die Anzahl der konfigurierten Ansichten. S73
Set ocRMListViewConfigs = ocRMProject.ViewConfigs 'Gibt die konfigurierten Ansichten im aktuellen Projekt als Objekt vom Typ ListViewConfigs zurück. S81

Dim oApplicants
Set oApplicants = ocRMListViewConfigs.ItemByName(CStr(sMailView))
If oApplicants Is Nothing Then
	MsgBox "Die Ansicht " & sMailView & " kann nicht geöffnet werden: " & sMailView, vbInformation, "combit eMail Autopilot Script"
	ReleaseObjects()
	WScript.Quit
End If

' Dim ocRMViewConfig 'Repräsentiert eine aktive (geöffnete) Ansicht in der combit Anwendung. Ab S151
' Set ocRMViewConfig = ocRMListViewConfigs.ItemByName(CStr(sMailView)) ' Ansicht in der die eMail-Adressen stehen
' If ocRMViewConfig Is Nothing Then
'	MsgBox "Die Ansicht in der die eMail-Adressen enthalten sind kann nicht geöffnet werden: " & sMailView, vbInformation, "combit eMail Autopilot Script"
'	ReleaseObjects()
'	WScript.Quit
' End If

Dim opCurrentRecordSet 'Liefert ein Objekt vom Typ RecordSet zurück. S154 CHECKEN, OB DAS SINNVOLL IST!

Dim ocRMProject
Set ocRMProject = cRM.CurrentProject

Dim oMailViewRecordSet
Set oMailViewRecordSet = ocRMViewConfig.CreateRecordSet

' Logdatei
Dim oFSO, oLogFile, otmpfolder
Const TemporaryFolder = 2
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set otmpfolder = oFSO.GetSpecialFolder(TemporaryFolder)
Set oLogFile = oFSO.OpenTextFile (otmpfolder & sLogFile, 8, true) 
If oLogFile Is Nothing Then
	MsgBox "Logdatei kann nicht erstellt werden: " & sLogFile, vbInformation, "combit eMail Autopilot Script"
	ReleaseObjects()
	WScript.Quit
End If

' Statuszeile
SetStatus("Neue eMail zur Bearbeitung eingetroffen. Bearbeitungszeitpunkt: " & CStr(Now))

' eMail-Betreff und -Text
Dim sMailtext, sStatusTextLine
sMailtext = GetFieldsMailText(sMailBody)
If sMailtext = "" Then
	Msgbox "BEGIN:FIELDS - END:FIELDS fehlt in der eMail! Script wird beendet.", vbInformation, "combit eMail Autopilot Script"
	ReleaseObjects()
	WScript.Quit
End If

' Die Variablen tauchen nur hier auf. Ich seh absolut keine Verwendung dafür bisher
' Dim sFormula4File, bExit 
' bExit = False
' sFormula4File = "x"

' Main:
Call Main()
SetStatus("============================")
' Objekte freigeben:
Call ReleaseObjects()

Function Main()
Dim nReturnFilter, oMailViewRecord		

If InStr(1, sMailSubject, "Kontaktformular") > 0 Then
	SetStatus("Aktion: Kontaktformular")
	nReturnFilter = DSFilterForEMail(oApplicants)
		If nReturnFilter >= 0 Then
			If nReturnFilter = 1 Then
				' Kein Datensatz anhand der eMail-Adresse gefunden, neuen Datensatz anlegen:
				Set oMailViewRecord = oMailViewRecordSet.NewRecord()
			Else
					' Datensatz gefunden:
				Set oMailViewRecord = oMailViewRecordSet.CurrentRecord
			End If
		End If
		
		If WriteCRMFieldsKontakt = True Then
			SetStatus("Datensatz erfolgreich geschrieben")
		End If
End If

If InStr(1, sMailSubject, "Registrierungsformular") > 0 Then
	Setstatus("Aktion: Registrierungsformular")
	nReturnFilter = DSFilterForEMail(oApplicants)
		If nReturnFilter >= 0 Then
			If nReturnFilter = 1 Then
				' Kein Datensatz anhand der eMail-Adresse gefunden, neuen Datensatz anlegen:
				Set oMailViewRecord = oMailViewRecordSet.NewRecord()
			Else
					' Datensatz gefunden:
				Set oMailViewRecord = oMailViewRecordSet.CurrentRecord
			End If
		End If
		
		If WriteCRMFieldsReg = True Then
			SetStatus("Datensatz erfolgreich geschrieben")
		End If

Else
	Setstatus("Information: Die E-Mail ist weder Kontaktformular noch Registrierung.")
End If
Exit Function 

'--------------------------------------------------------------------------------------------------------------
' Funktion: Hilfsfunktion filtern von Datensätzen anhand der eMail-Adresse:
Function DSFilterForEMail(ByVal oRMConfig)
'--------------------------------------------------------------------------------------------------------------
	
	Dim nFldType, i
	Dim bNewFilter
	DSFilterForEMail = -1
	SetStatus("Datensatz anhand der eMail-Adresse suchen.")
	
	For i = 0 To oRMConfig.FldCount
		' eMail-Feld = 12:
		If oRMConfig.FldType(i) = 12 Then
			Dim sMAILField, sValue
			sMAILField = oRMConfig.FldName(i)
			
			If ParseFormula(sMAILField, sValue) = True Then
				Dim sFilterExpression
				sFilterExpression = GetEMAILFilter(sMAILField, sValue)
				
				If sFilterExpression <> "" Then					
					' Filter im cRM erstellen:
					bNewFilter = oMailViewRecordSet.SetFilter(CStr(sFilterExpression))
					If oMailViewRecordSet.MoveFirst Then
						
						Dim nRecCount
						nRecCount = oMailViewRecordSet.RecCount
						
						If nRecCount > 1 Then
							SetStatus("Es wurden mehrere Datensätze gefunden.")
							' Es wurden mehrere DS mit der eMail-Adresse gefunden. 
							DSFilterForEMail = 2 
						ElseIf nRecCount = 1 Then					
							
							SetStatus("Der Datensatz wurde gefunden.")
							DSFilterForEMail = 0													
							
						End If 
					Else
						SetStatus("Kein Datensatz gefunden - Neuanlage.")
						DSFilterForEMail = 1						
					End If
				End If
			End If
		End If
	Next	
End Function

'--------------------------------------------------------------------------------------------------------------
' Funktion: Status für den combit eMail AutoPiloten schreiben
Function SetStatus(ByVal sStatus)
'--------------------------------------------------------------------------------------------------------------
	sStatusTextLine = sStatusTextLine & vbCrLf & sStatus
	HostApp.SetOutputString(sStatusTextLine)
	
	oLogFile.Write sStatus & vbCrLf

End Function

'--------------------------------------------------------------------------------------------------------------
' Funktion: 
Function WriteCRMFieldsKontakt(ByRef oFilterRecordSet)
'--------------------------------------------------------------------------------------------------------------
	
	Dim sField, sValue, i
	Dim oRecord
	Set oRecord = oFilterRecordSet.CurrentRecord
	If oRecord Is Nothing Then
		Set oRecord = oFilterRecordSet.NewRecord()
	End If
	
	' Datensatz zum Schreiben sperren:
	oRecord.Lock
	For i = 1 To oApplicants.FldCount
		' Felder dürfen nicht vom Typ: 21 Datensatz-ID, 14 Globale Eindeutige ID, 24 Automatische Nr. sein
		If oApplicants.FldType(i) <> 14 And oApplicants.FldType(i) <> 24 And oApplicants.FldType(i) <> 21 Then
			sField = oApplicants.FldName(i)
			If ParseFormula(sField, sValue) = True Then
				
				Dim sCurrentFieldValue 
				sCurrentFieldValue = oRecord.GetContentsByName(CStr(sField))
				
				' Überprüfen, ob der Inhalt des Feldes leer ist oder bei Logischen Feldern auf False gesetzt ist:
				If sCurrentFieldValue = "" Then
					
					SetStatus("Neuer Feldinhalt: "  &   sField & " - " & sValue)
					
					' Schreiben:
					oRecord.SetContentsByName CStr(sField), CStr(sValue)
					
				End If
				
			End If		
		End If
		
	Next
	
	If Not oRecord.Save Then
		
		SetStatus("Datensatz konnte nicht geschrieben werden!")
		Set oRecord = Nothing
		WriteCRMFields False
		Exit Function
	End If	
	
	oRecord.Unlock
	
	' Datensatz-ID ermittlen:
	sPrimaryKey = oRecord.GetContentsByName(CStr(cPrimaryField))
	oFilterRecordSet.SetFilterByPrimaryKey(CStr(sPrimaryKey))
	
	Set oRecord = Nothing
	WriteCRMFields = True

End Function

'--------------------------------------------------------------------------------------------------------------
' Funktion: 
Function WriteCRMFieldsReg(ByRef oFilterRecordSet)
'--------------------------------------------------------------------------------------------------------------
	
	Dim sField, sValue, i
	Dim oRecord
	Set oRecord = oFilterRecordSet.CurrentRecord
	If oRecord Is Nothing Then
		Set oRecord = oFilterRecordSet.NewRecord()
	End If
	
	' Datensatz zum Schreiben sperren:
	oRecord.Lock
	For i = 1 To oApplicants.FldCount
		' Felder dürfen nicht vom Typ: 21 Datensatz-ID, 14 Globale Eindeutige ID, 24 Automatische Nr. sein
		If oApplicants.FldType(i) <> 14 And oApplicants.FldType(i) <> 24 And oApplicants.FldType(i) <> 21 Then
			sField = oApplicants.FldName(i)
			If ParseFormula(sField, sValue) = True Then
				
				Dim sCurrentFieldValue 
				sCurrentFieldValue = oRecord.GetContentsByName(CStr(sField))
				
				' Überprüfen, ob der Inhalt des Feldes leer ist oder bei Logischen Feldern auf False gesetzt ist:
				If sCurrentFieldValue = "" Then
					
					SetStatus("Neuer Feldinhalt: "  &   sField & " - " & sValue)
					
					' Schreiben:
					oRecord.SetContentsByName CStr(sField), CStr(sValue)
					
				End If
				
			End If		
		End If
		
	Next
	
	If Not oRecord.Save Then
		
		SetStatus("Datensatz konnte nicht geschrieben werden!")
		Set oRecord = Nothing
		WriteCRMFields False
		Exit Function
	End If	
	
	oRecord.Unlock
	
	' Datensatz-ID ermittlen:
	sPrimaryKey = oRecord.GetContentsByName(CStr(cPrimaryField))
	oFilterRecordSet.SetFilterByPrimaryKey(CStr(sPrimaryKey))
	
	Set oRecord = Nothing
	WriteCRMFields = True

End Function

'--------------------------------------------------------------------------------------------------------------
' Funktion: Hilfsfunktion zum Parsen einer Formel / Werts:
Function ParseFormula(ByVal sSearchKey, ByRef sValue)
'--------------------------------------------------------------------------------------------------------------
	
	sValue = ""
	Dim nPos, nPosEnd
	nPos = Instr(1, sMailtext, sSearchKey & ":")
	
	If nPos <> 0 Then
		nPos = nPos + Len(sSearchKey & ":")
		nPosEnd = Instr(nPos, sMailtext, vbCrLf)
		sValue = Mid(sMailtext, nPos, nPosEnd - nPos)			
		
		If Len(Trim(sValue)) > 0 And sValue <> vbCrLf Then
			
			ParseFormula = True
			Exit Function
		End If
	End If

	ParseFormula = False
	
End Function

'--------------------------------------------------------------------------------------------------------------
' Funktion: Objekte freigeben:
Sub ReleaseObjects()
'--------------------------------------------------------------------------------------------------------------

	' Objekte freigeben:
	Set oMailViewRecordSet = Nothing
	Set oSubscribeRecordSet = Nothing
	Set ocRMViewConfig = Nothing
	Set oApplicants = Nothing
	Set ocRMListViewConfigs = Nothing	
	Set ocRMProject = Nothing
	Set oMailObject = Nothing
	
End Sub