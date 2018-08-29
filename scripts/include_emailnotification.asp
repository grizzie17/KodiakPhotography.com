<!--#include file="mailserver.asp"-->
<!--#include file="mailsend.asp"-->
<!--#include file="remind_motorcycles.asp"-->
<%

DIM g_subscriptions_sDataFolder

FUNCTION findSubscriptionsFolder()
	DIM	sFolder
	
	IF "" = g_subscriptions_sDataFolder THEN
		sFolder = findFolder( "database\emailsubscriptions" )
		IF "" = sFolder THEN
			sFolder = findFolder( "cgi-bin\emailsubscriptions" )
		END IF
		IF "" <> sFolder THEN
			findSubscriptionsFolder = sFolder
			g_subscriptions_sDataFolder = sFolder
		ELSE
			findSubscriptionsFolder = ""
		END IF
	ELSE
		findSubscriptionsFolder = g_subscriptions_sDataFolder
	END IF
END FUNCTION


DIM	g_notification_sFile
g_notification_sFile = ""

FUNCTION findNotificationFile()

	IF "" <> g_notification_sFile THEN
		findNotificationFile = g_notification_sFile
	ELSE

		findNotificationFile = ""
		DIM	sFolder
		DIM	sFile
		
		sFile = ""
		sFolder = findSubscriptionsFolder()
		IF "" <> sFolder THEN
			sFile = g_oFSO.BuildPath(sFolder,"lastemailed.txt")
			g_notification_sFile = sFile
			IF g_oFSO.FileExists(sFile) THEN
				findNotificationFile = sFile
			END IF
		END IF
	END IF
	
END FUNCTION


FUNCTION findNotificationErrorFile()

		findNotificationErrorFile = ""
		DIM	sFolder
		DIM	sFile
		
		sFile = ""
		sFolder = findSubscriptionsFolder()
		IF "" <> sFolder THEN
			sFile = g_oFSO.BuildPath(sFolder,"emailerror.txt")
			IF g_oFSO.FileExists(sFile) THEN
				findNotificationErrorFile = sFile
			END IF
		END IF
	
END FUNCTION


SUB writeNotificationErrorFile( sError )
	DIM	sFolder
	DIM	sFile
	
	sFolder = findSubscriptionsFolder()
	IF "" <> sFolder THEN
		sFile = g_oFSO.BuildPath( sFolder, "emailerror.txt" )
		IF g_oFSO.FileExists(sFile) THEN
			g_oFSO.DeleteFile(sFile)
		END IF
		DIM	oFile
		SET oFile = g_oFSO.CreateTextFile( sFile, TRUE )
		oFile.WriteLine sError
		oFile.Close
		SET oFile = Nothing
	END IF
END SUB


SUB clearNotificationErrorFile()
	DIM	sFolder
	DIM	sFile
	
	sFolder = findSubscriptionsFolder()
	IF "" <> sFolder THEN
		sFile = g_oFSO.BuildPath( sFolder, "emailerror.txt" )
		IF g_oFSO.FileExists(sFile) THEN
			g_oFSO.DeleteFile(sFile)
		END IF
	END IF
END SUB


FUNCTION needsNotification()
	needsNotification = FALSE
	
	DIM	sFile
	sFile = findNotificationFile()
	IF "" = sFile THEN
		needsNotification = TRUE
	ELSEIF "" <> findNotificationErrorFile() THEN
		needsNotification = TRUE
	ELSE
		DIM	oFile
		SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
		IF NOT Nothing IS oFile THEN
			DIM	sLine
			sLine = oFile.ReadLine
			oFile.Close
			SET oFile = Nothing
			
			DIM	d
			d = CDATE(sLine)
			IF d <> DATE THEN
				needsNotification = TRUE
			END IF
		END IF
	END IF
END FUNCTION


SUB	writeNotificationDate()

	DIM	sFile
	sFile = findNotificationFile()
	IF "" <> sFile THEN
		IF g_oFSO.FileExists(sFile) THEN
			g_oFSO.DeleteFile(sFile)
		END IF
		DIM	oFile
		SET oFile = g_oFSO.CreateTextFile( sFile, TRUE )
		oFile.WriteLine DATE
		oFile.Close
		SET oFile = Nothing
	END IF
END SUB



FUNCTION findSubscriptionsFile()

	findSubscriptionsFile = ""
	DIM	sFolder
	DIM	sFile

	sFile = ""
	sFolder = findSubscriptionsFolder()
	IF "" <> sFolder THEN
		sFile = g_oFSO.BuildPath(sFolder,"subscriptions.txt")
		IF g_oFSO.FileExists(sFile) THEN
			findSubscriptionsFile = sFile
		END IF
	END IF

END FUNCTION


SUB processNotification()

	nDateBegin = jdateFromVBDate( NOW )
	nDateEnd = nDateBegin + 3
	
	gHtmlOption_encodeEmailAddresses = FALSE
	
	SET oCalendar = loadMotorcycleFiles( nDateBegin, nDateEnd, "email,email2", "holiday,none", FALSE )
	IF NOT Nothing IS oCalendar THEN
	
		DIM	sTomorrow
		sTomorrow = "/calendar/event[($any$ category='email') && (date='" & logDateFromJDate( nDateBegin +1 ) & "')]"
		'Response.Write "<p>" & Server.HTMLEncode(sTomorrow) & "</p>"
	
		IF NOT Nothing IS oCalendar.xmldom.documentElement.selectSingleNode(sTomorrow) THEN
	
			'Load the XSL
			DIM	oXSL
			DIM	oXML
			SET oXSL = Server.CreateObject("msxml2.DOMDocument")
			oXSL.async = false
			oXSL.load(Server.MapPath("scripts/remind.xslt"))
			
			SET oXML = oCalendar.xmldom
			
			DIM	sEvents
			sEvents = oXML.transformNode(oXSL)
			'Response.Write sEvents
			'EXIT SUB
			
			DIM	sFile
			sFile = findSubscriptionsFile()
			IF "" <> sFile THEN
		
				DIM oSMTP
			
				SET oSMTP = mailMakeSMTP()
				IF NOT Nothing IS oSMTP THEN
					oSMTP.Server = "mail.rocketcitywings.org"
					'oSMTP.Server = mailServer()
					
					DIM	sAccount
					DIM	sPW
					
					sAccount = "Notifications@rocketcitywings.org"
					sPW = "bear1701"
					
					IF "" <> sAccount  AND  "" <> sPW THEN
						oSMTP.User = sAccount
						oSMTP.PW = sPW
					END IF
					
					
					DIM	bSubscriptions
					DIM	sLine
					DIM	oFile
					SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
					
					bSubscriptions = FALSE
					
					IF NOT Nothing IS oFile THEN
					
						oSMTP.From = "Notifications@RocketCityWings.org"
						IF "localhost" <> LCASE(Request.ServerVariables("HTTP_HOST")) THEN
							oSMTP.FromName = "Event Notification"
							DO UNTIL oFile.AtEndOfStream
								sLine = TRIM(oFile.ReadLine)
								IF "" <> sLine  AND  1 < INSTR(sLine,"@") THEN
									oSMTP.AddRecipientBCC sLine
									bSubscriptions = TRUE
								END IF
							LOOP
						ELSE
							oSMTP.FromName = "Local RocketCityWings.org"
							oSMTP.AddRecipientBCC "Webmaster@RocketCityWings.org"
							bSubscriptions = TRUE
						END IF
						oFile.Close
						SET oFile = Nothing
					END IF
					
					IF bSubscriptions THEN
					

						DIM	sTextCSS
						sTextCSS = ""
						sFile = findRemindFile( "remind.css" )
						IF "" <> sFile THEN
							SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
							IF NOT oFile IS Nothing THEN
								sTextCSS = oFile.ReadAll
								oFile.Close
								SET oFile = Nothing
							END IF
						END IF
						
						DIM	i
						DIM	sBaseURL
						sBaseURL = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")
						i = INSTRREV( sBaseURL, "/" )
						IF 0 < i THEN sBaseURL = LEFT(sBaseURL,i)


						oSMTP.Subject = "Rocket City Wings upcoming events"
						
						oSMTP.Body = "" _
							&	"<html>" & vbCRLF _
							&	"<head>" & vbCRLF _
							&	"<style type=""text/css"">" & vbCRLF _
							&	vbCRLF _
							&	"h2" & vbCRLF _
							&	"{" & vbCRLF _
							&	"color: #FFFFFF;" & vbCRLF _
							&	"background-color: #990000;" & vbCRLF _
							&	"font-family: Arial,Helvetica,Verdana,sans-serif;" & vbCRLF _
							&	"border-bottom: 6px solid #FFCC00;" & vbCRLF _
							&	"padding-left: 0.25em;" & vbCRLF _
							&	"}" & vbCRLF _
							&	vbCRLF _
							&	".noreply" & vbCRLF _
							&	"{" & vbCRLF _
							&	"color: #999999; " & vbCRLF _
							&	"}" & vbCRLF _
							&	vbCRLF _
							&	sTextCSS & vbCRLF _
							&	vbCRLF _
							&	"</style>" & vbCRLF _
							&	"<base href=""" & sBaseURL & """>" & vbCRLF _
							&	"</head>" & vbCRLF _
							&	"<body>" & vbCRLF _
							&	"<h2>Rocket City Wings <i>Upcoming Events</i></h2>" & vbCRLF _
							&	"<table>" & vbCRLF _
							&	"<tr><td>" & vbCRLF _
							&	sEvents & vbCRLF _
							&	"</td></tr>" & vbCRLF _
							&	"</table>" & vbCRLF _
							&	"<p><br></p>" & vbCRLF _
							&	"<p>For more events visit " _
							&		"<a href=""http://www.RocketCityWings.org/"">" _
							&		"www.RocketCityWings.org" _
							&		"</a>" _
							&	"</p>" & vbCRLF _
							&	"<p class=""noreply"">Do not reply to this email</p>" & vbCRLF _
							&	"</body>" & vbCRLF _
							&	"</html>" & vbCRLF
					
					
						'send it
						DIM	sError
						sError = oSMTP.Send
						IF "0:" = LEFT(sError,2) THEN
							clearNotificationErrorFile
							writeNotificationDate
						ELSE
							writeNotificationErrorFile sError
						END IF
					END IF
					
					
				END IF
		
				SET oSMTP = Nothing
			
			END IF
		ELSE
			clearNotificationErrorFile
			writeNotificationDate
		END IF
	END IF

	
END SUB


SUB emailNotification()

	'IF "localhost" <> LCASE(Request.ServerVariables("HTTP_HOST")) THEN
		IF needsNotification() THEN
			writeNotificationDate
			processNotification
		END IF
	'END IF

END SUB



IF Response.Buffer THEN Response.Flush

emailNotification


%>
