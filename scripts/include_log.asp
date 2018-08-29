<!--#include file="mailserver.asp"-->
<!--#include file="mailsend.asp"-->
<%


FUNCTION logPage_isUser( sAgent )

	IF "" <> sAgent THEN
		DIM	sAgentLC
		sAgentLC = LCASE(sAgent)
		
		logPage_isUser = TRUE
		IF 0 < INSTR(sAgentLC,"slurp") THEN
			logPage_isUser = FALSE
		ELSEIF 0 < INSTR(sAgentLC,"teoma") THEN
			logPage_isUser = FALSE
		END IF
	END IF

END FUNCTION


SUB logPage()
	IF Response.Buffer THEN Response.Flush
	
	IF FALSE  AND  "localhost" <> LCASE(Request.ServerVariables("HTTP_HOST")) THEN
	
		DIM	sAgent
		sAgent = Request.ServerVariables("HTTP_USER_AGENT")
		
		IF logPage_isUser( sAgent ) THEN

			DIM oSMTP
		
			SET oSMTP = mailMakeSMTP()
			IF NOT Nothing IS oSMTP THEN
				oSMTP.Server = "mail.rocketcitywings.org"
				'oSMTP.Server = mailServer()
				
				DIM	sAccount
				DIM	sPW
				
				sAccount = "admin@rocketcitywings.org"
				sPW = "bear1701"
				
				IF "" <> sAccount  AND  "" <> sPW THEN
					oSMTP.User = sAccount
					oSMTP.PW = sPW
				END IF
				
				DIM	sURL
				sURL = Request.ServerVariables("URL")
				DIM	sQuery
				sQuery = Request.ServerVariables("QUERY_STRING")
				
				DIM	sPage
				DIM	i
				
				i = INSTRREV( sURL, "/" )
				IF 0 < i THEN
					sPage = MID( sURL, i+1 )
				ELSE
					sPage = sURL
				END IF
	
				mailLoadNames oSMTP, "admin@rocketcitywings.org", "", "John@GrizzlyWeb.com", "", ""
				
				oSMTP.Subject = "Visitor to RocketCityWings.org" & " - " & sPage
				
				oSMTP.Body = "" _
						&	"Referer = " & Request.ServerVariables("HTTP_REFERER") & vbCRLF _
						&	"URL = " & sURL & vbCRLF _
						&	"Query = " & sQuery & vbCRLF _
						&	sAgent
				
				
				'send it
				DIM	sError
				sError = oSMTP.Send
			END IF
	
			SET oSMTP = Nothing
		
		END IF

	END IF
END SUB

logPage


%>
