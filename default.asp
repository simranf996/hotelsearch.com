
<%

	'Database Variables
	Dim strDSN, strCityName, strHotelName

	CONST_DB_SERVER = "ISCWPC0QN373\SQLEXPRESS04"
	strInitialCatalog = "HotelBookingSystem"
	strUserId = "farheen;"
	strPasswd = "farheen;"

	strDSN = "Provider = SQLOLEDB.1; Password = " & strPasswd & " User ID = " & strUserId & " Initial Catalog = " & strInitialCatalog & "; Data Source = " & CONST_DB_SERVER

  strHotelName = Trim(Request.Form("txtHotelName"))
	if strHotelName = "" THEN
	strHotelName = Trim(Request.QueryString("txtHotelName"))
	end if

 strCityName = Trim(Request.Form("cboCityName"))
 if strCityName = "" THEN
 strCityName = Trim(Request.QueryString("cboCityName"))
 end if

 Set dbConn = Server.CreateObject("ADODB.Connection")
	dbConn.Open strDSN

	Sub PopulateCityNames()

			Set dbConn = Server.CreateObject("ADODB.Connection")
			dbConn.Open strDSN

	 		strSQLItems = " SELECT	DISTINCT City " & _
	 									"	FROM			Hotels "
			Set rsItems = dbConn.Execute( strSQLItems )

			Response.Write "	<SELECT NAME=""cboCityName"" STYLE=""WIDTH:150PX; Height: 50px;""  CLASS=""clsSelect""> "
	 		Response.Write "		<OPTION VALUE="""">--Select City--</OPTION> "

			If Not rsItems.EOF Then
	 			Do While Not rsItems.EOF
	 				strItem = Trim( rsItems( "City" ) )
					if strItem <> "" Then
	 				If strCityName =  strItem Then
	 					Response.Write "	<OPTION VALUE=""" & strItem & """ SELECTED> " & strItem & " </OPTION> "
	 				Else
	 					Response.Write "	<OPTION VALUE=""" & strItem & """> " & strItem & " </OPTION> "
	 				End If
					End If
	 			rsItems.MoveNext
 			Loop
	 		End If

	 		If IsObject( rsItems ) Then
	 			Set rsItems = Nothing
	 	  End If

	End Sub

Sub BuildHotelList()

		strSQLHotelList =	"SELECT  HotelName, City, Area, TelephoneNo, TotalRooms " & _
											"FROM	Hotels " & _
											"WHERE "
    strWhere = ""

		IF strHotelName <> "" THEN
		sqlQueryName =  " HotelName LIKE '%" & strHotelName & "%'"
			if strWhere <> "" then
	         	strWhere = strWhere & " and " & sqlQueryName
	    else
	         	strWhere = strWhere & sqlQueryName
	    End if
		End If

		IF strCityName <> "" THEN
		sqlQueryCity =  " City = '" & strCityName & "'"
			if strWhere <> "" then
						strWhere = strWhere & " and " & sqlQueryCity
			else
						strWhere = strWhere & sqlQueryCity
			End if
	 End If

		strSQLHotelList = strSQLHotelList & strWhere

		if strHotelName <> "" or strCityName <> "" Then
         strSQLHotelList = strSQLHotelList & " ORDER	BY	HotelName "

		Set rsHotelList = dbConn.Execute( strSQLHotelList )

					Response.Write "<TABLE BORDER=0 WIDTH=""50%"" ALIGN=""CENTER"" CELLSPACING=1 CELLPADDING=2 CLASS=""clsTable1"">"
					Response.Write	"<TR>"
					Response.Write			"<TH BGCOLOR=""#584b4f"" WIDTH=""5%""><FONT CLASS=""clsLabel"">#</FONT></TH>"
					Response.Write			"<TH BGCOLOR=""#584b4f"" WIDTH=""35%""><FONT CLASS=""clsLabel"">Hotel Name</FONT></TH>"
					Response.Write			"<TH BGCOLOR=""#584b4f"" WIDTH=""20%""><FONT CLASS=""clsLabel"">City</FONT></TH>"
					Response.Write			"<TH BGCOLOR=""#584b4f"" WIDTH=""20%""><FONT CLASS=""clsLabel"">Area</FONT></TH>"
					Response.Write			"<TH BGCOLOR=""#584b4f"" WIDTH=""20%""><FONT CLASS=""clsLabel"">Contact Number</FONT></TH>"
					Response.Write			"<TH BGCOLOR=""#584b4f"" WIDTH=""20%""><FONT CLASS=""clsLabel"">Total Rooms</FONT></TH>"
					Response.Write	"</TR>"


						If Not rsHotelList.EOF Then
								intSlNo = 1
								Do While Not rsHotelList.EOF
									strHotelName = Trim( rsHotelList( "HotelName" ) )
									strCity 		 = Trim( rsHotelList( "City" ) )
									strArea 		 = Trim( rsHotelList( "Area" ) )
									strTelephone  = Trim( rsHotelList( "TelephoneNo" ) )
									strRooms  = Trim( rsHotelList( "TotalRooms" ) )

									Response.Write "	<TR>"
								  Response.Write "		<TD  Class =""clsData"" BGCOLOR=""#584b4f"" ALIGN=""CENTER"" VALIGN=""TOP""> " & intSlNo & " </TD> "
									Response.Write "		<TD  Class =""clsData"" BGCOLOR=""#584b4f"" VALIGN=""TOP""> " & strHotelName & " </TD> "
								  Response.Write "		<TD  Class =""clsData"" BGCOLOR=""#584b4f"" VALIGN=""TOP""> " & strCity  & " </TD> "
								  Response.Write "		<TD  Class =""clsData"" BGCOLOR=""#584b4f"" VALIGN=""TOP""> " & strArea  & " </TD> "
									Response.Write "		<TD  Class =""clsData"" BGCOLOR=""#584b4f"" VALIGN=""TOP""> " & strTelephone  & " </TD> "
									Response.Write "		<TD  Class =""clsData"" BGCOLOR=""#584b4f"" VALIGN=""TOP""> " & strRooms  & " </TD> "
									Response.Write "	</TR> "

								intSlNo = intSlNo + 1
								rsHotelList.MoveNext
								Loop
								Response.Write "</TABLE>"

						End If

						If IsObject( rsHotelList ) Then
								Set rsHotelList = Nothing
						End If
  end if
End Sub

%>
<HTML>
<HEAD>
	<link href="style.css" type="text/css" rel="stylesheet">
	<link href="https://fonts.googleapis.com/css?family=Qwigley|Ruda" type="text/css" rel="stylesheet">
	<link href="https://fonts.googleapis.com/css?family=Lato" rel="stylesheet">
	<title> Hotel Booking System </title>

</HEAD>

<BODY>
 <div align ="center">
	 <H1 class ="clsHeader"> Search Hotels </H1>
 </div>
 <br/>

 <div align="center">
	<Form name="frmHotelsList">
		<Table>
		<TR>

			<TD>

			<% Call PopulateCityNames() %>
		  </TD>
		<TD>
		<INPUT TYPE="TEXT" size="48" MAXLENGTH="80" NAME="txtHotelName" CLASS="clsFormInput" VALUE= "<% = strHotelName %>" onMouseOver="Javascript:settooltip()"/>
		</TD>



		<TD>
	<INPUT TYPE="BUTTON" NAME="btnGenerateHotelsList"  CLASS="clsButton" VALUE="Search"	ONCLICK="JavaScript:GenerateHotelsList()">
	</TD>
		<TR/>


		</Table>
	</Form>
 </div>

 <br/>
 <br/>
 <br/>

		<%
		IF strHotelName <> "" or strCityName <> "" Then
		Call BuildHotelList()
		else
			Response.Write "	<p> To search for a hotel, type in a name and click on ""Search"" </p>"
			Response.Write "	<p> Alternatively, you can search all hotels in a city by selecting it clicking on ""Search"". </p> "
		End If
		%>

	<br>
</BODY>


</HTML>

<Script language="Javascript">

function settooltip() {

			document.getElementById("txtHotelName").title = 'Please enter the hotel name to search';
			document.getElementById("cboCityName").title = 'Please enter the city to search';

	}

	function GenerateHotelsList()
		{


				document.frmHotelsList.method = "post"
				document.frmHotelsList.action = "default.asp"
				document.frmHotelsList.target = "content"
				document.frmHotelsList.submit()
		}

</Script>

<%
	'Destroy the Database and Recordset Connections
	If IsObject( dbConn ) Then
		dbConn.Close
		Set dbConn = Nothing
	End If

	If IsObject( rsGetUserName ) Then
		Set rsGetUserName = Nothing
	End If
%>
