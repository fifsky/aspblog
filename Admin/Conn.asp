
<%
	dim conn
	dim connstr
	dim db
	db="../Database/data.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)
	conn.Open connstr
%>
