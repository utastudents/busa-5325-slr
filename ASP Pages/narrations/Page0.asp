
<html>
<head>
<%
    dim strPrefixNumber
    dim strPrefix
    dim intNumberSlides
    dim intDashPos
    dim strPath

    strPrefixNumber = request.form("go")

    intDashPos = instr(strPrefixNumber,"-")

    strPrefix = left(strprefixnumber, intDashPos-1)
    intNumberSlides = right(strPrefixNumber,len(strprefixnumber)-intDashPos)
    strPath=strPrefix & "/" & strPrefix
%>
<bgsound src="<%=strPath%>-000.wav">
</head>

<body>
<h4> At the Far Right, Please Select Either the Previous or Next Radio Button and Press the Enter Key </h4>

<TAble border=0>
<tr>
<TD>


<img src="<%=strpath%>-000.gif">

</td>
<td>
<Form action="changefield.asp" method="post">
     Slide Number:<BR>
     <INPUT NAME="txtSlideNumber" VALUE="0"><p>
     Slides Numbered 0 to <%=intNumberSlides%><p>
     <INPUT TYPE=RADIO NAME=GO VALUE=PRE>Previous 
      <INPUT TYPE=RADIO NAME=GO VALUE=NXT>Next <p>
     <INPUT TYPE=submit VALUE="Go">
     <Input type=reset value="Cancel">
     <input Name="Prefix" value="<%=strPrefix%>" Type = "Hidden">
     <input Name="intMaxNum" value = "<%=intNumberSlides%>" type="hidden">
</Form> 
</td>
</tr>
</table>

</body>
</html>
