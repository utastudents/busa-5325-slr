<html>
<body>
<% temp=request.form("intMaxNum")
   intMaxNumber = Cint(request.form("intMaxNum"))
   strPrefix = request.form("prefix")
   strPath = strPrefix & "/" & strPrefix & "-" 
   SlideNumber = Request.form("txtSlideNumber")

  

   if isnumeric(slidenumber) then
     intSlideNumber = Cint(trim(slidenumber))
   else
     intslidenumber = 0
   end if
 
   'Check for Values outside range
   if intSlidenumber<=0 then
       intslidenumber=0
   end if

   if intslidenumber>=intMaxNumber then
      intslidenumber=intmaxnumber
   end if
  
   'Request direction and go forward or back 
   strDirection = request.form("GO")
      
   'If previous chosen and there are some left, go to previous slide   
   if (strDirection = "PRE") and (intslidenumber>0) then
        intslidenumber = intslidenumber-1
   end if

   'If Next chosen and there are some left, go to previous slide   
     if (strdirection = "NXT") and (intslidenumber<intmaxnumber) then
        intSlideNumber=intSlidenumber+1
   end if

%>

<TABLE BORDER=0>
<TR>
<TD>
<%
   if intSlideNumber<10 then
     strFileName = strpath & "00" & intSlideNumber
   else
     strFileName = strPath & "0" & intSlideNumber
   end if
%>

<img src="<%=strFilename%>.gif">
<embed src="<%=strFileName%>.wav" hidden="TRUE" >
<bgsound src="<%=strFileName%>.wav" >


</TD>
<TD>
<FORM ACTION="changefieldNet.asp" method="post">
     Slide Number:<BR>
     <INPUT NAME="txtSlideNumber" VALUE="<%=intSlideNumber%>"  size=3><p>
     Slides Numbered from 0 to <%=intMaxNumber%><p> 
     <INPUT TYPE=RADIO NAME=GO VALUE=PRE>Previous 
     <INPUT TYPE=RADIO NAME=GO VALUE=NXT>Next<p> 
     <input Name="Prefix" value="<%=strPrefix%>" Type = "Hidden">
     <input Name="intMaxNum" value = "<%=intMaxNumber%>" type="hidden">
     <INPUT TYPE=submit VALUE="GO">
     <Input type=reset value="Cancel"> 
  </FORM></p>
</TD>
</TR>
</TABLE>

</body>
</html>
