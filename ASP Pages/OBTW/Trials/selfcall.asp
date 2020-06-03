<%
Sub showform
%>
<form method="post">
     Your Name: <input type="text" name="FirstName" /><br />
     <input type="submit" name="Submit" value="Submit" />
</form>
<%
End Sub

If Request.Form("Submit") = "" Then
    call showform()
Else  %>

<p><!--#include file = "includeTry.inc"--> </p>
<%

End If
%>