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
Else
    If Len(Trim(Request.Form("FirstName")))>0 Then
       Response.Write "Your name is " & Trim(Request.Form("FirstName"))
       Response.Write "<br />Thanks!!"
    Else
       Response.Write "You forgot to enter your name"
       call showform()
   End If
End If
%>