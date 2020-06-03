   <% 
   If Len(Trim(Request.Form("FirstName")))>0 Then
       Response.Write "Your name is " & Trim(Request.Form("FirstName"))
       Response.Write "<br />Thanks!!"
    Else
       Response.Write "You forgot to enter your name"
       call showform()
   End If
   %>