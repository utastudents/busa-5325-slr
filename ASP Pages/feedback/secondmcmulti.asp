<html >
<head>

<title>My Online Survey</title>

</head>

<%
    Dim cid
    dim ans
    dim f, fs,fname

   'Check for correct variable lengths and type
    cid = Request.Form("cid")
    ans =  Request.Form("ans")
    
   
     if len(ans) <>0 then
        select case ans
         case "1", "2", "3", "4", "5"
         case else
           ans="0"
        end select
        ans = chr(64 + ans)
     else
        ans="missing"
        response.write("Please go back and choose an answer. It is missing at the moment. <br> ")
     end if

    if isnumeric(cid)=false then 
      response.write("Please go back and enter a computer number. It is missing or not a number at the moment. <br> ")
      cid="missing"
    else

       if (csng(cid)-cint(cid) <> 0 ) or (cint(cid)>72) or (cint(cid) < 1)  then 
          response.write( "<br> You entered " & cid & " for your computer number. <br>")
          response.write( "The computer number should be an integer from 1 to 72. <br>")
          response.write( " Please go back and enter a computer number with the correct value. <br> <br>")
          cid = "missing"
       end if
    end if


    if (ans<>"missing") and (cid <> "missing") then 
         Response.write("Your entered a computer number of " & cid & " and an answer of " & ans & "<br> <br>"  ) 
    end if

   'Send responses to text files
 

  
   if (ans<>"missing") and (cid <>"missing") then
     set fs=Server.CreateObject("Scripting.FileSystemObject" )
     set f=fs.OpenTextFile(Server.MapPath("/faculty/eakin/asps/feedback/SentItems/" & cid & ".txt"),8,true)
     f.write(cid)
     f.write(" ")
     if ans="" then ans="."
     f.writeline(ans)
     f.Close
     response.write("The value " & cid & " and the value " & ans & " were written <br> <br>")
     set f=Nothing
     set fs=Nothing
   end if
%>

Thank You <br>

</body>

</html>