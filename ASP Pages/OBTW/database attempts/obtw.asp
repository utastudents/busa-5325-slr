<% Option Explicit %>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Test</title>
</head>
<body>
<form action="" method=POST enctype="application/x-www-form-urlencoded">
<% 

      
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      'VARIABLE DECLARATION
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

      Dim objConn
      Dim strPath
      Dim strConn
      Dim objRecSetPlayed
      
      Dim arrType
      Dim strSQL
      Dim rsTest
      
      call sbOpen("OBTW","Played")

   
     

      'put playercard in a column to left of game cards
       prhtml("<table border = " & chr(34) & "1" & chr(34) & ">" )
          prhtml("<tr> <td>")
            
             Call CreatePlayerCard(3)
          prhtml("</td> <td> ... </td> <td>")
             Call CreateTable()
       prhtml("</td> </tr> </table>")
      prhtml("</table>")
      
    '==================== 'subroutines start here '======================  
      
    sub PrHtml(strHtml)
       response.write(strhtml)
    end sub      


      sub sbOpen(strDataBaseName,strTableName)
           dim strQuery
           Set objConn = CreateObject("ADODB.Connection") 
                 
           strPath = Server.MapPath(strdatabasename & ".accdb")
          ' PrHTML("Database Connection Path: " & strPath & "</br>")
           strConn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strPath & ";"
           objConn.Open strConn  
           set objRecSetPlayed = createobject("ADODB.Recordset")
      
           'strQuery = "SELECT * from " & strTableName & " where ID = " & "1" 
           strQuery = "SELECT * from " & strTableName 
        
           objRecSetplayed.open strQuery, objConn     
      end sub



      sub CreateTable()
           dim intRow
           dim intCol
           dim strQuery
           dim strRowName
           dim strCheckBoxDefinition 
           dim intChecked 
           dim strCheckedStatus 

           objrecsetPlayed.movefirst
           PrHTML("<Table border=" & chr(34) & "1" & chr(34) & ">")
              'First Row contains the Labels
               PrHTML("<tr>")
                  'First column of first Row is empty
                    PrHTML("<td> </td>")
                   'next columns are the player ids
                    For intCol = 1 to 5
                       PrHTML("<td> Player" & intcol & " </td>")
                    Next 
               PrHTML("</tr>")

           Do While NOT objRecSetPlayed.Eof
               PrHTML("<tr>")
              'first column are the row labels
                strRowName = objrecsetplayed("Label")
               PrHTML("<td> " & strRowName & "</td>")
              'the other columns indicate if a checkbox has been set 1 if yes 0 if no
               for intcol=1 to 5
                intchecked = objrecsetPlayed("player"& intcol)
                strCheckedStatus = "... "
                if intchecked = 1 then strcheckedstatus = "&#10004"

               ' PrHTML("<td>" & strcheckedstatus & " </td>")
                'strCheckBoxDefinition ="<input type=" & chr(34) & "checkbox" & chr(34) 
                'strCheckBoxDefinition = strCheckBoxDefinition & "name=" & chr(34) & "vehicle" &  chr(34) & " value=" 
                'strCheckBoxDefinition=strCheckBoxDefinition & chr(34) &  "Car" & chr(34) & strcheckedstatus & ">"
                PrHTML("<td>" & strcheckedstatus & " </td>")
                'example
                '	<input type="checkbox" name="vehicle" value="Car" checked>
              next 
              PrHTML("</tr>")
              objRecSetPlayed.MoveNext
           Loop


          PrHTML("</Table>")

      

      end sub

   sub CreatePlayerCard(intPlayerNumber)
           dim intRow
           dim intCol
           dim strQuery
           dim strRowName
           dim strCheckBoxDefinition 
           dim intChecked 
           dim strCheckedStatus 
           
           objrecsetplayed.close

           strQuery = "SELECT * from Played"  
        
           objRecSetplayed.open strQuery, objConn  
             
         
           objRecSetplayed.movefirst

           PrHTML("<Table border=" & chr(34) & "1" & chr(34) & ">")
              'First Row contains the Labels
               PrHTML("<tr>")
                  'First column is player's name
                   PrHTML("<td> Player" & intPlayerNumber & " </td>")
                     
               PrHTML("</tr>")

            prhtml("<tr> <td>")
            prhtml("<input type=" & Chr(34) & "radio"        & Chr(34) & _
                " checked=" & Chr(34) & "checked"      & Chr(34) & _
                  " name=" & Chr(34) & "Choice"   & Chr(34)  & _
                  " value=" & Chr(34) & 99 & Chr(34) & "> Choose one below <br>") 
            
            
           for introw =1 to 4
               'prhtml("<tr> <td>")
                strRowName = objrecsetplayed("Label")
                prhtml("<input type=" & Chr(34) & "radio"        & Chr(34) & _
                  " name=" & Chr(34) & "Choice"   & Chr(34)  & _
                  " value=" & Chr(34) & intRow & Chr(34) & "> " & strrowname & "<br>") 
                'prhtml("</td> </tr>")
              
              ' PrHTML("<tr>")
              'first column are the row labels
               ' strRowName = objrecsetplayed("Label")
               'PrHTML("<td> " & strRowName & "</td>")
            
               'the other column indicates if a checkbox has been set 1 if checked and  0 if no
                            
   
                'intchecked = objrecsetplayed("player"& intPlayerNumber)
                'strCheckedStatus = "... "
                'if intchecked = 1 then strcheckedstatus = "Checked"

               ' PrHTML("<td>" & strcheckedstatus & " </td>")
              '  strCheckBoxDefinition ="<input type=" & chr(34) & "checkbox" & chr(34) 
              '  strCheckBoxDefinition = strCheckBoxDefinition & "name=" & chr(34) & "vehicle" &  chr(34) & " value=" 
              '  strCheckBoxDefinition=strCheckBoxDefinition & chr(34) &  "Car" & chr(34) & strcheckedstatus & ">"
               ' PrHTML("<td>" & strCheckBoxDefinition & " </td>")
                'example
                '	<input type="checkbox" name="vehicle" value="Car" checked>
               
               'PrHTML("</tr>")
              objRecSetPlayed.MoveNext
          next 
          prhtml("</td> </tr>")
       PrHTML("</Table>"  ) 
   end sub

   objRecSetPlayed.Close
   Set objRecSetPlayed = nothing
   objConn.Close 
   Set objConn = nothing

%>


</body>
