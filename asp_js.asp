<%
   ' console.log
   function callConsoleLog(_value)
      response.write "<script>"
      response.write "console.log('"& _value &"');"
      response.write "</script>"
      response.end
   End Function 

   ' alert
   function callAlert(_value)
      response.write "<script>"
      response.write "alert('"& _value &"');"
      response.write "</script>"
      response.end
   End Function    
%>