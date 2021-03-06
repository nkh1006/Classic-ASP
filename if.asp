<%	
   dim a, b
   a = true,
   b = 1

   if a = true then 
		response.write "true"
	else
		response.write "false"
	end if

   if b > 0 then 
		response.write "1"
	elseif a <> true then
		response.write "2"
   else
      response.write "3"
	end if   
%>   