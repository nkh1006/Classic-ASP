<%
   ' classic asp에서의 mssql dbconnect 내용
   ' 해당 아이피 주소를 넣어주고 

   dim dbIp = "xx.xxx.xxx.xxx,port" 
   dim dbName = "sampleDb" 
   dim dbId = "samplId" 
   dim dbPwd = "samplePwd" 

   ' db connect
   Set db = Server.CreateObject("ADODB.Connection")
   strconnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID="&dbId&";Initial Catalog="&dbName&";Data Source="&dbIp&";Password="&dbPwd
   db.Open strconnect

   ' query 날리기
   query = "SELECT * FROM tbl_name"

   ' rs에 넣어줌 rs는 record set인데 아무거나 다른 이름으로 해도 됨
   set rs = server.createobject("adodb.recordset") 

   ' 정상적으로 실행
   rs.open query,db,1 

   ' 데이터가 있으면 false, 없으면 true
   if Not rs.eof then
      ' 타언어에도 하나씩 있는 기본 출력문
      response.write "no record" 
      else
      ' query리 문의 결곽값 안에있는 필드명을 적어준다.
      data1 = rs("field_name1")
      data2 = rs("field_name2")
      response.write data1, data2       
   end if 

   rs.close 
%>