<head>
<title>localhost_Macro21 - 디지털 산업단지</title>
<link rel="stylesheet" href="text.css" type="text/css" title="macro21 css">
<link rel="stylesheet" href="top.css" type="text/css">

<%
'Response.Write session("userlevel") & session("id")
'Response.End 
if session("code") <> "" then 

elseif session("id") = "" then 
	%>
	<script>
		window.open ("/login/login.asp?path=reload", "NewWindow",'Width=520,Height=300,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0');
	</script>
		
<% response.end %>
<%else
    if session("userlevel") = "7" then

		session("id") = "bbang"          '----> 한전 관리자 메뉴 
		session("userid") = "D000285"    '----> 수주 관리(견적/발주) 처리를 위해 
		session("DealerCode") = "0303"   '----> 제품 관리(입력/수정) 관련
	
	'elseif session("supply") = true then 
	
	elseif session("kepco") <> true then %>
		<script>
			alert('승인전 혹은 탈퇴한 회원이십니다.');
			//parent.frames("tops").location.reload();
			//window.navigate('/ke/index.asp');
			parent.location.href ="/login/logout.asp";
		</script>
		<% Response.End 
		
	end if 
 
end if 

main = request("main")%>
<div class="wrapper" >
  <div class="content">
	<frameset cols="170,*" border="0" align = "center">
		<%if session("code") <> "" then%>
		<frame src="menu1.asp" name="left" scrolling="auto" marginwidth="0" marginheight="0">
			<%if main = "m" then%>
			<frame src="/product/modify_form.asp?catalogno=<%=request("catalogno")%>" name="right" scrolling="auto" marginwidth="0" marginheight="0">
			<%else%>
			<frame src="/product/insert_form.asp" name="right" scrolling="auto" marginwidth="0" marginheight="0">
			<%end if%>
		<%else%>
		<frame src="menu.asp" name="left" scrolling="auto" marginwidth="0" marginheight="0">
			<%if main = "P" then%>
			<frame src="/point/event.asp" name="right" scrolling="auto" marginwidth="0" marginheight="0">
			<%elseif main = "T" then%>
			<frame src="/reorder.asp" name="right" scrolling="auto" marginwidth="0" marginheight="0">
			<%elseif main = "" then%>
			<frame src="main.asp" name="right" scrolling="auto" marginwidth="0" marginheight="0">
			<%end if%>
		<%end if%>
	</frameset>
	</div>
</div>
</html>