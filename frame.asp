<%@ Language=VBScript %>
<%

main_url = request("main_url")
'Response.Write main_url


'if session("userlevel") = "7" then
		'Response.Write "test"

		'session("id") = "bbang"          '----> 한전 관리자 메뉴 
		'session("userid") = "D000285"    '----> 수주 관리(견적/발주) 처리를 위해 
		'session("DealerCode") = "0303"   '----> 제품 관리(입력/수정) 관련
		
'elseif session("kepco") <> true then %>
		<script>
			//alert('회원가입하셔서 통합서비스를 받으시기 바랍니다');
			//parent.frames("tops").location.reload();
			//window.navigate('/ke/index.asp');s
			//parent.location.href ="/offer/write_guest.asp";
		</script>
		<% 'Response.End 
		
'end if 
 %>

<html>
<head>
	<title> (주)매크로21 - MRO 전용매장</title>
</head>

<frameset rows="81,*" border="0">
<frame src="index_top.asp" name="tops" scrolling="auto" marginwidth="0" marginheight="0">
<frame src="<%=main_url%>" name="content" scrolling="auto" marginwidth="0" marginheight="0">
</frameset>

</html>
