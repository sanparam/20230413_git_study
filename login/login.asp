<%@ Language=VBScript %>
<% 
Response.Buffer = True
Response.Expires = 0
'---------------------------------------------------------------------
	'아이디와 패스워드 입력받는 부분과 처리부분을 나눈다.
	'(login_process.asp에서 처리) - cerry
'---------------------------------------------------------------------

url = request("path")

'Response.Write url


%>

<html>
<head>
<title>매크로21 - 사용자 확인</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style>
<!--
a:link {text-decoration:none; color:#666666; font-size: 9pt;}
a:visited {text-decoration:none; color:#666666; font-size: 9pt;}
a:hover {text-decoration:underline; color:#005699; font-size: 9pt;}
a:active {text-decoration:underline; color:#005699; font-size: 9pt;}

tr{	font-family: "verdana","Tahoma", "돋움", "돋움체";
	font-size: 9pt;
	line-height: 150%;
	color: #666666;
	letter-spacing: -1px;
}
-->
</style>
</head>

<script LANGUAGE="javascript">
<!--
function CheckForm(form)
{
	if (document.Trans.LogID.value == "")
	{
		alert ("[아이디]를  입력하셔야 합니다! ");document.Trans.LogID.focus ();return ;
	}
	if (document.Trans.Password.value == "")
	{
		alert ("[비밀번호]를  입력하셔야 합니다! ");document.Trans.Password.focus ();return;
	}
	document.Trans.submit ();
	//self.close();
}
//-->
</script>



<!--id/password분실-->
<script LANGUAGE="javascript">

function loseid()
{
win= window.open ('https://localhost/manage/member/member_pswd/loseid.asp','losewin','width=425,height=410,left=100,top=100,status=yes,toolbar=yes,scrollbars=no,resizable=0');
}
function Member_IN(){
	//alert("확인")
	opener.parent.content.location.href ="/manage/member_inSE_ke.asp" ;
	//opener.document.
	self.close()
}
	
</script>

<body bgcolor="#FFFFFF" onload="window.document.Trans.LogID.focus();" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="Trans" action="/login/login_process.asp" method="post" onsubmit="CheckForm(this);return false;">
<input type="hidden" name="path" value="<%=url%>">
<table width="520" border="0" cellspacing="0" cellpadding="0" height="300">
  <tr>
    <td background="images/bg.gif">
      <table width="520" border="0" cellspacing="0" cellpadding="0" height="300">
        <tr> 
          <td height="80">&nbsp;</td>
        </tr>
        <tr> 
          <td align="right" valign="top"> 
            <table width="320" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="30">&nbsp;</td>
              </tr>
              <tr> 
                <td height="30"><img src="images/icon01.gif" width="9" height="9"> 
                  아이디와 비밀번호를 입력하고 로그인 해 주세요.</td>
              </tr>
              <tr> 
                <td height="80"> 
                  <table width="250" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="100" align="right">회원 아이디&nbsp;&nbsp;&nbsp;&nbsp; 
                      </td>
                      <td width="100"> 
                        <input type="text" name="LogID" size="12" tabindex="1">
                      </td>
                      <td rowspan="2" width="50"><input type="image" border="0" name="imageField2" src="images/button-login.gif" width="46" height="46" tabindex="3"></td>
                    </tr>
                    <tr> 
                      <td width="100" align="right">비밀번호&nbsp;&nbsp;&nbsp;&nbsp; 
                      </td>
                      <td width="100"> 
                        <input type="password" name="Password" size="12" tabindex="2">
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td valign="bottom" align="right">
                  <table width="80%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="20"><img src="images/icon01.gif" width="9" height="9"><a href="/manage/member_pswd/loseid.asp">아이디, 
                        비밀번호 찾기</a></td>
                    </tr>
                    <tr>
                      <td height="20"><img src="images/icon01.gif" width="9" height="9"><a href="javascript:Member_IN();">회원가입</a></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>