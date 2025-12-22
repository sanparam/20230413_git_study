<%@ Language=VBScript %>
<!--#include virtual="/sql_injection.asp"-->
<% Response.Buffer = true
   Response.Expires = 0 
   
   msort=request("sel")
 
   ' request가 member_in_join.asp로 부터 넘어오지 않음 - 수정 필요...
   '==============================================================
	'한전-기업회원가입창
    '==============================================================
  
    SET Conn=Server.CreateObject("ADODB.Connection")
		Conn.Open Application("MACRO21_DSN_ConnectionString")
	
	
	
	
	UpTeaSql="select  dcode, dcodenm from dcode where mcode ='0003' and use_yn='Y'"
	UpjongSql="select  dcode, dcodenm from dcode where mcode ='0004' and use_yn='Y'"
	AreaSql="select  dcode, dcodenm from dcode where mcode ='0005' and use_yn='Y'"
	
	sql2="select gradecd, chummageamt as camt, assuranceamt as aamt, monthamt as mamt from dealergrade where gubun='D'"
   
	sql = "select sitecd,sitenm from bmark"
   
   lsortSql="SELECT LSORTCD,LSORTNM FROM LSORTCD WHERE use_yn = 'Y'"
   
%>
<html>
<head>
<title>MRO전용매장 회원가입</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="javascript" src="member_e.js"></script>
<link rel="stylesheet" href="../text.css" type="text/css">
<script LANGUAGE="javascript">
<!--
  function list_sch(li)
    { 
	 //alert(li);
	 document.Trans.sel_msortcd[0].text="k"
	 //window.location.href="member_inSP.asp?sel=" + li  ;
	 // alert("/search/advance_search.asp?sel=" + li );
    }
   
   function EnterCheck(i) 
   {
     if(event.keyCode ==13 && i==1) 
     { 
      document.Trans.Passwd.focus(); 
     }
     if(event.keyCode ==13 && i==2) 
     { 
      Trans.submit();
      //   alert("폼이 전송되었습니다");
     } 
   }

//-->
</script>
<script language="javascript">

function check()
{
  var str = document.Trans.IdentifyID2.value.length;
  if(str == 6)
    document.Trans.IdentifyID2_1.focus();
}
function encheck()
{
  var str = document.Trans.identifyid1.value.length;
  var str1 = document.Trans.identifyid1_2.value.length;
  if(str == 3)
    document.Trans.identifyid1_2.focus();
  if(str1 == 2)
    document.Trans.identifyid1_3.focus();
}    

  function click_email() {
		if ( document.Trans.e_mail.value == '') {
			document.Trans.e_mail.value = document.Trans.customid.value + "@mail.macro21.com";
		}
	
   }
  function supply()
  {
   window.location.href="../pds/contents.asp?num=17"
   self.close();
   }
  </script>

<script LANGUAGE="javascript">
<!--
      function list_sch2(li2) { 
           
	   window.location.href="member_inSP.asp?sel=" + li2  ;
	   // alert("/search/advance_search.asp?sel=" + li );
        }
   
//-->
</script>
<!--소스나라 학습예제http://www.sourcenara.com-->
<script language="javascript">
<!--
function charType(){
var str=document.Trans.customid.value;
var i;          
	  if(str.charAt(i) >= 'ㄱ' && str.charAt(i) <= 'ㅎ' || str.charAt(i) >= 'ㅏ' && str.charAt(i) <= 'ㅣ' || str.charAt(i) >= '가' && str.charAt(i) <= '항'){
          alert('한글은 입력하지 못합니다. 영문으로 입력하십시오.');
              }
 }  

//function re(){
//document.Trans.customid.value="";
//}

</script>
</head>

<body bgcolor="#FFFFFF" marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" onload="this.document.Trans.customid.focus();" onload="jusoCallBack('roadAddrPart1','addrDetail','zipNo');" >
<!--<body bgcolor="#FFFFFF" marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" onload="this.document.Trans.customid.focus();">-->

	
<script language="javascript">
// opener관련 오류가 발생하는 경우 아래 주석을 해지하고, 사용자의 도메인정보를 입력합니다. ("팝업API 호출 소스"도 동일하게 적용시켜야 합니다.)
document.domain = "localhost";

function  
jusoCallBack(roadAddrPart1,addrDetail,zipNo){
	document.getElementById('address1').value = roadAddrPart1;
	document.getElementById('address3').value = addrDetail;
	/*document.getElementById('roadAddrPart2').value = roadAddrPart2;
	document.getElementById('engAddr').value = engAddr;
	document.getElementById('jibunAddr').value = jibunAddr; */
	document.getElementById('Post').value = zipNo;
	//let addressEl = document.querySelector("#address1"); // id를 찾는것이다.
	//addressEl.value = roadAddrPart1;
	//let addressE2 = document.querySelector("#address3"); // id를 찾는것이다.
	//addressE2.value = addrDetail;
	//let addressE3 = document.querySelector("#Post"); // id를 찾는것이다.
	//addressE3.value = zipNo;
	//**document.form.roadFullAddr.value = roadFullAddr;
	//document.formTrans.address1.value = roadAddrPart1;
	//document.formTrans.address3.value = addrDetail;
	//document.form.roadAddrPart2.value = roadAddrPart2;
	//document.form.engAddr.value = engAddr;
	//document.form.jibunAddr.value = jibunAddr;
	//document.formTrans.Post.value = zipNo; 
	/**document.form.admCd.value = admCd;
	document.form.rnMgtSn.value = rnMgtSn;
	document.form.bdMgtSn.value = bdMgtSn;
	document.form.detBdNmList.value = detBdNmList;
	/**2017년 2월 추가 제공 **/
	/**document.form.bdNm.value = bdNm;
	document.form.bdKdcd.value = bdKdcd;
	document.form.siNm.value = siNm;
	document.form.sggNm.value = sggNm;
	document.form.emdNm.value = emdNm;
	document.form.liNm.value = liNm;
	document.form.rn.value = rn;
	document.form.udrtYn.value = udrtYn;
	document.form.buldMnnm.value = buldMnnm;
	document.form.buldSlno.value = buldSlno;
	document.form.mtYn.value = mtYn;
	document.form.lnbrMnnm.value = lnbrMnnm;
	document.form.lnbrSlno.value = lnbrSlno;
	/**2017년 3월 추가 제공 **/
	/**document.form.emdNo.value = emdNo;*/
	//document.formTrans.submit();
 }

function goPopup(){
	// 주소검색을 수행할 팝업 페이지를 호출합니다.
	// 호출된 페이지(jusopopup.asp)에서 실제 주소검색URL(https://business.juso.go.kr/addrlink/addrLinkUrl.do)를 호출하게 됩니다.
	var pop = window.open("/asp_sample/jusoPopup.asp","pop","width=570,height=420, scrollbars=yes, resizable=yes"); 
	
	// 모바일 웹인 경우, 호출된 페이지(jusopopup.asp)에서 실제 주소검색URL(https://business.juso.go.kr/addrlink/addrMobileLinkUrl.do)를 호출하게 됩니다.
    //var pop = window.open("/jusoPopup.asp","pop","scrollbars=yes, resizable=yes"); 
}


</script>

<table width="750" border="0" cellspacing="15" cellpadding="0">
  <tr valign="top" align="center"> 
    <td class="txt-33"> 
      <table width="700" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="700" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="images/memberin_icon06.gif" width="500" height="26"></td>
              </tr>
              <tr bordercolor="#E8E8E8"> 
                <td class="txt-33"> 
                  <table width="700" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="20">&nbsp;</td>
                      <td width="7000" class="txt-33">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td class="txt-33">회원가입신청을 해주셔서 감사합니다.<br><br>
					  <font color=red><b>Macro21은 MRO B2B 전문회사로서, 기업용 사무소모품 및 각종 공구류등의 산업기자재를 통합구매 및 판매합니다.</b></font><br><br>
					  
					  <!--font color=red>MRO 전용매장에서는 업무용 전산문구류 및 각종 공구류등의 산업기자재를 종합 판매합니다.</font><br>
					  <br-->
					  <font color=red><b> 첫 거래이신 경우, 추가하여 <a href="mailto:webmaster@macro21.com"><font color ="red"><b>E-mail</b></font></a> 혹은 FAX(031-479-7439)로 사업자등록증을 보내주시기 바랍니다.<br>
					  </b></font><br>
					  <font color=red>*</font> 가 붙은 곳은 <font color=red>필수항목</font>으로 반드시 기재하여 주십시오.
					  </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td height="30"> 
                  <hr width="700" size="1" align="center">
                </td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="1" cellpadding="3" class="txt-33">
              <form name="Trans" id="form" action="member_inPG_SU.asp" method="post" onsubmit="CheckForm(this);return false;">
				 <input type='hidden' name='who' value='enterprise'>
                 <INPUT TYPE="HIDDEN" NAME="GUBUN" VALUE="G">
                 <INPUT TYPE="HIDDEN" NAME="STATUS" VALUE="A">
                <tr valign="bottom" bgcolor="#c0defa" align="center"> 
                  <td class="txt-33" colspan="4"><b>[ MRO 전용매장 회원가입신청 ]</b></td>
                </tr>
                <tr valign="bottom"> 
                  <td class="txt-33" colspan="4" height="30"><b>【 기본 정보 】</b></td>
                </tr>
                <tr> 
                  <td bgcolor="#c0defa" width="100"><font face="Courier New, Courier, mono">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">*&nbsp;</font>ID</font></td>
                  <td colspan="3" bgcolor="#EBEBEB"> 
                    
                    <input type="text" name="customid" size="10" maxlength="10" onKeypress="charType()" onKeyup="charType()" STYLE="ime-mode:inactive">
                    <a href="javascript:IdCheck_win(document.Trans.customid.value)"><img src="images/id_check.gif" border="0" WIDTH="100" HEIGHT="21" align="absmiddle"></a>
                    *영문,숫자포함 <b>10자리이내.</b> </td>
                </tr>
                <tr> 
                  <td bgcolor="#c0defa" width=100>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">*&nbsp;</font>비밀번호</td>
                  <td class="body" bgcolor="#EBEBEB" width=200> 
                    <input type="password" name="Passwd" size="8" maxlength="8">
                  </td>
                  <td width="200" align="center" bgcolor="#c0defa"><font color="#FF0000">*&nbsp;</font>비밀번호확인 
                  </td>
                  <td bgcolor="#EBEBEB" width=100>
                    <input type="password" name="RePasswd" size="8" maxlength="8">
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#c0defa">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">*&nbsp;</font>회사</td>
                  <td colspan="3" bgcolor="#EBEBEB"> 
                    <input type="text" name="customnm" size="25 maxlength="30" STYLE="ime-mode:active">
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#c0defa" class="bodybrown">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">*&nbsp;</font>전화</td>
                  <td bgcolor="#EBEBEB" class="bodybrown"> 
                    <input name="FTelNO1" maxlength="3" size="3" ONKEYPRESS="valuecheck()">
                    &nbsp; 
                    <input name="MTelNO1" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
                    - 
                    <input name="LTelNO1" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
                  </td>
                  <td bgcolor="#c0defa" class="bodybrown" align="center" width="100"><font face="Courier New, Courier, mono">FAX</font></td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input name="FTelNO2" maxlength="3" size="3" ONKEYPRESS="valuecheck()">
					&nbsp; 
				  <input name="MTelNO2" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
					- 
				  <input name="LTelNO2" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
                  </td>
                </tr>
                <!--<tr> 
                  <td bgcolor="#c0defa" align="center"><font color="#FF0000">*</font>사업자등록번호</td>
                  <td colspan="3" bgcolor="#EBEBEB"> 
                    <input name="identifyid1" maxlength="3" size="3" OnKeyPress="encheck()" ONKEYPRESS="valuecheck()">
                - 
                <input type="text" name="identifyid1_2" size="2" maxlength="2" OnKeyPress="encheck()" ONKEYPRESS="valuecheck()">
                -
                <input type="text" name="identifyid1_3" size="5" maxlength="5" ONKEYPRESS="valuecheck()">
                  </td>
                </tr>-->
				<tr> 
                  <!--<td bgcolor="#c0defa">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">*&nbsp;</font>소재지</td>
                  <td colspan="3" bgcolor="#EBEBEB"> 
                 <% if session("userlevel") = "7" then %>   <a href="javascript:AddrChk_win()"><img src="images/zipcode.gif" border="0" WIDTH="85" HEIGHT="21" align="absmiddle"></a>
					<input name="Post" maxlength="7" size="7" maxlength="7" readonly>우편번호 &nbsp;&nbsp;* 우편번호검색버튼을 이용해주세요.<br> 
					<% end if %>
					<input name="Address1" size="35" maxlength="60" ><br>
					<% if session("userlevel") = "7" then %>
					<input name="Address3" size="35" maxlength="60">* 동 이하번지는 이곳에 입력해주세요.<br>
					<% else %>
					<input name="Address3" size="35" maxlength="60">* 도로명 주소를 이곳에 입력해주세요.<br>
					<% end if %>
                  </td>-->
				  <td bgcolor="#c0defa"align="center"><font color="#FF0000">*</font>사업장소재지</td>
				 
                  <td colspan="3" bgcolor="#EBEBEB"> 					
						<a href="javascript:goPopup();"><img src="images/zipcode.gif" border="0" WIDTH="85" HEIGHT="21" align="absmiddle"></a>
						<!--<input type="button" value="주소검색" onclick="goPopup();">-->
						<input type="text" name="Post" maxlength="7" size="7" maxlength="7" id="Post" readonly ONKEYPRESS="valuecheck()">&nbsp;* 우편번호검색버튼을 이용해주세요.<br>
						<input type="text" name="address1" size="35" maxlength="60" id="address1" readonly ONKEYPRESS="valuecheck()"><br>
						<input type="text" name="address3" size="35" maxlength="60" id="address3" ONKEYPRESS="valuecheck()">* 동 이하번지는 이곳에 입력해주세요.<br>
								
						  <!--<form name="rdnAddr">
							<!--<input id ="roadFullAddr">-->
							<!--<input id ="roadAddrPart1">
							<input id ="addrDetail">
							<!--<input id ="roadAddrPart2">
							<input id ="engAddr">
							<input id ="jibunAddr">-->
							<!--<input id ="zipNo">
							<!--<input id ="admCd">
							<input id ="rnMgtSn">
							<input id ="bdMgtSn">
							<input id ="detBdNmList">-->
							<!--2017년 2월 추가 제공-->
							<!--<input id ="bdNm">
							<input id ="bdKdcd">
							<input id ="siNm">
							<input id ="sggNm">
							<input id ="emdNm">
							<input id ="liNm">
							<input id ="rn">
							<input id ="udrtYn">
							<input id ="buldMnnm">
							<input id ="buldSlno">
							<input id ="mtYn">
							<input id ="lnbrMnnm">
							<input id ="lnbrSlno">
							<input id ="emdNo">-->	
						<!--</form>-->
	
				  </td>
				
                </tr>
				<tr> 
                  <td bgcolor="#c0defa" class="bodybrown">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="Courier New, Courier, mono">Homepage</font></td>
                  <td colspan="3" bgcolor="#EBEBEB" class="body"> 
                    <input type="text" name="homepage" size="35" maxlength="60" value="http://">
                     &nbsp; <br>
                   
                  </td> </tr>
				  <tr>
                  <td bgcolor="#c0defa" class="bodybrown">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="Courier New, Courier, mono"><b>공급사 등록신청</b></font></td>
                  <td colspan="3" bgcolor="#EBEBEB" class="body"> 
                     <!--//<input type="radio" onclick="supply();">-->
                     <input type="radio" onClick="location.href='mailto:webmaster@macro21.com'">
                     &nbsp; 공급사 가입신청의 경우, 별도의 절차가 있습니다. <font color="#FF0000">&nbsp;체크&nbsp;</font>하여 주십시오.<br>
                   
                  </td> 
				  
                </tr>
                <tr valign="bottom"> 
                  <td class="bodybrown" height="30" colspan="4"><b>【 기타 정보 】 
                    </b></td>
                </tr>
                <tr> 
                  <td bgcolor="#c0defa" class="bodybrown">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"></font>담당자 성함</td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input type="text" name="name2" size="15" maxlength="30" STYLE="ime-mode:active">
                  </td>
                  <td bgcolor="#c0defa" class="bodybrown" align="center"><font color="#FF0000"></font><font face="Courier New, Courier, mono">E-mail</font></td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input type="text" name="e_mail" size="25" maxlength="30" STYLE="ime-mode:inactive">
                  </td>
                </tr>
                
                <input type=hidden name=identifyID2 value="111111">
                <input type=hidden name=identifyID2_1 value="1111119">
                <input type=hidden name=buse value="*">
                <input type=hidden name=position value="*">
                <input type=hidden name=FTTELNO1 value="111">
                <input type=hidden name=MTTELNO1 value="111">
                <input type=hidden name=LTTELNO1 value="1111">
                <input type=hidden name=FTTELNO2 value="111">
                <input type=hidden name=MTTELNO2 value="111">
                <input type=hidden name=LTTELNO2 value="1111">

              <!--  <input type=hidden name=post value="111-111">-->
                 <input type=hidden name=identifyID1 value="000">
                <input type=hidden name=identifyID1_2 value="00">
                <input type=hidden name=identifyID1_3 value="00000">
               <!-- <input type=hidden name=Address1 value="미기입">
                <input type=hidden name=Address3 value="X">-->
                <input type=hidden name=upteanm value="*">
                <input type=hidden name=upjongnm value="*">
                
                
                <!--tr> 
                  <td bgcolor="#c0defa" class="bodybrown" align="center"><font color="#FF0000">*</font>이름</td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input type="text" name="name2" size="15" maxlength="30">
                  </td>
                  <td bgcolor="#c0defa" class="bodybrown" align="center" width="100"><font color="#FF0000">*</font>주민등록번호</td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input type="text" name="IdentifyID2" size="6" MAXLENGTH="6" OnKeyPress="check()" ONKEYPRESS="valuecheck(window.document.Trans.IdentifyID2.value)">
                    - 
                    <input type="text" name="IdentifyID2_1" size="7" MAXLENGTH="7" ONKEYPRESS="valuecheck()">
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#c0defa" class="bodybrown" align="center"><font color="#FF0000">*</font>부서</td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input type="text" name="buse" size="15" maxlength="30">
                  </td>
                  <td bgcolor="#c0defa" class="bodybrown" align="center" width="100"><font color="#FF0000">*</font>직급</td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input type="text" name="position" size="15" maxlength="20">
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#c0defa" class="bodybrown" align="center"><font color="#FF0000">*</font>전화번호</td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input name="FTTelNO1" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
                    &nbsp; 
                    <input name="MTTelNO1" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
                    - 
                   <input name="LTTelNO1" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
                  </td>
                  <td bgcolor="#c0defa" class="bodybrown" align="center" width="100"><font color="#FF0000">*</font>휴대전화</td>
                  <td bgcolor="#EBEBEB" class="body"> 
                    <input name="FTTelNO2" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
					&nbsp; 
				  <input name="MTTelNO2" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
					- 
				  <input name="LTTelNO2" maxlength="4" size="4" ONKEYPRESS="valuecheck()">
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#c0defa" class="bodybrown" align="center"><font color="#FF0000">*</font><font face="Courier New, Courier, mono">E-mail</font></td>
                  <td bgcolor="#EBEBEB" class="body" colspan="3"> 
                    <input type="text" name="e_mail" size="35" maxlength="60">
                  </td>
                </tr-->
                
                <!--tr> 
                  <td bgcolor="#CCCC99" class="bodybrown" align="center">e-mail 서비스 
                  </td>
                  <td bgcolor="#EBEBEB" class="bodybrown" colspan='3'> 
                    <input type="checkbox" name="mail_check" VALUE="1" onclick="javascript:click_email();">
                    신청 &nbsp;&nbsp;* 부여받는 메일 주소는 [신청한 ID]@mail.macro21.com입니다. 
                    <a href="/macro_help/email_help.htm">도움말</a>
                  </td>
                </tr--> 
                <tr>
                  <td class="body" colspan="4" bgcolor="#c0defa">&nbsp;</td>
                </tr>
                <!--tr valign="bottom" bgcolor="#F6F6F6"> 
                  <td class="bodybrown" colspan="4" height="30"><b>【 주관심 분야 】</b></td>
                </tr>
               
                 
                <% '====================================================================
					' 관심분야 선택 폼 수정
				'=======================================================================
				%>
                <tr>
				  <td align=center colspan="4" class="body">
					  <table border=0 cellpadding=1 cellspacing=0>
						
					    <tr bgcolor=#2b4671>
						  <td colspan="4" > 
						    <table border=0 cellpadding=5 cellspacing=0>
						    <% set rs=Conn.Execute (lsortSql)
						    i = 0 %>
						    <% do while not rs.EOF %>
								<% if i mod 3 = 0 then %>
									<% if i = 0 then %>
										<tr>
									<% else %>
										</tr>
										<tr>
									<% end if %>  
								<% end if %>
							  <td bgcolor=#CCCC99>
								<input type=checkbox name ="lsortcd" value="<%=rs("lsortcd")%>"></td>
							  <td bgcolor="#F6F6F6" width=180 class="body"><%=rs("lsortnm") %></td>
							   
							   <%i=i+1 
								rs.MoveNext ()
							      loop
							      rs.close
                         set rs=nothing %>
							  </tr>
							</table>
						  </td>
						</tr>
					  </table>
					</td>
				</tr>
				 <% '====================================================================
					' 관심분야 선택 폼 수정 끝
				'=======================================================================
				%>
				 <tr valign="bottom" bgcolor="#F6F6F6"> 
					<td class="bodybrown" colspan="4" height="30"><b>【 즐겨찾는 페이지 】최대 3개까지 선택 가능</b></td>
						
				</tr>
				<tr>
				  <td align=center colspan="4" class="body">
					  <table border=0 cellpadding=1 cellspacing=0>
						
					    <tr bgcolor=#2b4671>
					      <td colspan="4" >
					       <table border=0 cellpadding=5 cellspacing=0>
						    <% set rs2=Conn.Execute (sql)
						    i = 0 %>
						    <% do while not rs2.EOF %>
								<% if i mod 3 = 0 then %>
									<% if i = 0 then %>
										<tr>
									<% else %>
										</tr>
										<tr>
									<% end if %>  
								<% end if %>
							  <td bgcolor=#CCCC99>
								<input type=checkbox name ="sitecd" value="<%=rs2("sitecd")%>"></td>
							  <td bgcolor="#F6F6F6" width=180 class="body"><%=rs2("sitenm") %></td>
							   
							   <%i=i+1 
								rs2.MoveNext ()
							      loop
							      rs2.close
                         set rs2=nothing %>
							  </tr>
							</table>
						  </td>
						</tr>
					  </table>
					</td>
				</tr>
				
                 <tr> 
                  <td bgcolor="#CCCC99" class="bodybrown" align="center">뉴스레터신청
                  </td>
                  <td bgcolor="#EBEBEB" class="bodybrown" colspan="3"> 
                    <input type="checkbox" name="newsletter" VALUE="Y">
                    신청 &nbsp; * 매크로21에서 제공하는 다양한 정보를 메일로 받아보시려면 신청해주세요</td>
                </tr>
                <tr> 
                  <td class="body" colspan="4" bgcolor="#CCCC99">&nbsp;</td>
                </tr-->
                <tr valign="bottom"> 
                  <td colspan="4" height="30"><b>【 약 관 】</b></td>
                </tr>
                <tr align="center" bgcolor="#EBEBEB"> 
                  <td colspan="4"> 
                 <textarea name="textfield14" cols="100" rows="10" class="body" readonly><!--#include file="M_stipulation_txt.htm" -->
				 
				 </textarea>
                    <br>
                    <input type="radio" name="radio" checked>
                    동의함 
                    <input type="radio" name="radio">
                    동의안함 </td>
                </tr>
               <tr>
				  <td align=right class="body" bgcolor="#F6F6F6" colspan="4"><a href="http://www.macro21.com/manage/이용약관.zip"><img src="images/down.gif" width=35 height=29 alt="" border="0"><font color=#24c4e6><b>::회원 이용약관 다운로드::</b></font></a></td>
				</tr>
                <tr align="center"> 
                 <td class="body" colspan="4" height="30" bgcolor="#c0defa">
                  <input class="text" type="submit" value="[&nbsp;&nbsp;등록하기&nbsp;&nbsp;]" id="submit1" name="submit1">
                  <input class="text" type="button" onclick="history.back(-1);" value="[&nbsp;&nbsp;취소하기&nbsp;&nbsp;]"  name="submit3">
                </td>
                </tr>
                <!--tr align="center" bgcolor="#FFFFFF"> 
                   <td class="body" colspan="4" height="60" bgcolor="#ffffff" align=center valign="middle"> 
                  <a href="member_contract.htm">회원 이용약관</a> | <a href="member_info.htm">개인정보보호정책</a> | <a href="../../../aboutus/aboutus_main.asp">회사소개</a> <br><hr width="600" size="1" align="right">
                </td>
                </tr-->
				<tr align="center" bgcolor="#FFFFFF"> 
                  <td class="body" colspan="4"><hr width="700" size="1" align="right"></td>
                </tr>
                <tr align="center" bgcolor="#FFFFFF"> 
                  <td class="body" colspan="4"><!--#include file="../include/macro21.htm"--></td>
                </tr>
              </form>
            </table>
          </td>
  </tr>
</table>

</table>
</body>
</html>
