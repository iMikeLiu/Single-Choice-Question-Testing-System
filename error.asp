<html>
<%
Session("error")=true
%>
<script>
alert("您的信息不正确（验证码、学号、姓名、班级）！");
function sleep(delay) {
  var start = (new Date()).getTime();
  while ((new Date()).getTime() - start < delay) {
    continue;
  }
 
}
sleep(100);
 location.href="index.asp"

</script>

</html>