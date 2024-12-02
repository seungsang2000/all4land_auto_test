<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>엑셀 파일 업로드</title>
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script>
$(function(){
    $("form").on("submit", function(event){
        event.preventDefault(); // 기본 submit 이벤트를 막습니다.
        
        var formData = new FormData(this); // 폼 데이터 가져오기

        $.ajax({
            url: "/uploadExcel.do",  // 서버의 업로드 처리 URL
            type: "POST",
            data: formData,
            processData: false,  // 파일 데이터 전송 시 false로 설정
            contentType: false,  // contentType을 자동으로 설정하게끔
            xhrFields: {
                responseType: 'blob'  // 응답을 Blob으로 설정
            },
            success: function(response) {
                // 서버에서 반환된 엑셀 파일을 다운로드 처리
                var blob = new Blob([response], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                var link = document.createElement("a");
                link.href = URL.createObjectURL(blob);
                link.download = "테스트 결과.xlsx"; // 다운로드 파일명
                link.click();
            },
            error: function(xhr, status, error) {
                alert("파일 업로드 실패: " + error);
            }
        });
    });
});
</script>
</head>
<body>
<h3>엑셀 업로드</h3>
    <!-- 엑셀 파일 업로드 폼 -->
    <form name="popForm" method="post" enctype="multipart/form-data">
         <p><label for="excelFile">파일찾기</label></p>

         <p><input name="excelFile" id="excelFile" type="file" size="30"></p>

        <!-- 업로드 버튼 -->
        <div class="btn-area">
            <button type="submit" class="btn btn-yellow btn-ok">업로드</button>
            <button type="button" onclick="self.close();" class="btn btn-yellow btn-cancel">창닫기</button>
        </div>
    </form>
</body>
</html>