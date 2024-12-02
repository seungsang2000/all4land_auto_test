<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>개방해 OpenAPI 테스트</title>
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script>
$(function(){
    $("form").on("submit", function(event){
        event.preventDefault(); 
        
        var formData = new FormData(this); 

        $.ajax({
            url: "/uploadExcel.do", 
            type: "POST",
            data: formData,
            processData: false,  
            contentType: false,  
            xhrFields: {
                responseType: 'blob'  
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
<h3>개방해 OpenAPI 테스트</h3>
    <!-- 엑셀 파일 업로드 폼 -->
    <form name="popForm" method="post" enctype="multipart/form-data">
         <p><label for="excelFile">파일찾기</label></p>

         <p><input name="excelFile" id="excelFile" type="file" size="30"></p>

        <!-- 업로드 버튼 -->
        <div class="btn-area">
            <button type="submit" class="btn btn-yellow btn-ok">테스트</button>
            <button type="button" onclick="self.close();" class="btn btn-yellow btn-cancel">창닫기</button>
        </div>
    </form>
</body>
</html>