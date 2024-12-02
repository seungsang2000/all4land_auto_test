package egovframework.kss.main.controller;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import egovframework.kss.main.model.LayerData;
import egovframework.kss.main.service.ExcelService;

@Controller
public class ExcelController {

	@Autowired
	private ExcelService excelService;

	@PostMapping("/uploadExcel.do")
	public void uploadExcel(@RequestParam("excelFile") MultipartFile file) {
		try {
			if (file == null || file.isEmpty()) {
				System.out.println("파일이 전달되지 않았습니다.");
				return;
			}
			System.out.println("파일 이름: " + file.getOriginalFilename());
			// 파일 처리
			List<LayerData> dataList = excelService.parseExcelFile(file);
			System.out.println(dataList); // 데이터를 출력해서 확인

		} catch (Exception e) {
			System.out.println("파일 처리 중 오류 발생: " + e.getMessage());
			e.printStackTrace();
		}
	}
}
