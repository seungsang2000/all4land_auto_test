package egovframework.kss.main.service;

import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import egovframework.kss.main.model.LayerData;

@Service
public class ExcelService {

	@SuppressWarnings("resource")
	public List<LayerData> parseExcelFile(MultipartFile file) throws Exception {
		List<LayerData> dataList = new ArrayList<>();

		if (file.isEmpty()) {
			throw new IllegalArgumentException("파일이 비어 있습니다.");
		}

		try (InputStream inputStream = file.getInputStream()) {
			Workbook workbook;

			// 파일 확장자에 따라 XSSFWorkbook (xlsx) 또는 HSSFWorkbook (xls) 사용
			if (file.getOriginalFilename().endsWith(".xlsx")) {
				try {
					if (inputStream.available() == 0) {
						throw new IllegalArgumentException("파일이 비어 있습니다.");
					}
					workbook = new XSSFWorkbook(inputStream);  // .xlsx 파일 처리
				} catch (Exception e) {
					throw new IllegalArgumentException("엑셀 파일(.xlsx)을 파싱하는 데 실패했습니다. 파일이 손상되었을 수 있습니다.");
				}
			} else if (file.getOriginalFilename().endsWith(".xls")) {
				try {
					workbook = new HSSFWorkbook(inputStream);  // .xls 파일 처리
				} catch (Exception e) {
					throw new IllegalArgumentException("엑셀 파일(.xls)을 파싱하는 데 실패했습니다. 파일이 손상되었을 수 있습니다.");
				}
			} else {
				throw new IllegalArgumentException("유효하지 않은 파일 형식입니다. 엑셀 파일을 업로드해 주세요.");
			}

			Sheet sheet = workbook.getSheetAt(0);  // 첫 번째 시트 선택

			// 데이터는 3번째 행 (인덱스 2)부터 시작
			for (int i = 2; i < sheet.getLastRowNum() && i < 12; i++) {
				Row row = sheet.getRow(i);
				if (row == null)
					continue;

				String layerName = getCellValue(row.getCell(3));  // D열
				String layerEnglishName = getCellValue(row.getCell(4));  // E열
				String url1 = getCellValue(row.getCell(10));  // K열
				String url2 = getCellValue(row.getCell(11));  // L열
				String url3 = getCellValue(row.getCell(12));  // M열

				// URL1 테스트 및 결과 저장
				url1 = !url1.isEmpty() ? (callApi(url1) ? "O" : "X") : "입력값 없음";

				// URL2 테스트 및 결과 저장
				url2 = !url2.isEmpty() ? (callApi(url2) ? "O" : "X") : "입력값 없음";

				// URL3 테스트 및 결과 저장
				url3 = !url3.isEmpty() ? (callApi(url3) ? "O" : "X") : "입력값 없음";

				System.out.println(layerName + "= WMS : " + url1 + ", WMS 이미지 : " + url2 + ", WFS : " + url3);

				dataList.add(new LayerData(i - 1, layerName, layerEnglishName, url1, url2, url3));
			}
		} catch (Exception e) {
			System.out.println("에러: " + e.getMessage());
			throw new IllegalArgumentException("엑셀 파일을 읽는 중 오류가 발생했습니다: " + e.getMessage());
		}

		return dataList;
	}

	private String getCellValue(Cell cell) {
		if (cell == null)
			return "";
		switch (cell.getCellType()) {
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				return String.valueOf(cell.getNumericCellValue());
			case BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case FORMULA:
				return cell.getCellFormula();
			default:
				return "";
		}
	}

	public static boolean callApi(String apiUrl) {
		HttpURLConnection connection = null;
		try {

			URL url = new URL(apiUrl);

			// HttpURLConnection 열기
			connection = (HttpURLConnection) url.openConnection();
			connection.setRequestMethod("GET");
			connection.setConnectTimeout(30000);
			connection.setReadTimeout(30000);

			// 요청 실행 및 응답 코드 확인
			int responseCode = connection.getResponseCode();
			return responseCode >= 200 && responseCode < 300;  // 2xx 응답 성공
		} catch (Exception e) {
			System.err.println("API 호출 실패: " + apiUrl + " - " + e.getMessage());
			return false;
		} finally {
			if (connection != null) {
				connection.disconnect();  // 연결 해제
			}
		}
	}
}
