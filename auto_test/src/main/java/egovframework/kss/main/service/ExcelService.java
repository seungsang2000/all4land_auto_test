package egovframework.kss.main.service;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

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

			Row headerRow = sheet.getRow(1); // 헤더 뽑아오기
			String layerNameHeader = getCellValue(headerRow.getCell(3)); // D열
			String layerEnglishHeader = getCellValue(headerRow.getCell(4)); // E열
			String url1Header = getCellValue(headerRow.getCell(10)); // K열
			String url2Header = getCellValue(headerRow.getCell(11)); // L열
			String url3Header = getCellValue(headerRow.getCell(12)); // M열

			if (!layerNameHeader.equals("명칭") || !layerEnglishHeader.equals("레이어명") || !url1Header.equals("WMS 예시") || !url2Header.equals("WMS 이미지") || !url3Header.equals("WFS")) { // 액셀 파일 형식 검사
				throw new Exception("액셀 파일의 양식이 다릅니다"); // 액셀 헤더가 양식과 다른 경우, 예외처리
			}

			// 데이터는 3번째 행 (인덱스 2)부터 시작
			for (int i = 2; i < sheet.getLastRowNum() && i < 12; i++) {
				boolean comma = false; // 쉼표 찍을 건지
				Row row = sheet.getRow(i);
				if (row == null)
					continue;

				String layerName = getCellValue(row.getCell(3));  // D열
				String layerEnglishName = getCellValue(row.getCell(4));  // E열
				String url1 = getCellValue(row.getCell(10));  // K열
				String url2 = getCellValue(row.getCell(11));  // L열
				String url3 = getCellValue(row.getCell(12));  // M열
				String note = "";

				if (layerName.isEmpty() || layerEnglishName.isEmpty()) { //레이어의 명칭이나 레이어명이 없는 항목은 테스트하지 않고 넘어감. url 틀릴 경우 대비
					if (layerName.isEmpty()) {
						note += "레이명 명칭 없음";
						comma = true;
					}

					if (layerEnglishName.isEmpty()) {
						if (comma) {
							note += ",";
						}
						note += "레이어명 없음";
					}
					dataList.add(new LayerData(i - 1, layerName, layerEnglishName, "", "", "", note));
					continue;
				}

				// URL1 테스트 및 결과 저장
				url1 = !url1.isEmpty() ? (callWMS(url1, layerEnglishName) ? "O" : "X") : "입력값 없음";

				// URL2 테스트 및 결과 저장
				url2 = !url2.isEmpty() ? (callWMSImage(url2) ? "O" : "X") : "입력값 없음";

				// URL3 테스트 및 결과 저장
				url3 = !url3.isEmpty() ? (callWFS(url3, layerEnglishName) ? "O" : "X") : "입력값 없음";

				System.out.println(layerName + "= WMS : " + url1 + ", WMS 이미지 : " + url2 + ", WFS : " + url3);

				dataList.add(new LayerData(i - 1, layerName, layerEnglishName, url1, url2, url3, note));
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

	/*public boolean callApi(String apiUrl) {
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
			return responseCode >= 200 && responseCode < 300;  // 2xx 응답 성공 -> 이걸로 안된다. 정규식으로 바꿀 것 
		} catch (Exception e) {
			System.err.println("API 호출 실패: " + apiUrl + " - " + e.getMessage());
			return false;
		} finally {
			if (connection != null) {
				connection.disconnect();  // 연결 해제
			}
		}
	}*/

	//WMS 호출 검사 코드
	public boolean callWMS(String apiUrl, String layerCode) {
		HttpURLConnection connection = null;
		try {

			URL url = new URL(apiUrl);

			// HttpURLConnection 열기
			connection = (HttpURLConnection) url.openConnection();
			connection.setRequestMethod("GET");
			connection.setConnectTimeout(20000);
			connection.setReadTimeout(20000);

			// 요청 실행 및 응답 코드 확인
			int responseCode = connection.getResponseCode();
			if (responseCode != HttpURLConnection.HTTP_OK) {
				return false;
			}
			BufferedReader in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
			StringBuilder response = new StringBuilder();
			String inputLine;
			while ((inputLine = in.readLine()) != null) {
				response.append(inputLine);
			}
			in.close();

			// 응답 본문 리턴
			String WMS_response = response.toString();

			String regex = "^var OtmsWmsLayer = new ol\\.source\\.TileWMS\\(\\{url:'http://www\\.khoa\\.go\\.kr/oceanmap/wmsdata\\.do', serverType:'mapserver', transition: 0, params:\\{ServiceKey:'[^']+',LAYERS:'" + Pattern.quote(layerCode) + "'\\}\\}\\);$";
			Pattern pattern = Pattern.compile(regex);
			Matcher matcher = pattern.matcher(WMS_response);

			// 매칭된 결과 반환
			return matcher.matches();

		} catch (Exception e) {
			System.err.println("API 호출 실패: " + apiUrl + " - " + e.getMessage());
			return false;
		} finally {
			if (connection != null) {
				connection.disconnect();  // 연결 해제
			}
		}
	}

	//WMS 이미지 호출 검사 코드
	public boolean callWMSImage(String apiUrl) { //후에 수정해야 할거 : 이미지 파일을 가져오되, 레이어가 찍혀있는지...?

		HttpURLConnection connection = null;
		try {

			URL url = new URL(apiUrl);

			// HttpURLConnection 열기
			connection = (HttpURLConnection) url.openConnection();
			connection.setRequestMethod("GET");
			connection.setConnectTimeout(20000);
			connection.setReadTimeout(20000);

			// 요청 실행 및 응답 코드 확인
			int responseCode = connection.getResponseCode();
			if (responseCode != HttpURLConnection.HTTP_OK) {
				return false;
			}

			// Content-Type 헤더 확인
			String contentType = connection.getContentType();
			if (contentType != null && contentType.startsWith("image/png")) {
				return true; // png 파일임
			} else {
				return false; // png 파일이 아님
			}

		} catch (Exception e) {
			System.err.println("API 호출 실패: " + apiUrl + " - " + e.getMessage());
			return false;
		} finally {
			if (connection != null) {
				connection.disconnect();  // 연결 해제
			}
		}
	}

	// WFS 호출. 정규식이 아닌 XML 파싱 사용
	public boolean callWFS(String apiUrl, String layerCode) {
		HttpURLConnection connection = null;
		try {

			DocumentBuilderFactory dbFactoty = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactoty.newDocumentBuilder();
			Document doc = dBuilder.parse(apiUrl);

			doc.getDocumentElement().normalize();
			System.out.println("Root element: " + doc.getDocumentElement().getNodeName());

			NodeList nList = doc.getElementsByTagName("mapprime:" + layerCode.toLowerCase()); // xml 파싱시 대소문자 구분하므로, 소문자로 변환
			System.out.println("파싱할 리스트 수 : " + nList.getLength()); // 이거 이용해볼까? 0개 이상일때 성공시켜도 될거 같은데.

			if (nList.getLength() > 0) {
				return true;
			} else {
				return false; // 이거도 후에 수정
			}

		} catch (Exception e) {
			System.err.println("API 호출 실패: " + apiUrl + "-" + e.getMessage());
			return false;
		} finally {
			if (connection != null) {
				connection.disconnect();
			}
		}

	}

}
