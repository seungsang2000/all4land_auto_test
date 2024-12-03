package egovframework.kss.main.model;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
public class LayerData {
	private int order;
	private String layerName;
	private String layerEnglishName;
	private String url1;
	private String url2;
	private String url3;
	private String XMLUrl;
	private String JSONUrl;
	private String note;

	// 생성자
	public LayerData(int order, String layerName, String layerEnglishName, String url1, String url2, String url3, String XMLUrl, String JSONUrl, String note) {
		this.order = order;
		this.layerName = layerName;
		this.layerEnglishName = layerEnglishName;
		this.url1 = url1;
		this.url2 = url2;
		this.url3 = url3;
		this.XMLUrl = XMLUrl;
		this.JSONUrl = JSONUrl;
		this.note = note;
	}
}
