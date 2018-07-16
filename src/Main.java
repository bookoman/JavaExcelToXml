import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Properties;

import org.xml.sax.InputSource;

import com.util.ExcelToXmlUtil;

public class Main {
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Properties prop = new Properties();
//		InputStream in = Object.class.getResourceAsStream("config.properties");
		FileInputStream fileInput = null;
		try {
			fileInput = new FileInputStream(new File("config.properties"));
			prop.load(fileInput);
			
			String excelDir = prop.getProperty("excel_dir").trim();
			String xmlDir = prop.getProperty("xml_dir").trim();
			
			System.out.println("excleDir:" + excelDir);
			System.out.println("xmlDir:" + xmlDir);
			new ExcelToXmlUtil(excelDir, xmlDir);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
	
	
	

}
