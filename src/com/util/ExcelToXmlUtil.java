package com.util;

import java.io.File;
import java.util.ArrayList;

public class ExcelToXmlUtil {
//	private static final DecimalFormat DECIMAL_FORMAT_PERCENT = new DecimalFormat("##.00%");//格式化分比格式，后面不足2位的用0补齐
	
//	private static final DecimalFormat df_per_ = new DecimalFormat("0.00%");//格式化分比格式，后面不足2位的用0补齐,比如0.00,%0.01%
	
//	private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd"); // 格式化日期字符串
	
//	private static final FastDateFormat FAST_DATE_FORMAT = FastDateFormat.getInstance("yyyy/MM/dd");
	
//	private static final DecimalFormat DECIMAL_FORMAT_NUMBER  = new DecimalFormat("0.00E000"); //格式化科学计数器
// 
//	private static final Pattern POINTS_PATTERN = Pattern.compile("0.0+_*[^/s]+"); //小数匹配
	
	
	
	public ArrayList<String> excels;
	
	public File excelDicFile;
	public File xmlDicFile;
	
	
	public ExcelToXmlUtil(String excelDir,String xmlDir) {
		super();
		
		this.excelDicFile = new File(excelDir);
		this.xmlDicFile = new File(xmlDir);
		
		this.excuteChange(xmlDir);
	}
	public void excuteChange(String xmlDir)
	{
		
		if  (!excelDicFile .exists()  && !excelDicFile .isDirectory())    
        {     
            excelDicFile .mkdir();  
        }
		if  (!xmlDicFile .exists()  && !xmlDicFile .isDirectory())    
		{     
			xmlDicFile .mkdir();  
		}
		
		File[] files = this.excelDicFile.listFiles();
		String fileName;
		ExcelToXmlItem item;
		File file;
		for (int i = 0;i < files.length;i++) {
			file = files[i];
			fileName = file.getName();
			if(fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
				item = new ExcelToXmlItem(file);
				item.excuteChange(xmlDir);
			}
		}
		
	}

}
