package com.util;

import java.io.File;
import java.util.ArrayList;

public class ExcelToXmlUtil {
//	private static final DecimalFormat DECIMAL_FORMAT_PERCENT = new DecimalFormat("##.00%");//��ʽ���ֱȸ�ʽ�����治��2λ����0����
	
//	private static final DecimalFormat df_per_ = new DecimalFormat("0.00%");//��ʽ���ֱȸ�ʽ�����治��2λ����0����,����0.00,%0.01%
	
//	private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd"); // ��ʽ�������ַ���
	
//	private static final FastDateFormat FAST_DATE_FORMAT = FastDateFormat.getInstance("yyyy/MM/dd");
	
//	private static final DecimalFormat DECIMAL_FORMAT_NUMBER  = new DecimalFormat("0.00E000"); //��ʽ����ѧ������
// 
//	private static final Pattern POINTS_PATTERN = Pattern.compile("0.0+_*[^/s]+"); //С��ƥ��
	
	
	
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
