package com.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToXmlItem {
	public File excelFile;
	public File xmlFile;
	public ExcelToXmlItem(File excelFile) {
		super();
		this.excelFile = excelFile;
	}
	
	public void excuteChange(String xmlDir)
	{
		String filePath = xmlDir + "\\" + excelFile.getName();
		InputStream is = null;  
		Workbook workbook = null;
		try {
			is = new FileInputStream(this.excelFile);
			if (filePath.endsWith(".xls")) {  
                workbook = new HSSFWorkbook(is);  
            } else if (filePath.endsWith(".xlsx")) {  
                workbook = new XSSFWorkbook(is);  
            }  
			filePath = xmlDir + "\\" + excelFile.getName().split(".xls")[0] + ".xml";
			this.xmlFile = new File(filePath);
//			int sheetsNumber = workbook.getNumberOfSheets();
			int sheetsNumber = 1;
			
			for (int n = 0; n < sheetsNumber; n++) {
				Sheet sheet = workbook.getSheetAt(n);
				Object value = null;
				Row row = null;
				Cell cell = null;
				String[] colNames = null;
				//0.不使用，备注信息 1.服务器使用 2.客户端使用 3.通用
				int[] userTypes = null;
				int[] dataTypes = null;
				int rowLen = sheet.getPhysicalNumberOfRows();
				int colLen;
				StringBuffer xmlStr = new StringBuffer("<?xml version='1.0' encoding='utf-8'?>\r<root>\r");
				for (int i = sheet.getFirstRowNum(); i <= rowLen; i++) { // 从第二行开始读取
					row = sheet.getRow(i);
					if (row == null) {
						continue;
					}
					colLen = row.getLastCellNum();
					if(i == 1) {
						colNames = new String[row.getLastCellNum()];
					}
					else if(i == 2)
					{
						userTypes = new int[row.getLastCellNum()];
					}
					else if(i == 3) {
						dataTypes = new int[row.getLastCellNum()];
					}
					if(i > 3){
						xmlStr.append("\t<element ");
					}
					for (int j = row.getFirstCellNum(); j <= colLen; j++) {
						
						cell = row.getCell(j);
						if (cell == null) {
							continue;
						}
						if(i == 0)
						{
							//描述
							value = getCellValue(cell,-1);
						}
						else if(i == 1)
						{
							//列名
							value = cell.getStringCellValue();
							colNames[j] = String.valueOf(value);
						}
						else if(i == 2)
						{
							//客户端服务器用到类型
							value = (int)cell.getNumericCellValue();
							userTypes[j] = (int)value;
						}
						else if(i == 3)
						{
							value = (int)cell.getNumericCellValue();
							dataTypes[j] = (int)value;
						}
						else {
							value = getCellValue(cell,dataTypes[j]);
							//0.不使用，备注信息 1.服务器使用 2.客户端使用 3.通用
							if(userTypes[j] == 2 || userTypes[j] == 3) {
								xmlStr.append(colNames[j] + "='" + value + "' ");
							}
						}
//						System.out.print(value+" ");
					}
					if(i > 3)
					{
						xmlStr.append("/>\r");
					}
//					System.out.println(xmlStr);
				}
				xmlStr.append("</root>");
				this.saveXmlFile(xmlStr);
//				System.out.println(xmlStr);
			}
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		finally {
			IOUtils.closeQuietly(workbook);
			IOUtils.closeQuietly(is);
		}
		
		
	}
	
	/**
	 * 获取excel 单元格数据
	 * 
	 * @param cell
	 * @param dataType 0.数字 1.字符串 2.数组 3.长整型
	 * @return
	 */
	private Object getCellValue(Cell cell,int dataType) {
		
		Object value = null;
		switch (cell.getCellTypeEnum()) {
		case _NONE:
			break;
		case STRING:
			value = cell.getStringCellValue();
			break;
		case NUMERIC:
			value = (int)cell.getNumericCellValue();
			break;
		case BOOLEAN:
			value = cell.getBooleanCellValue();
			break;
		case BLANK:
			//value = ",";
			break;
		default:
			value = cell.toString();
		}
		//0.数字 1.字符串 2.数组 3.长整型
		switch (dataType) {
		case 0:
			return (int)value;
		case 1:
			return String.valueOf(value);
		case 2:
			return String.valueOf(value);
		case 3:
			return (Long)value;
		default:
			return String.valueOf(value);
		}
	}
	/**
	 * 保存xml
	 * @param sb
	 */
	private void saveXmlFile(StringBuffer sb) {
		FileOutputStream out;
		try {
			out = new FileOutputStream(this.xmlFile);
			OutputStreamWriter outWrite = new OutputStreamWriter(out, "utf-8");
			outWrite.write(sb.toString());
			outWrite.flush();
			outWrite.close();
			System.out.println(this.xmlFile.getName() + " -----success");
		} catch (Exception e) {
			// TODO: handle exception
			System.out.println(this.xmlFile.getName() + " -----error");
			e.printStackTrace();
		}
	}
	
}
