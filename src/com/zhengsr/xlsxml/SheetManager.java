package com.zhengsr.xlsxml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import javax.swing.text.html.parser.Entity;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


import com.zhengsr.xlsxml.bean.CusRow;
import com.zhengsr.xlsxml.read.ReadXlsToXml;
/**
 * @author zhengshaorui 2018/6/24
 */
public class SheetManager {

	// 这两个要一一对应
	private static final String[] LANGUAGE_NAMES = new String[] { 
		"简体中文/繁中","繁体中文/繁中", "English", "Czech",
		"Danish", "Dutch", "Spanish","Finnish",
		"Portuguese", "French", "Deutsch", "Greek",
		"Italiano/Italian","日语/Japanese", "Norwegian", "Polski/Polish", 
		"Romanian", "Russian", "Swedish","Turkish",
		"Arabic","Chinese (Simple)","Chinese (Traditional)","Hungarian",
		"Thai","Persian","Vietnam/Vietnamese","Korea/Korean",
		"Deutsch (German)"};

	private static final String[] LANGUAGE_FLODERS = new String[] {
		"values-zh-rCN", "values-zh-rTW", "values", "values-cs-rCZ",
		"values-da-rDK", "values-nl", "values-es", "values-fi-rFI",
		"values-pt", "values-fr", "values-de", "values-el-rGR",
		"values-it", "values-ja-rJP", "values-nb-rNO", "values-pl-rPL",
		"values-ro-rRO", "values-ru-rRU", "values-sv-rSE", "values-tr-rTR",
		"values-ar","values-zh-rCN","values-zh-rTW","values-hu-rHU",
		"values-th-rTH","values-fa","values-vi-rVN","values-ko-rKR",
		"values-de"};

	private static Set<Entry<String, String>> LANGMAP;
	private static Set<Entry<String, String>> FLOADER;
	private boolean isLastItemString = false;
	private String mStringKey = "";

	private static class Holder {
		static SheetManager INSTANCE = new SheetManager();
	}

	public static SheetManager getInstance() {
		return Holder.INSTANCE;
	}

	private SheetManager() {
		int length = LANGUAGE_NAMES.length;
		Map<String, String> language_map = new HashMap<>();
		Map<String, String> floder_map = new HashMap<>();
		for (int i = 0; i < length; i++) {
			language_map.put(LANGUAGE_NAMES[i], LANGUAGE_FLODERS[i]);
			floder_map.put(LANGUAGE_FLODERS[i], LANGUAGE_NAMES[i]);
			LANGMAP = language_map.entrySet();
			FLOADER = floder_map.entrySet();
		}
	}

	/**
	 * 获取xml最大行数
	 * 
	 * @param sheet
	 * @return
	 */
	public int getMaxRow(Sheet sheet) {
		int firstrow = sheet.getFirstRowNum();
		int lastrow = sheet.getLastRowNum();
		int num = 0;
		for (int i = firstrow; i < lastrow; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				num++;
			}
		}
		return num;
	}
	/**
	 * 获取当前列的最大行数
	 * @return
	 */
	public int getCellMaxRow(Sheet sheet,int cIndex){
		int firstrow = sheet.getFirstRowNum();
		int lastrow = sheet.getLastRowNum();
		int num = 0;
		for (int i = firstrow; i < lastrow; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				Cell cell = row.getCell(cIndex);
				if(cell != null){
					num++;
				}
			}
		}
		return num;
	}

	/**
	 * 通过语言，获取对应的文件夹名称
	 * @param value
	 * @return
	 */
	public String getFolderName(String value) {
		for (Map.Entry<String, String> entry : LANGMAP) {
			if (entry.getKey().contains(value.trim())) {
				return entry.getValue();
			}
		}
		return null;
	}
	
	/**
	 * 通过语言，获取对应的文件夹名称
	 * @param value
	 * @return
	 */
	public String getLanduageName(String flodername) {
		for (Map.Entry<String, String> entry : FLOADER) {
			if (entry.getKey().equals(flodername.trim())) {
				return entry.getValue();
			}
			
		}
		return null;
	}
	
	/**
	 * 创建文件夹
	 * @param path
	 * @return
	 */
	public String createFloder(String rootpath,String path) {
		String[] paths = path.split("/");
		int length = paths.length;
		String currentPath = rootpath;
		for (int i = 0; i < length; i++) {
			String dir = paths[i];
			File file = new File(currentPath, dir);
			if (!file.exists()) {
				file.mkdir();
			}
			currentPath = file.getAbsolutePath();
		}
		return currentPath;
	}
	
	/**
	 * 创建文件，然后把 string 的通用格式写上
	 * @param path
	 */
	public void createFileAndData(String path){
		File file = new File(path,ReadXlsToXml.STRING_NAME);
		if (file.exists()) {
			file.delete();
		}
		FileOutputStream fos = null;
		try {
			file.createNewFile();
			fos = new FileOutputStream(file);
			StringBuilder sb = new StringBuilder();
			sb.append(
					"<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n"
							+ "<resources>").append("\r\n");
			

			fos.write(sb.toString().getBytes("utf-8"));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (fos != null) {
				try {
					fos.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}
	
	/**
	 * 把值写进字符串中
	 * @param key string 中的key
	 * @param value string 中的 value
	 * @param maxRow 行的最大值
	 * @param index 列
	 * @param dir  文件夹的路径
	 * @param ignoreRow 要忽略的行，即不要翻译的行
	 */
	public void writeValueToString(int cIndex,Cell cell,int maxRow, int rowIndex,
			String dir,int ignoreRow) {
		String value = "";
		if(cell != null){
			value = cell.toString();
			if(cIndex == 0){ //先保留key值
				mStringKey = value;
				
			}else{ //然后把数据写进去
				int type = cell.getCellType();
				//日期需要特殊处理
				if(type == HSSFCell.CELL_TYPE_NUMERIC){
					SimpleDateFormat sdf;
					if (cell.getCellStyle().getDataFormat() == HSSFDataFormat  
	                        .getBuiltinFormat("h:mm")) {  
	                    sdf = new SimpleDateFormat("HH:mm");  
	                } else {// 日期  
	                    sdf = new SimpleDateFormat("yyyy-MM-dd");  
	                }  
	                java.util.Date date =  cell.getDateCellValue();  
	                value = sdf.format(date);  
				}else{
					value = cell.toString();
				}
				File file = new File(dir, ReadXlsToXml.STRING_NAME);
				if (file.exists()) {
					FileOutputStream fos = null;
					try {
						fos = new FileOutputStream(file,true);
						StringBuilder builder = new StringBuilder();
						builder.append("\t")
								// 有个小空格
								.append("<string name=\"").append(mStringKey)
								.append("\">").append(value).append("</string>")
								.append("\r\n");
	
						if ((maxRow - ignoreRow) == rowIndex) { 
																	
							builder.append("</resources>\r\n");
						}
						
						fos.write(builder.toString().getBytes("utf-8"));
					} catch (Exception e) {
						e.printStackTrace();
					} finally {
						if (fos != null) {
							try {
								fos.close();
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
						}
					}
				}
			}
		}
		
		
		
	}
	
	
	
	/**
	 * 写数据到 array 中
	 * @param cIndex 列
	 * @param cell 单元格
	 * @param maxCellNum 最大的列数
	 * @param maxRowNum 最大的行数
	 * @param rowIndex 行
	 * @param pathMap 文件夹的map
	 * @param ignoreRow 要忽略的行，即不要翻译的行
	 */
	public void writeValueToArray(int cIndex,Cell cell,int maxCellNum,int maxRowNum,int rowIndex,
			Map<Integer, String> pathMap,int ignoreRow){
		String value = "";
		StringBuilder sb = null;
		if(cell != null){
			value = cell.toString();
			
			
			if(cIndex == 0){  //第一列
				//第一次我们先把这个
				sb = new StringBuilder();
				
				if(!isLastItemString){
					sb.append("\t")
					  .append("<string-array=\"").append(value).append("\">");
					
					for (int i = 2; i < maxCellNum; i++) {
						//把能获取到文件夹的列，写入 array 字符串
						writeDatatoFile(i,sb.toString(),pathMap);
					}
					
				}else{
					
					//补全上一次
					sb = new StringBuilder();
					sb.append("\n")
					  .append("\t")
					  .append("</string-array>");
					for (int i = 2; i < maxCellNum; i++) {
						//把能获取到文件夹的列，写入 array 字符串
						writeDatatoFile(i,sb.toString(),pathMap);
					}
					System.out.println("两次: "+isLastItemString);
					sb = new StringBuilder();
					sb.append("\n")
					  .append("\t")
					  .append("<string-array=\"").append(value).append("\">");
					for (int i = 2; i < maxCellNum; i++) {
						//把能获取到文件夹的列，写入 array 字符串
						System.out.println("sd: "+pathMap.get(i));
						writeDatatoFile(i,sb.toString(),pathMap);
					}
					isLastItemString = false;
				}
			}else if(cIndex>1){
				//写数据
				sb = new StringBuilder();
				sb.append("\n")
				  .append("\t\t")
				  .append("<item>")
				  .append(value)
				  .append("</item>");
				
				
				isLastItemString = true;
				 	
				writeDatatoFile(cIndex, sb.toString(),pathMap);
				
				if(rowIndex == maxRowNum - ignoreRow
						&& cIndex == maxCellNum - 1){
					sb = new StringBuilder();
					sb.append("\n")
					  .append("\t")
					  .append("</string-array>\r\n")
					  .append("</resources>\r\n");
					for (int i = 2; i < maxCellNum; i++) {
						//把能获取到文件夹的列，写入 array 字符串
						writeDatatoFile(i,sb.toString(),pathMap);
					}
				}
			}
		}
	}
	

	public List<CusRow> parseStringXml(String path,String stringName,CreationHelper createHelper) {
		List<CusRow> lists = new ArrayList<>();
        try {
        	File file = new File(path,stringName); //以 strings为标准，如果客制化字符串，可以在这里改
        	if (file.exists()) {
        		 InputStream is = new FileInputStream(file);
        		 DocumentBuilder documentBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        		 Document document = documentBuilder.parse(is);
                 NodeList nodeList = document.getElementsByTagName("string");
                 int length = nodeList.getLength();
                 for (int i = 0; i < length; i++) {
					//String name = nodeList.item(i).getAttributes().getNamedItem("name").getNodeName();
                	 CusRow cusRow = new CusRow();
                	 cusRow.key = nodeList.item(i).getAttributes().getNamedItem("name").getNodeValue();
                	 cusRow.value= nodeList.item(i).getTextContent();
                	 lists.add(cusRow);
				}
			}
           
        } catch (Exception e) {
            e.printStackTrace();
        }
        return lists;
    }
	
	private  void writeDatatoFile(int cIndex,String value,Map<Integer, String> pathMap){

		File file = new File(pathMap.get(cIndex), ReadXlsToXml.STRING_NAME);
		if (file.exists()) {
			FileOutputStream fos = null;
			
			try {
				fos = new FileOutputStream(file,true);
				fos.write(value.getBytes("utf-8"));
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				if (fos != null) {
					try {
						fos.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		}
	}
	

}
