package com.zhengsr.xlsxml.write;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import com.zhengsr.xlsxml.SheetManager;
import com.zhengsr.xlsxml.bean.CusRow;
import com.zhengsr.xlsxml.bean.XlsWriteBean;
import com.zhengsr.xlsxml.method.WriteXlsManager;
import com.zhengsr.xlsxml.read.ReadXlsToXml;


/**
 * @author zhengshaorui 2018/6/24
 */
public class WriteXmlToXls {
	private static String VALUE_PATH = "test";  //value文件夹存放的路径
	private static String ROOT_PATH; // 当前路径
	private static String XLS_NAME = "workbook.xlsx"; //要生成的名字

	
	public static void main(String[] args) {
		File file = new File("");
		ROOT_PATH = file.getAbsolutePath();
		XlsWriteBean bean = new XlsWriteBean.Builder()
				.setRootPath(ROOT_PATH)
				.setFileFloderName(VALUE_PATH)
				.setXlsName(XLS_NAME)
				.builder();
		WriteXlsManager.getInstance().startWrite(bean.getBuilder());
		System.out.println("在 "+ROOT_PATH+File.separator+VALUE_PATH+"生成文件啦!!");
	}


}	
