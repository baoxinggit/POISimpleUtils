package com.bx.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Exell表格导出类
 * 
 * @author 暴兴
 */
public class ExellExportUtil {
	/**
	 * 新建Excel文件
	 * @param path
	 * @param fileName
	 * @param sheetName
	 * @param flag
	 */
	private static Workbook createExell(String path,String fileName,String sheetName,boolean flag) {
		Workbook wb = null;
		InputStream is = null;
		if(flag){
			try {
				File file = new File(path + File.separator + fileName + ".xlsx");
				if(file.exists()){
					is = new FileInputStream(file);
					wb = WorkbookFactory.create(is);
					if(wb.iterator().hasNext()){
						return wb;
					} 
				}
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (EncryptedDocumentException e) {
				e.printStackTrace();
			} catch (InvalidFormatException e) {
				e.printStackTrace();
			} finally {
				try {
					if(is != null)
						is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		wb = new XSSFWorkbook();
		wb.createSheet(sheetName);
		return wb;
	}

	/**
	 * 将对象导出到文件
	 * 
	 * @param fileName
	 *            文件名
	 * @param sheetName
	 *            表名
	 * @param cellName
	 *            列名
	 * @param path
	 *            路径
	 * @param clazz
	 *            对象的类型
	 * @param list
	 *            需要导出的对象的集合
	 * @param flag 
	 * 			  是否是在后面继续添加
	 */
	public static void writeToFile(String fileName, String sheetName,
			String[] cellName, String path, Class<?> clazz, List<?> list,
			boolean flag) {
		Workbook wb = createExell(path,fileName,sheetName,flag);
		Sheet sheet = wb.getSheetAt(0);
		Row row = null;
		int rows = 0;
		if(flag){
			rows = sheet.getLastRowNum();
		}
		row = sheet.createRow(0);
		for (int i = 0; i < cellName.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(cellName[i].toString());
		}
		Field[] fields = clazz.getDeclaredFields();
		for (int i = rows; i < list.size() + rows; i++) {
			Row row2 = sheet.createRow(i + 1);
			try {
				for (int j = 0; j < fields.length; j++) {
					Cell cell = row2.createCell(j);
					fields[j].setAccessible(true);
					cell.setCellValue(fields[j].get(list.get(i - rows)).toString());
					fields[j].setAccessible(false);
				}
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
		}
		write(wb, path, fileName);
	}
	
	/**
	 * 将对象写出
	 * @param wb
	 * @param path
	 * @param fileName
	 */
	private static void write(Workbook wb, String path, String fileName) {
		OutputStream os = null;
		try {
			os = new FileOutputStream(path + File.separator + fileName
					+ ".xlsx");
			wb.write(os);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (os != null)
					os.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
