/**
 * @author DADAO
 * 写的规则遵循逐页逐行逐列
 */
package com.tops.common.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelUtil {
	private OutputStream _ous = null;
	
	private Workbook _wb = null;
	private Sheet _sheet = null;
	private Row _row = null;
	private Cell _cell = null;
	private CellStyle _cs = null;
	private Font _font = null;
	
	public ExcelUtil(String pathname){
		try {
			_ous = new FileOutputStream(new File(pathname));
			_wb = new HSSFWorkbook();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	
	public void setSheetName(String sheetName){
		int sheetId = _wb.getSheetIndex(_sheet);
		_wb.setSheetName(sheetId, sheetName);
	}
	
	public Sheet getSheetAt(int sheetId){
		_sheet = _wb.getSheetAt(sheetId);
		return _sheet;
	}
	
	public Row getRow(int rowId){
		_row = _sheet.getRow(rowId);
		return _row;
	}
	
	public Cell getCell(int cellId){
		_cell = _row.getCell(cellId);
		return _cell;
	}
	
	public Sheet createSheet(){
		_sheet = _wb.createSheet();
		return _sheet;
	}
	
	public Row createRow(int rowId){
		_row = _sheet.createRow(rowId);
		return _row;
	}
	
	public void createCell(int cellId, String content){
		_cell = _row.createCell(cellId);
		_cell.setCellValue(content);
	}
	
	public CellStyle getCellStyle(){
		return _cs;
	}
	
	public Font getFont(){
		return _font;
	}
	
	public void createCellStyle(){
		_cs = _wb.createCellStyle();
		_font = _wb.createFont();
	}
	
	public void setCellStyle(){
		_cs.setFont(_font);
		_cell.setCellStyle(_cs);
	}
	
	public void clearCellStyle(){
		_font = null;
		_cs = null;
	}
	
	public void clearFont(){
		_font = null;
	}
	
	public void write() throws IOException{
		_wb.write(_ous);
	}
	
	public void close(){
		try {
			if(_wb!=null) _wb.close();
			if(_ous!=null) _ous.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public int getLastRowNum(){
		return _sheet.getLastRowNum();
	}
	
	public int getLatColNum(){
		return _row.getLastCellNum();
	}
	
	public void setColWidth(int col, int width){//列宽是按照字符个数计算的
		_sheet.setColumnWidth(col, width*256);
	}
	
	public void setRowHight(int height){
		_row.setHeightInPoints((float)(height/2));
	}
	
	public void addMergedRegion(int startRow, int endRow, int startCol, int endCol){
		_sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startCol, endCol));
	}
	
	public void setBorder(){
		_cs.setBorderBottom(BorderStyle.THIN);
		_cs.setBorderLeft(BorderStyle.THIN);
		_cs.setBorderRight(BorderStyle.THIN);
		_cs.setBorderTop(BorderStyle.THIN);
	}
	
	public void setAlignment(){
		_cs.setAlignment(HorizontalAlignment.CENTER);
		_cs.setVerticalAlignment(VerticalAlignment.CENTER);
	}

}
