package com.tops.common.util;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class TestExcel {
	
	public static void writeExcel(){
		ExcelUtil xls = new ExcelUtil("E:/1.xls");
		try {
			xls.createSheet();
			xls.setSheetName("第一页");
			for(int i = 0; i <= 10; i++){
				xls.setColWidth(i, 20);
			}
			
			for(int i = 0; i <= 5; i++){
				xls.createRow(i);
				xls.setRowHight(40);
				for(int j = 0; j <= 10; j++){
					xls.createCell(j, "第"+i+"行:第"+j+"列.");
					xls.createCellStyle();
					xls.setAlignment();
					xls.setBorder();
					xls.setCellStyle();
					xls.clearCellStyle();
				}
			}
			
			xls.createSheet();
			xls.setSheetName("第二页");
			xls.createRow(0);
			xls.createCell(0, "你好哇,李银河.");
			
			xls.createCellStyle();
			Font font = xls.getFont();
			font.setFontHeightInPoints((short)24);
			xls.setCellStyle();
			xls.clearCellStyle();
			
			xls.createRow(1);
			
			for(int j = 0; j <= 10; j++){
				xls.createCell(j, "");
				xls.createCell(j, "第"+j+"列");
				xls.createCellStyle();
				xls.setAlignment();
				xls.setBorder();
				xls.setCellStyle();
				xls.clearCellStyle();
			}
			
			xls.addMergedRegion(1, 1, 0, 3);
			xls.addMergedRegion(1, 2, 4, 4);
			xls.addMergedRegion(1, 2, 5, 6);
			
			xls.createRow(2);
			for(int j = 0; j <= 10; j++){
				xls.createCell(j, "");
				xls.createCell(j, "第"+j+"列");
				xls.createCellStyle();
				xls.setAlignment();
				xls.setBorder();
				xls.setCellStyle();
			}
			
			
			xls.createRow(3);
			for(int i = 4; i <= 10; i++){
				xls.createRow(i);
				for(int j = 0; j <= 10; j++){
					xls.createCell(j, "第"+i+"行:第"+j+"列.");
				}
			}
			xls.write();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			xls.close();
		}
	}
	
	public static void main(String[] args) throws Exception {
//		writeExcel();
		readExcel();
	}
	
	@SuppressWarnings("resource")
	public static void readExcel(){
		
	    try {
	    	InputStream ins = new FileInputStream("E:/1.xls");
			HSSFWorkbook wb = new HSSFWorkbook(ins);
			Sheet sheet = wb.getSheetAt(0);
			
			for(int i = 0; i <= sheet.getLastRowNum(); i++){
				Row row = sheet.getRow(i);
				for(int j = 0; j < row.getLastCellNum(); j++){
					Cell cell = row.getCell(j);
					String str = getCellValue(cell);
					System.out.println(str);
				}
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static String getCellValue(Cell cell){
		Object result = "";
		switch (cell.getCellTypeEnum()) {
		case STRING:
			result = cell.getStringCellValue();
			break;
		case NUMERIC:
			result = cell.getNumericCellValue();
			break;
		case BOOLEAN:
			result = cell.getBooleanCellValue();
			break;
		case FORMULA:
			result = cell.getCellFormula();
			break;
		case ERROR:
			result = cell.getErrorCellValue();
			break;
		case BLANK:
			break;
		default:
			break;
		}
		return result.toString();
	}
	
/*	private void createActExcel(ActionContext ac, Workbook wb, String org, String year, String mon) throws Exception{
		TagBiz tbz = BizFactory.getBiz(TagBiz.class, ac);
		ArrayList<TagPO> taglst = tbz.listTag();
		ArrayList<CameraVO> data = getBookActData(ac);
		OrgCache oc = CacheFactory.getCache(OrgCache.class);
		
		int curRow = 0;//创建行
		int k = 3;//监控功能类型--即标签起始位置
		int tagLength = taglst.size();
		
		Sheet sheet = wb.createSheet();
		wb.setSheetName(0, "变动表");
		
		Row row = sheet.createRow(curRow++);//第一行-标题
		CellStyle cs = wb.createCellStyle();
		Font font = wb.createFont();
		
		font.setFontHeightInPoints((short)24);//字体
		row.setHeightInPoints((short)34);//行高
		
		cs.setFont(font);
		cs.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
		cs.setAlignment(HorizontalAlignment.CENTER);//水平居中
		
		Cell cell = row.createCell(0);
		cell.setCellValue("视频监控变动表");
		cell.setCellStyle(cs);
		
		row = sheet.createRow(curRow++);//第二行-信息
		cs = wb.createCellStyle();
		font = wb.createFont();
		cs.setFont(font);
		cs.setVerticalAlignment(VerticalAlignment.CENTER);
		cs.setAlignment(HorizontalAlignment.CENTER);
		
		cell = row.createCell(0);
		cell.setCellValue("报表单位:");
		
		cell = row.createCell(1);
		cell.setCellValue(oc.getOrgName(org));
		
		cell = row.createCell(k + tagLength + 2);
		cell.setCellValue("建账月份:");
		
		cell = row.createCell(k + tagLength + 3);
		cell.setCellValue(year + "年" + mon + "月");
		
		row = sheet.createRow(curRow++);//第三行-表头
		cs = wb.createCellStyle();
		font = wb.createFont();
		cs.setFont(font);
		cs.setVerticalAlignment(VerticalAlignment.CENTER);
		cs.setAlignment(HorizontalAlignment.CENTER);
		
		cs.setBorderBottom(BorderStyle.THIN);//边框
		cs.setBorderLeft(BorderStyle.THIN);
		cs.setBorderRight(BorderStyle.THIN);
		cs.setBorderTop(BorderStyle.THIN);
		
		cell = row.createCell(0);
		cell.setCellValue("设区市");
		cell.setCellStyle(cs);
		
		cell = row.createCell(1);
		cell.setCellValue("县市区");
		cell.setCellStyle(cs);
		
		cell = row.createCell(2);
		cell.setCellValue("监控点名称");
		cell.setCellStyle(cs);
		
		for(int i = k; i < tagLength + k; i++){
			cell = row.createCell(i);
			if(i == k){
				cell.setCellValue("监控功能类型");
				cell.setCellStyle(cs);
			}else{
				cell.setCellValue("");
				cell.setCellStyle(cs);
			}
		}
		
		cell = row.createCell(k + tagLength);
		cell.setCellValue("线路编号");
		cell.setCellStyle(cs);
		
		cell = row.createCell(k + tagLength + 1);
		cell.setCellValue("桩号");
		cell.setCellStyle(cs);
		
		cell = row.createCell(k + tagLength + 2);
		cell.setCellValue("运营商");
		cell.setCellStyle(cs);
		
		cell = row.createCell(k + tagLength + 3);
		cell.setCellValue("异动情况");
		cell.setCellStyle(cs);
		
		row = sheet.createRow(curRow++);//第四行-既标签行
		cs = wb.createCellStyle();
		font = wb.createFont();
		cs.setFont(font);
		cs.setVerticalAlignment(VerticalAlignment.CENTER);
		cs.setAlignment(HorizontalAlignment.CENTER);
		
		cs.setBorderBottom(BorderStyle.THIN);
		cs.setBorderLeft(BorderStyle.THIN);
		cs.setBorderRight(BorderStyle.THIN);
		cs.setBorderTop(BorderStyle.THIN);
		
		cell = row.createCell(0);
		cell.setCellStyle(cs);
		
		cell = row.createCell(1);
		cell.setCellStyle(cs);
		
		cell = row.createCell(2);
		cell.setCellStyle(cs);
		
		for (TagPO tpo : taglst) {
			cell = row.createCell(k++);
			cell.setCellValue(tpo.getName());
			cell.setCellStyle(cs);
		}
		k = 3;//初始化为3
		
		cell = row.createCell(k + tagLength);
		cell.setCellStyle(cs);
		
		cell = row.createCell(k + tagLength + 1);
		cell.setCellStyle(cs);
		
		cell = row.createCell(k + tagLength + 2);
		cell.setCellStyle(cs);
		
		cell = row.createCell(k + tagLength + 3);
		cell.setCellStyle(cs);
		
		//数据行
		cs = wb.createCellStyle();
		font = wb.createFont();
		cs.setFont(font);
		cs.setVerticalAlignment(VerticalAlignment.CENTER);
		cs.setAlignment(HorizontalAlignment.CENTER);
		
		cs.setBorderBottom(BorderStyle.THIN);
		cs.setBorderLeft(BorderStyle.THIN);
		cs.setBorderRight(BorderStyle.THIN);
		cs.setBorderTop(BorderStyle.THIN);
		
		for(CameraVO vo: data){
			row = sheet.createRow(curRow++);
			int curCol = 0;
			cell = row.createCell(curCol++);
			cell.setCellValue((String)vo.getValue("CITY"));
			cell.setCellStyle(cs);
			
			cell = row.createCell(curCol++);
			cell.setCellValue((String)vo.getValue("DISTRICT"));
			cell.setCellStyle(cs);
			
			cell = row.createCell(curCol++);
			cell.setCellValue(vo.getName());
			cell.setCellStyle(cs);
			
			for (TagPO tpo : taglst) {
				cell = row.createCell(curCol++);
				if((Integer)vo.getValue(tpo.getId())==1){
					cell.setCellValue("✔");
				}
				cell.setCellStyle(cs);
			}
			
			cell = row.createCell(curCol++);
			cell.setCellValue(vo.getRouteNo());
			cell.setCellStyle(cs);
			
			cell = row.createCell(curCol++);
			cell.setCellValue((String)vo.getValue("STAKE"));
			cell.setCellStyle(cs);
			
			cell = row.createCell(curCol++);
			cell.setCellValue(vo.getNetProvider());
			cell.setCellStyle(cs);
			
			cell = row.createCell(curCol++);
			cell.setCellValue((String)vo.getValue("LISTTYPE"));
			cell.setCellStyle(cs);
		}
		
		//excel表头合并
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, k + tagLength + 3));
		sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(2, 3, 1, 1));
		sheet.addMergedRegion(new CellRangeAddress(2, 3, 2, 2));
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, k + tagLength - 1));
		sheet.addMergedRegion(new CellRangeAddress(2, 3, k + tagLength, k + tagLength));
		sheet.addMergedRegion(new CellRangeAddress(2, 3, k + tagLength + 1, k + tagLength + 1));
		sheet.addMergedRegion(new CellRangeAddress(2, 3, k + tagLength + 2, k + tagLength + 2));
		sheet.addMergedRegion(new CellRangeAddress(2, 3, k + tagLength + 3, k + tagLength + 3));
		
		//列宽设置
		sheet.setColumnWidth(2, 30*256);
		sheet.setColumnWidth(k + tagLength + 3, 10*256);
	}*/
}
