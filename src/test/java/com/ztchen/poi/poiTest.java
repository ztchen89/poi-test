package com.ztchen.poi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class poiTest
{
	@Test
	public void testReadFromExcel()
	{
		try
		{
			File file = new File("excel/test.xlsx");
			System.out.println(file.getAbsolutePath());
			Workbook wb = WorkbookFactory.create(file);
			Sheet s = wb.getSheetAt(0);//第一个sheet
			Row row = s.getRow(0);//第一行数据
			Cell cell = row.getCell(0);//上面那行的一列数据
			String str = cell.getStringCellValue();
			System.out.println(cell.getCellType());
			System.out.println(str);
			
			
		} catch (InvalidFormatException e)
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
	}
	
	//将各种类型转换成String类型
	private String getCellValue(Cell c)
	{
		String o = null;
		switch (c.getCellType())
		{
		case Cell.CELL_TYPE_BLANK:
			o = ""; 
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			o = String.valueOf(c.getBooleanCellValue()); 
			break;
		case Cell.CELL_TYPE_FORMULA://公式类型
			o = String.valueOf(c.getCellFormula()); 
			break;
		case Cell.CELL_TYPE_NUMERIC:
			//o = String.valueOf((int)c.getNumericCellValue());
			o = new DecimalFormat("#").format(c.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING:
			o = c.getStringCellValue();
			break;
		default:
			o = null; 
			break;
		}
		
		return o;
	}
	
	
	public void testShowListFromExcel()
	{
		try
		{
			File file = new File("excel/test.xlsx");
			Workbook wb = WorkbookFactory.create(file);
			Sheet sheet = wb.getSheetAt(0);//第一个sheet
			System.out.println(sheet.getLastRowNum());
			Row row = null;
			
			for(int i = 0; i < sheet.getLastRowNum(); i++)
			{
				row = sheet.getRow(i);
				for(int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++)
				{
					System.out.print(getCellValue(row.getCell(j)) + "-----");
				}
				
				System.out.println();
			}
			
		} catch (InvalidFormatException e)
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
	}
	
	
	/*
	 * 这种方式不常用，因为excel表的第一行不一定是数据，有可能是标题之类的东西
	 * 而且结束的数据也不一定是最后一行
	 */
	@Test
	public void testShowListFromExcel2()
	{
		try
		{
			File file = new File("excel/test.xlsx");
			Workbook wb = WorkbookFactory.create(file);
			Sheet sheet = wb.getSheetAt(0);//第一个sheet

			for (Row row : sheet)
			{
				for (Cell cell : row)
				{
					System.out.print(getCellValue(cell) + "---");
				}
				
				System.out.println();
			}
			
		} catch (InvalidFormatException e)
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
	}
	
	@Test
	public void testCreateExcel()
	{
		Workbook wb = new HSSFWorkbook();
		FileOutputStream fos = null;
		
		try
		{
			fos = new FileOutputStream("excel/test2.xls");
			Sheet sheet = wb.createSheet("测试01");
			Row row = sheet.createRow(0);
			//设置样式,一般不用这种方式，一般会定义模板
			/*
			row.setHeightInPoints(30);//设置行高
			CellStyle cs = wb.createCellStyle();
			cs.setAlignment(CellStyle.ALIGN_CENTER);//设置居中
			cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			*/
			
			Cell cell = row.createCell(0);
			cell.setCellValue("用户名");
			cell = row.createCell(1);
			cell.setCellValue("密码");
			
			wb.write(fos);
		} catch (FileNotFoundException e)
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}finally{
			try
			{
				if(null != fos)
					fos.close();
			} catch (IOException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
	}
	
	
	
	
	
	
	
}
