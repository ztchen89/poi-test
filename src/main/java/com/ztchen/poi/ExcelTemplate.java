package com.ztchen.poi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelTemplate
{
	private static final String DATA_LINE = "datas";
	private static ExcelTemplate template = new ExcelTemplate();
	
	private Workbook wb;
	private Sheet sheet;
	private int initColIndex;//初始化列下标
	private int initRowIndex;//初始化行下标
	private int curColIndex;//当前列下标
	private int curRowIndex;//当前行下标
	private Row curRow;//当前行对象
	
	private ExcelTemplate()
	{
	}
	
	public static ExcelTemplate getInstance()
	{
		return template;
	}
	
	//1.读取相应的模板文档,有两种读取方式
	
	/*
	 * 第一种是在classpath下读取
	 */
	public void readTemplateByClasspath(String path)
	{
		try
		{
			wb = WorkbookFactory.create(ExcelTemplate.class.getResourceAsStream(path));
			initTemplate();
		} catch (InvalidFormatException e)
		{
			e.printStackTrace();
			throw new RuntimeException("读取模板格式有错！");
		} catch (IOException e)
		{
			e.printStackTrace();
			throw new RuntimeException("读取模板不存在！");
		}
	}
	
	/*
	 * 第二种是直接文件路径
	 */
	public void readTemplateByPath(String path)
	{
		try
		{
			wb = WorkbookFactory.create(new File(path));
			initTemplate();
		} catch (InvalidFormatException e)
		{
			e.printStackTrace();
			throw new RuntimeException("读取模板格式有错！");
		} catch (IOException e)
		{
			e.printStackTrace();
			throw new RuntimeException("读取模板不存在！");
		}
	}
	
	private void initTemplate()
	{
		sheet = wb.getSheetAt(0);
		initConfigData();
		curRow = sheet.getRow(curRowIndex);
		createNewRow();
	}

	//找到要插入数据的位置,几行几列
	private void initConfigData()
	{
		boolean findData = false;
		for(Row row : sheet)
		{
			if(findData) break;//如果找到要插入的位置，则不需要往下运行
			for (Cell cell : row)
			{
				//判断如果定位的那一列的数据类型不是String就跳过
				if(cell.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String str = cell.getStringCellValue();
				if(str.endsWith(DATA_LINE))
				{
					initColIndex = cell.getColumnIndex();
					initRowIndex = row.getRowNum();
					curColIndex = initColIndex;
					curRowIndex = initRowIndex;
					findData = true;
					break;
				}
			}
		}
		
		System.out.println(curColIndex + "," + curRowIndex);
		
	}
	
	/*
	 * 定位到当前行，顺序填充数据到每一列上
	 */
	public void createCell(String value)
	{
		curRow.createCell(curColIndex).setCellValue(value);
		curColIndex++;
	}
	
	/*
	 * 创建新的一行
	 */
	public void createNewRow()
	{
		curRow = sheet.createRow(curRowIndex);
		curRowIndex++;
		curColIndex = initColIndex;//将列重新定位到初始化列
	}
	
	
	/*
	 * 写入文件方式
	 * 根据模板填充数据后写入到一个excel中，并输出到一个位置上
	 */
	public void writeToFile(String filepath)
	{
		FileOutputStream fos = null;
		try
		{
			fos = new FileOutputStream(filepath);
			wb.write(fos);
		} catch (FileNotFoundException e)
		{
			e.printStackTrace();
			throw new RuntimeException("写入文件不存在");
		} catch (IOException e)
		{
			e.printStackTrace();
			throw new RuntimeException("写入数据失败" + e.getMessage());
		}finally{
			try
			{
				if(fos != null)
					fos.close();
			} catch (IOException e)
			{
				e.printStackTrace();
			}
		}
	}
	
	/*
	 * 写入流方式
	 * 根据模板填充数据后写入到一个excel中，并输出到一个位置上
	 */
	public void writeToStream(OutputStream os)
	{
		try
		{
			wb.write(os);
		} catch (IOException e)
		{
			e.printStackTrace();
			throw new RuntimeException("写入流失败" + e.getMessage());
		}
	}
	
	
	
	
	
	
	
	
}
