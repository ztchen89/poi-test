package com.ztchen.poi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/*
 * 该类实现基于模板的导出
 * 如果导出序号，需要在excel中定义一个标识为sernums
 * 如果需要替换信息，需要传入一个map,map中存储要替换的值，在excel中通过#开头
 */
public class ExcelTemplate
{
	private static final String DATA_LINE = "datas";
	private static final String DEFAULT_STYLE = "defaultStyles";
	private static final String STYLE = "styles";
	private static final String SER_NUM = "sernums";
	private static ExcelTemplate template = new ExcelTemplate();
	
	private Workbook wb;
	private Sheet sheet;
	private int initColIndex;//初始化列下标
	private int initRowIndex;//初始化行下标
	private int curColIndex;//当前列下标
	private int curRowIndex;//当前行下标
	private Row curRow;//当前行对象
	private int lastRowIndex;//最后一行下标
	private CellStyle defaultStyle;//默认样式
	private float rowHeight;//默认行高
	private Map<Integer, CellStyle> styles;//存储某一行所对应的样式
	private int serColIndex;//序号的列
	
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
	public ExcelTemplate readTemplateByClasspath(String path)
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
		
		return this;
	}
	
	/*
	 * 第二种是直接文件路径
	 */
	public ExcelTemplate readTemplateByPath(String path)
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
		
		return this;
	}
	
	private void initTemplate()
	{
		sheet = wb.getSheetAt(0);
		initConfigData();
		lastRowIndex = sheet.getLastRowNum();
		//curRow = sheet.getRow(curRowIndex);
		//createNewRow();
	}

	//找到要插入数据的位置,几行几列
	private void initConfigData()
	{
		boolean findData = false;
		boolean findSer = false;
		for(Row row : sheet)
		{
			if(findData) break;//如果找到要插入的位置，则不需要往下运行
			for (Cell cell : row)
			{
				//判断如果定位的那一列的数据类型不是String就跳过
				if(cell.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String str = cell.getStringCellValue().trim();
				if(str.equals(SER_NUM))
				{
					serColIndex = cell.getColumnIndex();
					findSer = true;
				}
				
				if(str.equals(DATA_LINE))
				{
					initColIndex = cell.getColumnIndex();
					initRowIndex = row.getRowNum();
					curColIndex = initColIndex;
					curRowIndex = initRowIndex;
					defaultStyle = cell.getCellStyle();//初始化默认样式
					rowHeight = row.getHeightInPoints();//初始化行高
					findData = true;
					initStyles();
					break;
				}
			}
			if(!findSer)
			{
				initSer();
			}
		}
		
		System.out.println(curColIndex + "," + curRowIndex);
		
	}
	
	/*
	 * 初始化序号位置
	 */
	private void initSer()
	{
		for(Row row : sheet)
		{
			for (Cell cell : row)
			{
				if(cell.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String str = cell.getStringCellValue().trim();
				if(str.equals(SER_NUM))
				{
					serColIndex = cell.getColumnIndex();
				}
			}
		}
	}

	private void initStyles()
	{
		styles = new HashMap<Integer, CellStyle>();
		for(Row row : sheet)
		{
			for (Cell cell : row)
			{
				//判断如果定位的那一列的数据类型不是String就跳过
				if(cell.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String str = cell.getStringCellValue().trim();
				if(str.equals(DEFAULT_STYLE))
				{
					defaultStyle = cell.getCellStyle();//初始化默认样式
				}
				
				if(str.equals(STYLE))
				{
					//存储每一列所对应的样式
					styles.put(cell.getColumnIndex(), cell.getCellStyle());
				}
			}
		}
	}

	/*
	 * 定位到当前行，顺序填充数据到每一列上
	 */
	public void createCell(String value)
	{
		Cell cell = curRow.createCell(curColIndex);
		cell.setCellValue(value);
		setCellStyle(cell);
		
		curColIndex++;
	}
	
	public void createCell(int value)
	{
		Cell cell = curRow.createCell(curColIndex);
		cell.setCellValue(value);
		setCellStyle(cell);
		
		curColIndex++;
	}
	
	public void createCell(Date value)
	{
		Cell cell = curRow.createCell(curColIndex);
		cell.setCellValue(value);
		setCellStyle(cell);
		
		curColIndex++;
	}
	
	public void createCell(double value)
	{
		Cell cell = curRow.createCell(curColIndex);
		cell.setCellValue(value);
		setCellStyle(cell);
		
		curColIndex++;
	}
	
	public void createCell(boolean value)
	{
		Cell cell = curRow.createCell(curColIndex);
		cell.setCellValue(value);
		setCellStyle(cell);
		
		curColIndex++;
	}
	
	public void createCell(Calendar value)
	{
		Cell cell = curRow.createCell(curColIndex);
		cell.setCellValue(value);
		setCellStyle(cell);
		
		curColIndex++;
	}

	private void setCellStyle(Cell cell)
	{
		/*
		 * 判断在map中包含列下标，就设置其存储的样式，否则设置为默认样式
		 */
		if(styles.containsKey(curColIndex))
		{
			cell.setCellStyle(styles.get(curColIndex)); 
		}else {
			cell.setCellStyle(defaultStyle);//每次创建一列，设置该列样式
		}
	}
	
	/*
	 *	创建新行，在使用时只要添加完一行，需要调用该方法创建
	 */
	public void createNewRow()
	{
		if(lastRowIndex > curRowIndex && curRowIndex != initRowIndex)
		{
			sheet.shiftRows(curRowIndex, lastRowIndex, 1, true, true);
			lastRowIndex++;
		}
		curRow = sheet.createRow(curRowIndex);
		curRow.setHeightInPoints(rowHeight);//每次创建一行，设置行高
		curRowIndex++;
		curColIndex = initColIndex;//将列重新定位到初始化列
	}
	
	/*
	 * 插入序号，会自动找相应的序号标示的位置完成插入
	 */
	public void insertSer()
	{
		int index = 1;
		Row row = null;
		Cell cell = null;
		for(int i = initRowIndex; i < curRowIndex; i++)
		{
			row = sheet.getRow(i);
			cell = row.createCell(serColIndex);
			setCellStyle(cell);
			cell.setCellValue(index++);
		}
	}
	
	/*
	 * 根据map替换相应的常量，替换以#号开头的常量
	 */
	public void replaceConstantData(Map<String, String> datas)
	{
		if(datas == null)
			return;
		for(Row row : sheet)
		{
			for (Cell cell : row)
			{
				//判断如果定位的那一列的数据类型不是String就跳过
				if(cell.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String str = cell.getStringCellValue().trim();
				if(str.startsWith("#"))
				{
					if(datas.containsKey(str.substring(1)))
					{
						cell.setCellValue(datas.get(str.substring(1)));
					}
				}
			}
		}
	}
	
	
	/*
	 * 写入文件方式
	 * 根据模板填充数据后写入到一个excel中，将文件写到相应的路径下
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
	 * 写入输出流方式
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
