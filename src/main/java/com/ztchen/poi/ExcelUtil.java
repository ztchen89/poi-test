package com.ztchen.poi;

import java.beans.Beans;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil
{
	private static ExcelUtil eu = new ExcelUtil();
	
	private ExcelUtil()
	{
	}
	
	public static ExcelUtil getInstance()
	{
		return eu;
	}
	
	/*
	 * 处理对象输出到excel的方法
	 */
	private ExcelTemplate handlerObj2Excel(Map<String,String> datas, String templatePath,List objList, Class clazz,boolean isClasspath)
	{
		ExcelTemplate template = null;
		try
		{
			template = ExcelTemplate.getInstance();
			if(isClasspath)
			{
				template.readTemplateByClasspath(templatePath);
			}else {
				template.readTemplateByPath(templatePath);
			}
			
			List<ExcelHeader> headers = getHeaderLisr(clazz);
			Collections.sort(headers);
			//输出标题
			template.createNewRow();
			for (ExcelHeader excelHeader : headers)
			{
				template.createCell(excelHeader.getTitle());
			}
			//输出值
			for(Object obj : objList)
			{
				template.createNewRow();
				for(ExcelHeader excelHeader : headers)
				{
//					String methodName = excelHeader.getMethodName();
//					Method m = clazz.getDeclaredMethod(methodName);
//					m.invoke(obj);
//					Object rel = m.invoke(obj);
					template.createCell(getObj(excelHeader, clazz, obj));
				}
			}
			
			template.replaceConstantData(datas);
		} catch (SecurityException e)
		{
			e.printStackTrace();
		} catch (IllegalArgumentException e)
		{
			e.printStackTrace();
		}
		return template;
	}
	
	
	private String getObj(ExcelHeader excelHeader,Class clazz,Object obj)
	{
		try
		{
			String methodName = excelHeader.getMethodName();
			Method m = clazz.getDeclaredMethod(methodName);
			m.invoke(obj);
			Object rel = m.invoke(obj);
			return rel.toString();
		} catch (NoSuchMethodException e)
		{
			e.printStackTrace();
		} catch (SecurityException e)
		{
			e.printStackTrace();
		} catch (IllegalAccessException e)
		{
			e.printStackTrace();
		} catch (IllegalArgumentException e)
		{
			e.printStackTrace();
		} catch (InvocationTargetException e)
		{
			e.printStackTrace();
		}
		
		return null;
	}
	/*
	 * 将excel输出到文件
	 */
	public void exportObj2ExcelByTemplate(Map<String,String> datas, String templatePath,String outputPath, List objList, Class clazz,boolean isClasspath)
	{
		ExcelTemplate template = handlerObj2Excel(datas, templatePath, objList, clazz, isClasspath);
		template.writeToFile(outputPath);
			
	}
	
	/*
	 * 将excel输出到一个流
	 */
	public void exportObj2ExcelByTemplate(Map<String,String> datas, String templatePath,OutputStream os, List objList, Class clazz,boolean isClasspath)
	{
		ExcelTemplate template = handlerObj2Excel(datas, templatePath, objList, clazz, isClasspath);
		template.writeToStream(os);
	}
	
	/*
	 * 不基于模板，将对象输出到excel中
	 */
	private Workbook handleExportObj2Excel(List objList, Class clazz,boolean isXssf)
	{
		Workbook wb = null;
		if(isXssf)
		{
			wb = new XSSFWorkbook();
		}else {
			wb = new HSSFWorkbook();
		}
		
		Sheet sheet = wb.createSheet();
		Row row = sheet.createRow(0);
		List<ExcelHeader> headers = getHeaderLisr(clazz);
		Collections.sort(headers);//排序很重要
		//写标题
		for (int i = 0; i < headers.size(); i++)
		{
			row.createCell(i).setCellValue(headers.get(i).getTitle());
		}
		
		//写数据
		Object obj = null;
		for (int i = 0; i < objList.size(); i++)
		{
			row = sheet.createRow(i + 1);
			obj = objList.get(i);
			for (int j = 0; j < headers.size(); j++)
			{
				row.createCell(j).setCellValue(getObj(headers.get(j), clazz, obj));
			}
		}
		
		return wb;
		
	}
	
	/*
	 * 不基于模板的输出，输出到一个路径
	 */
	public void exportObj2ExcelByPath(String outputPath, List objList, Class clazz,boolean isXssf)
	{
		Workbook wb = handleExportObj2Excel(objList, clazz, isXssf);
		FileOutputStream fos = null;
		try
		{
			fos = new FileOutputStream(outputPath);
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
				if(fos != null)
					fos.close();
			} catch (IOException e)
			{
				e.printStackTrace();
			}
		}
	}
	
	/*
	 * 不基于模板的输出，输出到一个流
	 */
	public void exportObj2ExcelByStream(OutputStream os, List objList, Class clazz,boolean isXssf)
	{
		try
		{
			Workbook wb = handleExportObj2Excel(objList, clazz, isXssf);
			wb.write(os);
		} catch (IOException e)
		{
			e.printStackTrace();
		}
		
	}
	
	public List<ExcelHeader> getHeaderLisr(Class clazz)
	{
		List<ExcelHeader> headers = new ArrayList<ExcelHeader>();
		Method[] methods = clazz.getDeclaredMethods();
		for (Method method : methods)
		{
			String methodName = method.getName();
			if(methodName.startsWith("get"))
			{
				if(method.isAnnotationPresent(ExcelResources.class))
				{
					ExcelResources er = method.getAnnotation(ExcelResources.class);
					headers.add(new ExcelHeader(er.title(), er.order(), methodName));
				}
			}
		}
		
		return headers;
	}
	
	
	//从excel中读取数据对象
	private Map<Integer, String> getHeaderMap(Row titleRow, Class clazz)
	{
		List<ExcelHeader> headers = getHeaderLisr(clazz);
		Map<Integer, String> map = new HashMap<Integer, String>();
		for (Cell cell : titleRow)
		{
			String title = cell.getStringCellValue();
			for(ExcelHeader excelHeader : headers)
			{
				if(excelHeader.getTitle().equals(title))
				{
					map.put(cell.getColumnIndex(), excelHeader.getMethodName().replace("get", "set"));
					break;
				}
			}
		}
		
		return map;
	}
	
	//通过classpath读取文件
	public List<Object> readExcel2ObjsByClasspath(String path,Class clazz, int readLine,int tailLine)
	{
		Workbook wb = null;
		try
		{
			wb = WorkbookFactory.create(ExcelUtil.class.getResourceAsStream(path));
			return handlerExcel2Obj(wb, clazz, readLine,tailLine);
		} catch (InvalidFormatException e)
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
		
		return null;
	}
	
	//通过路径读取文件
	public List<Object> readExcel2ObjsByPath(String path,Class clazz, int readLine,int tailLine)
	{
		Workbook wb = null;
		try
		{
			wb = WorkbookFactory.create(new File(path));
			return handlerExcel2Obj(wb, clazz, readLine,tailLine);
		} catch (InvalidFormatException e)
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
		
		return null;
	}
	
	//通过classpath读取文件，从零行开始读
	public List<Object> readExcel2ObjsByClasspath(String path,Class clazz)
	{
		return this.readExcel2ObjsByClasspath(path, clazz, 0,0);
	}
	
	//通过路径读取文件，从零行开始读
	public List<Object> readExcel2ObjsByPath(String path,Class clazz)
	{
		return this.readExcel2ObjsByPath(path, clazz, 0,0);
	}
	
	private List<Object> handlerExcel2Obj(Workbook wb,Class clazz,int readLine,int tailLine)
	{
		Sheet sheet = wb.getSheetAt(0);
		List<Object> objList = null;
		try
		{
			Row row = sheet.getRow(readLine);
			Map<Integer, String> map = getHeaderMap(row, clazz);
			objList = new ArrayList<Object>();
			
			Object obj = null;
			for(int i = readLine + 1;i <= sheet.getLastRowNum() - tailLine ;i++)
			{
				row = sheet.getRow(i);
				obj = clazz.newInstance();
				for(Cell cell : row)
				{
					int ci = cell.getColumnIndex();
					String methodName = map.get(ci).substring(3);
					methodName = methodName.substring(0, 1).toLowerCase() + methodName.substring(1);
					BeanUtils.copyProperty(obj, methodName, this.getCellValue(cell));
				}
				
				objList.add(obj);
			}
		} catch (InstantiationException e)
		{
			e.printStackTrace();
		} catch (IllegalAccessException e)
		{
			e.printStackTrace();
		} catch (InvocationTargetException e)
		{
			e.printStackTrace();
		}
		
		return objList;
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
}
