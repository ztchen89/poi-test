package com.ztchen.poi;

import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

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
	private ExcelTemplate handlerObj2Excel(Map<String,String> datas, String templatePath,List objList, Class clazz,boolean isClasspath) throws NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException
	{
		ExcelTemplate template = ExcelTemplate.getInstance();
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
				String methodName = excelHeader.getMethodName();
				Method m = clazz.getDeclaredMethod(methodName);
				Object rel = m.invoke(obj);
				template.createCell(rel.toString());
			}
		}
		
		template.replaceConstantData(datas);
		return template;
	}
	
	/*
	 * 将excel输出到文件
	 */
	public void exportObj2ExcelByTemplate(Map<String,String> datas, String templatePath,String outputPath, List objList, Class clazz,boolean isClasspath)
	{
		try
		{
			ExcelTemplate template = handlerObj2Excel(datas, templatePath, objList, clazz, isClasspath);
			template.writeToFile(outputPath);
			
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
	}
	
	/*
	 * 将excel输出到一个流
	 */
	public void exportObj2ExcelByTemplate(Map<String,String> datas, String templatePath,OutputStream os, List objList, Class clazz,boolean isClasspath)
	{
		try
		{
			ExcelTemplate template = handlerObj2Excel(datas, templatePath, objList, clazz, isClasspath);
			template.writeToStream(os);
			
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
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
