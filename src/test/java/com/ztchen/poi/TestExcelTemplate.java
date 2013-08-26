package com.ztchen.poi;


import java.util.HashMap;
import java.util.Map;

import org.junit.Test;

import com.ztchen.poi.ExcelTemplate;

public class TestExcelTemplate
{
	@Test
	public void testReadTemplateByClasspath()
	{
		//ExcelTemplate.getInstance().readTemplateByClasspath("excel/default.xls");
		ExcelTemplate template = ExcelTemplate.getInstance().readTemplateByPath("excel/default.xls");
		template.createNewRow();
		template.createCell("111");
		template.createCell("222");
		template.createCell("333");
		template.createCell("444");
		template.createNewRow();
		template.createCell("111");
		template.createCell("222");
		template.createCell("333");
		template.createCell("444");
		template.createNewRow();
		template.createCell("111");
		template.createCell("222");
		template.createCell("333");
		template.createCell("444");
		template.createNewRow();
		template.createCell("111");
		template.createCell("222");
		template.createCell("333");
		template.createCell("444");
		template.createNewRow();
		template.createCell("111");
		template.createCell("222");
		template.createCell("333");
		template.createCell("444");
		template.createNewRow();
		template.createCell("111");
		template.createCell("222");
		template.createCell("333");
		template.createCell("444");
		template.createNewRow();
		template.createCell("111");
		template.createCell("222");
		template.createCell("333");
		template.createCell("444");
		
		Map<String, String> datas = new HashMap<String, String>();
		datas.put("title", "测试用户信息");
		datas.put("date", "2013-08-26");
		datas.put("dep", "kmust");
		template.replaceConstantData(datas);//替换常量
		template.insertSer();
		template.writeToFile("E:/test01.xls");
		
	}
	
}
