package com.ztchen.poi;


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
		template.writeToFile("E:/test01.xls");
		
	}
	
}
