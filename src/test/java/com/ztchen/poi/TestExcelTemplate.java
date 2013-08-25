package com.ztchen.poi;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

import com.ztchen.poi.ExcelTemplate;

public class TestExcelTemplate
{
	@Test
	public void testReadTemplateByClasspath()
	{
		//ExcelTemplate.getInstance().readTemplateByClasspath("excel/default.xls");
		ExcelTemplate.getInstance().readTemplateByPath("excel/default.xls");
		
	}
	
}
