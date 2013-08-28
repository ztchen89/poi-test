package com.ztchen.poi;


import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;

import com.ztchen.model.User;
import com.ztchen.poi.ExcelTemplate;

public class TestExcelTemplate
{
	@Test
	public void testReadTemplateByClasspath()
	{
		ExcelTemplate template = ExcelTemplate.getInstance().readTemplateByClasspath("/excel/default.xls");
		//ExcelTemplate template = ExcelTemplate.getInstance().readTemplateByPath("excel/default.xls");
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
		template.createCell(11);
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
	
	@Test
	public void testObj2Excel()
	{
		List<User> users = new ArrayList<User>();
		users.add(new User(1, "ztchen", 22));
		users.add(new User(2, "ztchen", 22));
		users.add(new User(3, "ztchen", 22));
		users.add(new User(4, "ztchen", 22));
		Map<String, String> datas = new HashMap<String, String>();
		datas.put("title", "测试用户信息");
		datas.put("date", "2013-08-26");
		datas.put("dep", "kmust");
		ExcelUtil.getInstance().exportObj2ExcelByTemplate(datas,"/excel/user.xls", "e:/user.xls",users, User.class, true);
	}
	
	@Test
	public void testObj2ExcelByDefault()
	{
		List<User> users = new ArrayList<User>();
		users.add(new User(1, "ztchen", 22));
		users.add(new User(2, "ztchen", 22));
		users.add(new User(3, "ztchen", 22));
		users.add(new User(4, "ztchen", 22));
		ExcelUtil.getInstance().exportObj2ExcelByPath("e:/testNoTemplate2.xls", users, User.class, false);
	}
}
