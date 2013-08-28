package com.ztchen.model;

import com.ztchen.poi.ExcelResources;

public class User
{
	private int id;
	private String username;
	private int age;
	
	public User()
	{
		super();
	}
	
	public User(int id, String username, int age)
	{
		super();
		this.id = id;
		this.username = username;
		this.age = age;
	}

	@ExcelResources(title="用户标识",order=1)
	public int getId()
	{
		return id;
	}
	
	public void setId(int id)
	{
		this.id = id;
	}
	
	@ExcelResources(title="用户名",order=2)
	public String getUsername()
	{
		return username;
	}
	
	public void setUsername(String username)
	{
		this.username = username;
	}
	
	@ExcelResources(title="用户年龄",order=3)
	public int getAge()
	{
		return age;
	}
	
	public void setAge(int age)
	{
		this.age = age;
	}

	@Override
	public String toString()
	{
		return "User [id=" + id + ", username=" + username + ", age=" + age
				+ "]";
	}
	
}
