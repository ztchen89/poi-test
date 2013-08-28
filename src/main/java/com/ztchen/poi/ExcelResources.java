package com.ztchen.poi;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

//RUNTIME表示编译程序将Annotation存储于class文件中，可由VM读入，能通过反射读取到 
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelResources
{
	String title();//属性的标题名称
	int order() default 9999;//在excel中的顺序
}
