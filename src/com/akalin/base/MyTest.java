package com.akalin.base;

import java.io.File;
import java.util.List;

public class MyTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		File file=new File("E:\\上市新书.xls");
		//ExcelOpt.readExcel(file);
		String[] str={"定价","出版社"};
		ExcelOpt excelOpt=new ExcelOpt();
		List<List<Object>> list=excelOpt.readExcel(file,str);
		for(List<Object> ls:list){
			for(Object obj:ls){
				System.out.print(obj+"\t");
			}
			System.out.println();
		}
	}

}
