package com.bossy.base;

import java.io.File;

public class Test {
	public static void main(String[] args){
		/*ExcelOpt opt=new ExcelOpt();
		ExcelOpt.writeExcel("E:\\testExcel.xls");
		System.out.println("ok!");
		ArrayList list=new ArrayList();
		for(int i=0;i<10;i++){
			BookVO book=new BookVO();
			book.setBookName("WebWork in action+"+i);
			book.setBookAuthor("唐勇+"+i);
			book.setBookPrice("39 元+"+i);
			book.setBookConcern("飞思科技+"+i);
			list.add(book);
		}
		ExcelOpt.writeExcelBo("E:\\上市新书.xls",list);
		System.out.println("Book ok!!");*/
		File file=new File("E:\\上市新书.xls");
		//ExcelOpt.readExcel(file);
		ExcelOpt excelOpt=new ExcelOpt();
		System.out.print(ExcelOpt.readExcel(file));
	}

}
