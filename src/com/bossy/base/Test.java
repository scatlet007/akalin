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
			book.setBookAuthor("����+"+i);
			book.setBookPrice("39 Ԫ+"+i);
			book.setBookConcern("��˼�Ƽ�+"+i);
			list.add(book);
		}
		ExcelOpt.writeExcelBo("E:\\��������.xls",list);
		System.out.println("Book ok!!");*/
		File file=new File("E:\\��������.xls");
		//ExcelOpt.readExcel(file);
		ExcelOpt excelOpt=new ExcelOpt();
		System.out.print(ExcelOpt.readExcel(file));
	}

}
