package com.akalin.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.bossy.base.BookVO;

public class ExcelOpt {
	/**
	*生成一个Excel文件 jxl
	*@param fileName 要生成的Excel文件名
	*@param list 数据源，属性为List<List<Object>>
	*@jxl.jar 版本 ：2.6
	*/
	public static void writeExcel(String fileName,List<List<Object>> list){
		WritableWorkbook wwb=null;
		try{
			//首先要用Workbook类的工厂方法创建一个可写入的工作簿(Workbook)对象
			wwb=Workbook.createWorkbook(new File(fileName));
		}catch(IOException e){
			e.printStackTrace();
		}
		if(wwb!=null){
			//创建一个可写入的工作表
			//Workbook 的createheet方法有两个参数，第一个是工作标的名称，第二个是工作表在工作簿的位置
			WritableSheet ws=wwb.createSheet("工作表的名称",0);
			//下面开始添加单元格
			for(int i=0;i<list.size();i++){
				for(int j=0;j<list.get(i).size();j++){
					//这里要注意的是，在Excel中，第一个参数表示列，第二个参数表示行
					Label labelC=new Label(j,i,(String)list.get(1+i).get(j)+1);
					try{
						//将生成的单元格添加到工作表中
						ws.addCell(labelC);
					}catch(RowsExceededException e){
						e.printStackTrace();
					}catch(WriteException e){
						e.printStackTrace();
					}
				}
			}
			try{
				//从内存中写入文件中
				wwb.write();
				//关闭资源，释放内存
				wwb.close();
			}catch(IOException e){
				e.printStackTrace();
			}catch(WriteException e){
				e.printStackTrace();
			}
		}
		
	}
	/**
	*生成一个Excel文件POI
	*@param inputFile 输入模版文件路径
	*@param outputFile 输入文件存放于服务器路径
	*@param dataList 待导出数据
	*@throws Exception
	*@roseuid:
	*/
	public static void exceportExcelFile(String inputFile,String outputFile,List dataList) throws Exception{
		//用模版文件构造poi
		POIFSFileSystem fs=new POIFSFileSystem(new FileInputStream(inputFile));
		//创建模版工作表
		HSSFWorkbook teamplatewb=new HSSFWorkbook(fs);
		//直接取模版的第一个sheet对象
		HSSFSheet templateSheet=teamplatewb.getSheetAt(1);
		//得到模版的第一个sheet的第一行对象 为了得到模版样式
		HSSFRow templateRow=templateSheet.getRow(0);
		
		//取得Excel文件的总列数
		int columns=templateSheet.getRow((short)0).getPhysicalNumberOfCells();
		//创建样式数组
		HSSFCellStyle styleArray[]=new HSSFCellStyle[columns];
		//一次性创建所有列的样式放在数组里
		for(int s=0;s<columns;s++){
			//得到数组实例
			styleArray[s]=teamplatewb.createCellStyle();
		}
		//循环对每一个单元格进行赋值
		//定位行
		for(int rowId=1;rowId<dataList.size();rowId++){
			//依次取出第rowId行的数据 每一个数据是valueList
			List valueList=(List)dataList.get(rowId -1);
			//定位列
			for(int columnId=0;columnId<columns;columnId++){
				//依次取出对应与columnId列的值
				//每一个单元格的值
				String dataValue=(String)valueList.get(columnId);
				//取出columnId列的style
				//模版每一列的样式
				HSSFCellStyle style=styleArray[columnId];
				//去模版第columnId列的单元格对象
				//模版单元格对象
				HSSFCell templateCell=templateRow.getCell((short) columnId);
				//创建一个新的rowId行对象
				//新建的行对象
				HSSFRow hssfRow=templateSheet.createRow(rowId);
				//差un构建新的rowId行 columnId列 单元格对象
				//新建的单元格对象
				HSSFCell cell=hssfRow.createCell((short)columnId);
				//对应的模版单元格 样式为非锁定
				if(templateCell.getCellStyle().getLocked()==false){
					//设置此列style为非锁定
					style.setLocked(false);
					//设置到新的单元格上
					cell.setCellStyle(style);
				}else{
					//否则样式为锁定
					style.setLocked(true);
					//设置到新单元格上
					cell.setCellStyle(style);
				}
				//设置编码
				//cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				//设置值 统一为String
				cell.setCellValue(dataValue);
			}
		}
		//设置输入流
		FileOutputStream fOut=new FileOutputStream(outputFile);
		//将模版的内容写到输出文件上
		teamplatewb.write(fOut);
		fOut.flush();
		//操作结束，关闭文件
		fOut.close();
	}
	
	/**
	*导出数据为XLS格式
	*@param fos生成Excel文件path
	*@param columnName 表单表头值
	*@param ve 要导出的数据值
	*/
	public static void writeExcelBo(String fos,String[] columnName,List<List<Object>> ve){
		jxl.write.WritableWorkbook wwb;
		try{
			wwb=Workbook.createWorkbook(new File(fos));
			jxl.write.WritableSheet ws=wwb.createSheet("上市新书",10);
			//导出表头值
			for(int i=0;i<columnName.length;i++){
				ws.addCell(new jxl.write.Label(i,1,columnName[i]));
			}
			//循环读取每一行的值
			for(int i=0;i<ve.size();i++){
				//循环读取每一单元格的值
				for(int j=0;j<ve.get(i).size();j++){
					//向外写单元格的值
					ws.addCell(new jxl.write.Label(j,i+2,(String)ve.get(i).get(j)));
				}
			}
			wwb.write();
			//关闭工作簿对象
			wwb.close();
		}catch(IOException e){
			e.printStackTrace();
		}catch(RowsExceededException e){
			e.printStackTrace();
		}catch(WriteException e){
			e.printStackTrace();
		}
	}
	
	/**
	 * 读取Excel的内容
	 * @param file 待读取文件
	 * @param 列名数组
	 * @return 返回数据
	 */
	public static List<List<Object>> readExcel(File file,String[] columnName){
		//StringBuffer sb=new StringBuffer();
		List<List<Object>> list=new ArrayList<List<Object>>();
		Workbook wb=null;
		try{
			//构造Workbook对象
			wb=Workbook.getWorkbook(file);
		}catch(BiffException e){
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
		if(wb==null){
			return null;
		}
		//获得工作簿对象之后，就可以通过它得到工作表对象了
		Sheet[] sheet=wb.getSheets();
		if(sheet!=null&&sheet.length>0){
			//对每个工作表进行循环
			for(int i=0;i<sheet.length;i++){
				//得到当前工作表的行数
				int rowNum=sheet[i].getRows();
				//得到当前行的所有单元格
				/*for(int j=0;j<sheet[i].getColumns();j++){
					Cell[] cells=sheet[i].getRow(j);
					if(cells!=null&&cells.length>0){
						//对每个单元格进行循环
						for(int k=0;k<cells.length;k++){
							//读取当前单元格的值
							String cellValue=cells[k].getContents();
							sb.append(cellValue+"\t");
						}
					}
					sb.append("\r\n");
				}*/
				//匹配表头值，得到要写入的列下标
				
				//sb.append("\r\n");
			}
		}
		//最后关闭资源，释放内存
		wb.close();
		return list;
	}
}
