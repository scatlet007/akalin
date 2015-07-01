package com.bossy.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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
public class ExcelOpt {
	/**
	*����һ��Excel�ļ� jxl
	*@param fileName Ҫ���ɵ�Excel�ļ���
	*@jxl.jar �汾 ��2.6
	*/
	public static void writeExcel(String fileName){
		WritableWorkbook wwb=null;
		try{
			//����Ҫ��Workbook��Ĺ�����������һ����д��Ĺ�����(Workbook)����
			wwb=Workbook.createWorkbook(new File(fileName));
		}catch(IOException e){
			e.printStackTrace();
		}
		if(wwb!=null){
			//����һ����д��Ĺ�����
			//Workbook ��createheet������������������һ���ǹ���������ƣ��ڶ����ǹ������ڹ�������λ��
			WritableSheet ws=wwb.createSheet("�����������",0);
			//���濪ʼ��ӵ�Ԫ��
			for(int i=0;i<10;i++){
				for(int j=0;j<5;j++){
					//����Ҫע����ǣ���Excel�У���һ��������ʾ�У��ڶ���������ʾ��
					Label labelC=new Label(j,i,"�����ǵ�"+(i+1)+"��,��"+(j+1)+"��");
					try{
						//�����ɵĵ�Ԫ����ӵ���������
						ws.addCell(labelC);
					}catch(RowsExceededException e){
						e.printStackTrace();
					}catch(WriteException e){
						e.printStackTrace();
					}
				}
			}
			try{
				//���ڴ���д���ļ���
				wwb.write();
				//�ر���Դ���ͷ��ڴ�
				wwb.close();
			}catch(IOException e){
				e.printStackTrace();
			}catch(WriteException e){
				e.printStackTrace();
			}
		}
		
	}
	/**
	*����һ��Excel�ļ�POI
	*@param inputFile ����ģ���ļ�·��
	*@param outputFile �����ļ�����ڷ�����·��
	*@param dataList ����������
	*@throws Exception
	*@roseuid:
	*/
	public static void exceportExcelFile(String inputFile,String outputFile,List dataList) throws Exception{
		//��ģ���ļ�����poi
		POIFSFileSystem fs=new POIFSFileSystem(new FileInputStream(inputFile));
		//����ģ�湤����
		HSSFWorkbook teamplatewb=new HSSFWorkbook(fs);
		//ֱ��ȡģ��ĵ�һ��sheet����
		HSSFSheet templateSheet=teamplatewb.getSheetAt(1);
		//�õ�ģ��ĵ�һ��sheet�ĵ�һ�ж��� Ϊ�˵õ�ģ����ʽ
		HSSFRow templateRow=templateSheet.getRow(0);
		
		//ȡ��Excel�ļ���������
		int columns=templateSheet.getRow((short)0).getPhysicalNumberOfCells();
		//������ʽ����
		HSSFCellStyle styleArray[]=new HSSFCellStyle[columns];
		//һ���Դ��������е���ʽ����������
		for(int s=0;s<columns;s++){
			//�õ�����ʵ��
			styleArray[s]=teamplatewb.createCellStyle();
		}
		//ѭ����ÿһ����Ԫ����и�ֵ
		//��λ��
		for(int rowId=1;rowId<dataList.size();rowId++){
			//����ȡ����rowId�е����� ÿһ��������valueList
			List valueList=(List)dataList.get(rowId -1);
			//��λ��
			for(int columnId=0;columnId<columns;columnId++){
				//����ȡ����Ӧ��columnId�е�ֵ
				//ÿһ����Ԫ���ֵ
				String dataValue=(String)valueList.get(columnId);
				//ȡ��columnId�е�style
				//ģ��ÿһ�е���ʽ
				HSSFCellStyle style=styleArray[columnId];
				//ȥģ���columnId�еĵ�Ԫ�����
				//ģ�浥Ԫ�����
				HSSFCell templateCell=templateRow.getCell((short) columnId);
				//����һ���µ�rowId�ж���
				//�½����ж���
				HSSFRow hssfRow=templateSheet.createRow(rowId);
				//��un�����µ�rowId�� columnId�� ��Ԫ�����
				//�½��ĵ�Ԫ�����
				HSSFCell cell=hssfRow.createCell((short)columnId);
				//��Ӧ��ģ�浥Ԫ�� ��ʽΪ������
				if(templateCell.getCellStyle().getLocked()==false){
					//���ô���styleΪ������
					style.setLocked(false);
					//���õ��µĵ�Ԫ����
					cell.setCellStyle(style);
				}else{
					//������ʽΪ����
					style.setLocked(true);
					//���õ��µ�Ԫ����
					cell.setCellStyle(style);
				}
				//���ñ���
				//cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				//����ֵ ͳһΪString
				cell.setCellValue(dataValue);
			}
		}
		//����������
		FileOutputStream fOut=new FileOutputStream(outputFile);
		//��ģ�������д������ļ���
		teamplatewb.write(fOut);
		fOut.flush();
		//�����������ر��ļ�
		fOut.close();
	}
	
	/**
	*��������ΪXLS��ʽ
	*@param fos����Excel�ļ�path
	*@param bo Ҫ���������
	*/
	public static void writeExcelBo(String fos,java.util.List ve){
		jxl.write.WritableWorkbook wwb;
		try{
			wwb=Workbook.createWorkbook(new File(fos));
			jxl.write.WritableSheet ws=wwb.createSheet("��������",10);
			ws.addCell(new jxl.write.Label(0,1,"����"));
			ws.addCell(new jxl.write.Label(1,1,"����"));
			ws.addCell(new jxl.write.Label(2,1,"����"));
			ws.addCell(new jxl.write.Label(3,1,"������"));
			int bookSize=ve.size();
			BookVO book=new BookVO();
			for(int i=0;i<bookSize;i++){
				book=(BookVO)ve.get(i);
				ws.addCell(new jxl.write.Label(0,i+2,book.getBookName()));
				ws.addCell(new jxl.write.Label(1,i+2,book.getBookAuthor()));
				ws.addCell(new jxl.write.Label(2,i+2,""+book.getBookPrice()));
				ws.addCell(new jxl.write.Label(3,i+2,book.getBookConcern()));
			}
			ws.addCell(new jxl.write.Label(0,0,"2007��07�¼�����������!"));
			wwb.write();
			//�رչ���������
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
	 * ��ȡExcel������
	 * @param file ����ȡ�ļ�
	 * @return ��������
	 */
	public static String readExcel(File file){
		StringBuffer sb=new StringBuffer();
		Workbook wb=null;
		try{
			//����Workbook����
			wb=Workbook.getWorkbook(file);
		}catch(BiffException e){
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
		if(wb==null){
			return null;
		}
		//��ù���������֮�󣬾Ϳ���ͨ�����õ������������
		Sheet[] sheet=wb.getSheets();
		if(sheet!=null&&sheet.length>0){
			//��ÿ�����������ѭ��
			for(int i=0;i<sheet.length;i++){
				//�õ���ǰ�����������
				int rowNum=sheet[i].getRows();
				//�õ���ǰ�е����е�Ԫ��
				for(int j=0;j<rowNum;j++){
					Cell[] cells=sheet[i].getRow(j);
					if(cells!=null&&cells.length>0){
						//��ÿ����Ԫ�����ѭ��
						for(int k=0;k<cells.length;k++){
							//��ȡ��ǰ��Ԫ���ֵ
							String cellValue=cells[k].getContents();
							sb.append(cellValue+"\t");
						}
					}
					sb.append("\r\n");
				}
				sb.append("\r\n");
			}
		}
		//���ر���Դ���ͷ��ڴ�
		wb.close();
		return sb.toString();
	}
}
