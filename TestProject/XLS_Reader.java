package main;

import java.awt.List;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLS_Reader {
	
	public static String FilePath="C:\\Users\\B1019\\Documents\\Training\\LearningPOI\\ExcelSheet";
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fout = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;
	
	public XLS_Reader(String path) throws Exception{
		this.path= path;
		fis=new FileInputStream(path);
		workbook= new XSSFWorkbook(fis);
		sheet=workbook.getSheetAt(0);
		//System.out.println(sheet.getSheetName());
		fis.close();
		
	}
	
	
	public void CreateWorkbook() throws Exception{
		fout= new FileOutputStream(FilePath+"\\Test3.xlsx");
		workbook= new XSSFWorkbook();
		sheet= workbook.createSheet("FirstSheet");
		workbook.write(fout);
		fout.close();
	}
	
	public String GetRowCount(String Path, String SheetName) throws Exception{
		System.out.println("------ Get Row Count ------");
		fis=new FileInputStream(path);
		int SheetIndex=workbook.getSheetIndex(SheetName);
		//System.out.println(SheetIndex);
		if(SheetIndex==-1){
			return "No Sheet Found";
		}
		else{
		sheet=workbook.getSheetAt(SheetIndex);
		System.out.println(sheet.getSheetName());
	    int LastRowNo= sheet.getPhysicalNumberOfRows();
	    System.out.println(LastRowNo);
	    return "Sheet Found";	
		}
		
	}
	
	public void ReadData(String SheetName){
		System.out.println("------ Read Data ------");
		int SheetIndex=workbook.getSheetIndex(SheetName);
		sheet=workbook.getSheetAt(SheetIndex);
		Iterator<Row> RowIterator= sheet.iterator();
		while(RowIterator.hasNext()){
			Row row=RowIterator.next();
			Iterator<Cell> CellIterator=row.cellIterator();
			while(CellIterator.hasNext()){
				Cell cell=CellIterator.next();
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.println(cell.getNumericCellValue()+"\t");
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.println(cell.getStringCellValue()+"\t");
					break;
				}
			}
		}

		
	}
	
	/*for(int c=0; c< row.getPhysicalNumberOfCells(); c++){
	cell=row.getCell(c);
	System.out.println(cell + "\t");
		}*/
	
	public void ReadSetUsedData(String SheetName, String Site, String C1, String C2){
		
		System.out.println("------ Read Data and Set Used Data ------");
		int SheetIndex=workbook.getSheetIndex(SheetName);
		sheet=workbook.getSheetAt(SheetIndex);
		for(int r=0; r< sheet.getPhysicalNumberOfRows(); r++){
			row=sheet.getRow(r);
			if(row.getCell(0).toString().contains(Site)){
				if(row.getCell(2).toString().contains("0.0")){
					if(row.getCell(4).toString().contains("0.0")){
						Cell IP= cell=row.getCell(1);
						Cell UP= cell.getRow().getCell(3);
						System.out.println("Issue Point " + IP );
						System.out.println("Usage Point " + UP );
				
					}
				}	
			}
		}
		
	}
	


	public static void main(String[] args) {
		// TODO Auto-generated method stub
		

	}

}
