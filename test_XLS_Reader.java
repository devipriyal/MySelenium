package test;

import main.XLS_Reader;

public class test_XLS_Reader {
	public static XLS_Reader xls;
	public static String path = "C:\\Users\\B1019\\Documents\\Training\\LearningPOI\\ExcelSheet\\Test1.xlsx";
	

	public static void main(String[] args) throws Exception {
		
		//xls.CreateWorkbook();
		xls= new XLS_Reader(path);
		//System.out.println(xls.GetRowCount(path,"FirstSheet"));
		//System.out.println("\n");
		//xls.ReadData("FirstSheet");
		//xls.ReadSetUsedData("FirstSheet", "CCAD", "IP", "UP");
		boolean b1, b2;
		String s1=null;
		String s2="true";
		b1= Boolean.valueOf(s1);
		b2= Boolean.valueOf(s2);
		System.out.print(b1+ "\n");
		System.out.print(b2);
		

	}

}
