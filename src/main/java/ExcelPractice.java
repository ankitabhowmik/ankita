    import java.io.FileInputStream;
	import java.io.FileNotFoundException;
	import java.io.IOException;

	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class ExcelPractice {
			
			
			public String path;
			public FileInputStream fs=null;
			private XSSFWorkbook workbook=null;
			private XSSFSheet sheet=null;
			private XSSFRow row=null;
			private XSSFCell cell=null;
			
			
			public ExcelPractice(String path)
			{
				this.path=path;
				try {
					fs=new FileInputStream(path);
					workbook=new XSSFWorkbook(fs);
					sheet=workbook.getSheetAt(0);
				    fs.close();
					
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			
			public boolean isSheetAvailable(String sheetname)
			{
				int index=workbook.getSheetIndex(sheetname);
				if(index==-1)
				{
					index=workbook.getSheetIndex(sheetname.toUpperCase());
					if(index==-1)
					{
						return false;
					}
					System.out.println("sheet is available");
					return true;
				}
				System.out.println("Sheet is available");
				return true;
			}
			
			public int getrowcount(String sheetname)
			{
				if(!isSheetAvailable(sheetname))
				{
					return 0;
					
				}
				sheet=workbook.getSheet(sheetname);
				int rownum=sheet.getLastRowNum();
				
				int rowcount=rownum+1;
				System.out.println(rowcount);
				return rowcount;
				
				
			}
			
			public int getcolcount(String sheetname) {
				if(!isSheetAvailable(sheetname))
				{
					return 0;
					
				}
				sheet=workbook.getSheet(sheetname);
				row=sheet.getRow(0);
				int colnum=row.getLastCellNum();
				System.out.println(colnum);
				return colnum;
				
			}
			
			public String getcelldata(String sheetname, String colname, int rownum)
			{
				int index=workbook.getSheetIndex(sheetname);
				sheet=workbook.getSheetAt(index);
				row=sheet.getRow(0);
				int totalcolno=getcolcount(sheetname);
				int colnum=-1;
				for(int i=0;i<totalcolno;i++)
				{
					if(row.getCell(i).getStringCellValue().equalsIgnoreCase(colname))
					{
				       colnum=i;
				       System.out.println(colnum);
					}
				}
			
				if(rownum<=0)
					return "";
				else
				row=sheet.getRow(rownum);
				cell=row.getCell(colnum);
				String cellvalue="";
				
				switch(cell.getCellType())
				{
				case STRING:
					cellvalue=cell.getStringCellValue();
				System.out.println(cellvalue);
				break;
				
				case NUMERIC:
					cellvalue=String.valueOf(cell.getNumericCellValue());
				System.out.println(cellvalue);
				break;
				
				case BOOLEAN:
					cellvalue=String.valueOf(cell.getBooleanCellValue());
				System.out.println(cellvalue);
				break;
				
				}
				return cellvalue;
			}
			
			

			public static void main(String[] args) {
				// TODO Auto-generated method stub
				
				ExcelPractice excel=new ExcelPractice("C:\\Users\\sony\\Documents\\PracticeStart1\\src\\test\\resources\\Excel\\test.xlsx");
					
				String sheetname="Sheet1";
				excel.isSheetAvailable(sheetname);
				excel.getrowcount(sheetname);
				excel.getcolcount(sheetname);
				excel.getcelldata(sheetname, "password", 3);
				

			}

		
	}





