import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class Activity {

	public static void main(String[] args) {
		readExcel();
		writeExcel();
	}
	public static void readExcel()  {
        FileInputStream fileInputStream=null ;
        POIFSFileSystem poifsFileSystem=null ;
        HSSFWorkbook hssfWorkbook=null ;
        String filePath = "excel/read.xls";
        try
        {
			fileInputStream = new FileInputStream(filePath);
			poifsFileSystem = new POIFSFileSystem( fileInputStream );
			hssfWorkbook = new HSSFWorkbook(poifsFileSystem);
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0); 

			Iterator<Row> rowIterator=hssfSheet.iterator();
			
			while(rowIterator.hasNext()) {
				Row currentRow=rowIterator.next();
				Iterator<Cell> cellIterator=currentRow.iterator();
				while(cellIterator.hasNext()) {
					Cell currentCell=cellIterator.next();
	
					if(currentCell.getCellTypeEnum() !=CellType.BLANK)
						System.out.print(currentCell.toString()+"\t");
				}
				System.out.println();
			}

			fileInputStream.close();
        }catch(java.io.IOException e)
        {
          e.printStackTrace();
        }

	}
	
	public static void writeExcel() {
		FileOutputStream fileOutputStream =null;
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook() ;
        String filePath = "excel/write.xls";
        HSSFSheet hssfSheet=hssfWorkbook.createSheet("leave");
        Object[][] datatypes = {
                {"Datatype", "Type", "Size(in bytes)"},
                {"int", "Primitive", 2},
                {"float", "Primitive", 4},
                {"double", "Primitive", 8},
                {"char", "Primitive", 1},
                {"String", "Non-Primitive", "No fixed size"}
        };
        
        int rowNum=0;
        
        for(Object[] dataType:datatypes) {
        	Row row=hssfSheet.createRow(rowNum++);
        	int columnNum=0;
        	
        	for(Object field:dataType) {
        		Cell cell=row.createCell(columnNum++);
        		if(field instanceof String)
        			cell.setCellValue((String) field);
        		else if(field instanceof Integer)
        			cell.setCellValue((Integer)field);
        	}
        }
        try {
        	fileOutputStream=new FileOutputStream(filePath);
        	hssfWorkbook.write(fileOutputStream);
        	hssfWorkbook.close();
        } catch (FileNotFoundException e) {
			// TODO: handle exception
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
