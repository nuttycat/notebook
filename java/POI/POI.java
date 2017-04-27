import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class POI {

	public static Workbook getWookbook(String filename){
		InputStream inp = null;
		try {
			inp = new FileInputStream(filename);
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(inp);
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return wb;
	}
	
	public static Map<String,Row> getUsefulMap(Workbook wb,int sheetIndex){
		Sheet sheet = wb.getSheetAt(sheetIndex);
		int rowNum = sheet.getLastRowNum();
		//System.out.println(rowNum);
		Map<String, Row> mapNo2Row = new HashMap<String, Row>();
		for(int i=1;i<=rowNum;i++){
			Row row = sheet.getRow(i);
			Cell cellNo = row.getCell(23);
			if(cellNo==null){
				continue;
			}
			String no=cellNo.getStringCellValue().trim();
			int index = no.indexOf('-');
			if(index>=0){
				//System.out.println("before:" + no);
				no = no.substring(0,index);
				//System.out.println("after:" + no);
			}
			Cell cellNum=row.getCell(10);
			if(cellNum==null){
				continue;
			}
			double num = cellNum.getNumericCellValue();
			Cell cellType = row.getCell(7);
			if(cellType==null){
				continue;
			}
			String type=cellType.getStringCellValue().trim();
			type = type.replaceAll("[^0-9a-zA-Z]","").toLowerCase();
			
			Cell cellDate = row.getCell(0);
			if(cellDate==null||!cellDate.getStringCellValue().startsWith("2017")){
				continue;
			}
			
			String key = no + "--" + num + "--" + type;
			key = key.trim();
			if(!mapNo2Row.containsKey(key)){
				mapNo2Row.put(key, row);
				//System.out.println(key);
			}else{
				System.out.println("num repeat:" + key);
			}
		}
		return mapNo2Row;
	}
	
	public static void main(String[] args) {
		Workbook wb = getWookbook("nianfu.xls");
		Map<String,Row> mapNo2Row = getUsefulMap(wb, 0);
		wb = getWookbook("duizhangdan.xlsx");
		writeByOrder(mapNo2Row,wb,"对账单","output.xlsx");
	}
 
	 
	private static void writeByOrder(Map<String, Row> mapNo2Row, Workbook wb,
			String sheetName,String outputFile) {
		Sheet sheet = wb.getSheetAt(0);
		wb.createSheet(sheetName);
		Sheet sheetOut = wb.getSheetAt(1);
		
		int rowNum = sheet.getLastRowNum();
		for(int i=6;i<=rowNum;i++){
			Row row = sheet.getRow(i);
			if(row==null){
				continue;
			}
			Cell cellNo = row.getCell(1);
			if(cellNo==null){
				continue;
			}
			String no=cellNo.getStringCellValue().trim();
			int index = no.indexOf('-');
			if(index>=0){
				//System.out.println("before:" + no);
				no = no.substring(0,index);
				//System.out.println("after:" + no);
			}
			
			Cell cellNum=row.getCell(4);
			if(cellNum==null){
				continue;
			}
			double num = cellNum.getNumericCellValue();
			Cell cellType = row.getCell(3);
			if(cellType==null){
				continue;
			}
			String type=cellType.getStringCellValue().trim();
			type = type.replaceAll("[^0-9a-zA-Z]","").toLowerCase();
			
			String key = no + "--" + num + "--" + type;
			key = key.trim();
			if(mapNo2Row.containsKey(key)){
				addRow(sheetOut,mapNo2Row.get(key),i,key);
				System.out.println("contains:" + key);
			}else{
				addRow(sheetOut,null,i,key);
				System.out.println("not contains:" + key);
			}
		}
		
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(outputFile);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			wb.write(fileOut);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		try {
			fileOut.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static void addRow(Sheet sheetOut, Row row, int rowNo,String key) {
		Row r = sheetOut.createRow(rowNo);
		Cell c = r.createCell(24);
		c.setCellValue(key);
		if(row==null){
			return ;
		}
		for(int i=0;i<=23;i++){
			c = r.createCell(i);
			Cell fromCell = row.getCell(i);
			if(fromCell==null){
				continue;
			}
			System.out.println(i);
			System.out.println(fromCell.getCellTypeEnum());
			//c.setCellType(fromCell.getCellTypeEnum());
			
			if(fromCell.getCellTypeEnum().equals(CellType.STRING)){
				c.setCellValue(fromCell.getStringCellValue());
			}else if(fromCell.getCellTypeEnum().equals(CellType.NUMERIC)){
				c.setCellValue(fromCell.getNumericCellValue());
			}
		}
	}
}
