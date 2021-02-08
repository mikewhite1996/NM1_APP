import java.awt.BorderLayout;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Map;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JProgressBar;
import javax.swing.JTextField;
import javax.swing.WindowConstants;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;

public class GUI_Startup {
	static JProgressBar b = new JProgressBar();
	static String barMessage;
	static int barProgress;
	static fillProgress bar = new fillProgress();
	static JTextField textFile = new JTextField(20);
	static JFrame f=new JFrame();
	String filePath = "C:/Users/mikew/Documents/";
	ImageIcon logo;
	JLabel pic;
	static long startTime = System.currentTimeMillis();
	public void GUI() throws IOException{
		b.setBounds(135,220,100,20);
		f.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
		f.setTitle("NM1 Inventory Allocations");
		JButton chooseFile=new JButton("Choose File");
		JButton Submit=new JButton("Submit");
		Submit.setBounds(290,150,75,20);
		chooseFile.setBounds(135,175,100, 40);
		textFile.setBounds(85, 150, 200, 20);
		textFile.setEditable(false);
		JLabel fileLabel = new JLabel("File Location: ");
		fileLabel.setBounds(20, 150, 100, 20);
		fileLabel.setLabelFor(textFile);
		
		chooseFile.addActionListener(new ActionListener(){
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
					FileSelection();
				}
			});
	
		
		Submit.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent arg0){
				try{
					FileNameCheck();
					writeExcel(filePath, textFile.getText(), "Updated");
					readExcel(filePath, textFile.getText(), "Updated");
					Desktop.getDesktop().open(new File(filePath+textFile.getText()));
					long elapsedTime = System.currentTimeMillis() - startTime;
					System.out.print(elapsedTime);
					
				}
				catch(FileNameException | IOException | ExcelDataFormatError e){
					System.out.println("what the fuck");
					FileSelection();
				}
			}
		});
		f.add(b);
		f.add(textFile);
		f.add(fileLabel);          
		f.add(Submit);//adding button in JFrame
		f.add(chooseFile);
		ImageLoader(); 
		
		f.setSize(400,500);//400 width and 500 height  
		f.setLayout(null);//using no layout managers
		f.setLocationRelativeTo(null);
		f.setVisible(true);//making the frame visible  
	}
	public static void FileSelection(){
		JFileChooser chooser = new JFileChooser();
	    FileNameExtensionFilter filter = new FileNameExtensionFilter(
	        "Excel File", "xlsx");
	    chooser.setFileFilter(filter);
	    int returnVal = chooser.showOpenDialog(f);
	    if(returnVal == JFileChooser.APPROVE_OPTION) {
	    		textFile.setText(chooser.getSelectedFile().getName());
	    	}
		
	}
	public static void writeExcel(String filePath, String fileName, String sheetName) throws IOException{
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName);
		
		for (int r=0;r < 140000; r++ )
		{
			XSSFRow row = sheet.createRow(r);
			//iterating c number of columns
			for (int c=0;c < 5; c++ )
			{
				XSSFCell cell = row.createCell(c);
				if(c==2){
					//System.out.print("stepped into");
					cell.setCellValue("Test");
					//System.out.println(cell.getRawValue());
				}else{
					cell.setCellValue("Cell");
				}
				
			}
		}
		barMessage = "Writing Excel";
		barProgress = 30;
		bar.update(barMessage, barProgress, b);
		
		
		FileOutputStream fileOut = new FileOutputStream(filePath+fileName);

		//write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}
	public static void filterExcel(String filePath, String fileName, XSSFWorkbook wb, XSSFSheet sheet, String search) throws IOException{
		/* Step-1: Get the CTAutoFilter Object */
		CTAutoFilter sheetFilter=sheet.getCTWorksheet().getAutoFilter();                             
		/* Step -2: Add new Filter Column */
		CTFilterColumn  myFilterColumn=sheetFilter.insertNewFilterColumn(0);
		/* Step-3: Set Filter Column ID */
		myFilterColumn.setColId(1L);
		/* Step-4: Add new Filter */
		CTFilter myFilter=myFilterColumn.addNewFilters().insertNewFilter(0);
		/* Step -5: Define Auto Filter Condition - We filter Brand with Value of "A" */
		myFilter.setVal(search);                           
		XSSFRow r1;
		/* Step-6: Loop through Rows and Apply Filter */
		for(Row r : sheet) {
		        for (Cell c : r) {
		                if (c.getColumnIndex()==1 && !c.getStringCellValue().equals(search)) {
		                	System.out.println("steps in");
		                        r1=(XSSFRow) c.getRow();
		                        //System.out.println(r1);
		                        if (r1.getRowNum()!=0) { /* Ignore top row */
		                                /* Hide Row that does not meet Filter Criteria */
		                                r1.getCTRow().setHidden(true); }
		                                }                              
		        }
		  }
		barMessage = "Filtering Done";
		barProgress = 50;
		bar.update(barMessage, barProgress, b);
		FileOutputStream out = new FileOutputStream(filePath+fileName);
		wb.write(out);
		out.close();
	}
	@SuppressWarnings({ "rawtypes", "incomplete-switch" })
	public static void readExcel(String filePath, String fileName, String sheetName) throws IOException, ExcelDataFormatError{
		InputStream ExcelFileToRead = new FileInputStream(filePath+fileName);
		XSSFWorkbook  wb = new XSSFWorkbook(ExcelFileToRead);
		
		XSSFWorkbook test = new XSSFWorkbook(); 
		
		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow row; 
		XSSFCell cell, valueCell, keyCell;

		Iterator rows = sheet.rowIterator();
		
		sheet.setAutoFilter(CellRangeAddress.valueOf("A1:E140000"));
		
		filterExcel(filePath, fileName, wb, sheet, "Test");
		
		while (rows.hasNext())
		{
			row=(XSSFRow) rows.next();
			Iterator cells = row.cellIterator();
			while (cells.hasNext())
			{
				cell=(XSSFCell) cells.next();
				valueCell = row.getCell(3);
				keyCell = row.getCell(2);
				
				String value = valueCell.getStringCellValue().trim();
				String key = keyCell.getStringCellValue().trim();
					  
				//Putting key & value in dataMap
				//dataMap.put(key, value);
					  
				//Putting dataMap to excelFileMap
				//excelFileMap.put("DataSheet", dataMap);
				  
				switch(cell.getCellType()){
				case STRING:
					if(cell.getStringCellValue().equals("Test")){
						
					}
					else if(cell.getStringCellValue().equals("Cell")){
		
					}else{
						throw new ExcelDataFormatError("Incorrect Data Format", f);
					}
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue()+" ");
					break;
				case BOOLEAN:
					System.out.print("error");
					break;
				}
			}
			barMessage = "Updates Completed";
			barProgress = 100;
			bar.update(barMessage, barProgress, b);
			System.out.println();
		}
	
	}
	
	public static void ImageLoader() throws IOException{
		ImageIcon logo = new ImageIcon("src/bell_logo.png");
		JLabel pic = new JLabel("Bell Logo");
		pic.setOpaque(true);
		pic.setIcon(logo);
		pic.setBounds(80, 0, 225, 100);
		f.add(pic, BorderLayout.SOUTH);
	}
	
	public static void FileNameCheck() throws FileNameException{
		if(textFile.getText().equals("JavaBooks.xlsx")){
			System.out.print("correct");
		}
		else{
			throw new FileNameException("Incorrect File Name - Please select correct file",f);
		}
	}
	
	
/*	public static Map<String, Map<String, String>> setMapData(String filePath, String fileName, String sheetName) throws IOException{
		FileInputStream fis = new FileInputStream(filePath+fileName);
		
		Workbook workbook = new XSSFWorkbook(fis);
		
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		
		 for(int i=0; i<=lastRow; i++){
			  
			  Row row = sheet.getRow(i);
			  
			  //1st Cell as Value
			  Cell valueCell = row.getCell(1);
				  
			  //0th Cell as Key
			  Cell keyCell = row.getCell(0);
				  
			  String value = valueCell.getStringCellValue().trim();
			  String key = keyCell.getStringCellValue().trim();
				  
			  //Putting key & value in dataMap
			  dataMap.put(key, value);
				  
			  //Putting dataMap to excelFileMap
			  excelFileMap.put("DataSheet", dataMap);
		  }
		  
		return excelFileMap; 
		
	} */
	}

