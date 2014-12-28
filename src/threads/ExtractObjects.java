package threads;

import indexing.ChemicalIndexing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Shape;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

import dto.ShapesDTO;

/*
 * This class serves the main purpose of extracting embedded objects from word and ppt files.
 * It extends Thread class and hence the codes can be executed in separate threads.
 */
public class ExtractObjects extends Thread{
	private File file;
	private ArrayList<ShapesDTO> shapesDTOList;
	
	/* Constructor to initialize the object */
	public ExtractObjects(File file){
		System.err.close();
		this.file = file;
		shapesDTOList = new ArrayList<>();
		System.out.println("Reading objects from file: "+file.getName());
	}
	
	/* Run the thread */
	@Override
	public void run() {
		/* Call two different section Apache POI based on file type */
		if(this.file.getName().endsWith(".doc") || this.file.getName().endsWith(".docx")){
			getObjectsHWPF();
		}else{
			getObjectsHSSF();
		}
		saveReport();
	}
	
	/*
	 * Generally required for files having .doc extension 
	 */
	public void getObjectsHWPF(){
		HWPFDocument doc=null;
		try {
			doc = new HWPFDocument(new FileInputStream(new File(this.file.getAbsolutePath())));
			List<org.apache.poi.hwpf.usermodel.Picture> pictureList = new ArrayList<>();
			pictureList = doc.getPicturesTable().getAllPictures();
			for(org.apache.poi.hwpf.usermodel.Picture picture : pictureList){
				ShapesDTO shapesDTO = new ShapesDTO();
		        shapesDTO.setLocation(picture.getStartOffset());
		        shapesDTO.setType(picture.getMimeType());
		        this.shapesDTOList.add(shapesDTO);
			}
		}catch(Exception e){
			System.out.println("String readig with XWPF...");
		}finally{
			if(this.shapesDTOList.isEmpty()){
				getObjectsXWPF(); //Call the different section of POI if HWPF fails
				saveReport();
			}
		}
	}
	
	/*
	 * Generally required for files having .docx extension 
	 */
	public void getObjectsXWPF(){
		XWPFDocument doc=null;
		try {
			doc = new XWPFDocument(new FileInputStream(new File(this.file.getAbsolutePath())));
			List<XWPFPictureData> pictureList = doc.getXWPFDocument().getAllPackagePictures();
			for(XWPFPictureData picture : pictureList){
				ShapesDTO shapesDTO = new ShapesDTO();
		        shapesDTO.setType(picture.getPackagePart().getContentType());
		        this.shapesDTOList.add(shapesDTO);
			}
		}catch(Exception e){
			System.out.println("Some error occured while reading file: "+this.file.getName());
		}
	}
	
	/*
	 * Save the report in destination folder with detail information about embedded objects
	 */
	public void saveReport(){
		try{
			int fileCounter = 0;
			int pos = this.file.getName().lastIndexOf(".");
			/* Get the filename with same name as source file name. If duplicate found increase the counter. */
			String reportName = pos > 0 ? this.file.getName().substring(0, pos)+"_"+this.file.getName().substring(pos+1, this.file.getName().length()) : this.file.getName();
			if(new File(ChemicalIndexing.destDir.replaceAll("/", "\\")+"\\"+reportName).exists()){
				while(new File(ChemicalIndexing.destDir.replaceAll("/", "\\")+"\\"+reportName+"_"+fileCounter).exists()){
					fileCounter++;
				}
				reportName = reportName+"_"+fileCounter;
			}
			
			Workbook wb = new HSSFWorkbook();
		    Sheet sheet1 = wb.createSheet("Embeded Objects");
		    Row row = sheet1.createRow(0);
		    row.createCell(0).setCellValue("Object Type");
		    row.createCell(1).setCellValue("Location");
		    for(ShapesDTO shapesDTO :  shapesDTOList){
		    	row = sheet1.createRow(sheet1.getLastRowNum()+1);
			    row.createCell(0).setCellValue(shapesDTO.getType());
			    row.createCell(1).setCellValue(shapesDTO.getLocation());
		    }
		    /* save the report in Excel file */
		    FileOutputStream fileOut = new FileOutputStream(ChemicalIndexing.destDir.replaceAll("/", "\\")+"\\"+reportName+".xls");
		    wb.write(fileOut);
		    fileOut.close();
		}catch(Exception e){
			System.out.println("Some error occured while saving the report for file: "+this.file.getName());
		}
	}
	
	/*
	 * Generally required for files having .ppt extension 
	 */
	public void getObjectsHSSF(){
		SlideShow ppt = null;
		try {
			ppt = new SlideShow(new HSLFSlideShow(this.file.getAbsolutePath()));
		  //get slides 
		  Slide[] slide = ppt.getSlides();
		  for (int i = 0; i < slide.length; i++){
		    Shape[] sh = slide[i].getShapes();
		    for (int j = 0; j < sh.length; j++){		        
		        ShapesDTO shapesDTO = new ShapesDTO();
		        shapesDTO.setLocation(i);
		        shapesDTO.setType(sh[j].getShapeName());
		        this.shapesDTOList.add(shapesDTO);
		    }
		  }
		}catch (IOException e) {
			System.out.println("Start reading with XSSF...");
		}
		finally{
			if(this.shapesDTOList.isEmpty()){
				System.out.println("Start reading with XSSF...");
				getObjectsXSSF();	//Call the different section of POI if HSSF fails
				saveReport();
			}
		}
	}
	
	/*
	 * Generally required for files having .pptx extension 
	 */
	public void getObjectsXSSF(){
	      File file = new File(this.file.getAbsolutePath());
	      XMLSlideShow ppt = null;
		try {
			  ppt = new XMLSlideShow(new FileInputStream(file));	      
		      //get slides 
		      XSLFSlide[] slide = ppt.getSlides();
		      //getting the shapes in the presentation
		      for (int i = 0; i < slide.length; i++){
		         XSLFShape[] sh = slide[i].getShapes();
		         for (int j = 0; j < sh.length; j++){
			        ShapesDTO shapesDTO = new ShapesDTO();
			        shapesDTO.setType(sh[j].getShapeName());
			        shapesDTO.setLocation(i);
			        this.shapesDTOList.add(shapesDTO);
		         }
		      }
		}catch(Exception e){
			System.out.println("Some error occured while reading file: "+this.file.getName());
		}
	}

}
