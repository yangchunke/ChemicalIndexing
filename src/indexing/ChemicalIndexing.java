package indexing;

import static constants.UserConstants.MAX_THREADS;

import java.io.File;
import java.io.FileFilter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Scanner;

import threads.ExtractObjects;

public class ChemicalIndexing {

	/**
	 * @param args
	 * @throws IOException 
	 */
	private static ArrayList<File> fileList=null;
	private static Thread[] threads = new Thread[MAX_THREADS];
	public static String sourceDir = null;
	public static String destDir = null;
	
	/* Main method to begin the program */
	public static void main(String[] args) throws IOException {
		System.out.print("Please provide path to source directory: ");
		@SuppressWarnings("resource")
		Scanner scanner = new Scanner(System.in);
		sourceDir = scanner.next();
		File file = new File(sourceDir);
		fileList = new ArrayList<>();
		/* Get all files recursively present in the provided directory */
		new ChemicalIndexing().getAllFiles(file);
		/* Create destination folder Report if not present */
		new ChemicalIndexing().createDestinationDir();
		
		/* Run different threads to fetch the embeded objects. MAX_THREADS can be configured in UserConstants.java */
		while(!fileList.isEmpty()){
			for(int i=0;i<MAX_THREADS;i++){
				if(threads[i]==null || !threads[i].isAlive()){
					threads[i] = new ExtractObjects(fileList.remove(0));
					threads[i].start();	//Begin the thread
				}
			}
		}
	}
	
	/*
	 * This method recursively read through all files and finds only those files which have .ppt/.pptx/.doc/.docx extension
	 */
	public void getAllFiles(File file) {
	    if(file.isDirectory()){
		    File[] children = file.listFiles(new FileFilter() {
				@Override
				public boolean accept(File pathname) {
					return (pathname.getName().endsWith(".ppt") || pathname.getName().endsWith(".pptx") || pathname.getName().endsWith(".doc") || pathname.getName().endsWith(".docx"));
				}
			});
		    for (File child : children) {
		    	getAllFiles(child); //Recursive call
		    }
	    }
	    else{
	    	fileList.add(file);
	    }
	}
	
	/*
	 * Create destination directory with current timestamp
	 */
	public void createDestinationDir(){
		String timeStamp = new SimpleDateFormat("dd-MMM-yyyy HH.mm.ss").format(Calendar.getInstance().getTime());
		if(!new File(sourceDir.replaceAll("/", "\\")+"\\Report").exists()){
			new File(sourceDir.replaceAll("/", "\\")+"\\Report").mkdir();
		}
		destDir = sourceDir.replaceAll("/", "\\")+"\\Report\\"+timeStamp;
		new File(destDir).mkdir();
	}
	
	
}
