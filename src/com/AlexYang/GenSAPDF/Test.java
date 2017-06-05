package com.AlexYang.GenSAPDF;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Test {
	
	private static ModifyWordTemplate template;
	private static final String DOCX_MODEL_PATH = "doc/Sales Advice Form Template.docx";
	private static final String DOCX_FILE_WRITE = "doc/output.docx";
	private static Map<String,String> map = new HashMap<String, String>();

	public static void init(){
		File file = new File(DOCX_MODEL_PATH);
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream(file);
			template = new ModifyWordTemplate(fileInputStream);
		} catch(IOException exception){
			exception.printStackTrace();
		}
		
		//Data test
		map.put("Filenumber", "12345");
		map.put("FirstName1", "San");
		map.put("FirstName2", "Si");
		map.put("Surname1", "Zhang");
		map.put("Surname2", "Li");
		map.put("Date1", "2");
		map.put("Month1", "10");
		map.put("Year1", "1987");
        map.put("Date2", "22");
        map.put("Month2", "1");
        map.put("Year2", "1977");
        map.put("Mobile1", "0451737373");
        map.put("Mobile2", "0412732373");
		map.put("Address1", "202 Priss Road, Wynyard, NSW 2500");
        map.put("Address2", "212 Pxcvriss Road, Rodbes, NSW 2500");
        map.put("Email1", "ZhangSan@aabc.com");
        map.put("Email2", "LiSi@aabc.com");
        map.put("LocalResident1", "Local");
        map.put("LocalResident2", "PR");
        map.put("ProjectName", "ABC Project");
        map.put("ParkingNo", "33");
        map.put("Address", "1233 Naval Road, Parsley, NSW 2567");
        map.put("StorageNo", "33");
        map.put("Price", "1230000");
        map.put("PurchaserType", "Cash");
        map.put("ApartNo", "33");
        map.put("Notes", "This is a test notes.This is a test notes.This is a test notes.This is a test notes.");
        map.put("SolicitorName", "LvShi");
        map.put("LawFirm", "NewLvshihang");
        map.put("SolicitorMobile", "0412345678");
        map.put("SolicitorEmail", "LvShi@NewLvshihang.com");
        map.put("SolicitorAddress", "3 Hahaha Road, Rodbes, NSW 2500");
        map.put("Reservationfee", "5000");
        map.put("ModeofPayment", "Cash");
	}
	
	public static void testReplaceTag()
	{
		template.replaceTag(map);
	}
	

	public static void templateWrite()
    {
		testReplaceTag();
		File file = new File(DOCX_FILE_WRITE);
		FileOutputStream out;
		try {
			out = new FileOutputStream(file);
			BufferedOutputStream bos = new BufferedOutputStream(out);
			template.write(bos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

    public static void main(String[] args)
    {
	    init();
	    templateWrite();
	    Docx2PDF.docx2PDF();
    }
}
