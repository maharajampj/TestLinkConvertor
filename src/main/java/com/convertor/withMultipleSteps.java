package com.convertor;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.setUp.setUp;

public class withMultipleSteps {

	public static void main(String[] args) throws IOException {
		setUp set = new setUp();
		Properties prop = set.envSetUp();
		String suiteName = prop.getProperty("suiteName");
		String details = prop.getProperty("details");
		String excelFile = prop.getProperty("excelFileName");
		String xmlFile = prop.getProperty("xmlFileName");
		//System.out.println(readData(excelFile));
		writeXML(readData(excelFile), suiteName, details, xmlFile);

	}

	public static List<String> readData(String name) throws IOException {
		List<String> list = new ArrayList<String>();
		FileInputStream file = new FileInputStream(new File(System.getProperty("user.dir") + "//Data//" + name));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		int sheetNo=workbook.getNumberOfSheets();
		for(int i=0;i<sheetNo;i++)
		{
		XSSFSheet sheet = workbook.getSheetAt(i);
		XSSFRow row = sheet.getRow(0);
		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();
		int minRowIx = sheet.getFirstRowNum();
		int maxRowIx = sheet.getLastRowNum();

		for (int rowIx = minRowIx + 1; rowIx <= maxRowIx; rowIx++)
		{

			for (int colIx = minColIx; colIx <= maxColIx; colIx++) 
			{
				XSSFCell cell = sheet.getRow(rowIx).getCell(colIx);
			
				if(cell!=null)
				{
					if (cell.getCellType() != 3) 
					{
						if(cell.getCellType()==1)
						{
						list.add(String.valueOf(cell.getStringCellValue()));
						}
						else if(cell.getCellType()==0)
						{
							list.add(String.valueOf((Math.round(cell.getNumericCellValue()))));
						}
					} 
					else 
					{
						list.add("null");
					}

				}

			}

		}
		}
		workbook.close();
		return list;
	}

	public static void writeXML(List<String> list, String suiteName, String suitedetails, String xmlName) 
	{
		try {
			DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder documentBuilder = documentBuilderFactory.newDocumentBuilder();
			Document document = documentBuilder.newDocument();

			Element testsuite = document.createElement("testsuite");
			document.appendChild(testsuite);
			Attr attr = document.createAttribute("name");
			attr.setValue(suiteName);
			testsuite.setAttributeNode(attr);

			Element details = document.createElement("details");
			details.appendChild(document.createTextNode(suitedetails));
			testsuite.appendChild(details);

			for (int i = 0; i < list.size();) 
			{
				boolean flag = true;
				Element testcase = document.createElement("testcase");
				testsuite.appendChild(testcase);
				Attr attr1 = document.createAttribute("name");
				attr1.setValue(list.get(i));
				testcase.setAttributeNode(attr1);

				Element summary = document.createElement("summary");
				summary.appendChild(document.createTextNode(list.get(i + 1)));
				testcase.appendChild(summary);

				Element preconditions = document.createElement("preconditions");
				preconditions.appendChild(document.createTextNode(list.get(i + 2)));
				testcase.appendChild(preconditions);
				
				Element execution_type = document.createElement("execution_type");
				execution_type.appendChild(document.createTextNode(list.get(i + 3)));
				testcase.appendChild(execution_type);
				
				Element importance = document.createElement("importance");
				importance.appendChild(document.createTextNode(list.get(i + 4)));
				testcase.appendChild(importance);
				
				Element status = document.createElement("status");
				status.appendChild(document.createTextNode(list.get(i + 5)));
				testcase.appendChild(status);
				
				

				Element steps = document.createElement("steps");
				testcase.appendChild(steps);

				do {

					Element step = document.createElement("step");
					steps.appendChild(step);

					Element step_number = document.createElement("step_number");
					step_number.appendChild(document.createTextNode(list.get(i + 6)));
					step.appendChild(step_number);

					Element actions = document.createElement("actions");
					actions.appendChild(document.createTextNode(list.get(i + 7)));
					step.appendChild(actions);

					Element expectedresults = document.createElement("expectedresults");
					expectedresults.appendChild(document.createTextNode(list.get(i + 8)));
					step.appendChild(expectedresults);

					if (flag)
					{
						Element keywords = document.createElement("keywords");
						testcase.appendChild(keywords);

						Element keyword1 = document.createElement("keyword");
						keywords.appendChild(keyword1);
						Attr attr3 = document.createAttribute("name");
						attr3.setValue(list.get(i + 9));
						keyword1.setAttributeNode(attr3);
						
						Element keyword2 = document.createElement("keyword");
						keywords.appendChild(keyword2);
						Attr attr4 = document.createAttribute("name");
						attr4.setValue(list.get(i + 10));
						keyword2.setAttributeNode(attr4);
						
						Element keyword3 = document.createElement("keyword");
						keywords.appendChild(keyword3);
						Attr attr5 = document.createAttribute("name");
						attr5.setValue(list.get(i + 11));
						keyword3.setAttributeNode(attr5);

					}
					flag = false;
					i = i + 12;
					
					 if (list.get(i).equals("break")) 
					 {
						 break; 
					 }
					 

				} while ((list.get(i) + list.get(i + 1) + list.get(i + 2)).equalsIgnoreCase("nullnullnull"));

				
				 if (list.get(i).equals("break")) 
				 {
					 break; 
				 }
				 

			}

			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			DOMSource source = new DOMSource(document);
			StreamResult streamResult = new StreamResult(
					new File(System.getProperty("user.dir") + "//Data//" + xmlName));
			transformer.transform(source, streamResult);
			System.out.println("Test Cases Written Successfully in " + xmlName);

		} catch (Exception e) {
			System.out.println("Test Cases are Failed to Written in " + xmlName);
			e.printStackTrace();
		}

	}

}
