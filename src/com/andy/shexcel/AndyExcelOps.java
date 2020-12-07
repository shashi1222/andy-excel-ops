package com.andy.shexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.reflect.generics.tree.Tree;

public class AndyExcelOps
{
	private static Map<String, Object[]> shdata = new TreeMap<String, Object[]>();
	public static void main(String[] args) 
	{
		try
		{
			/*Enter the file name here - with extension i.e. .xlsx - see sample code below*/
			/*FileInputStream file = new FileInputStream(new File("APUnit5cAfter.xlsx"));*/
			FileInputStream file = new FileInputStream(new File("APUnit5cAfterIn.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			System.out.println("Opened file" + file);

			XSSFSheet sheet = workbook.getSheet("Sheet1");
			/*XSSFSheet sheetWrite = workbook.createSheet("shresult");*/

			System.out.println("Opened Sheet" + sheet);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			Integer index=0;
			int i = 0;
			int singleEntryCount = 0;
			Map<Integer, Integer> calcTime= new HashMap<>();
			Map<Integer, Double> calcMarks = new HashMap<>();

			int totalSec = 0;
			double totalMarks = 0;
			String currName, prevName = null;

			while (rowIterator.hasNext()) 
			{
				Row row = rowIterator.next();

				if(i > 0) {

					if(i==1) {
						currName = prevName = row.getCell(0).getStringCellValue() +
								row.getCell(1).getStringCellValue();
						shdata.put(index.toString(), new Object[] {"Name", "TotalTime", "Last5QTime",
								"Last5QAve", "Last5Qx0.5"});
						index++;
					}

					if(row.getFirstCellNum() == 0 &&
							row.getCell(4, Row.RETURN_BLANK_AS_NULL) != null ) {
						currName = row.getCell(0).getStringCellValue() +
								row.getCell(1).getStringCellValue();
						String quizDate = null;
						double shMarks = 0;
						if (row.getCell(4, Row.RETURN_BLANK_AS_NULL) != null) {
							quizDate = row.getCell(4).getDateCellValue().toString();
							shMarks = row.getCell(2).getNumericCellValue();
							/*System.out.println("Marks is: " + shMarks);*/
						}
						if(currName.equals(prevName)) {

							totalSec = totalSec + calculateSec(quizDate);
							totalMarks = totalMarks + shMarks;
							singleEntryCount++;
							calcMarks.put(singleEntryCount, shMarks);
							calcTime.put(singleEntryCount, calculateSec(quizDate));
						} else {
							int effective_count = 0;
							int start_index = 0;

							int[] startIndexArr = {start_index};

							effective_count = calcStartIndexAndEffectiveCount(
									singleEntryCount,
									startIndexArr);
							start_index = startIndexArr[0];

							System.out.println("start index and effect count is: "
							+ start_index + "," + effective_count);

							int total_effective_time = 0;
							int total_effective_marks = 0;
							for (int j =start_index; j< (start_index + effective_count); j++) {

								total_effective_marks+= calcMarks.get(j);
								total_effective_time+= calcTime.get(j);
							}
							float marks_avg = (float)total_effective_marks/50;
							System.out.println("marks avg is: " + marks_avg);
							System.out.println("Andy expected: " + marks_avg/2);
							System.out.println("Total Count " + singleEntryCount + "\t Name:"+ prevName + "\t Total Min:"
									+ totalSec/60 + " \t and Sec:" + totalSec%60 + "\t and Marks:" + totalMarks);
							System.out.println("Effective Count " + effective_count + "\t Name:"+ prevName + "\t Avg Min:"
									+ total_effective_time/60 + " \t and Sec:" + total_effective_time%60 + "\t and Avg Marks:" + marks_avg);
							/*shdata.put(index.toString(), new Object[] {prevName, totalSec/60, totalSec%60});*/
							shdata.put(index.toString(), new Object[] {prevName, totalSec/60 + ":" + totalSec%60,
							total_effective_time/60 + ":" + total_effective_time%60,
							marks_avg, marks_avg/2});
							index++;
							singleEntryCount = 1;

							System.out.println("Starting for Name" + "\t" + currName);
							totalSec = calculateSec(quizDate);
							totalMarks = shMarks;
							calcMarks.put(singleEntryCount, shMarks);
							calcTime.put(singleEntryCount, calculateSec(quizDate));
							prevName = currName;
						}

					} else {

					}
				}
				i++;
			}

			// this is for the last value..
			 System.out.println("Total Count" + singleEntryCount + "\t Name:" + prevName + "\t Total Min:" +
					 totalSec/60 + " \t and Sec:" + totalSec%60 + "\t and Marks:" + totalMarks);
			int effective_count = 0;
			int start_index = 0;

			int[] startIndexArr = {start_index};

			effective_count = calcStartIndexAndEffectiveCount(
					singleEntryCount,
					startIndexArr);
			start_index = startIndexArr[0];

			System.out.println("start index and effect count is: "
					+ start_index + "," + effective_count);

			int total_effective_time = 0;
			int total_effective_marks = 0;
			for (int j =start_index; j< (start_index + effective_count); j++) {

				total_effective_marks+= calcMarks.get(j);
				total_effective_time+= calcTime.get(j);
			}
			float marks_avg = (float)total_effective_marks/50;
			System.out.println("marks avg is: " + marks_avg);
			System.out.println("Andy expected: " + marks_avg/2);
			System.out.println("Effective Count " + effective_count + "\t Name:"+ prevName + "\t Avg Min:"
					+ total_effective_time/60 + " \t and Sec:" + total_effective_time%60 + "\t and Avg Marks:" + marks_avg);

			shdata.put(index.toString(), new Object[] {prevName, totalSec/60 + ":" + totalSec%60,
					total_effective_time/60 + ":" + total_effective_time%60,
					marks_avg, marks_avg/2});
			writeToExcelFile();
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}

	private static int calcStartIndexAndEffectiveCount(int singleEntryCount, int[] startIndexArr) {

		int temp_effective_count = 0;
		if(singleEntryCount <= 10) {
			// make both as 0;
		}
		else if(singleEntryCount > 10 && singleEntryCount < 15){
			startIndexArr[0] = 11;
			temp_effective_count = singleEntryCount - 10;

		} else if(singleEntryCount >= 15) {
			startIndexArr[0] = (singleEntryCount - 5) + 1;
			temp_effective_count = 5;
		}
		return temp_effective_count;
	}

	private static void writeToExcelFile() {

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheetWrite = workbook.createSheet("shresult");
		Set<String> shkeyset = shdata.keySet();
		int rownum = 0;
		for (String key : shkeyset)
		{
			Row row = sheetWrite.createRow(rownum++);
			Object [] objArr = shdata.get(key);
			int cellnum = 0;

			for (Object obj : objArr)
			{
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
				else if(obj instanceof Float)
					cell.setCellValue((Float)obj);
			}
		}

		try
		{
			//Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File("APUnit5cAfterOut.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("APUnit5cAfter.xlsx written successfully on disk.");
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	private static int calculateSec(String quizDate) {

		String[] strArr = new String[5];
		strArr = quizDate.split(":");
		int min = Integer.parseInt(strArr[1]);
		int sec;
		sec = Integer.parseInt(strArr[2].substring(0,2));
		return (min*60+sec);
	}
}
