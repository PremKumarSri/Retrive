package com.candidjava;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataReterive {

	public static void main(String[] args) throws Throwable {
		List<String> datafromExcel = getDatafromExcel("Number", "76543212");
		System.out.println(datafromExcel);
	}

	public static List<String> getDatafromExcel(String filter, String nameNum) throws Throwable {
		try {
			File f = new File(System.getProperty("user.dir") + "\\Excel\\Test.xlsx");
			FileInputStream fin = new FileInputStream(f);
			File of = new File(System.getProperty("user.dir") + "\\Excel\\Output.xlsx");
			Workbook oxw = new XSSFWorkbook(new FileInputStream(of));
			Sheet sheetAt = oxw.getSheetAt(0);
			for (int i = 1; i < sheetAt.getPhysicalNumberOfRows(); i++) {
				Row row = sheetAt.getRow(i);
				for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
					Cell cell = row.getCell(j);
					cell.setCellValue("");
				}
			}
			oxw.write(new FileOutputStream(of));
			Workbook xw = new XSSFWorkbook(fin);
			int ocount = 0;
			List<String> lidata = new ArrayList<String>();
			for (int s = 0; s < xw.getNumberOfSheets(); s++) {
				Sheet sheet = xw.getSheetAt(s);
				if (sheet.getSheetName().equals("Expense") || sheet.getSheetName().equals("Trainer Payment")) {
					continue;
				}
				for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
					Row row = sheet.getRow(i);
					if (row == null) {
						continue;
					}
					Cell cell = null;
					if (filter.equalsIgnoreCase("Name")) {
						cell = row.getCell(1);
					} else if (filter.equalsIgnoreCase("Number")) {
						cell = row.getCell(2);

					} else {
						System.out.println("Wrong filter");
					}
					if (cell == null) {
						continue;
					}
					CellType cellType = cell.getCellType();
					String headerValue = null;
					if (cellType.equals(CellType.STRING)) {
						headerValue = cell.getStringCellValue();

					} else if (cellType.equals(CellType.NUMERIC)) {
						double numericCellValue = cell.getNumericCellValue();
						long l = (long) numericCellValue;
						headerValue = String.valueOf(l);

					} else if (cellType.equals(CellType.FORMULA)) {
						double numericCellValue = cell.getNumericCellValue();
						long l = (long) numericCellValue;
						headerValue = String.valueOf(l);

					} else if (cellType.equals(CellType.BLANK)) {
						continue;
					}

					if (headerValue.equalsIgnoreCase(nameNum)) {

						Cell date = row.getCell(0);
						CellType dateType = date.getCellType();
						String dateValue = null;
						if (dateType.equals(CellType.NUMERIC)) {
							Date dateCellValue = date.getDateCellValue();
							dateValue = dateCellValue.toString().substring(0, 10);
						}

						Cell name = row.getCell(1);
						CellType nameType = name.getCellType();
						String nameValue = null;
						if (nameType.equals(CellType.STRING)) {
							nameValue = name.getStringCellValue();
						} else if (nameType.equals(CellType.NUMERIC)) {
							double numericCellValue = name.getNumericCellValue();
							long l = (long) numericCellValue;
							nameValue = String.valueOf(l);
						} else if (cellType.equals(CellType.BLANK)) {
							continue;
						}
						Cell contact = row.getCell(2);
						CellType contactType = contact.getCellType();
						String contactValue = null;
						if (contactType.equals(CellType.STRING)) {
							contactValue = contact.getStringCellValue();
						} else if (contactType.equals(CellType.NUMERIC)) {
							double numericCellValue = contact.getNumericCellValue();
							long l = (long) numericCellValue;
							contactValue = String.valueOf(l);
						} else if (cellType.equals(CellType.BLANK)) {
							continue;
						}
						Cell course = row.getCell(3);
						CellType courseType = course.getCellType();
						String courseValue = null;
						if (courseType.equals(CellType.STRING)) {
							courseValue = course.getStringCellValue();
						} else if (courseType.equals(CellType.NUMERIC)) {
							double numericCellValue = course.getNumericCellValue();
							long l = (long) numericCellValue;
							courseValue = String.valueOf(l);
						} else if (cellType.equals(CellType.BLANK)) {
							continue;
						}

						Cell totalfee = row.getCell(4);
						CellType totalfeeType = totalfee.getCellType();
						String totalfeeValue = null;
						if (totalfeeType.equals(CellType.STRING)) {
							totalfeeValue = totalfee.getStringCellValue();
						} else if (totalfeeType.equals(CellType.NUMERIC)) {
							double numericCellValue = totalfee.getNumericCellValue();
							long l = (long) numericCellValue;
							totalfeeValue = String.valueOf(l);
						} else if (cellType.equals(CellType.BLANK)) {
							continue;
						}
						Cell feepaid = row.getCell(5);
						CellType feepaidType = feepaid.getCellType();
						String feepaidValue = null;
						if (feepaidType.equals(CellType.STRING)) {
							feepaidValue = feepaid.getStringCellValue();
						} else if (feepaidType.equals(CellType.NUMERIC)) {
							double numericCellValue = feepaid.getNumericCellValue();
							long l = (long) numericCellValue;
							feepaidValue = String.valueOf(l);
						} else if (cellType.equals(CellType.BLANK)) {
							continue;
						}
						Cell balance = row.getCell(6);
						CellType balanceType = balance.getCellType();
						String balanceValue = null;
						if (balanceType.equals(CellType.STRING)) {
							balanceValue = balance.getStringCellValue();
						} else if (balanceType.equals(CellType.NUMERIC)) {
							double numericCellValue = balance.getNumericCellValue();
							long l = (long) numericCellValue;
							balanceValue = String.valueOf(l);
						} else if (balanceType.equals(CellType.FORMULA)) {
							double numericCellValue = balance.getNumericCellValue();
							long l = (long) numericCellValue;
							balanceValue = String.valueOf(l);
						} else if (cellType.equals(CellType.BLANK)) {
							continue;
						}

						ocount++;
						Row Createrow = oxw.getSheetAt(0).createRow(ocount);

						oxw.getSheetAt(0).getRow(ocount).createCell(1).setCellValue(nameValue);

						System.out.println("Name :" + nameValue);
						oxw.getSheetAt(0).getRow(ocount).createCell(0).setCellValue(dateValue.toString());
						System.out.println("Date :" + dateValue);
						oxw.getSheetAt(0).getRow(ocount).createCell(2).setCellValue(contactValue);

						System.out.println("Contact :" + contactValue);
						oxw.getSheetAt(0).getRow(ocount).createCell(3).setCellValue(courseValue);

						System.out.println("Course :" + courseValue);
						oxw.getSheetAt(0).getRow(ocount).createCell(4).setCellValue(totalfeeValue);

						System.out.println("Total Fee :" + totalfeeValue);
						oxw.getSheetAt(0).getRow(ocount).createCell(5).setCellValue(feepaidValue);

						System.out.println("Fee Paid :" + feepaidValue);
						oxw.getSheetAt(0).getRow(ocount).createCell(6).setCellValue(balanceValue);

						System.out.println("Balance :" + balanceValue);
						System.out.println("Write completed");

						lidata.add(dateValue + "&" + nameValue + "&" + contactValue + "&" + courseValue + "&"
								+ totalfeeValue + "&" + feepaidValue + "&" + balanceValue);

					}
				}

			}
			FileOutputStream fout = new FileOutputStream(of);
			oxw.write(fout);
			oxw.close();
			return lidata;
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			throw new Exception();
		}
	}
}
