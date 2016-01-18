package com.francis.excel;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * Hello world!
 *
 */
public class SignIn {
	public static void main(String[] args) {

		String path = "D:\\考勤整理\\12月份8楼原始数据.xls";
		// String path = "E:\\7-11月考勤明细\\7-11月考勤明细\\9月\\9月8楼原始1.xls";

		try {
			readXls(path);

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void readXls(String path) throws Exception {

		InputStream is = new FileInputStream(path);

		HSSFWorkbook hssfworkbook = new HSSFWorkbook(is);

		int t1 = 0;
		int t2 = 0;
		int t3 = 0;
		int t4 = 0;
		int t5 = 0;
		int t6 = 0;
		int t7 = 0;

		int tt1 = 0;
		int tt2 = 0;
		int tt3 = 0;
		int tt4 = 0;
		int tt5 = 0;
		int tt6 = 0;
		int tt7 = 0;
		int tt8 = 0;
		int tt9 = 0;
		int tt10 = 0;
		int tt11 = 0;
		int tt12 = 0;
		int tt13 = 0;

		int tt15 = 0; // 加班至凌晨

		String lastName = "王伟鹏";

		System.out.println("姓名" + "\t" + "8:00-8：30" + "\t" + "8:30-9：00"
				+ "\t" + "9:00-9:30" + "\t" + "9:30-10:00" + "\t"
				+ "10:00-10:30" + "\t" + "10:30-11:00" + "\t" + "11点以后" + "\t"
				+ "18:00-18:30" + "\t" + "18:30-19:00" + "\t" + "19:00-19:30"
				+ "\t" + "19:30-20:00" + "\t" + "20:00-20:30" + "\t"
				+ "20:30-21:00" + "\t" + "21:00-21:30" + "\t" + "21:30-22:00"
				+ "\t" + "22:00-22:30" + "\t" + "22:30-23:00" + "\t"
				+ "23:00-23:30" + "\t" + "23:30-24:00" + "\t" + "24点后");

		// List<List<String>> result = new ArrayList<List<String>>();

		for (int numsheet = 0; numsheet < hssfworkbook.getNumberOfSheets(); numsheet++) {

			HSSFSheet hssfsheet = hssfworkbook.getSheetAt(numsheet);

			if (hssfsheet == null) {
				continue;
			}

			for (int rownum = 0; rownum < hssfsheet.getLastRowNum(); rownum++) {
				HSSFRow hssfrow = hssfsheet.getRow(rownum);
				int minColIx = hssfrow.getFirstCellNum();
				int maxColIx = hssfrow.getLastCellNum();
				List<String> rowlist = new ArrayList<String>();

				HSSFCell hssFcell0 = hssfrow.getCell(0);

				String curname = getStringVal(hssFcell0);

				if (!curname.equals(lastName)) {

					System.out.println(lastName + "\t" + t1 + "\t" + t2 + "\t"
							+ t3 + "\t" + t4 + "\t" + t5 + "\t" + t6 + "\t"
							+ t7 + "\t" + tt1 + "\t" + tt2 + "\t" + tt3 + "\t"
							+ tt4 + "\t" + tt5 + "\t" + tt6 + "\t" + tt7 + "\t"
							+ tt8 + "\t" + tt9 + "\t" + tt10 + "\t" + tt11
							+ "\t" + tt12 + "\t" + tt13

					);

					t1 = 0;
					t2 = 0;
					t3 = 0;
					t4 = 0;
					t5 = 0;
					t6 = 0;
					t7 = 0;

					tt1 = 0;
					tt2 = 0;
					tt3 = 0;
					tt4 = 0;
					tt5 = 0;
					tt6 = 0;
					tt7 = 0;
					tt8 = 0;
					tt9 = 0;
					tt10 = 0;
					tt11 = 0;
					tt12 = 0;
					tt13 = 0;

				}

				HSSFCell hssFcell1 = hssfrow.getCell(1);
				if (hssFcell1 != null && !"".equals(hssFcell1)) {
					String uptime = getStringVal(hssFcell1);

					if ("加班至凌晨".equals(uptime) || "年中会".equals(uptime)) {
						tt15++;
					}

					if (!"".equals(uptime) && !"忘打卡".equals(uptime)
							&& !"加班至凌晨".equals(uptime) && !"年中会".equals(uptime)) {

						int hour = new SimpleDateFormat("HH:mm:ss").parse(
								uptime).getHours();

						int mins = new SimpleDateFormat("HH:mm:ss").parse(
								uptime).getMinutes();

						if (hour == 8 && mins <= 30) {
							t1++;
						} else if (hour == 8 && mins > 30) {
							t2++;
						} else if (hour == 9 && mins <= 30) {
							t3++;
						} else if (hour == 9 && mins > 30) {
							t4++;
						} else if (hour == 10 && mins <= 30) {
							t5++;
						} else if (hour == 10 && mins > 30) {
							t6++;
						} else if (hour == 11 || hour == 12 || hour == 13
								|| hour == 14 || hour == 15 || hour == 16
								|| hour == 17) {
							t7++;
						}

					}

				}

				HSSFCell hssFcell2 = hssfrow.getCell(2);
				if (hssFcell2 != null && !"".equals(hssFcell2)) {
					String downtime = getStringVal(hssFcell2);

					if ("加班至凌晨".equals(downtime) || "年中会".equals(downtime)) {
						tt15++;
					}

					if (!"".equals(downtime) && !"忘打卡".equals(downtime)
							&& !"加班至凌晨".equals(downtime)
							&& !"年中会".equals(downtime)) {

						int hour = new SimpleDateFormat("HH:mm:ss").parse(
								downtime).getHours();

						int mins = new SimpleDateFormat("HH:mm:ss").parse(
								downtime).getMinutes();

						if (hour == 18 && mins <= 30) {
							tt1++;
						} else if (hour == 18 && mins > 30) {
							tt2++;
						} else if (hour == 19 && mins <= 30) {
							tt3++;
						} else if (hour == 19 && mins > 30) {
							tt4++;
						} else if (hour == 20 && mins <= 30) {
							tt5++;
						} else if (hour == 20 && mins > 30) {
							tt6++;
						} else if (hour == 21 && mins <= 30) {
							tt7++;
						} else if (hour == 21 && mins > 30) {
							tt8++;
						} else if (hour == 22 && mins <= 30) {
							tt9++;
						} else if (hour == 22 && mins > 30) {
							tt10++;
						} else if (hour == 23 && mins <= 30) {
							tt11++;
						} else if (hour == 23 && mins > 30) {
							tt12++;
						} else if (hour == 00 || hour == 01 || hour == 02
								|| hour == 03 || hour == 04 || hour == 05
								|| hour == 06 || hour == 07) {
							tt13++;
						}

					}

				}

				lastName = curname;

				// System.out.println(getStringVal(hssFcell));

				// for (int colIx = minColIx; colIx < maxColIx; colIx++) {
				// HSSFCell hssFcell = hssfrow.getCell(colIx);
				// if (hssFcell == null) {
				// System.out.println("null");
				// }else{
				//
				// System.out.println(getStringVal(hssFcell));
				// }
				//
				// //rowlist.add(getStringVal(hssFcell));
				// }

				// result.add(rowlist);
			}

		}

		// return result;

		System.out.println(tt15);

	}

	public static String getStringVal(HSSFCell hssfcell) {
		switch (hssfcell.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN:
			return hssfcell.getBooleanCellValue() ? "True" : "FALSE";
		case Cell.CELL_TYPE_FORMULA:
			return hssfcell.getCellFormula();
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(hssfcell)) {
				DateFormat sdf = new SimpleDateFormat("HH:mm:ss");
				return sdf.format(HSSFDateUtil.getJavaDate(hssfcell
						.getNumericCellValue()));
			}
			hssfcell.setCellType(Cell.CELL_TYPE_STRING);
			return hssfcell.getStringCellValue();
		case Cell.CELL_TYPE_STRING:
			return hssfcell.getStringCellValue();

		default:
			return "";
		}
	}

}
