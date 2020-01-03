package com.exe.fau;

import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JTextField;
import javax.swing.TransferHandler;

//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.exe.entity.Person;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class FauCalculator {
	private static JTextField field;
	private static String filepath;
	private static IntStream intStream = IntStream.range(300000, 305000);
	private static int[] allCompanyCode = intStream.toArray();
	private static Map<Integer, List<Person>> allPeople = new HashMap<>();

	static JFrame frame = new JFrame();
	public static void CopyPathToTextField() {
		frame.setTitle("��ק�ļ����ı�����ʾ�ļ�·��");
		frame.setSize(500, 300);
		frame.setLocationRelativeTo(null);
		frame.setLayout(null);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		field = new JTextField();
		field.setBounds(50, 50, 300, 30);

		field.setTransferHandler(new TransferHandler() {
			private static final long serialVersionUID = 1L;

			@Override
			public boolean importData(JComponent comp, Transferable t) {
				try {
					Object o = t.getTransferData(DataFlavor.javaFileListFlavor);

					filepath = o.toString();
					if (filepath.startsWith("[")) {
						filepath = filepath.substring(1);
					}
					if (filepath.endsWith("]")) {
						filepath = filepath.substring(0, filepath.length() - 1);
					}
					// System.out.println("CopyPathToTheField" + filepath);
					field.setText(filepath);
					return true;
				} catch (Exception e) {
					e.printStackTrace();
				}
				return false;
			}

			@Override
			public boolean canImport(JComponent comp, DataFlavor[] flavors) {
				for (int i = 0; i < flavors.length; i++) {
					if (DataFlavor.javaFileListFlavor.equals(flavors[i])) {
						return true;
					}
				}
				return false;
			}
		});

		frame.add(field);
		frame.setVisible(true);
	}

	public static List readExcel(File file) {
		try {
			// ��������������ȡExcel
			InputStream is = new FileInputStream(file.getAbsolutePath());
			// jxl�ṩ��Workbook��
			Workbook wb = Workbook.getWorkbook(is);
			// Excel��ҳǩ����
			// int sheet_size = wb.getNumberOfSheets();
			int sheet_size = 1;
			for (int index = 0; index < sheet_size; index++) {
				List<List> outerList = new ArrayList<List>();
				// ÿ��ҳǩ����һ��Sheet����
				Sheet sheet = wb.getSheet(index);
				// sheet.getRows()���ظ�ҳ��������
				for (int i = 0; i < sheet.getRows(); i++) {
					List innerList = new ArrayList();
					// sheet.getColumns()���ظ�ҳ��������
					for (int j = 0; j < sheet.getColumns(); j++) {
						String cellinfo = sheet.getCell(j, i).getContents();
						if (cellinfo.isEmpty()) {
							innerList.add("");
						} else {
							innerList.add(cellinfo);
						}
					}
					outerList.add(i, innerList);
				}
				// System.out.println(outerList);
				return outerList;
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}
	
//	public static List readExcel(File file) {
//		try {
//			InputStream xlsxFile = new FileInputStream(file.getAbsolutePath());
//			XSSFWorkbook workbook = new XSSFWorkbook(xlsxFile);
//			
//			List<List> outerList = new ArrayList<List>();
//			XSSFSheet sheet = workbook.getSheetAt(0);
//			
//			int rows = sheet.getLastRowNum() + 1;
//			int cols = 7;
//			
//			for (int i = 0; i < rows; i++) {
//				XSSFRow row = sheet.getRow(i);
//				List innerList = new ArrayList();
//				for (int j = 0; j < cols; j++) {
//					XSSFCell cell = row.getCell(j);
//					cell.setCellType(Cell.CELL_TYPE_STRING);
//					String cellValue = cell.getStringCellValue();
//					if (cellValue.isEmpty()) {
//						innerList.add("");
//					} else {
//						innerList.add(cellValue);
//					}
//				}
//				outerList.add(i, innerList);
//			}
////			System.out.println(outerList);
//			return outerList;
//		} catch (FileNotFoundException e) {
//			e.printStackTrace();
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
//		return null;
//	}

	public static void insert_in(String filePath) {
		File file = new File(filePath); // ��Ҫ������excel��,�ļ���׺������Ϊxlsx
		List excelList = readExcel(file);

		for (int i = 0; i < allCompanyCode.length; i++) { // ��ʼ��allPeople
			List<Person> persons = new ArrayList<>();
			allPeople.put(allCompanyCode[i], persons);
		}

		for (int i = 0; i < excelList.size(); i++) { // ѭ������List�е�����
			List list = (List) excelList.get(i);
			int j = 0;
			if (j < list.size()) {
				Person person = new Person(); // �½�һ�����󣬽�����ѭ����ֵ������
				person.setCompanyCode(Integer.parseInt(list.get(0).toString()));
				person.setEndDate(list.get(1).toString());
				person.setName(list.get(2).toString());

				if (list.get(3).equals("��")) { // ����Ů�Ա�ת��Ϊ0��1���������㷽��Ҫ����ԡ�2
					person.setSex(div(Double.parseDouble("1"), Math.sqrt(2)));
				} else if (list.get(3).equals("Ů")) {
					person.setSex(div(Double.parseDouble("0"), Math.sqrt(2)));
				} else {
					allPeople.remove(person.getCompanyCode());
					allCompanyCode[codePosition(person.getCompanyCode())] = 0;
					continue;
				}

				if (list.get(4).toString().equals("")) { // �������䣬�������㷽��Ҫ�����10
					allPeople.remove(person.getCompanyCode());
					allCompanyCode[codePosition(person.getCompanyCode())] = 0;
					continue;
				} else {
					person.setAge(div(Double.parseDouble(list.get(4).toString()), 10));
				}

				if (list.get(5).toString().equals("")) { // ����ѧ�����������㷽��Ҫ����ԡ�2
					allPeople.remove(person.getCompanyCode());
					allCompanyCode[codePosition(person.getCompanyCode())] = 0;
					continue;
				} else {
					person.setEducation(div(Double.parseDouble(list.get(5).toString()), Math.sqrt(2)));
				}

				if (list.get(6).toString().equals("")) { // ���빤�����ޣ��������㷽��Ҫ�����12
					allPeople.remove(person.getCompanyCode());
					allCompanyCode[codePosition(person.getCompanyCode())] = 0;
					continue;
				} else {
					person.setWorkMonth(div(Double.parseDouble(list.get(6).toString()), 12));
				}

				if (codeExist(person.getCompanyCode())) {
					allPeople.get(person.getCompanyCode()).add(person);
					// System.out.println(person);
				}
				j = 2147483647;
			}
		}
		// for (Object key : allPeople.keySet()) {
		// System.out.println(key + ": " + allPeople.get(key));
		// }
	}

	public static boolean codeExist(int code) {
		for (int i : allCompanyCode) {
			if (code == i) {
				return true;
			}
		}
		return false;
	}

	public static int codePosition(int code) {
		for (int i = 0; i < allCompanyCode.length; i++) {
			if (code == allCompanyCode[i]) {
				return i;
			}
		}
		return 0;
	}

	/**
	 * ��double���ݽ���ȡ����.
	 * 
	 * @param value
	 *            double����.
	 * @param scale
	 *            ����λ��(������С��λ��).
	 * @param roundingMode
	 *            ����ȡֵ��ʽ.
	 * @return ���ȼ���������.
	 */
	public static double round(double value, int scale, int roundingMode) {
		BigDecimal bd = new BigDecimal(value);
		bd = bd.setScale(scale, roundingMode);
		double d = bd.doubleValue();
		bd = null;
		return d;
	}

	/**
	 * double ���
	 * 
	 * @param d1
	 * @param d2
	 * @return
	 */
	public static double sum(double d1, double d2) {
		BigDecimal bd1 = new BigDecimal(Double.toString(d1));
		BigDecimal bd2 = new BigDecimal(Double.toString(d2));
		return bd1.add(bd2).doubleValue();
	}

	/**
	 * double ���
	 * 
	 * @param d1
	 * @param d2
	 * @return
	 */
	public static double sub(double d1, double d2) {
		BigDecimal bd1 = new BigDecimal(Double.toString(d1));
		BigDecimal bd2 = new BigDecimal(Double.toString(d2));
		return bd1.subtract(bd2).doubleValue();
	}

	/**
	 * double �˷�
	 * 
	 * @param d1
	 * @param d2
	 * @return
	 */
	public static double mul(double d1, double d2) {
		BigDecimal bd1 = new BigDecimal(Double.toString(d1));
		BigDecimal bd2 = new BigDecimal(Double.toString(d2));
		return bd1.multiply(bd2).doubleValue();
	}

	/**
	 * double ����
	 * 
	 * @param d1
	 * @param d2
	 * @return
	 */
	public static double div(double d1, double d2) {
		// ��Ȼ�ڴ�֮ǰ����Ҫ�жϷ�ĸ�Ƿ�Ϊ0��
		// Ϊ0����Ը���ʵ����������Ӧ�Ĵ���
		if (d2 == 0) {
			return 0;
		}
		BigDecimal bd1 = new BigDecimal(Double.toString(d1));
		BigDecimal bd2 = new BigDecimal(Double.toString(d2));
		return bd1.divide(bd2, 5, BigDecimal.ROUND_HALF_UP).doubleValue();
	}

	public static double avg(double[] ds) {
		BigDecimal sum = new BigDecimal(0);
		BigDecimal length = new BigDecimal(ds.length);
		for (double d : ds) {
			BigDecimal bd = new BigDecimal(d);
			sum = sum.add(bd);
		}
		return sum.divide(length, 5, BigDecimal.ROUND_HALF_UP).doubleValue();
	}

	/**
	 * ͨ���ݹ�ķ�ʽ��������ϴ�����Ӧ�����ݽṹ
	 * 
	 * @param dataList����ʼ����
	 * @param dataIndex����ʼ������ʼ�±�
	 * @param result����ʼ�������
	 * @param resultIndex����ʼ����������ʼ�±�
	 * @param resultList:
	 *            ������е������ʽ
	 */
	public static void combinationSelect(int[] dataList, int dataIndex, int[] result, int resultIndex,
			List<int[]> resultList) {
		int resultLen = result.length;
		int resultCount = resultIndex + 1;
		if (resultCount > resultLen) { // ȫ��ѡ����ʱ�������Ͻ��
			int[] addToList = Arrays.copyOf(result, resultLen);
			// System.out.println(Arrays.toString(result));
			resultList.add(addToList);
			return;
		}

		// �ݹ�ѡ����һ��
		for (int i = dataIndex; i < dataList.length + resultCount - resultLen; i++) {
			result[resultIndex] = dataList[i];
			combinationSelect(dataList, i + 1, result, resultIndex + 1, resultList);
		}
	}

	public static boolean isInGroup1(int m, int[] a) {
		for (int i : a) {
			if (m == i) {
				return true;
			}
		}
		return false;
	}

	public static double subgroupCharacteristicAverage(double[] datas) {
		double avg_Data;
		double sumData = 0;
		for (double d : datas) {
			sumData = sum(sumData, d);
		}
		avg_Data = div(sumData, datas.length);
		return avg_Data;
	}

	public static double betweenGroupSumOfSquaresForCharacteristic(double subgroupCharacteristicAverage,
			double avgOfCurrentData, int sizeOfCurrentGroup) {
		double subNum = sub(subgroupCharacteristicAverage, avgOfCurrentData);
		double squareNum = mul(subNum, subNum);
		double mulNum = mul(squareNum, sizeOfCurrentGroup);
		return mulNum;
	}

	public static void main(String[] args) {
		CopyPathToTextField();
		try {
			Thread.sleep(10000);
		} catch (InterruptedException e) {
			// System.out.println("sleep error");
			e.printStackTrace();
		}
		// System.out.println("main" + filepath);
		insert_in(filepath);

		/**
		 * ��ȡ�����е�ͳ�ƽ�ֹʱ�������ļ����������ļ�
		 */
		String endDate = null;
		for (Integer key : allPeople.keySet()) {
			List<Person> list = allPeople.get(key); // ����Ӧ��˾����Աȡ������list
			if (list.size() == 0) {
				continue;
			} else {
				endDate = new String(list.get(0).getEndDate());
				break;
			}
		}
		String createFilePath;
		createFilePath = filepath.substring(0, filepath.lastIndexOf("\\") + 1); // �õ������ļ���·������ͬ��Ŀ¼�´�������ļ�
//		System.out.println(createFilePath);
		File fileToCreate = new File(createFilePath + endDate + ".xlsx"); // ��ͳ�ƽ�ֹ����λ�ļ��������ļ�
		String[] titles = { "֤ȯ����"/*1*/, "��˾������"/*2*/, "�������"/*3*/,	 // ����
				""/*4*/, 
				"Total_Sum_Of_Squares_Sex12AndAge"/*5*/, 
				"Sex1"/*6*/, 
				"Sex2"/*7*/, 
				"Age"/*8*/, 
				"Average_Sex1"/*9*/, 
				"Average_Sex2"/*10*/,
				"Average_Age"/*11*/, 
				"Subgroup_Characteristic_Average_Sex1"/*12*/, 
				"Subgroup_Characteristic_Average_Sex2"/*13*/,	
				"Subgroup_Characteristic_Average_Age"/*14*/, 
				"Between_Group_Sum_Of_Squares_For_Characteristic_Sex1"/*15*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Sex2"/*16*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Age"/*17*/, 
				"Subgroup_Between_SS_Sex12AndAge"/*18*/,
				"Total_Between_SS_Sex12AndAge"/*19*/, 
				"Fau_g_Sex12AndAge"/*20*/, 
				""/*21*/, 
				"Total_Sum_Of_Squares_EducationAndWorkMonth"/*22*/,
				"Education1"/*23*/, 
				"Education2"/*24*/, 
				"Education3"/*25*/, 
				"Education4"/*26*/, 
				"Education5"/*27*/, 
				"Education6"/*28*/, 
				"Education7"/*29*/, 
				"WorkMonth"/*30*/, 
				"Average_Education1"/*31*/, 
				"Average_Education2"/*32*/, 
				"Average_Education3"/*33*/, 
				"Average_Education4"/*34*/, 
				"Average_Education5"/*35*/, 
				"Average_Education6"/*36*/, 
				"Average_Education7"/*37*/, 
				"Average_WorkMonth"/*38*/,
				"Subgroup_Characteristic_Average_Education1"/*39*/, 
				"Subgroup_Characteristic_Average_Education2"/*40*/, 
				"Subgroup_Characteristic_Average_Education3"/*41*/, 
				"Subgroup_Characteristic_Average_Education4"/*42*/, 
				"Subgroup_Characteristic_Average_Education5"/*43*/, 
				"Subgroup_Characteristic_Average_Education6"/*44*/, 
				"Subgroup_Characteristic_Average_Education7"/*45*/, 
				"Subgroup_Characteristic_Average_WorkMonth"/*46*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Education1"/*47*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Education2"/*48*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Education3"/*49*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Education4"/*50*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Education5"/*51*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Education6"/*52*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_Education7"/*53*/,
				"Between_Group_Sum_Of_Squares_For_Characteristic_WorkMonth"/*54*/,
				"Subgroup_Between_SS_Education1_7AndWorkMonth"/*55*/, 
				"Total_Between_SS_Education1_7AndWorkMonth"/*56*/,
				"Fau_g_Education1_7AndWorkMonth"/*57*/ };

		try {
			fileToCreate.createNewFile();
			WritableWorkbook workbook = Workbook.createWorkbook(fileToCreate); // ����������
			WritableSheet sheet = workbook.createSheet("sheet1", 0); // ����sheet

			/**
			 * ��һ����������
			 */
			for (int i = 0; i < titles.length; i++) {
				sheet.addCell(new Label(i, 0, titles[i]));
			}

			int line = 1; // ��¼���뵥Ԫ������������������λ��0��
			
//			XSSFWorkbook newWorkBook = new XSSFWorkbook();
//			XSSFSheet newSheet = newWorkBook.createSheet("sheet1");
//			int line = 0;
//			XSSFRow row = newSheet.createRow(line);
//			line++;
//			for (int i = 0; i < titles.length; i++) {
//				XSSFCell cell = row.createCell(i);
//				cell.setCellValue(titles[i]);
//			}
//			
//			FileOutputStream fileOutputStream;
			
			/**
			 * ��ÿ����˾���м���
			 */
			for (Integer key : allPeople.keySet()) { // ���ݹ�˾����ѡ�����˾���м���
				List<Person> list = allPeople.get(key); // ����Ӧ��˾����Աȡ������list
				if (list.size() == 0) {
					continue;
				}

				double[] sex1 = new double[list.size()], 
						sex2 = new double[list.size()], 
						age = new double[list.size()],
						education1 = new double[list.size()], 
						education2 = new double[list.size()], 
						education3 = new double[list.size()], 
						education4 = new double[list.size()], 
						education5 = new double[list.size()], 
						education6 = new double[list.size()], 
						education7 = new double[list.size()], 
						workMonth = new double[list.size()];
				double avg_Sex1, avg_Sex2, avg_Age, 
						avg_Education1, avg_Education2, avg_Education3, 
						avg_Education4, avg_Education5, avg_Education6, 
						avg_Education7, avg_WorkMonth;
				double[] avg_Sex12AndAge = new double[3]; // ������avg�������飬�������ļ���
				double[] avg_Education1_7AndWorkMonth = new double[8];
				double tSOS_Sex12AndAge = 0, tSOS_Education1_7AndWorkMonth = 0;

				/**
				 * ����ǰ��˾�ĸ������ݴ����Ӧ�ı�����
				 */
				for (int i = 0; i < list.size(); i++) {
					if (list.get(i).getSex() == div(1, Math.sqrt(2))) {
						sex1[i] = list.get(i).getSex();
						sex2[i] = 0;
					} else {
						sex1[i] = 0;
						sex2[i] = list.get(i).getSex();
					}
					
					age[i] = list.get(i).getAge();
					
					if (list.get(i).getEducation() == div(1, Math.sqrt(2))) {
						education1[i] = div(1, Math.sqrt(2));
						education2[i] = 0;
						education3[i] = 0;
						education4[i] = 0;
						education5[i] = 0;
						education6[i] = 0;
						education7[i] = 0;
					} else if (list.get(i).getEducation() == div(2, Math.sqrt(2))) {
						education1[i] = 0;
						education2[i] = div(1, Math.sqrt(2));
						education3[i] = 0;
						education4[i] = 0;
						education5[i] = 0;
						education6[i] = 0;
						education7[i] = 0;
					} else if (list.get(i).getEducation() == div(3, Math.sqrt(2))) {
						education1[i] = 0;
						education2[i] = 0;
						education3[i] = div(1, Math.sqrt(2));
						education4[i] = 0;
						education5[i] = 0;
						education6[i] = 0;
						education7[i] = 0;
					} else if (list.get(i).getEducation() == div(4, Math.sqrt(2))) {
						education1[i] = 0;
						education2[i] = 0;
						education3[i] = 0;
						education4[i] = div(1, Math.sqrt(2));
						education5[i] = 0;
						education6[i] = 0;
						education7[i] = 0;
					} else if (list.get(i).getEducation() == div(5, Math.sqrt(2))) {
						education1[i] = 0;
						education2[i] = 0;
						education3[i] = 0;
						education4[i] = 0;
						education5[i] = div(1, Math.sqrt(2));
						education6[i] = 0;
						education7[i] = 0;
					} else if (list.get(i).getEducation() == div(6, Math.sqrt(2))) {
						education1[i] = 0;
						education2[i] = 0;
						education3[i] = 0;
						education4[i] = 0;
						education5[i] = 0;
						education6[i] = div(1, Math.sqrt(2));
						education7[i] = 0;
					} else if (list.get(i).getEducation() == div(7, Math.sqrt(2))) {
						education1[i] = 0;
						education2[i] = 0;
						education3[i] = 0;
						education4[i] = 0;
						education5[i] = 0;
						education6[i] = 0;
						education7[i] = div(1, Math.sqrt(2));
					}
					
					workMonth[i] = list.get(i).getWorkMonth();
				}
				avg_Sex1 = avg_Sex12AndAge[0] = avg(sex1);
				avg_Sex2 = avg_Sex12AndAge[1] = avg(sex2);
				avg_Age = avg_Sex12AndAge[2] = avg(age);
				avg_Education1 = avg_Education1_7AndWorkMonth[0] = avg(education1);
				avg_Education2 = avg_Education1_7AndWorkMonth[1] = avg(education2);
				avg_Education3 = avg_Education1_7AndWorkMonth[2] = avg(education3);
				avg_Education4 = avg_Education1_7AndWorkMonth[3] = avg(education4);
				avg_Education5 = avg_Education1_7AndWorkMonth[4] = avg(education5);
				avg_Education6 = avg_Education1_7AndWorkMonth[5] = avg(education6);
				avg_Education7 = avg_Education1_7AndWorkMonth[6] = avg(education7);
				avg_WorkMonth = avg_Education1_7AndWorkMonth[7] = avg(workMonth);

				/**
				 * ����totalSumOfSquares
				 */
				for (int q = 0; q < list.size(); q++) {
					tSOS_Sex12AndAge = sum(tSOS_Sex12AndAge, mul(sub(sex1[q], avg_Sex1), sub(sex1[q], avg_Sex1)));
					tSOS_Sex12AndAge = sum(tSOS_Sex12AndAge, mul(sub(sex2[q], avg_Sex2), sub(sex2[q], avg_Sex2)));
					tSOS_Sex12AndAge = sum(tSOS_Sex12AndAge, mul(sub(age[q], avg_Age), sub(age[q], avg_Age)));

					tSOS_Education1_7AndWorkMonth = sum(tSOS_Education1_7AndWorkMonth,
							mul(sub(education1[q], avg_Education1), sub(education1[q], avg_Education1)));
					tSOS_Education1_7AndWorkMonth = sum(tSOS_Education1_7AndWorkMonth,
							mul(sub(education2[q], avg_Education2), sub(education2[q], avg_Education2)));
					tSOS_Education1_7AndWorkMonth = sum(tSOS_Education1_7AndWorkMonth,
							mul(sub(education3[q], avg_Education3), sub(education3[q], avg_Education3)));
					tSOS_Education1_7AndWorkMonth = sum(tSOS_Education1_7AndWorkMonth,
							mul(sub(education4[q], avg_Education4), sub(education4[q], avg_Education4)));
					tSOS_Education1_7AndWorkMonth = sum(tSOS_Education1_7AndWorkMonth,
							mul(sub(education5[q], avg_Education5), sub(education5[q], avg_Education5)));
					tSOS_Education1_7AndWorkMonth = sum(tSOS_Education1_7AndWorkMonth,
							mul(sub(education6[q], avg_Education6), sub(education6[q], avg_Education6)));
					tSOS_Education1_7AndWorkMonth = sum(tSOS_Education1_7AndWorkMonth,
							mul(sub(education7[q], avg_Education7), sub(education7[q], avg_Education7)));
					tSOS_Education1_7AndWorkMonth = sum(tSOS_Education1_7AndWorkMonth,
							mul(sub(workMonth[q], avg_WorkMonth), sub(workMonth[q], avg_WorkMonth)));
				}

				/**
				 * �Թ�˾��Ա�������м���
				 */
//				fileOutputStream = new FileOutputStream(fileToCreate);
				for (int i = 0; i < list.size() / 2; i++) {
					List<int[]> combinationResult = new ArrayList<>();
					IntStream listNum = IntStream.range(0, list.size()); // �����ù�˾�������ȵ��±����飬�����������
					int[] num = listNum.toArray();
					combinationSelect(num, 0, new int[i + 1], 0, combinationResult); // ����ǰ������ʽ����������������combinationResult��

					/**
					 * ������в�ͬ������������������ÿһ��������м���
					 */
					for (int n = 0; n < combinationResult.size(); n++) {
						double fau_g_Sex12AndAge, fau_g_EducationAndWorkMonth;
						double[] group1_Sex1 = new double[i + 1], 
								group1_Sex2 = new double[i + 1],
								group1_Age = new double[i + 1], 
								group1_Education1 = new double[i + 1],
								group1_Education2 = new double[i + 1],
								group1_Education3 = new double[i + 1],
								group1_Education4 = new double[i + 1],
								group1_Education5 = new double[i + 1],
								group1_Education6 = new double[i + 1],
								group1_Education7 = new double[i + 1],
								group1_WorkMonth = new double[i + 1];
						double[] group2_Sex1 = new double[list.size() - i - 1],
								group2_Sex2 = new double[list.size() - i - 1],
								group2_Age = new double[list.size() - i - 1],
								group2_Education1 = new double[list.size() - i - 1],
								group2_Education2 = new double[list.size() - i - 1],
								group2_Education3 = new double[list.size() - i - 1],
								group2_Education4 = new double[list.size() - i - 1],
								group2_Education5 = new double[list.size() - i - 1],
								group2_Education6 = new double[list.size() - i - 1],
								group2_Education7 = new double[list.size() - i - 1],
								group2_WorkMonth = new double[list.size() - i - 1];

						double group1_SubgroupCharacteristicAverage_Sex1, 
								group1_SubgroupCharacteristicAverage_Sex2,
								group1_SubgroupCharacteristicAverage_Age,
								group1_SubgroupCharacteristicAverage_Education1,
								group1_SubgroupCharacteristicAverage_Education2,
								group1_SubgroupCharacteristicAverage_Education3,
								group1_SubgroupCharacteristicAverage_Education4,
								group1_SubgroupCharacteristicAverage_Education5,
								group1_SubgroupCharacteristicAverage_Education6,
								group1_SubgroupCharacteristicAverage_Education7,
								group1_SubgroupCharacteristicAverage_WorkMonth;
						double group2_SubgroupCharacteristicAverage_Sex1, 
								group2_SubgroupCharacteristicAverage_Sex2,
								group2_SubgroupCharacteristicAverage_Age,
								group2_SubgroupCharacteristicAverage_Education1,
								group2_SubgroupCharacteristicAverage_Education2,
								group2_SubgroupCharacteristicAverage_Education3,
								group2_SubgroupCharacteristicAverage_Education4,
								group2_SubgroupCharacteristicAverage_Education5,
								group2_SubgroupCharacteristicAverage_Education6,
								group2_SubgroupCharacteristicAverage_Education7,
								group2_SubgroupCharacteristicAverage_WorkMonth;
						
						double[] group1_SCA_Sex12AndAge = new double[3]; // ��Ÿ���subgroupCharacteristicAverageֵ���������ļ���
						double[] group1_SCA_Education1_7AndWorkMonth = new double[8];
						double[] group2_SCA_Sex12AndAge = new double[3];
						double[] group2_SCA_Education1_7AndWorkMonth = new double[8];

						double group1_BetweenGroupSumOfSquaresForCharacteristic_Sex1,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Sex2,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Age,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Education1,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Education2,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Education3,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Education4,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Education5,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Education6,
								group1_BetweenGroupSumOfSquaresForCharacteristic_Education7,
								group1_BetweenGroupSumOfSquaresForCharacteristic_WorkMonth;
						double group2_BetweenGroupSumOfSquaresForCharacteristic_Sex1,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Sex2,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Age,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Education1,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Education2,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Education3,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Education4,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Education5,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Education6,
								group2_BetweenGroupSumOfSquaresForCharacteristic_Education7,
								group2_BetweenGroupSumOfSquaresForCharacteristic_WorkMonth;
						
						double[] group1_BGSOSFC_Sex12AndAge = new double[3]; // ��Ÿ���betweenGroupSumOfSquaresForCharacteristicֵ���������ļ���
						double[] group1_BGSOSFC_Education1_7AndWorkMonth = new double[8];
						double[] group2_BGSOSFC_Sex12AndAge = new double[3];
						double[] group2_BGSOSFC_Education1_7AndWorkMonth = new double[8];

						double group1_SBSS_Sex12AndAge = 0, group1_SBSS_Education1_7AndWorkMonth = 0;
						double group2_SBSS_Sex12AndAge = 0, group2_SBSS_Education1_7AndWorkMonth = 0;

						double tBSS_Sex12AndAge, tBSS_Education1_7AndWorkMonth;

						int[] currentCombinationGroup1 = combinationResult.get(n); // ��һ�������±����������
						for (int j = 0; j < currentCombinationGroup1.length; j++) { // ��һ�������
							group1_Sex1[j] = sex1[currentCombinationGroup1[j]];
							group1_Sex2[j] = sex2[currentCombinationGroup1[j]];
							group1_Age[j] = age[currentCombinationGroup1[j]];
							group1_Education1[j] = education1[currentCombinationGroup1[j]];
							group1_Education2[j] = education2[currentCombinationGroup1[j]];
							group1_Education3[j] = education3[currentCombinationGroup1[j]];
							group1_Education4[j] = education4[currentCombinationGroup1[j]];
							group1_Education5[j] = education5[currentCombinationGroup1[j]];
							group1_Education6[j] = education6[currentCombinationGroup1[j]];
							group1_Education7[j] = education7[currentCombinationGroup1[j]];
							group1_WorkMonth[j] = workMonth[currentCombinationGroup1[j]];
						}

						int[] currentCombinationGroup2 = new int[num.length - currentCombinationGroup1.length]; // �ڶ��������±�����
						for (int k = 0, p = 0; k < num.length && p < currentCombinationGroup2.length; k++) { // �ڶ��������±�ֵ�����±�����
							if (isInGroup1(num[k], currentCombinationGroup1)) {
								continue;
							} else {
								currentCombinationGroup2[p] = num[k];
								p++;
							}
						}
						for (int j = 0; j < currentCombinationGroup2.length; j++) { // �ڶ��������
							group2_Sex1[j] = sex1[currentCombinationGroup2[j]];
							group2_Sex2[j] = sex2[currentCombinationGroup2[j]];
							group2_Age[j] = age[currentCombinationGroup2[j]];
							group2_Education1[j] = education1[currentCombinationGroup2[j]];
							group2_Education2[j] = education2[currentCombinationGroup2[j]];
							group2_Education3[j] = education3[currentCombinationGroup2[j]];
							group2_Education4[j] = education4[currentCombinationGroup2[j]];
							group2_Education5[j] = education5[currentCombinationGroup2[j]];
							group2_Education6[j] = education6[currentCombinationGroup2[j]];
							group2_Education7[j] = education7[currentCombinationGroup2[j]];
							group2_WorkMonth[j] = workMonth[currentCombinationGroup2[j]];
						}

						/**
						 * ���� subgroupCharacteristicAverage
						 */
						group1_SubgroupCharacteristicAverage_Sex1 = group1_SCA_Sex12AndAge[0] = subgroupCharacteristicAverage(
								group1_Sex1);
						group1_SubgroupCharacteristicAverage_Sex2 = group1_SCA_Sex12AndAge[1] = subgroupCharacteristicAverage(
								group1_Sex2);
						group1_SubgroupCharacteristicAverage_Age = group1_SCA_Sex12AndAge[2] = subgroupCharacteristicAverage(
								group1_Age);
						group1_SubgroupCharacteristicAverage_Education1 = group1_SCA_Education1_7AndWorkMonth[0] = subgroupCharacteristicAverage(
								group1_Education1);
						group1_SubgroupCharacteristicAverage_Education2 = group1_SCA_Education1_7AndWorkMonth[1] = subgroupCharacteristicAverage(
								group1_Education2);
						group1_SubgroupCharacteristicAverage_Education3 = group1_SCA_Education1_7AndWorkMonth[2] = subgroupCharacteristicAverage(
								group1_Education3);
						group1_SubgroupCharacteristicAverage_Education4 = group1_SCA_Education1_7AndWorkMonth[3] = subgroupCharacteristicAverage(
								group1_Education4);
						group1_SubgroupCharacteristicAverage_Education5 = group1_SCA_Education1_7AndWorkMonth[4] = subgroupCharacteristicAverage(
								group1_Education5);
						group1_SubgroupCharacteristicAverage_Education6 = group1_SCA_Education1_7AndWorkMonth[5] = subgroupCharacteristicAverage(
								group1_Education6);
						group1_SubgroupCharacteristicAverage_Education7 = group1_SCA_Education1_7AndWorkMonth[6] = subgroupCharacteristicAverage(
								group1_Education7);
						group1_SubgroupCharacteristicAverage_WorkMonth = group1_SCA_Education1_7AndWorkMonth[7] = subgroupCharacteristicAverage(
								group1_WorkMonth);
						// System.out.println(group1_SubgroupCharacteristicAverage_Age);

						group2_SubgroupCharacteristicAverage_Sex1 = group2_SCA_Sex12AndAge[0] = subgroupCharacteristicAverage(
								group2_Sex1);
						group2_SubgroupCharacteristicAverage_Sex2 = group2_SCA_Sex12AndAge[1] = subgroupCharacteristicAverage(
								group2_Sex2);
						group2_SubgroupCharacteristicAverage_Age = group2_SCA_Sex12AndAge[2] = subgroupCharacteristicAverage(
								group2_Age);
						group2_SubgroupCharacteristicAverage_Education1 = group2_SCA_Education1_7AndWorkMonth[0] = subgroupCharacteristicAverage(
								group2_Education1);
						group2_SubgroupCharacteristicAverage_Education2 = group2_SCA_Education1_7AndWorkMonth[1] = subgroupCharacteristicAverage(
								group2_Education2);
						group2_SubgroupCharacteristicAverage_Education3 = group2_SCA_Education1_7AndWorkMonth[2] = subgroupCharacteristicAverage(
								group2_Education3);
						group2_SubgroupCharacteristicAverage_Education4 = group2_SCA_Education1_7AndWorkMonth[3] = subgroupCharacteristicAverage(
								group2_Education4);
						group2_SubgroupCharacteristicAverage_Education5 = group2_SCA_Education1_7AndWorkMonth[4] = subgroupCharacteristicAverage(
								group2_Education5);
						group2_SubgroupCharacteristicAverage_Education6 = group2_SCA_Education1_7AndWorkMonth[5] = subgroupCharacteristicAverage(
								group2_Education6);
						group2_SubgroupCharacteristicAverage_Education7 = group2_SCA_Education1_7AndWorkMonth[6] = subgroupCharacteristicAverage(
								group2_Education7);
						group2_SubgroupCharacteristicAverage_WorkMonth = group2_SCA_Education1_7AndWorkMonth[7] = subgroupCharacteristicAverage(
								group2_WorkMonth);

						/**
						 * ����betweenGroupSumOfSquaresForCharacteristic
						 */
						group1_BetweenGroupSumOfSquaresForCharacteristic_Sex1 = group1_BGSOSFC_Sex12AndAge[0] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Sex1, avg_Sex1, currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Sex2 = group1_BGSOSFC_Sex12AndAge[1] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Sex2, avg_Sex2, currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Age = group1_BGSOSFC_Sex12AndAge[2] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Age, avg_Age, currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Education1 = group1_BGSOSFC_Education1_7AndWorkMonth[0] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Education1, avg_Education1,
								currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Education2 = group1_BGSOSFC_Education1_7AndWorkMonth[1] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Education2, avg_Education2,
								currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Education3 = group1_BGSOSFC_Education1_7AndWorkMonth[2] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Education3, avg_Education3,
								currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Education4 = group1_BGSOSFC_Education1_7AndWorkMonth[3] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Education4, avg_Education4,
								currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Education5 = group1_BGSOSFC_Education1_7AndWorkMonth[4] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Education5, avg_Education5,
								currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Education6 = group1_BGSOSFC_Education1_7AndWorkMonth[5] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Education6, avg_Education6,
								currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_Education7 = group1_BGSOSFC_Education1_7AndWorkMonth[6] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_Education7, avg_Education7,
								currentCombinationGroup1.length);
						group1_BetweenGroupSumOfSquaresForCharacteristic_WorkMonth = group1_BGSOSFC_Education1_7AndWorkMonth[7] = betweenGroupSumOfSquaresForCharacteristic(
								group1_SubgroupCharacteristicAverage_WorkMonth, avg_WorkMonth,
								currentCombinationGroup1.length);

						group2_BetweenGroupSumOfSquaresForCharacteristic_Sex1 = group2_BGSOSFC_Sex12AndAge[0] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Sex1, avg_Sex1, currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Sex2 = group2_BGSOSFC_Sex12AndAge[1] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Sex2, avg_Sex2, currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Age = group2_BGSOSFC_Sex12AndAge[2] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Age, avg_Age, currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Education1 = group2_BGSOSFC_Education1_7AndWorkMonth[0] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Education1, avg_Education1,
								currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Education2 = group2_BGSOSFC_Education1_7AndWorkMonth[1] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Education2, avg_Education2,
								currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Education3 = group2_BGSOSFC_Education1_7AndWorkMonth[2] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Education3, avg_Education3,
								currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Education4 = group2_BGSOSFC_Education1_7AndWorkMonth[3] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Education4, avg_Education4,
								currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Education5 = group2_BGSOSFC_Education1_7AndWorkMonth[4] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Education5, avg_Education5,
								currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Education6 = group2_BGSOSFC_Education1_7AndWorkMonth[5] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Education6, avg_Education6,
								currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_Education7 = group2_BGSOSFC_Education1_7AndWorkMonth[6] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_Education7, avg_Education7,
								currentCombinationGroup2.length);
						group2_BetweenGroupSumOfSquaresForCharacteristic_WorkMonth = group2_BGSOSFC_Education1_7AndWorkMonth[7] = betweenGroupSumOfSquaresForCharacteristic(
								group2_SubgroupCharacteristicAverage_WorkMonth, avg_WorkMonth,
								currentCombinationGroup2.length);

						/**
						 * ����subgroupBetweenSS
						 */
						for (int q = 0; q < group1_SCA_Sex12AndAge.length; q++) {
							group1_SBSS_Sex12AndAge = sum(group1_SBSS_Sex12AndAge,
									mul(currentCombinationGroup1.length,
											mul(sub(group1_SCA_Sex12AndAge[q], avg_Sex12AndAge[q]),
													sub(group1_SCA_Sex12AndAge[q], avg_Sex12AndAge[q]))));
							group2_SBSS_Sex12AndAge = sum(group2_SBSS_Sex12AndAge,
									mul(currentCombinationGroup2.length,
											mul(sub(group2_SCA_Sex12AndAge[q], avg_Sex12AndAge[q]),
													sub(group2_SCA_Sex12AndAge[q], avg_Sex12AndAge[q]))));
						}
						for (int q = 0; q < group1_SCA_Education1_7AndWorkMonth.length; q++) {
							group1_SBSS_Education1_7AndWorkMonth = sum(group1_SBSS_Education1_7AndWorkMonth, mul(
									currentCombinationGroup1.length,
									mul(sub(group1_SCA_Education1_7AndWorkMonth[q], avg_Education1_7AndWorkMonth[q]), sub(
											group1_SCA_Education1_7AndWorkMonth[q], avg_Education1_7AndWorkMonth[q]))));
							group2_SBSS_Education1_7AndWorkMonth = sum(group2_SBSS_Education1_7AndWorkMonth, mul(
									currentCombinationGroup2.length,
									mul(sub(group2_SCA_Education1_7AndWorkMonth[q], avg_Education1_7AndWorkMonth[q]), sub(
											group2_SCA_Education1_7AndWorkMonth[q], avg_Education1_7AndWorkMonth[q]))));
						}

						/**
						 * ����totalBetweenSS
						 */
						tBSS_Sex12AndAge = group1_SBSS_Sex12AndAge + group2_SBSS_Sex12AndAge;
						tBSS_Education1_7AndWorkMonth = group1_SBSS_Education1_7AndWorkMonth
								+ group2_SBSS_Education1_7AndWorkMonth;

						/**
						 * ����Fau-g
						 */
						fau_g_Sex12AndAge = div(tBSS_Sex12AndAge, tSOS_Sex12AndAge);
						fau_g_EducationAndWorkMonth = div(tBSS_Education1_7AndWorkMonth, tSOS_Education1_7AndWorkMonth);

						/**
						 * ������д��excel����
						 */
//						// ��һ������
//						XSSFRow row1 = newSheet.createRow(line);
//						line++;
//						
//						row1.createCell(0).setCellValue("" + list.get(0).getCompanyCode());
//						row1.createCell(1).setCellValue("" + list.size());
//						row1.createCell(2).setCellValue(Arrays.toString(currentCombinationGroup1));
//						row1.createCell(3).setCellValue("");
//						row1.createCell(4).setCellValue("" + round(tSOS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(5).setCellValue(Arrays.toString(group1_Sex1));
//						row1.createCell(6).setCellValue(Arrays.toString(group1_Sex2));
//						row1.createCell(7).setCellValue(Arrays.toString(group1_Age));
//						row1.createCell(8).setCellValue("" + round(avg_Sex1, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(9).setCellValue("" + round(avg_Sex2, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(10).setCellValue("" + round(avg_Age, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(11).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Sex1, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(12).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Sex2, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(13).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Age, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(14).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Sex1, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(15).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Sex2, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(16).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Age, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(17).setCellValue("" + round(group1_SBSS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(18).setCellValue("" + round(tBSS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(19).setCellValue("" + round(fau_g_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(20).setCellValue("");
//						row1.createCell(21).setCellValue("" + round(tSOS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(22).setCellValue(Arrays.toString(group1_Education1));
//						row1.createCell(23).setCellValue(Arrays.toString(group1_Education2));
//						row1.createCell(24).setCellValue(Arrays.toString(group1_Education3));
//						row1.createCell(25).setCellValue(Arrays.toString(group1_Education4));
//						row1.createCell(26).setCellValue(Arrays.toString(group1_Education5));
//						row1.createCell(27).setCellValue(Arrays.toString(group1_Education6));
//						row1.createCell(28).setCellValue(Arrays.toString(group1_Education7));
//						row1.createCell(29).setCellValue(Arrays.toString(group1_WorkMonth));
//						row1.createCell(30).setCellValue("" + round(avg_Education1, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(31).setCellValue("" + round(avg_Education2, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(32).setCellValue("" + round(avg_Education3, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(33).setCellValue("" + round(avg_Education4, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(34).setCellValue("" + round(avg_Education5, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(35).setCellValue("" + round(avg_Education6, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(36).setCellValue("" + round(avg_Education7, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(37).setCellValue("" + round(avg_WorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(38).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Education1, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(39).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Education2, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(40).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Education3, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(41).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Education4, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(42).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Education5, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(43).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Education6, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(44).setCellValue("" + round(group1_SubgroupCharacteristicAverage_Education7, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(45).setCellValue("" + round(group1_SubgroupCharacteristicAverage_WorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(46).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education1, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(47).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education2, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(48).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education3, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(49).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education4, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(50).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education5, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(51).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education6, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(52).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education7, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(53).setCellValue("" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_WorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(54).setCellValue("" + round(group1_SBSS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(55).setCellValue("" + round(tBSS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row1.createCell(56).setCellValue("" + round(fau_g_EducationAndWorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						
//						// �ڶ�������
//						XSSFRow row2 = newSheet.createRow(line);
//						line++;
//						
//						row2.createCell(0).setCellValue("" + list.get(0).getCompanyCode());
//						row2.createCell(1).setCellValue("" + list.size());
//						row2.createCell(2).setCellValue(Arrays.toString(currentCombinationGroup2));
//						row2.createCell(3).setCellValue("");
//						row2.createCell(4).setCellValue("" + round(tSOS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(5).setCellValue(Arrays.toString(group2_Sex1));
//						row2.createCell(6).setCellValue(Arrays.toString(group2_Sex2));
//						row2.createCell(7).setCellValue(Arrays.toString(group2_Age));
//						row2.createCell(8).setCellValue("" + round(avg_Sex1, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(9).setCellValue("" + round(avg_Sex2, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(10).setCellValue("" + round(avg_Age, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(11).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Sex1, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(12).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Sex2, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(13).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Age, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(14).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Sex1, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(15).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Sex2, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(16).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Age, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(17).setCellValue("" + round(group2_SBSS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(18).setCellValue("");
//						row2.createCell(19).setCellValue("");
//						row2.createCell(20).setCellValue("");
//						row2.createCell(21).setCellValue("" + round(tSOS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(22).setCellValue(Arrays.toString(group2_Education1));
//						row2.createCell(23).setCellValue(Arrays.toString(group2_Education2));
//						row2.createCell(24).setCellValue(Arrays.toString(group2_Education3));
//						row2.createCell(25).setCellValue(Arrays.toString(group2_Education4));
//						row2.createCell(26).setCellValue(Arrays.toString(group2_Education5));
//						row2.createCell(27).setCellValue(Arrays.toString(group2_Education6));
//						row2.createCell(28).setCellValue(Arrays.toString(group2_Education7));
//						row2.createCell(29).setCellValue(Arrays.toString(group2_WorkMonth));
//						row2.createCell(30).setCellValue("" + round(avg_Education1, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(31).setCellValue("" + round(avg_Education2, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(32).setCellValue("" + round(avg_Education3, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(33).setCellValue("" + round(avg_Education4, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(34).setCellValue("" + round(avg_Education5, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(35).setCellValue("" + round(avg_Education6, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(36).setCellValue("" + round(avg_Education7, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(37).setCellValue("" + round(avg_WorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(38).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Education1, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(39).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Education2, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(40).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Education3, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(41).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Education4, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(42).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Education5, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(43).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Education6, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(44).setCellValue("" + round(group2_SubgroupCharacteristicAverage_Education7, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(45).setCellValue("" + round(group2_SubgroupCharacteristicAverage_WorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(46).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education1, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(47).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education2, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(48).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education3, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(49).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education4, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(50).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education5, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(51).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education6, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(52).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education7, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(53).setCellValue("" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_WorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(54).setCellValue("" + round(group2_SBSS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP));
//						row2.createCell(55).setCellValue("");
//						row2.createCell(56).setCellValue("");
						
						// ��һ������
						sheet.addCell(new Label(0, line, "" + list.get(0).getCompanyCode()));
						sheet.addCell(new Label(1, line, "" + list.size()));
						sheet.addCell(new Label(2, line, Arrays.toString(currentCombinationGroup1)));
						sheet.addCell(new Label(3, line, ""));
						sheet.addCell(new Label(4, line, "" + round(tSOS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(5, line, Arrays.toString(group1_Sex1)));
						sheet.addCell(new Label(6, line, Arrays.toString(group1_Sex2)));
						sheet.addCell(new Label(7, line, Arrays.toString(group1_Age)));
						sheet.addCell(new Label(8, line, "" + round(avg_Sex1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(9, line, "" + round(avg_Sex2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(10, line, "" + round(avg_Age, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(11, line, "" + round(group1_SubgroupCharacteristicAverage_Sex1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(12, line, "" + round(group1_SubgroupCharacteristicAverage_Sex2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(13, line, "" + round(group1_SubgroupCharacteristicAverage_Age, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(14, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Sex1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(15, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Sex2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(16, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Age, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(17, line, "" + round(group1_SBSS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(18, line, "" + round(tBSS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(19, line, "" + round(fau_g_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(20, line, ""));
						sheet.addCell(new Label(21, line, "" + round(tSOS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(22, line, Arrays.toString(group1_Education1)));
						sheet.addCell(new Label(23, line, Arrays.toString(group1_Education2)));
						sheet.addCell(new Label(24, line, Arrays.toString(group1_Education3)));
						sheet.addCell(new Label(25, line, Arrays.toString(group1_Education4)));
						sheet.addCell(new Label(26, line, Arrays.toString(group1_Education5)));
						sheet.addCell(new Label(27, line, Arrays.toString(group1_Education6)));
						sheet.addCell(new Label(28, line, Arrays.toString(group1_Education7)));
						sheet.addCell(new Label(29, line, Arrays.toString(group1_WorkMonth)));
						sheet.addCell(new Label(30, line, "" + round(avg_Education1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(31, line, "" + round(avg_Education2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(32, line, "" + round(avg_Education3, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(33, line, "" + round(avg_Education4, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(34, line, "" + round(avg_Education5, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(35, line, "" + round(avg_Education6, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(36, line, "" + round(avg_Education7, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(37, line, "" + round(avg_WorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(38, line, "" + round(group1_SubgroupCharacteristicAverage_Education1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(39, line, "" + round(group1_SubgroupCharacteristicAverage_Education2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(40, line, "" + round(group1_SubgroupCharacteristicAverage_Education3, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(41, line, "" + round(group1_SubgroupCharacteristicAverage_Education4, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(42, line, "" + round(group1_SubgroupCharacteristicAverage_Education5, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(43, line, "" + round(group1_SubgroupCharacteristicAverage_Education6, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(44, line, "" + round(group1_SubgroupCharacteristicAverage_Education7, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(45, line, "" + round(group1_SubgroupCharacteristicAverage_WorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(46, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(47, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(48, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education3, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(49, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education4, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(50, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education5, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(51, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education6, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(52, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_Education7, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(53, line, "" + round(group1_BetweenGroupSumOfSquaresForCharacteristic_WorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(54, line, "" + round(group1_SBSS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(55, line, "" + round(tBSS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(56, line, "" + round(fau_g_EducationAndWorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						
						line++;
						
						// �ڶ�������
						sheet.addCell(new Label(0, line, "" + list.get(0).getCompanyCode()));
						sheet.addCell(new Label(1, line, "" + list.size()));
						sheet.addCell(new Label(2, line, Arrays.toString(currentCombinationGroup2)));
						sheet.addCell(new Label(3, line, ""));
						sheet.addCell(new Label(4, line, "" + round(tSOS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(5, line, Arrays.toString(group2_Sex1)));
						sheet.addCell(new Label(6, line, Arrays.toString(group2_Sex2)));
						sheet.addCell(new Label(7, line, Arrays.toString(group2_Age)));
						sheet.addCell(new Label(8, line, "" + round(avg_Sex1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(9, line, "" + round(avg_Sex2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(10, line, "" + round(avg_Age, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(11, line, "" + round(group2_SubgroupCharacteristicAverage_Sex1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(12, line, "" + round(group2_SubgroupCharacteristicAverage_Sex2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(13, line, "" + round(group2_SubgroupCharacteristicAverage_Age, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(14, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Sex1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(15, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Sex2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(16, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Age, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(17, line, "" + round(group2_SBSS_Sex12AndAge, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(18, line, ""));
						sheet.addCell(new Label(19, line, ""));
						sheet.addCell(new Label(20, line, ""));
						sheet.addCell(new Label(21, line, "" + round(tSOS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(22, line, Arrays.toString(group2_Education1)));
						sheet.addCell(new Label(23, line, Arrays.toString(group2_Education2)));
						sheet.addCell(new Label(24, line, Arrays.toString(group2_Education3)));
						sheet.addCell(new Label(25, line, Arrays.toString(group2_Education4)));
						sheet.addCell(new Label(26, line, Arrays.toString(group2_Education5)));
						sheet.addCell(new Label(27, line, Arrays.toString(group2_Education6)));
						sheet.addCell(new Label(28, line, Arrays.toString(group2_Education7)));
						sheet.addCell(new Label(29, line, Arrays.toString(group2_WorkMonth)));
						sheet.addCell(new Label(30, line, "" + round(avg_Education1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(31, line, "" + round(avg_Education2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(32, line, "" + round(avg_Education3, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(33, line, "" + round(avg_Education4, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(34, line, "" + round(avg_Education5, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(35, line, "" + round(avg_Education6, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(36, line, "" + round(avg_Education7, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(37, line, "" + round(avg_WorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(38, line, "" + round(group2_SubgroupCharacteristicAverage_Education1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(39, line, "" + round(group2_SubgroupCharacteristicAverage_Education2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(40, line, "" + round(group2_SubgroupCharacteristicAverage_Education3, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(41, line, "" + round(group2_SubgroupCharacteristicAverage_Education4, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(42, line, "" + round(group2_SubgroupCharacteristicAverage_Education5, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(43, line, "" + round(group2_SubgroupCharacteristicAverage_Education6, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(44, line, "" + round(group2_SubgroupCharacteristicAverage_Education7, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(45, line, "" + round(group2_SubgroupCharacteristicAverage_WorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(46, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education1, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(47, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education2, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(48, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education3, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(49, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education4, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(50, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education5, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(51, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education6, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(52, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_Education7, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(53, line, "" + round(group2_BetweenGroupSumOfSquaresForCharacteristic_WorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(54, line, "" + round(group2_SBSS_Education1_7AndWorkMonth, 3, BigDecimal.ROUND_HALF_UP)));
						sheet.addCell(new Label(55, line, ""));
						sheet.addCell(new Label(56, line, ""));
						
						line++;
					}
				}
//				newWorkBook.write(fileOutputStream);
//				fileOutputStream.close();
			}
			workbook.write();
			workbook.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		frame.dispose();
	}

}
