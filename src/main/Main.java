package main;

import random.ExcelCellFinder;

public class Main {

	public static void main(String[] args) {
		ExcelCellFinder excelCellFinder = new ExcelCellFinder(4,3);
		for (int i = 1; i <=750 ; i++) {
//			String xCell = excelCellFinder.getCellIndexCountHorizontal(i);
//			String yCell = excelCellFinder.getCellIndexCountVertical(i);
//			System.out.println(xCell);
//			System.out.println(yCell);
			System.out.println(excelCellFinder.numberToLetter(i));
		}
	}

}
