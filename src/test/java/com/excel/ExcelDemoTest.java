package com.excel;

import java.io.IOException;
import java.util.ArrayList;

public class ExcelDemoTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		ExcelDemo exel = new ExcelDemo();
		ArrayList<String> data = exel.getData("Delivery Truck");

		for (int i = 0; i < data.size(); i++) {
			System.out.println(data.get(i));
		}

	}

}
