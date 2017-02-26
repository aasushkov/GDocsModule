package main;

import googleWorker.GDocsModule;

import java.util.Scanner;

public class Main {



	public static void main(String[] args) {
		GDocsModule gDocsModule = new GDocsModule();

		Scanner in = new Scanner(System.in);
		System.out.print("Enter Link Google Sheets: ");
		String LINK = in.nextLine();

		try {
			GDocsModule.beatSheets(LINK);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
