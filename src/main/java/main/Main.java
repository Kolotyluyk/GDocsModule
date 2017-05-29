package main;

import googleWorker.GDocsModule;

import java.util.Scanner;

public class Main {



	public static void main(String[] args) {
		GDocsModule gDocsModule = new GDocsModule();

		Scanner in = new Scanner(System.in);
        System.out.print("Enter Link Google Sheets: for example https://docs.google.com/spreadsheets/d/1YmXTemgS52vRo4f98-nGlnt9acP7bh3kzLfr9gpK9lA/edit#gid=806350235\n");
        String LINK = in.nextLine();
        System.out.print("Enter mounth of beggin period:\n");
        String beginPeriod = in.nextLine();
        System.out.print("\nEnter mounth of finish period:\n");
        String finishPeriod = in.nextLine();
        System.out.print("\nEnter year of finish period:\n");
        int beginYear = in.nextInt();
        System.out.print("\nEnter year of finish period: \n");
        int finishYear = in.nextInt();
        //String LINK = "https://docs.google.com/spreadsheets/d/1YmXTemgS52vRo4f98-nGlnt9acP7bh3kzLfr9gpK9lA/edit#gid=806350235";

		try {
			GDocsModule.beatSheets(LINK,beginPeriod,finishPeriod,beginYear,finishYear);
			System.out.println("Finish program");
		} catch (Exception e) {

			e.printStackTrace();
		}

	}
}
