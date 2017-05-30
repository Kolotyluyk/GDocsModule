package main;

import googleWorker.GDocsModule;

import java.util.Scanner;

public class Main {


    private static String[] formatingPeriod(String[] period){
        String[] inputToGDoc=new String[4];
        if(period.length==1) {
            inputToGDoc[2] = inputToGDoc[3] = period[0];
            inputToGDoc[0] = GDocsModule.mounth.get(0);
            inputToGDoc[1] = GDocsModule.mounth.get(11);
        } else
        if(period[1].contains("2")&&period[0]!=null) {
            inputToGDoc[3] = inputToGDoc[2] = period[1];
            inputToGDoc[1] = inputToGDoc[0] = period[0];
        }
        else
        if(period.length==3) {
            inputToGDoc[3] = inputToGDoc[2] = period[2];
            inputToGDoc[0] = period[0];
            inputToGDoc[1] = period[1];
        }
        else
        if(period.length==4)
            inputToGDoc=period;

        return inputToGDoc;
    }

	public static void main(String[] args) {
		GDocsModule gDocsModule = new GDocsModule();

		Scanner in = new Scanner(System.in);
      //  System.out.print("Enter Link Google Sheets: for example https://docs.google.com/spreadsheets/d/1YmXTemgS52vRo4f98-nGlnt9acP7bh3kzLfr9gpK9lA/edit#gid=806350235\n");
      //  String LINK = in.nextLine();
            String LINK = "https://docs.google.com/spreadsheets/d/1YmXTemgS52vRo4f98-nGlnt9acP7bh3kzLfr9gpK9lA/edit#gid=806350235";

        System.out.print("Enter period: ");
        String comand = in.nextLine();
        String[] period=null;
        //check input data
        switch (comand)
        {
            case "whole" :
                try {
                    period= GDocsModule.getWorkSheetPeriod(LINK);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;
            case "":
                try {
                    period= GDocsModule.getWorkSheetPeriod(LINK);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;
            default:
                period=formatingPeriod(period);

      break;
        }
            try {
                    GDocsModule.beatSheets(LINK,period[0],period[1],Integer.parseInt(period[2]),Integer.parseInt(period[3]));
                    System.out.println("Finish program");
                } catch (Exception e) {

                    e.printStackTrace();
                }




/*
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
	//		GDocsModule.beatSheets(LINK,period[0],finishPeriod,beginYear,finishYear);
			System.out.println("Finish program");
		} catch (Exception e) {

			e.printStackTrace();
		}

*/
	}

}
