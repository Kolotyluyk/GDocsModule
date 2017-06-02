package main;

import googleWorker.GDocsModule;

import java.time.LocalDate;
import java.time.YearMonth;
import java.util.Scanner;

public class Main {

    private static LocalDate getDatePeriod(String[] period, int defaultMonth){
        //check input  finish data
        int month =defaultMonth;
        int year = YearMonth.now().getYear();

        if(period.length==1 )
        {try {
            if(!period[0].isEmpty())
            month=Integer.parseInt(period[0]);
            }catch (Exception e){
                System.out.println("You write incorrect month finish period");
            }
           }
        else
        if (period.length==2){
            try {
                month=Integer.parseInt(period[0]);
                year=Integer.parseInt(period[1]);;
            }catch (Exception e){
                System.out.println("You write incorrect month start period");
            }
             }

        return LocalDate.of(year,month,1);
    }

	public static void main(String[] args) {
		GDocsModule gDocsModule = new GDocsModule();
		Scanner in = new Scanner(System.in);
        System.out.print("Enter Link Google Sheets: for example " +
                "https://docs.google.com/spreadsheets/d/1YmXTemgS52vRo4f98-nGlnt9acP7bh3kzLfr9gpK9lA/edit#gid=806350235\n");
       String link = in.nextLine();
        System.out.print("\npossible format:\n" +
                "StartMonth StartYear\n" +
                "StartMonth\n" +
                "nothing- in this way start date is first localMonth of current year \n");
        System.out.print("\nEnter start period:");
        String startPeriod = in.nextLine();
        System.out.print("\npossible format:\n" +
                "FinishMonth FinishYear\n" +
                "FinishMonth\n" +
                "nothing- in this way finish of period is first localMonth of current year \n");

        System.out.print("\nEnter finish period:");
        String finishPeriod = in.nextLine();


        LocalDate dateOfStartPeriod= getDatePeriod(finishPeriod.split(" "),1);
        LocalDate dateOfFinishPeriod= getDatePeriod(finishPeriod.split(" "),YearMonth.now().getMonthValue());



            try {
                 System.out.println("Start making report please wait\n");
                    gDocsModule.beatSheets(link,dateOfStartPeriod,dateOfFinishPeriod);
                    System.out.println("Finish program");
                } catch (Exception e) {
                    e.printStackTrace();
                }
	}

}
