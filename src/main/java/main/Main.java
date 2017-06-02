package main;

import googleWorker.GDocsModule;

import java.time.LocalDate;
import java.time.YearMonth;
import java.util.Scanner;

public class Main {


    private static LocalDate getStartDatePeriod(String[] period){
        //check input  start data
        int year = YearMonth.now().getYear();
        int month = 1;
        if(period.length==1 )
        {try {
            if(!period[0].isEmpty())
            month=Integer.parseInt(period[0]);
            }catch (Exception e){
                System.out.println("You write incorrect month start period");
            }
            return LocalDate.of(year, month,1);
        }
        else
        if (period.length==2){
            try {
                month=Integer.parseInt(period[0]);
            }catch (Exception e){
                System.out.println("You write incorrect month start period");
            }
            return LocalDate.of(Integer.parseInt(period[0]),month,1);
        }
        return LocalDate.of(year,1,1);
    }

    private static LocalDate getFinishDatePeriod(String[] period){
        //check input  finish data
        int month = YearMonth.now().getMonthValue();
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


        LocalDate dateOfStartPeriod= getStartDatePeriod(startPeriod.split(" "));
        LocalDate dateOfFinishPeriod= getFinishDatePeriod(finishPeriod.split(" "));


            try {
                 System.out.println("Start making report please wait\n");
                    gDocsModule.beatSheets(link,dateOfStartPeriod,dateOfFinishPeriod);
                    System.out.println("Finish program");
                } catch (Exception e) {
                    e.printStackTrace();
                }
	}

}
