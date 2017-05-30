package googleWorker;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class GDocsModule {

	private static final String APPLICATION_NAME = "GDocsModule";
	private static final File DATA_STORE_DIR = new File(
			System.getProperty("user.home"), ".credentials/sheets.googleapis");
	private static FileDataStoreFactory DATA_STORE_FACTORY;
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static HttpTransport HTTP_TRANSPORT;
	private static final List<String> SCOPES_SHEETS = Arrays.asList(SheetsScopes.SPREADSHEETS, DriveScopes.DRIVE);
	private static File file = new File("historysheets.txt");
	private static final String patternId = "d/(.*)/edit";
	private static final Pattern sheetId = Pattern.compile(patternId);
	private static String SHEET_ID = "";


	static String headerRange = "A3:T3";
	static String dataRange = "A4:Z1000";
	static String dateRange = "B1:B1";
	static String countOfDayRange = "T1:T1";
	static String exchangeRateRange = "T3:T3";

	//static String formulRange = "O4:Z4";
	public static List<String> mounth = Arrays.asList("Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень");


	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	private static Credential authorize() throws Exception {
		InputStream in = GDocsModule.class.getResourceAsStream("/client_secret.json");
		GoogleClientSecrets clientSecrets =
				GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
		GoogleAuthorizationCodeFlow flow =
				new GoogleAuthorizationCodeFlow.Builder(
						HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES_SHEETS)
						.setDataStoreFactory(DATA_STORE_FACTORY)
						.setAccessType("offline")
						.build();
		Credential credential = new AuthorizationCodeInstalledApp(
				flow, new LocalServerReceiver()).authorize("user");
		//System.out.println(
		//		"Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	private static Sheets getSheetsService() throws Exception {
		Credential credential = authorize();
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME)
				.build();
	}

	private static Drive getDriveService() throws Exception {
		Credential credential = authorize();
		return new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME)
				.build();
	}

	private static void deleteFile(Drive service, String fileId) {
		try {
			service.files().delete(fileId).execute();
		} catch (IOException e) {
			System.out.println("An error occurred: " + e);
		}
	}

	private static String getSheetId(String link) {
		Matcher matcher = sheetId.matcher(link);
		if (matcher.find()) {
			SHEET_ID = matcher.group(1);
		}
		return SHEET_ID;
	}

	private static ValueRange readValue(Sheets serviceSheets, String spreadsheetId, String range) throws IOException {
		return serviceSheets.spreadsheets().values()
				.get(spreadsheetId, range)
				.setPrettyPrint(true)
				.execute();
		}

	private static String createSpreadSheet(List<ValueRange> childList, List<Object> row, List<List<Object>> headerValues,
											FileWriter WRITER, Sheets serviceSheets, String headerRange) throws IOException {
		headerValues.get(0).set(19,"Кількіть днів");
		ValueRange headValueRange = new ValueRange().setRange(headerRange).setValues(headerValues);
		childList.add(headValueRange);

		Spreadsheet spreadsheet = new Spreadsheet().setProperties(new SpreadsheetProperties()
				.setTitle(row.get(1).toString()).setAutoRecalc("ON_CHANGE"));
		String childSpreadSheetId = serviceSheets
				.spreadsheets()
				.create(spreadsheet)
				.execute()
				.getSpreadsheetId();
		WRITER.write(row.get(1).toString() + "|" + childSpreadSheetId + "\n");
		WRITER.flush();
		System.out.println("Document " + row.get(1).toString() + " successfully create");
		return childSpreadSheetId;
	}


	private static void writeValueToSheet(List<Object> row,List<List<Object>> dateValues,List<List<Object>> countOfDayValues,
										  List<List<Object>> exchangeRateValues, List<ValueRange> childList, String dataRange,
										  Sheets serviceSheets, String childSpreadSheetId) throws IOException
		{
			List<List<Object>> pasteData = new ArrayList<>();
			pasteData.add(row);
			pasteData.get(0).set(0,dateValues.get(0).get(0));
			pasteData.get(0).add(exchangeRateValues.get(0).get(0));
			pasteData.get(0).add(countOfDayValues.get(0).get(0));

			List<List<Object>> dataValues = readValue(serviceSheets, childSpreadSheetId, dataRange).getValues();
			int size;
			boolean flag =false;
			if (dataValues==null) size=0;
			else
				{
					size = dataValues.size();
				    for (List<Object> list: dataValues) {
				    if(list.get(0).equals(dateValues.get(0).get(0)))
				    {flag=true;
						break;}
			}
			}

			if (!flag){
			ValueRange dataValueRange = new ValueRange().setRange("A" + String.valueOf(size + 4) + ":Z1000").setValues(pasteData);
			//header id sheets not exist

			childList.add(dataValueRange);
			BatchUpdateValuesRequest oRequest = new BatchUpdateValuesRequest()
					.setValueInputOption("RAW")
					.setData(childList);
			serviceSheets.spreadsheets().values().batchUpdate(childSpreadSheetId, oRequest)
					//.setPrettyPrint(true)
					.execute();
			System.out.println("" +
					"Write : "+row.get(1)+" " + dateValues.get(0).get(0));

		}
		else System.out.println( "Is present "+row.get(1)+" " + dateValues.get(0).get(0));
		}



	public static String[] getWorkSheetPeriod(String link) throws Exception {
		Sheets serviceSheets = getSheetsService();
		String spreadsheetId = getSheetId(link);
		Spreadsheet response1= serviceSheets.spreadsheets().get(spreadsheetId).setIncludeGridData (false).execute ();
		List<Sheet> workSheetList = response1.getSheets();
		String begginingPeriodMounth="Січень";
		String finishingPeriodMounth="Січень";
		String begginingPeriodYear="3000";
		String finishingPeriodYear=String.valueOf(Integer.SIZE);
		LocalDateTime beginingDate = LocalDateTime.of(Integer.parseInt(begginingPeriodYear),mounth.indexOf(begginingPeriodMounth)+1,1,1,1);
		LocalDateTime finishingDate = LocalDateTime.of(Integer.parseInt(finishingPeriodYear),mounth.indexOf(finishingPeriodMounth)+1,1,1,1);
		for (Sheet sheet : workSheetList) {
			String period=sheet.getProperties().getTitle();
			String[] splitedPeriod=period.split(" ");
			LocalDateTime date = LocalDateTime.of(Integer.parseInt(splitedPeriod[1]),mounth.indexOf(splitedPeriod[0])+1,1,1,1);
			if(date.isAfter(finishingDate)) {
				finishingPeriodMounth=splitedPeriod[0];
				finishingPeriodYear=splitedPeriod[1];
				finishingDate=date;
			}
			if(date.isBefore(beginingDate)) {
				begginingPeriodMounth=splitedPeriod[0];
				begginingPeriodYear=splitedPeriod[1];
				beginingDate=date;
			}
		}
		String[] result=new String[4];
		result[0]=begginingPeriodMounth;
		result[1]=finishingPeriodMounth;
		result[2]=begginingPeriodYear;
		result[3]=finishingPeriodYear;
			return result;
	}


	private static boolean isPresentSheet(List<Sheet> workSheetList,String nameSheet ){
	for (Sheet sheet : workSheetList){
		System.out.print(sheet.getProperties().getTitle());
		if(sheet.getProperties().getTitle().equals(nameSheet))
			return true;
	}
	return false;
}

	public static void beatSheets(String link, String startMounth, String lastMounth, int startYear, int lastYear) throws Exception {
		Drive serviceDrive = getDriveService();
		FileWriter WRITER = new FileWriter(file, true);
		Sheets serviceSheets = getSheetsService();
		String spreadsheetId = getSheetId(link);
			List<List<Object>> headerValues = readValue(serviceSheets, spreadsheetId, headerRange).getValues();


						for (int i = startYear; i <= lastYear; i++)
				for (int j = mounth.indexOf(startMounth); j <= mounth.indexOf(lastMounth); j++) {
					Spreadsheet response1= serviceSheets.spreadsheets().get(spreadsheetId).setIncludeGridData (false).execute ();
					List<Sheet> workSheetList = response1.getSheets();
					if(isPresentSheet(workSheetList,mounth.get(j)+" "+String.valueOf(i))){
				List<String>  oldSheets = Files.lines(Paths.get(String.valueOf(file)), StandardCharsets.UTF_8)
						.collect(Collectors.toList());
				Map<String, String> SpreedSheetNameId = oldSheets.stream().collect(Collectors.toMap
						(s -> s.substring(0, s.indexOf('|')), s -> s.substring(s.indexOf('|') + 1, s.length())));

				List<List<Object>> dataValues = readValue(serviceSheets,
						spreadsheetId, mounth.get(j)+" "+String.valueOf(i)+"!" + dataRange).getValues();
				List<List<Object>> dateValues = readValue(serviceSheets,
						spreadsheetId, mounth.get(j)+" "+String.valueOf(i)+"!" + dateRange).getValues();
					List<List<Object>> countOfDayValues = readValue(serviceSheets,
						spreadsheetId, mounth.get(j)+" "+String.valueOf(i)+"!" + countOfDayRange).getValues();
				List<List<Object>> exchangeRateValues = readValue(serviceSheets,
						spreadsheetId, mounth.get(j)+" "+String.valueOf(i)+"!" + exchangeRateRange).getValues();
				if (dataValues == null || dataValues.size() == 0) {
					System.out.println("No data found.");
				} else {
					dataValues.forEach(row -> {
						List<ValueRange> childList = new ArrayList<ValueRange>();
						String childId = null;
						if (SpreedSheetNameId.containsKey(row.get(1).toString()))
							childId = SpreedSheetNameId.get(row.get(1).toString());
						else try {
							childId = createSpreadSheet(childList, row, headerValues, WRITER, serviceSheets, headerRange);
						} catch (IOException e) {
							e.printStackTrace();
						}
						try {
							writeValueToSheet(row,dateValues,countOfDayValues,exchangeRateValues, childList, dataRange, serviceSheets, childId);
						} catch (IOException e) {
							e.printStackTrace();
						}
					});
				}
					}
						}
	}

}