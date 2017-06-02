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
import java.time.LocalDate;
import java.util.*;
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

//value of range
	static String headerRange = "A3:T3";
	static String dataRange = "A4:Z1000";
	static String dateRange = "B1:B1";
	static String countOfDayRange = "T1:T1";
	static String exchangeRateRange = "T3:T3";


	public static List<String> localMonth = Arrays.asList("Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень");

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}
//connect to google doc using client_id and client_secret from client_secret.json
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

	//read from google sheet
	private static ValueRange readValue(Sheets serviceSheets, String spreadsheetId, String range) throws IOException {
		return serviceSheets.spreadsheets().values()
				.get(spreadsheetId, range)
				.setPrettyPrint(true)
				.execute();
		}

	private static String createSpreadSheet(List<ValueRange> childList, List<Object> row, List<List<Object>> headerValues,
											FileWriter WRITER, Sheets serviceSheets, String headerRange
											) throws IOException {
		headerValues.get(0).set(19,"Кількіть днів");
		ValueRange headValueRange = new ValueRange().setRange(headerRange).setValues(headerValues);
		childList.add(headValueRange);
		String nameSheet=row.get(1).toString();
		Spreadsheet spreadsheet = new Spreadsheet().setProperties(new SpreadsheetProperties()
				.setTitle(nameSheet).setAutoRecalc("ON_CHANGE"));
		String childSpreadSheetId = serviceSheets
				.spreadsheets()
				.create(spreadsheet)
				.execute()
				.getSpreadsheetId();
		WRITER.write(nameSheet + "|" + childSpreadSheetId + "\n");
		WRITER.flush();
		System.out.println("Document " + nameSheet + " successfully create");
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
						size=dataValues.indexOf(list);
						break;}
			}
			}
		//	if (!flag){
			ValueRange dataValueRange = new ValueRange().setRange("A" + String.valueOf(size + 4) + ":Z1000").setValues(pasteData);
			childList.add(dataValueRange);
			BatchUpdateValuesRequest oRequest = new BatchUpdateValuesRequest()
					.setValueInputOption("RAW")
					.setData(childList);
			serviceSheets.spreadsheets().values().batchUpdate(childSpreadSheetId, oRequest)
						.execute();
			System.out.println("" +
					"Write : "+row.get(1)+" " + dateValues.get(0).get(0));
	//	}
	//	else System.out.println( "Is present "+row.get(1)+" " + dateValues.get(0).get(0));
		}



	public static LocalDate getStartWorkSheetPeriod(String link) throws Exception {
		Sheets serviceSheets = getSheetsService();
		String spreadsheetId = getSheetId(link);
		Spreadsheet response1= serviceSheets.spreadsheets().get(spreadsheetId).setIncludeGridData (false).execute ();
		List<Sheet> workSheetList = response1.getSheets();
		LocalDate beginingDate =LocalDate.of(3000, 1,1);
		for (Sheet sheet : workSheetList) {
			String period=sheet.getProperties().getTitle();
			String[] splitedPeriod=period.split(" ");
			LocalDate date = LocalDate.of(Integer.parseInt(splitedPeriod[1]), Integer.parseInt(splitedPeriod[0])+1,1);
			if(date.isBefore(beginingDate)) {
				beginingDate=date;
			}
		}
		return beginingDate;
	}

	public static LocalDate getFinishWorkSheetPeriod(String link) throws Exception {
		Sheets serviceSheets = getSheetsService();
		String spreadsheetId = getSheetId(link);
		Spreadsheet response1= serviceSheets.spreadsheets().get(spreadsheetId).setIncludeGridData (false).execute ();
		List<Sheet> workSheetList = response1.getSheets();
		LocalDate finishingDate =LocalDate.of(1,1,1);
		for (Sheet sheet : workSheetList) {
			String period=sheet.getProperties().getTitle();
			String[] splitedPeriod=period.split(" ");
			LocalDate date = LocalDate.of(Integer.parseInt(splitedPeriod[1]), Integer.parseInt(splitedPeriod[0])+1,1);
			if(date.isAfter(finishingDate)) {
				finishingDate=date;
			}
		}
		return finishingDate;
	}


	private static boolean isPresentSheet(List<Sheet> workSheetList,String nameSheet ){
	for (Sheet sheet : workSheetList){
			if(sheet.getProperties().getTitle().equals(nameSheet))
			return true;
	}
	return false;
}

	public  void beatSheets(String link, LocalDate dateOfStartPeriod,LocalDate dateOfFinishPeriod) throws Exception {
		getDriveService();
		FileWriter writer = new FileWriter(file, true);
		Sheets serviceSheets = getSheetsService();
		String spreadsheetId = getSheetId(link);
			List<List<Object>> headerValues = readValue(serviceSheets, spreadsheetId, headerRange).getValues();

		Spreadsheet response1= serviceSheets.spreadsheets().get(spreadsheetId).setIncludeGridData(false).execute();
		List<Sheet> workSheetList = response1.getSheets();
		Locale loc = Locale.forLanguageTag("ua");
		for (LocalDate intermidiateDate=dateOfStartPeriod;intermidiateDate.isBefore(dateOfFinishPeriod);intermidiateDate=intermidiateDate.plusMonths(1))
					{
					//	String localMonth=intermidiateDate.getMonth().getDisplayName(TextStyle.FULL_STANDALONE, loc);
						String month=localMonth.get(intermidiateDate.getMonth().getValue()-1);
						String year=String.valueOf(intermidiateDate.getYear());
						String sheetName=month +" "+year;
					if(isPresentSheet(workSheetList, month +" "+year)){
						List<String>  oldSheets = Files.lines(Paths.get(String.valueOf(file)), StandardCharsets.UTF_8)
								.collect(Collectors.toList());
						Map<String, String> SpreedSheetNameId = oldSheets.stream().collect(Collectors.toMap
						        (s -> s.substring(0, s.indexOf('|')), s -> s.substring(s.indexOf('|') + 1, s.length())));
						//read values from sheet in preset range
						List<List<Object>> dataValues = readValue(serviceSheets,
							spreadsheetId, sheetName+"!" + dataRange).getValues();
						List<List<Object>> dateValues = readValue(serviceSheets,
							spreadsheetId,  sheetName+"!" + dateRange).getValues();
						List<List<Object>> countOfDayValues = readValue(serviceSheets,
							spreadsheetId,  sheetName+"!" + countOfDayRange).getValues();
						List<List<Object>> exchangeRateValues = readValue(serviceSheets,
							spreadsheetId,  sheetName+"!" + exchangeRateRange).getValues();
						//check read data not empty
						if (dataValues == null || dataValues.size() == 0) {
							System.out.println("No data found.");
						} else
							{
							dataValues.forEach(row ->
							{
							List<ValueRange> childList = new ArrayList<ValueRange>();
							String childId = null;
							String nameSheet=row.get(1).toString();
							if (SpreedSheetNameId.containsKey(nameSheet)) childId = SpreedSheetNameId.get(nameSheet);
							else try {
								childId = createSpreadSheet(childList, row, headerValues, writer, serviceSheets,
									headerRange);
								} catch (IOException e) {
								e.printStackTrace();
								}
							try {
							writeValueToSheet(row,dateValues,countOfDayValues,exchangeRateValues, childList, dataRange,
									          serviceSheets, childId);
							} catch (IOException e) {
								e.printStackTrace();
							}
							});
						}
					}
				}
	}
}