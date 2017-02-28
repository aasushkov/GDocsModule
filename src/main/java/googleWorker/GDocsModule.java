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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class GDocsModule {

	private static final String APPLICATION_NAME = "Google Sheets API";
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

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}
	//авторизация в сервисах гугл и сохранение ключа сессии
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
		System.out.println(
				"Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}
	//создание сервиса для работы с гугл-таблицами
	private static Sheets getSheetsService() throws Exception {
		Credential credential = authorize();
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME)
				.build();
	}
	//создание сервиса для работы с гугл-Диском
	private static Drive getDriveService() throws Exception {
		Credential credential = authorize();
		return new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME)
				.build();
	}
	//метод для удаления старых таблиц
	private static void deleteFile(Drive service, String fileId) {
		try {
			service.files().delete(fileId).execute();
		} catch (IOException e) {
			System.out.println("An error occurred: " + e);
		}
	}
	// извлечение ID документа из ссылки 
	private static String getSheetId(String link){
		Matcher matcher = sheetId.matcher(link);
		if(matcher.find()){
			SHEET_ID = matcher.group(1);
		}
		return SHEET_ID;
	}
	//разбиение главной таблицы на подтаблицы
	public static void beatSheets(String link) throws Exception {
		//удаление таблиц созданых в процессе прошлого запуска
		List<String> oldSheets = Files.lines(Paths.get(String.valueOf(file)), StandardCharsets.UTF_8)
				.collect(Collectors.toList());
		Drive serviceDrive = getDriveService();
		if (oldSheets != null) {
			for(String sheet : oldSheets) {
				deleteFile(serviceDrive, sheet);
			}
		}
		FileWriter WRITER = new FileWriter(file);
		Sheets serviceSheets = getSheetsService();
		//получение ID таблицы
		String spreadsheetId = getSheetId(link);
		System.out.println(spreadsheetId);
		//диапазон заголовка таблицы
		String headerRange = "A1:T3";
		//Диапазон данных
		String dataRange = "A4:N1000";
		//Диапазон формул
		String formulRange = "O4:R4";
		//извлечение массива заголовка
		ValueRange headerRespone = serviceSheets.spreadsheets().values()
				.get(spreadsheetId, headerRange)
				.setPrettyPrint(true)
				.setValueRenderOption("FORMULA")
				.execute();
		//извлечение масива данных
		ValueRange dataResponse = serviceSheets.spreadsheets().values()
				.get(spreadsheetId, dataRange)
				.setValueRenderOption("FORMULA")
				.execute();
		//извлечение массива формул

		ValueRange formulRespone = serviceSheets.spreadsheets().values()
				.get(spreadsheetId, formulRange)
				.setPrettyPrint(true)
				.setValueRenderOption("FORMULA")
				.execute();

		
		List<List<Object>> headerValues = headerRespone.getValues();
		List<List<Object>> dataValues = dataResponse.getValues();
		List<List<Object>> formulValues = formulRespone.getValues();
		if (dataValues == null || dataValues.size() == 0) {
			System.out.println("No data found.");
		} else {
			//создание таблиц
			for (List<Object> row : dataValues) {
				//создание заголовка + данных для дочерней таблицы
				List<List<Object>> pasteData = new ArrayList<>();
				pasteData.add(row);
				ValueRange headValueRange = new ValueRange().setRange(headerRange).setValues(headerValues);
				ValueRange dataValueRange = new ValueRange().setRange(dataRange).setValues(pasteData);
				ValueRange formulValueRange = new ValueRange().setRange(formulRange).setValues(formulValues);
				List<ValueRange> childList = new ArrayList<>();
				childList.add(headValueRange);
				childList.add(dataValueRange);
				childList.add(formulValueRange);
				Spreadsheet spreadsheet = new Spreadsheet().setProperties(new SpreadsheetProperties()
						.setTitle(row.get(1).toString()).setAutoRecalc("ON_CHANGE"));
				//создание дочерней таблицы, получение ее ID
				String childSpreadSheetId = serviceSheets
						.spreadsheets()
						.create(spreadsheet)
						.execute()
						.getSpreadsheetId();
				WRITER.write(childSpreadSheetId + "\n");
				WRITER.flush();
				//запись данных в созданную таблицу
				BatchUpdateValuesRequest oRequest = new BatchUpdateValuesRequest()
						.setValueInputOption("USER_ENTERED")
						.setData(childList);
				serviceSheets.spreadsheets().values().batchUpdate(childSpreadSheetId, oRequest).setPrettyPrint(true).execute();
			}
		}
	}
}
