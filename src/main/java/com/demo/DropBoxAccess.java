/**
 * 
 */
package com.demo;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import com.dropbox.core.DbxApiException;
import com.dropbox.core.DbxException;
import com.dropbox.core.DbxRequestConfig;
import com.dropbox.core.v2.DbxClientV2;
import com.dropbox.core.v2.files.FileMetadata;
import com.dropbox.core.v2.files.FolderMetadata;
import com.dropbox.core.v2.files.ListFolderResult;
import com.dropbox.core.v2.files.Metadata;
import com.dropbox.core.v2.sharing.CreateSharedLinkWithSettingsErrorException;
import com.dropbox.core.v2.sharing.ListSharedLinksErrorException;
import com.dropbox.core.v2.sharing.ListSharedLinksResult;
import com.dropbox.core.v2.sharing.SharedLinkMetadata;
import com.dropbox.core.v2.users.FullAccount;

/**
 * @author Mehul
**/

public class DropBoxAccess {
	
	private static final String ACCESS_TOKEN = "h0ecIM0XbtAAAAAAAAAAD0gxk7UbSt9-dgRNw6bOs2jRH_UVpKzNyDDGUj6uBmRW";
	
	private static final String DESTINATION_PATH = "/home/neosoft/Generated_Files";
	
    //private static final String ACCESS_TOKEN = "20lXSdp3OgAAAAAAAAAFZQhH8xuyPBZvejQrrjkR-VVZeZpP2IV30gumDmFmtC9q";
    
	// Get extensions for all the file names

	private static String getExtension(Metadata metadata) {
		String ext = null;
		if (metadata instanceof FileMetadata) {
			ext = metadata.getName().substring(metadata.getName().lastIndexOf(".") + 1);
		} else if (metadata instanceof FolderMetadata) {
			ext = null;
		}

		return ext;
	}
	 		
	 	// Get share link of all files in the dropbox (if not avail, then create it and get it)
	 		
	private static String getShareLink(DbxClientV2 client, Metadata metadata)
			throws ListSharedLinksErrorException, DbxException {
		String shareLink = null;
		try {

			// following line is to creat the share link when it does not exist!

			shareLink = client.sharing().createSharedLinkWithSettings(metadata.getPathDisplay()).getUrl();

		} catch (CreateSharedLinkWithSettingsErrorException ex) {

			// to get the share link if it already exists!

			ListSharedLinksResult result = client.sharing().listSharedLinksBuilder()
					.withPath(metadata.getPathDisplay())
					.withDirectOnly(true).start();
			List<SharedLinkMetadata> shareLinkList = result.getLinks();
			for (SharedLinkMetadata sharedLinkMetadata : shareLinkList) {
				shareLink = sharedLinkMetadata.getUrl();
			}

		} catch (DbxException ex) {
			System.out.println(ex);
		}
		return shareLink;
	}
	 		
	/**
	 * @param metadata
	 * @return
	 */
	private static String getFileName(Metadata metadata) {
		String fileName = null;
		if (metadata instanceof FileMetadata) {
			String fullFileName = metadata.getName().substring(metadata.getName().lastIndexOf("/") + 1);
			int endIndex = fullFileName.lastIndexOf(".");
			fileName = fullFileName.substring(0, endIndex);
		}
		return fileName;
	}
	
	private static BufferedReader getConsoleReader() throws IOException {
		BufferedReader reader =
                new BufferedReader(new InputStreamReader(System.in));
		return reader;
	}
    
	public static void main(String[] args) throws DbxApiException, DbxException, IOException {

		// Create Dropbox client
		DbxRequestConfig config = DbxRequestConfig.newBuilder("dropbox/java-tutorial").build();
		DbxClientV2 client = new DbxClientV2(config, ACCESS_TOKEN);

		FullAccount account = client.users().getCurrentAccount();
		System.out.println(account.getName().getDisplayName());

		String source = null;
		String fileName = "generate";
		String destination = null;
		
		System.out.println("Source path instructions:");
		System.out.println("Example: if your source folder is a folder 'aa' inside Dropbox_home/Test Folder 1/Files and Folders inside Folder/");
		System.out.println("Then i/p sourse Path would be like: /Test Folder 1/Files and Folders inside Folder/a2");
		System.out.println("NOTE: please insert the valid path of your Dropbox directory, Any other input will be set to failure!");
		System.out.println("Now enter the Source Folder Path carefully: ");
		source = getConsoleReader().readLine();
		
		System.out.println("Default destination path is /home/neosoft/Generated_Files");
		System.out.println("If you want to go with default path, Press 'Y' or 'y' otherwise Press any key to give destination path: ");
		String isDefault = getConsoleReader().readLine();
		
		if(isDefault.equals("Y")||isDefault.equals("y")) {
			destination = DESTINATION_PATH;
		} else {
			System.out.println("Destination path instructions: ");
			System.out.println("Example:");
			System.out.println("In Linux system:  if you want to give destination path to 'Generated_Files' folder that is into your 'user' folder, "
					+ "\nthen your destination path will be like '/home/neosoft/Generated_Files'");
			System.out.println("NOTE: please insert the valid path of your local directory"
					+ "\n Any other input will be set to failure!");
			System.out.println("Now Enter the Destination Folder Path: ");
			destination = getConsoleReader().readLine();
		}
		
		System.out.println("File name instructions: ");
		System.out.println("Example: if you want to generate the file 'generated data.xlsx', then just give the 'generated data' as a file name");
		System.out.println("Enter the File Name: ");
		fileName = getConsoleReader().readLine();
		
		// Get files and folder metadata from Dropbox root directory
		
		ListFolderResult result = client.files().listFolderBuilder(source).withRecursive(true).start();
		List<JSONObject> listOfJsonObject = new ArrayList<>();
		int i = 1, j = 1;
		while (true) {
			for (Metadata metadata : result.getEntries()) {
				
				System.out.println("----------------------- total count ----------------------- " + i++);
				
				System.out.println("path ----- "+ metadata.getPathDisplay());
				System.out.println("file_name ----- "+ getFileName(metadata));
				System.out.println("file_ext ----- "+ getExtension(metadata));
				System.out.println("share_link ----- "+ getShareLink(client, metadata));

				if (getFileName(metadata) != null 
						&& getExtension(metadata) != null
						&& getShareLink(client, metadata) != null) {
					JSONObject jsonObject = new JSONObject();
					jsonObject.put("path", metadata.getPathDisplay());
					jsonObject.put("file_name", getFileName(metadata));
					jsonObject.put("file_ext", getExtension(metadata));
					jsonObject.put("share_link", getShareLink(client, metadata));

					listOfJsonObject.add(jsonObject);
					System.out.println("----------------------- file count ----------------------- " + j++);
				}

			}

			if (!result.getHasMore()) {
				break;
			}
			result = client.files().listFolderContinue(result.getCursor());
		}

		// Store the files into Excel file

		String excelFile = destination+"/"+fileName+".xlsx";

		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("Sheet1");

		int rowNum = 0;

		Row row0 = sheet.createRow(0); // add 0-th row manually!

		row0.createCell(0).setCellValue("path");
		row0.createCell(1).setCellValue("file_name");
		row0.createCell(2).setCellValue("file_ext");
		row0.createCell(3).setCellValue("share_link");

		for (JSONObject jsonObject : listOfJsonObject) {
			Row row = sheet.createRow(++rowNum);
			int colNum = 0;

			row.createCell(colNum++).setCellValue(jsonObject.getString("path"));
			row.createCell(colNum++).setCellValue(jsonObject.getString("file_name"));
			row.createCell(colNum++).setCellValue(jsonObject.getString("file_ext"));
			row.createCell(colNum++).setCellValue(jsonObject.getString("share_link"));
		}

		FileOutputStream outputStream = new FileOutputStream(excelFile);
		book.write(outputStream);

		book.close();

	}
	
}
