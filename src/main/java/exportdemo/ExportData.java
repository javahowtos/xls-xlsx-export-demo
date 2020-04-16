package exportdemo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportData {
	// CONSTANTS
	/*----------------------------------------------------------*/
	static final String FILE_SAVE_LOCATION = "C:\\reports\\";
	static final String FILE_NAME = "UserReport.xlsx";
	/*----------------------------------------------------------*/

	public static void main(String[] args) throws IOException {
		// creating data to be exported
		List<User> userList = new ArrayList<User>();
		userList.add(new User(1, "John", "Doe"));
		userList.add(new User(2, "Peter", "Peterson"));

		// creating workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		// creating sheet with name "Report" in workbook
		XSSFSheet sheet = workbook.createSheet("Report");
		// this method creates header for our table
		createHeader(sheet, workbook);

		int rowCount = 0;
		for (User user : userList) {
			// creating row
			Row row = sheet.createRow(++rowCount);

			// adding first cell to the row
			Cell idCell = row.createCell(0);
			idCell.setCellValue(user.id);

			// adding second cell to the row
			Cell nameCell = row.createCell(1);
			nameCell.setCellValue(user.firstName);

			// adding third cell to the row
			Cell statusCell = row.createCell(2);
			statusCell.setCellValue(user.lastName);

		}
		try (FileOutputStream outputStream = new FileOutputStream(FILE_SAVE_LOCATION + FILE_NAME)) {
			workbook.write(outputStream);
			// don't forget to close workbook to prevent memory leaks
			workbook.close();
		}

	}

	private static void createHeader(XSSFSheet sheet, XSSFWorkbook workbook) {
		Row headerRow = sheet.createRow(0);
		headerRow.createCell(0).setCellValue("User ID");
		headerRow.createCell(1).setCellValue("First name");
		headerRow.createCell(2).setCellValue("Last name");
	}

}

class User {
	Integer id;
	String firstName;
	String lastName;

	public User(Integer id, String firstName, String lastName) {
		this.id = id;
		this.firstName = firstName;
		this.lastName = lastName;
	}

}
