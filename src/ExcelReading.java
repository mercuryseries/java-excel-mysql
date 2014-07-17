import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExcelReading {

	@SuppressWarnings("rawtypes")
	public static void main(String[] args) {

		// Déclaration des variables
		ArrayList<String> values = new ArrayList<String>();
		Connection connection;
		Statement statement;

		// Récupération des données depuis le fichier excel
		try {

			InputStream input = new FileInputStream("users.xls");
			POIFSFileSystem fs = new POIFSFileSystem(input);
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);
			Iterator rows = sheet.rowIterator();

			while (rows.hasNext()) {

				values.clear();

				HSSFRow row = (HSSFRow) rows.next();

				Iterator cells = row.cellIterator();

				while (cells.hasNext()) {

					HSSFCell cell = (HSSFCell) cells.next();
					
					if (HSSFCell.CELL_TYPE_NUMERIC == cell.getCellType())
						values.add(Integer.toString((int) cell.getNumericCellValue()));
					else if (HSSFCell.CELL_TYPE_STRING == cell.getCellType())
						values.add(cell.getStringCellValue());
					
				}

				// Insertion en BDD
				try {
					Class.forName("com.mysql.jdbc.Driver").newInstance();

					connection = DriverManager.getConnection(
							"jdbc:mysql://127.0.0.1/java", "root", "");

					statement = connection.createStatement();

					String sql = String
							.format("INSERT INTO users(first_name, last_name, username, password, age) VALUES ('%s', '%s', '%s', '%s', %s)",
									values.get(0), values.get(1),
									values.get(2), values.get(3), values.get(4));

					int count = statement.executeUpdate(sql);

					// Message de succès
					if (count > 0) {
						System.out.println("Enregistrement effectué!");
					}

				} catch (Exception e) {
					e.printStackTrace();
				}

			}
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
