package excel.finance;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.ArrayList;
import java.util.HashMap;

public class ExcelParser {
    public static List<List<List<String>>> readFile(String filePath) throws IOException, InvalidFormatException {
	List<List<List<String>>> dataEntries = new ArrayList<>();

	// Create a workbook from an Excel file
	Workbook workbook = WorkbookFactory.create(new File(filePath));

	// Create a sheetIterator to iterate through all of the sheets
	Iterator<Sheet> sheetIterator = workbook.sheetIterator();
	while (sheetIterator.hasNext()) {
	    List<List<String>> sheetEntries = new ArrayList<>();
	    Sheet sheet = sheetIterator.next();
	    DataFormatter dataFormatter = new DataFormatter();

	    // Iterate through the cells
	    for (Row row : sheet) {
		List<String> rowEntries = new ArrayList<>();
		for (Cell cell : row) {
		    // Add each row into a separate ArrayList
		    rowEntries.add(dataFormatter.formatCellValue(cell));
		}
		// Add each ArrayList row into the 2D ArrayList
		sheetEntries.add(rowEntries);
	    }
	    // Add each 2D ArrayList sheet into the 3D ArrayList
	    dataEntries.add(sheetEntries);
	}
	// Always close after opening
	workbook.close();
	return dataEntries;
    }

    public static void createFile(String filePath, List<List<List<String>>> dataEntries)
	    throws IOException, InvalidFormatException {
	// Create workbook to write to
	Workbook workbook = new XSSFWorkbook();

	// Create the dates and their corresponding day of the week
	List<String> dates = generateDates(18, 19);
	HashMap<String, String> dayOfWeek = generateDayOfWeek(dates);
	List<Integer> daysInMonths = generateDaysInMonths(18, 19);

	for (int i = 0; i < dataEntries.size(); i++) {
	    // Create a sheet and get a sheet from list
	    Sheet sheet = workbook.createSheet("Sheet " + (i + 1));
	    List<List<String>> sheetList = dataEntries.get(i);
	    int rowIndex = 0;

	    int columns = 0;
	    for (int j = 0; j < sheetList.size(); j++) {
		// Get a row from sheetList
		List<String> rowList = sheetList.get(j);
		int incrementIndex = 1;
		int dateIncrementer = 0;
		String incrementType = "";
		int month = 0;

		// If empty cells were parsed
		if (rowList.get(0).isEmpty()) {
		    break;
		}

		// Find index of the date in the rowList and parse the dateIncrementer and save
		// type
		for (; incrementIndex < rowList.size(); incrementIndex++) {
		    if (rowList.get(incrementIndex - 1).matches("\\d{1,2}/\\d{1,2}/\\d{1,2}")) {
			if (rowList.get(incrementIndex).matches("\\d+")) {
			    dateIncrementer = Integer.parseInt(rowList.get(incrementIndex));
			    incrementType = "Explicit";
			    break;
			} else if (rowList.get(incrementIndex).equals("Annually")) {
			    dateIncrementer = 365;
			    incrementType = "Annually";
			    break;
			} else if (rowList.get(incrementIndex).equals("Monthly")) {
			    month = Integer.parseInt(rowList.get(incrementIndex - 1).split("/")[0]);
			    int year = Integer.parseInt(rowList.get(incrementIndex - 1).split("/")[2]);
			    dateIncrementer = daysInMonths.get(month - 1 + ((year - 18) * 12));
			    month = month - 1 + ((year - 18) * 12);
			    incrementType = "Monthly";
			    break;
			}
		    }
		}

		int dateIndexInDates = search(dates, rowList.get(incrementIndex - 1));
		if (dateIndexInDates < 0) {
		    break;
		}
		String currDate = dates.get(dateIndexInDates);

		while (dateIndexInDates < dates.size()) {
		    Row row = sheet.createRow(rowIndex++);
		    int rowInfoIndex = 0;
		    for (int k = 0; k < rowList.size(); k++) {
			if (rowList.get(k).isEmpty()) {
			    dateIndexInDates = dates.size();
			    break;
			}
			columns = Math.max(columns, k);
			// Iterate through each row and write to the workbook cell accordingly
			if (k == incrementIndex) {
			    continue;
			    // Print the date and update it
			} else if (k == incrementIndex - 1) {
			    currDate = dates.get(dateIndexInDates);
			    // Account for holidays
			    currDate = adjustHoliday(dates, currDate, dateIndexInDates);
			    // Account for weekends
			    int tempDateIndex = search(dates, currDate);
			    if (dayOfWeek.get(currDate).equals("Saturday")) {
				currDate = dates.get(tempDateIndex - 1);
			    } else if (dayOfWeek.get(currDate).equals("Sunday")) {
				currDate = dates.get(tempDateIndex - 2);
			    }

			    row.createCell(rowInfoIndex++).setCellValue(currDate);
			    dateIndexInDates += dateIncrementer;

			    // Special case for monthly where we update dateIncrementer
			    if (incrementType.equals("Monthly")) {
				month++;
				if (month < daysInMonths.size()) {
				    dateIncrementer = daysInMonths.get(month);
				}
			    }
			} else {
			    row.createCell(rowInfoIndex++).setCellValue(rowList.get(k));
			}
		    }
		}
	    }
	    // Auto adjust the size of all the columns so text fits
	    for (int l = 0; l < columns; l++) {
		sheet.autoSizeColumn(l);
	    }
	}
	// Write the output to a file and close stream
	int dotIndex = filePath.lastIndexOf('.');
	String newFilePath = filePath.substring(0, dotIndex);
	// System.out.println(newFilePath);
	FileOutputStream fileOut = new FileOutputStream(newFilePath + " GENERATED.xlsx");
	workbook.write(fileOut);
	fileOut.close();

	// Close workbook
	workbook.close();

    }

    private static int search(List<String> dates, String day) {
	// Find the index of the starting date
	for (int i = 0; i < dates.size(); i++) {
	    if (day.equals(dates.get(i))) {
		return i;
	    }
	}
	return -1;
    }

    private static String adjustHoliday(List<String> dates, String currDate, int index) {
	if (currDate.equals("1/1/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("1/1/18")) {
	    currDate = dates.get(++index);
	}
	if (currDate.equals("12/31/18") || currDate.equals("12/31/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("12/25/18") || currDate.equals("12/25/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("12/24/18") || currDate.equals("12/24/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("11/22/18") || currDate.equals("11/28/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("11/12/18") || currDate.equals("11/11/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("10/8/18") || currDate.equals("10/14/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("9/3/18") || currDate.equals("9/2/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("7/4/18") || currDate.equals("7/4/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("5/28/18") || currDate.equals("5/27/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("2/19/18") || currDate.equals("2/18/19")) {
	    currDate = dates.get(--index);
	}
	if (currDate.equals("1/15/18") || currDate.equals("1/21/19")) {
	    currDate = dates.get(--index);
	}
	return currDate;
    }

    private static List<Integer> generateDaysInMonths(int startYear, int endYear) {
	List<Integer> daysInMonths = new ArrayList<>();
	// Generate the number of days in all of the months from startYear to endYear
	// inclusive
	for (; startYear <= endYear; startYear++) {
	    for (int i = 1; i <= 12; i++) {
		if (i == 2 && startYear % 4 == 0) {
		    daysInMonths.add(29);
		} else if (i == 2) {
		    daysInMonths.add(28);
		} else if (i == 4 || i == 6 || i == 9 || i == 11) {
		    daysInMonths.add(30);
		} else {
		    daysInMonths.add(31);
		}
	    }
	}
	return daysInMonths;
    }

    private static List<String> generateDates(int startYear, int endYear) {
	ArrayList<String> dates = new ArrayList<>();

	// Generate all of the days from startYear to endYear inclusive, accounting for
	// leap years
	for (; startYear <= endYear; startYear++) {
	    for (int i = 1; i <= 12; i++) {
		for (int j = 1; j <= 31; j++) {
		    if (i == 2 && j == 29 && startYear % 4 == 0) {
			String tempDate = i + "/" + j + "/" + startYear;
			dates.add(tempDate);
			break;
		    } else if (i == 2 && j >= 29) {
			break;
		    } else if ((i == 4 || i == 6 || i == 9 || i == 11) && j == 31) {
			break;
		    }
		    String tempDate = i + "/" + j + "/" + startYear;
		    dates.add(tempDate);
		}
	    }
	}
	return dates;
    }

    private static HashMap<String, String> generateDayOfWeek(List<String> dates) {
	HashMap<String, String> dayOfWeek = new HashMap<>();
	String[] namesOfDays = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
	int weekDay = 1; // Start at 1 because 1/1/18 is a Monday

	Iterator<String> dateIter = dates.iterator();
	while (dateIter.hasNext()) {
	    // Hash the date with the corresponding day of the week
	    dayOfWeek.put(dateIter.next(), namesOfDays[weekDay++]);
	    if (weekDay >= 7) {
		weekDay = 0;
	    }
	}
	return dayOfWeek;
    }
}
