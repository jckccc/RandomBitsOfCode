package random;

/**
 * This program will calculate generic column names of an excel spreadsheet
 * 
 * @author Jacky
 *
 */
public class ExcelCellFinder {

	int numRows;
	int numColumns;
	int numCells;

	/**
	 * @param numColumns	number of columns in the selected Excel cells
	 * @param numRows		number of rows in the selected Excel cells
	 */
	public ExcelCellFinder(int numColumns, int numRows) {
		this.numColumns = numColumns;
		this.numRows = numRows;
		this.numCells = numRows * numColumns;
	}

	/**
	 * Gets the cell index by counting horizontally from the left
	 * 
	 * @param cellNumber	the nth cell to get the index of
	 * @return 				the index of the cell counted horizontally
	 */
	public String getCellIndexCountHorizontal(int cellNumber) {
		if (cellNumber > this.numCells || cellNumber <= 0) {
			return "";
		}

		int columnNumber = cellNumber % numColumns;
		if (columnNumber == 0) {
			columnNumber = numColumns;
		}

		int rowIndex = cellNumber / numColumns;
		if (cellNumber % numColumns != 0) {
			rowIndex++;
		}

		String columnIndex = numberToLetter(columnNumber);

		return columnIndex + rowIndex;
	}
	
	/**
	 * Gets the cell index by counting vertically from the top left
	 * 
	 * @param cellNumber	the nth cell to get the index of
	 * @return 				the index of the cell counted vertically
	 */
	public String getCellIndexCountVertical(int cellNumber) {
		if (cellNumber > this.numCells || cellNumber <= 0) {
			return "";
		}

		int columnNumber = cellNumber % numRows;
		if (columnNumber == 0) {
			columnNumber = numRows;
		}

		int rowIndex = cellNumber / numRows;
		if (cellNumber % numRows != 0) {
			rowIndex++;
		}

		String columnIndex = numberToLetter(columnNumber);

		return columnIndex + rowIndex;
	}

	/**
	 * Turns the column location into letters (now recursively!)
	 * 
	 * @param columnLocation	the nth column # that will be turned into letters
	 * @return					the letters corresponding to the column
	 */
	private String numberToLetter(int columnLocation) {
		if (columnLocation == 0) {
			return "";
		}
		int rightLetter = columnLocation % 26;
		if (rightLetter == 0) {
			rightLetter = 26;
		}
		char right = (char) (rightLetter - 1 + 'A');
		return numberToLetter((columnLocation - 1) / 26) + Character.toString(right); 
	}
	
	// Getters and Setters //////////////////////////////////

	/**
	 * Gets the number of columns for the selected excel cells
	 * 
	 * @return		number of columns
	 */
	public int getNumColumns() {
		return this.numColumns;
	}

	/**
	 * Gets the number of rows for the selected excel cells
	 * 
	 * @return		number of rows
	 */
	public int getNumRows() {
		return this.numRows;
	}

	/**
	 * Gets the total number of cells
	 * 
	 * @return		total number of cells
	 */
	public int getNumCells() {
		return this.numCells;
	}

}
