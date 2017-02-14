package random;

/**
 * This program will calculate generic column names of an excel spreadsxheet
 * 
 * @author Jacky
 *
 */
public class ExcelCellFinder {

	int numRows;
	int numColumns;
	int numCells;

	public ExcelCellFinder(int numColumns, int numRows) {
		this.numColumns = numColumns;
		this.numRows = numRows;
		this.numCells = numRows * numColumns;
	}

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

	public String numberToLetter(int columnLocation) {
		int rightLetter = columnLocation % 26;
		if (rightLetter == 0) {
			rightLetter = 26;
		}
		char right = (char) (rightLetter - 1 + 'A');

		int leftLetter = (columnLocation - 1) / 26;
		char left = (char) (leftLetter - 1 + 'A');

		return leftLetter == 0 
				? Character.toString(right)
				: Character.toString(left) + Character.toString(right); 
	}
	
	// Getters and Setters

	public int getNumColumns() {
		return this.numColumns;
	}

	public int getNumRows() {
		return this.numRows;
	}

	public int getNumCells() {
		return this.numCells;
	}

}
