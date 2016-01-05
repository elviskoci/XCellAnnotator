/**
 * 
 */
package de.tudresden.annotator.test;

/**
 * @author Elvis Koci
 */
public class RangeTests {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		String c1 = "AA", c2 = "ABZ";		
		System.out.println(countCellsBetweenColumns(c1, c2, true));
	}
	
	public static int countCellsBetweenColumns(String startColumn, String endColumn, boolean inclusive){
		
		int countFields = 0;	
		if(startColumn.length() < endColumn.length()){
			
			for (int i = 0; i < startColumn.length(); i++) {
				int diff = endColumn.charAt(i) - startColumn.charAt(i);
				if(diff!=0)
					countFields = countFields + (diff * 26) + 26;
			}

			for (int j = endColumn.length()-startColumn.length(); j < endColumn.length(); j++) {
				countFields = countFields + ((endColumn.charAt(j)-'A') + 1);
			}
		}
		
		if(startColumn.length()==endColumn.length()){
			
			if(startColumn.length()>1){
				for (int k = 0; k < startColumn.length(); k++) {
					int diff = endColumn.charAt(k) - startColumn.charAt(k);
					
					if(k==(startColumn.length()-1)){
						countFields = countFields + diff + 1;
						break;
					}
					
					if(diff!=0){
						countFields = countFields + (diff * 26);
					}
					
					System.out.println(countFields);
				}
			}else{
				countFields = endColumn.compareTo(startColumn) + 1; 
			}
		}
		
		return countFields;
	}
	
	public static boolean checkForContainment(String range1 , String range2){
				
		if(range1==null || range1.compareTo("")==0 || range1.length()<2 || !range1.matches("^[a-zA-Z\\$]{3,5}[0-9]{1,7}") // XFD1048576
				|| range2==null || range2.compareTo("")==0 || range2.length()<2 || !range2.matches("^[a-zA-Z\\$]{3,5}+[0-9]{1,7}$")){
			System.out.println("One or both of the string are not valid range addresses!!!");
			return false;
		}
		
		String r1Cells[] = range1.split(":");
		String r2Cells[] = range2.split(":");
		
		String r1TopLeft = null, r1DownRight = null, r2TopLeft = null, r2DownRight = null;	
		r1TopLeft =  r1Cells[0];
		if(r1Cells.length == 1){
			r1DownRight = r1Cells[0];
		}else{
			r1DownRight = r1Cells[1];
		}
		
		r2TopLeft = r2Cells[0];	
		if(r2Cells.length == 1){
			r2DownRight = r2Cells[0];
		}else{
			r2DownRight = r2Cells[1];
		}
			
		int topColComp = compareCellsByColumn(r1TopLeft, r2TopLeft);
		int topRowComp = compareCellsByRow(r1TopLeft, r2TopLeft);
			
		int downColComp = compareCellsByColumn(r1DownRight, r2DownRight);
		int downRowComp = compareCellsByRow(r1DownRight, r2DownRight);
				
		boolean downRightCellContained =  topColComp<=0 && topRowComp<=0;
		boolean topLeftCellContained = downColComp>=0 && downRowComp>=0; 
		
		if(downRightCellContained && topLeftCellContained)
			return true;
		
		return false;
	}
	
	
	public static boolean checkForPartialContainment(String range1 , String range2){
		
		
		if(range1==null || range1.compareTo("")==0 || range1.length()<2 || !range1.matches("^[a-zA-Z\\$]{3,5}[0-9]{1,7}") // XFD1048576
				|| range2==null || range2.compareTo("")==0 || range2.length()<2 || !range2.matches("^[a-zA-Z\\$]{3,5}+[0-9]{1,7}$")){
			System.out.println("One or both of the string are not valid range addresses!!!");
			return false;
		}
		
		
		
		String r1Cells[] = range1.split(":");
		String r2Cells[] = range2.split(":");
		
		String r1TopLeft = null, r1DownRight = null, r2TopLeft = null, r2DownRight = null;	
		r1TopLeft =  r1Cells[0];	
		if(r1Cells.length == 1){
			r1DownRight = r1Cells[0];
		}else{
			r1DownRight = r1Cells[1];
		}
		
		r2TopLeft = r2Cells[0];	
		if(r2Cells.length == 1){
			r2DownRight = r2Cells[0];
		}else{
			r2DownRight = r2Cells[1];
		}
			
		int topColComp = compareCellsByColumn(r1TopLeft, r2TopLeft);
		int topRowComp = compareCellsByRow(r1TopLeft, r2TopLeft);
			
		int downColComp = compareCellsByColumn(r1DownRight, r2DownRight);
		int downRowComp = compareCellsByRow(r1DownRight, r2DownRight);
		
		int topDownColComp = compareCellsByColumn(r1TopLeft, r2DownRight);
		int topDownRowComp = compareCellsByRow(r1TopLeft, r2DownRight);
		
		int downTopColComp = compareCellsByColumn(r1DownRight, r2TopLeft);
		int downTopRowComp = compareCellsByRow(r1DownRight, r2TopLeft);
		
		boolean downRightCellContained =  topColComp<=0 && topRowComp<=0 && downColComp>=0 && downRowComp>=0; 
		
		boolean columnWithinBorders =  ((topDownColComp<=0 && downColComp>=0) || (downTopColComp>=0 && topColComp<=0)) && r2Cells.length == 2 ;
		boolean rowWithinBorders =  ((topDownRowComp<=0 && downRowComp>=0) || (downTopRowComp>=0 && topRowComp<=0)) && r2Cells.length == 2 ;
				
		if(columnWithinBorders || rowWithinBorders || downRightCellContained)
			return true;
		
		return false; 
	}
	
	
	public static int compareCellsByColumn(String cell1Address, String cell2Address){
				
		String col1 =  cell1Address.replaceAll("[0-9\\$]+","");
		String col2 =  cell2Address.replaceAll("[0-9\\$]+","");
		
		return col1.compareTo(col2);
	}
	
	
	public static int compareCellsByRow(String cell1Address, String cell2Address){
		
		int row1 =  Integer.valueOf(cell1Address.replaceAll("[^0-9]+",""));
		int row2 =  Integer.valueOf(cell2Address.replaceAll("[^0-9]+",""));
				
		return row1 - row2;
	}
}
