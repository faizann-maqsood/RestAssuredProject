import java.io.IOException;

public class main {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		exceldataconfig excel = new exceldataconfig("C:\\exceldata\\empdata.xlsx");
		//exceldataconfig excel = null;
		int row=excel.getRowCount("C:\\exceldata\\empdata.xlsx", "Sheet1");
		int col=excel.getCellCount();
		System.out.println(row);
		System.out.println(col);


	}

}
