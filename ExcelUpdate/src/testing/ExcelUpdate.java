package testing;

import java.io.IOException;

public class ExcelUpdate {

	public static void main(String[] args) throws IOException, InterruptedException {
		Boolean b= false; String data= null; String x= null; String found = null;
		String mapData= "C:\\Users\\SambitKumarPanda\\eclipse-workspace\\ExcelUpdate\\src\\MapData.xlsx";
		String DSRfile= Functions.readData(1, 1, "MapData", mapData);
		String DSRsheet= Functions.readData(1, 2, "MapData", mapData);
		String DEfile= Functions.readData(2, 1, "MapData", mapData);
		String DEsheet= Functions.readData(2, 2, "MapData", mapData);
		int srcReadStart= Functions.readIntData(1, 3, "MapData", mapData);
		int srcReadEnd= Functions.readIntData(1, 4, "MapData", mapData);
		int targetReadStart= Functions.readIntData(2, 3, "MapData", mapData);
		int targetReadEnd= Functions.readIntData(2, 4, "MapData", mapData);
		int targetRow= 0;
		int targetCell= 0;
		
		for(int i=srcReadStart-1;i<=srcReadEnd-1;i++) {
			x= Functions.readData(i, 0, DSRsheet, DSRfile);
			if(x!=null) {
				targetRow= Functions.findRow(x, targetReadStart, targetReadEnd, DEsheet, DEfile);
				if(targetRow!=0) {
					found= "yes";
					for(int j=5;j<=29;j++) {
						data= Functions.readData(i, Functions.mapColumn(Functions.readData(1, j, "MapData", mapData)), DSRsheet, DSRfile);
						targetCell= Functions.mapColumn(Functions.readData(2, j, "MapData", mapData));
									
						b= Functions.setCellData(DEfile, DEsheet, targetRow, targetCell, data);
									
						if(b==true) {
								System.out.println(x+" - "+(j)+" Done");
						}
						else{
								System.out.println(x+" "+(j)+" Not Done");
						}
					}
				}
				else {
					found= "no";
				}
			}
			else {
				continue;
			}
		}
		if(found.equals("no")) {
			System.out.println(x+": Not found in the target sheet.");
		}
	}
}