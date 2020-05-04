import com.jayway.restassured.http.ContentType;
import com.jayway.restassured.response.Response;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import static com.jayway.restassured.RestAssured.given;

public class GeoCoder {

    @Test
    public void Geo() throws Exception {


        File src = new File("C:\\Users\\User\\Downloads\\Outlets-LatLon-Locality Correction (1).xlsx");
        FileInputStream fis = new FileInputStream(src);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sh1 = wb.getSheetAt(1);

        String api_key = "";
        int lastRowNum = sh1.getLastRowNum();

        for (int i = 455; i <= lastRowNum; i++) {
            try {

                String locality = sh1.getRow(i).getCell(11).getStringCellValue();
                String streetline1 = sh1.getRow(i).getCell(12).getStringCellValue();
                String streetline2 = sh1.getRow(i).getCell(13).getStringCellValue();
                String outletcity = sh1.getRow(i).getCell(14).getStringCellValue();
                String state = sh1.getRow(i).getCell(16).getStringCellValue();
                String country = sh1.getRow(i).getCell(17).getStringCellValue();
                String address;

                address = (locality.trim()) + " " + (streetline1.trim()) + " " + (streetline2.trim()) + " " + (outletcity.trim()) + " " + (state.trim()) + " " + (country.trim());
                String address1 = address.replaceAll(",", "");
                String act_address = address1.replaceAll("\\s+", "+");
                String a = act_address.substring(1);
                int b = act_address.length();
                if (a.equals("+")) {
                    act_address = act_address.substring(2, b);
                    System.out.println(act_address);
                }


                Response resp = given().
                        queryParam("address", act_address).
                        queryParam("key", api_key).
                        when().
                        contentType(ContentType.JSON).
                        post("https://maps.googleapis.com/maps/api/geocode/json");
//                String response = resp.asString();
                float lat;
                float longi;

                try {
                    lat = resp.then().contentType(ContentType.JSON).extract().path("results[0].geometry.location.lat");
                    longi = resp.then().contentType(ContentType.JSON).extract().path("results[0].geometry.location.lng");
                } catch (Exception ex) {
                    System.out.println(ex + "exception caught for "+ i);
                    lat = (float) 00.00;
                    longi = (float) 00.00;
                }

                sh1.getRow(i).createCell(18).setCellValue(lat);
                sh1.getRow(i).createCell(19).setCellValue(longi);
                System.out.println(i+" "+lat+" "+longi);
                Thread.sleep(1000);

            } catch (Exception e) {
                System.out.println(e);
                e.printStackTrace();
                System.out.println("Exception found. Sleeping for 30 secs. Please check NETWORK CONNECTION, or its simply null response ");
                Thread.sleep(30000);
            }
        }
        FileOutputStream fos = new FileOutputStream(src);
        wb.write(fos);
    }

    @Test
    public void worker() throws Exception{

        File src = new File("C:\\Users\\User\\Downloads\\Outlets-LatLon-Locality Correction.xlsx");
        FileInputStream fis = new FileInputStream(src);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sh4 = wb.getSheetAt(4);
        int lastRowNum = sh4.getLastRowNum();


        for (int i = 1; i <= lastRowNum; i++) {

            String custom = sh4.getRow(i).getCell(0).getStringCellValue();
            String[] latlng = custom.split("\\s+");
            sh4.getRow(i).createCell(2).setCellValue(latlng[1]);
            sh4.getRow(i).createCell(3).setCellValue(latlng[2]);
            int j = i+454;
            System.out.println(j+" "+latlng[1]);
        }
        FileOutputStream fos = new FileOutputStream(src);
        wb.write(fos);

    }
}
