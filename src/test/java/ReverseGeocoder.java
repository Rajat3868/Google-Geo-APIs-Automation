import com.jayway.restassured.http.ContentType;
import com.jayway.restassured.response.Response;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import static com.jayway.restassured.RestAssured.given;

public class ReverseGeocoder {

    @Test
    public void RevGeo() throws Exception {


        File src = new File("C:\\Users\\User\\Downloads\\OutletToFetchLocalityAddressFromLatLonRemainingAll.xlsx");
        FileInputStream fis = new FileInputStream(src);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sh1 = wb.getSheetAt(2);


        String api_key = "";
        int lastRowNum = sh1.getLastRowNum();

        for (int i = 0; i <= 32; i++) {

            String lat = sh1.getRow(i).getCell(18).getRawValue();
            String longi = sh1.getRow(i).getCell(19).getRawValue();

            try {
                Response resp = given().
                        queryParam("latlng", "" + lat + "," + longi + "").
                        queryParam("key", api_key).
                        post("https://maps.googleapis.com/maps/api/geocode/json");

                String cont_type = resp.getContentType();
                if (cont_type.equals("application/json; charset=UTF-8")) {
                    try {
                        System.out.println(lat + "  " + longi);
                        String locality = resp.then().contentType(ContentType.JSON).extract().path("results[0].address_components[2].long_name");
                        String street_add = resp.then().contentType(ContentType.JSON).extract().path("results[0].address_components[1].long_name");
                        String address = resp.then().contentType(ContentType.JSON).extract().path("results[0].formatted_address");
                        String status = resp.then().contentType(ContentType.JSON).extract().path("status");
                        System.out.println(status);
                        System.out.println(i + " done");
                        System.out.println(locality);
                        System.out.println(street_add);
                        System.out.println(address);

//                    System.out.println(resp.asString());
                        sh1.getRow(i).createCell(23).setCellValue(locality);
                        sh1.getRow(i).createCell(22).setCellValue(street_add);
                        sh1.getRow(i).createCell(21).setCellValue(address);
                    }
                    catch (Exception e){
                        String address = resp.then().contentType(ContentType.JSON).extract().path("plus_code.compound_code");
                        sh1.getRow(i).createCell(21).setCellValue(address);
                        sh1.getRow(i).createCell(20).setCellValue("INVALID");
                        System.out.println(address);
                    }
                }
                else {
                    System.out.println(i + cont_type + " Content type found");
                    System.out.println(resp.asString());
                    Thread.sleep(10000);
                }

                Thread.sleep(1000);
            }
            catch (Exception e){

                System.out.println(e);
                e.printStackTrace();
                System.out.println("Exception found. Sleeping for 2 secs. Please check NETWORK CONNECTION");
                Thread.sleep(2000);
            }
        }
        FileOutputStream fos = new FileOutputStream(src);
        wb.write(fos);
    }
}
