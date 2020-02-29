package datadriven;

import java.io.IOException;

import org.json.simple.JSONObject;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.restassured.RestAssured;
import io.restassured.http.Method;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;

public class DataDrivenTest_AddNewEmployees {
	
	@Test(dataProvider="empdataprovider")
	public void postNewEmployees(String ename, String esal, String eage)
	{
		RestAssured.baseURI = "http://dummy.restapiexample.com/api/v1";
		RequestSpecification httprequest = RestAssured.given();
	
		JSONObject reaparam = new JSONObject();
		reaparam.put("name", ename);
		reaparam.put("salary", esal);
		reaparam.put("age", eage);
		
		httprequest.header("Content-Type","application/json");
		
		httprequest.body(reaparam.toJSONString());
		Response response = httprequest.request(Method.POST,"/create");
		
		String responseBody = response.getBody().asString();
		
		System.out.println(responseBody);
		Assert.assertEquals(responseBody.contains(ename),true);
		Assert.assertEquals(responseBody.contains(esal),true);
		Assert.assertEquals(responseBody.contains(eage),true);
		
		int code = response.getStatusCode();
		
		Assert.assertEquals(code, 200);
	}
	
	@DataProvider(name="empdataprovider")
	public String[][] getEmpData() throws IOException
	{
		exceldataconfig excel = new exceldataconfig("C:\\exceldata\\empdata4.xlsx");
		String path = "C:\\exceldata\\empdata4.xlsx";
		//exceldataconfig excel = null;
		int row=excel.getRowCount("C:\\exceldata\\empdata4.xlsx", "Sheet1");
		int col=excel.getCellCount();
		System.out.println(row);
		System.out.println(col);
		
		String empdata[][] = new String[row][col];
		
		for(int i=1;i<=row;i++)
		{
			for(int j=0;j<col;j++)
			{
				empdata[i-1][j]= excel.getCellData(path,"Sheet1",i,j);
				
			}
		}
		
		//String empdata[][] = {{"ali","300","25"}, {"zain","500","21"}, {"noor","700","29"}};
		return empdata;
	}

}
