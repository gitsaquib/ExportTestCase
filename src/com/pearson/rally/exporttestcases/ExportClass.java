package com.pearson.rally.exporttestcases;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.rallydev.rest.RallyRestApi;
import com.rallydev.rest.request.QueryRequest;
import com.rallydev.rest.response.QueryResponse;
import com.rallydev.rest.util.Fetch;
import com.rallydev.rest.util.QueryFilter;

public class ExportClass {
	
	public static void main(String args[]) {
		RallyRestApi restApi;
		try {
			restApi = loginRally();
			System.out.println("Successfully logged in Rally");
			System.out.println("Starting export: "+new Date());
	    	try {
				Scanner sc=new Scanner(new FileReader("D:\\workspace\\retrievetestcases\\5.txt"));
				List testCaseIds = new ArrayList<String>();
	    		while (sc.hasNextLine()){
		        	String testCaseId = sc.nextLine();
		        	testCaseIds.add(testCaseId);
		        }
	    		exportTestCases(restApi, testCaseIds);
		        sc.close();
				System.out.println("Ending export: "+new Date());
			} catch (IOException | ParseException e) {
				System.out.println("Failed export: "+ e.getMessage());
				e.printStackTrace();
			}
		} catch (URISyntaxException e) {
			System.out.println("Failed login: "+e.getMessage());
			e.printStackTrace();
		}
	}
	
	private static void exportTestCases(RallyRestApi restApi, List<String> testcaseIds) throws IOException, ParseException {
    	String wsapiVersion = "1.43";
        restApi.setWsapiVersion(wsapiVersion);
    	QueryRequest testCaseRequest = new QueryRequest("TestCase");
    	/*
    	QueryFilter filter = new QueryFilter("Deprecated", "=", "false");
    	filter = filter.and(new QueryFilter("Type", "=", "Regression"));
    	filter = filter.and(new QueryFilter("CreationDate", ">=", "2014-11-05"));
    	filter = filter.and(new QueryFilter("LastRun", ">=", "2015-10-01"));
    	*/
    	
    	QueryFilter filter = null;
    	for(String testCaseId:testcaseIds) {
    		if(null == filter) {
    			filter = new QueryFilter("FormattedID", "=", testCaseId);
    		} else {
    			filter = filter.or(new QueryFilter("FormattedID", "=", testCaseId));
    		}
    	}
    	
        testCaseRequest.setQueryFilter(filter);
        testCaseRequest.setLimit(1);
        
        QueryResponse testCaseQueryResponse = restApi.query(testCaseRequest);
        JsonArray array = testCaseQueryResponse.getResults();
        int numberTestCaseResults = array.size();
        if(numberTestCaseResults >0) {
        	List<TestCaseDTO> testCases = new ArrayList<TestCaseDTO>();
        	for(int i=0; i<numberTestCaseResults; i++) {
        		TestCaseDTO testCase = new TestCaseDTO();
        		JsonObject object = testCaseQueryResponse.getResults().get(i).getAsJsonObject();
    	        String formattedId = object.get("FormattedID").getAsString();
    	        String name = object.get("Name").getAsString();
    	        String workProduct = "";
    	        if(null != object.get("WorkProduct") && !object.get("WorkProduct").isJsonNull()) {
           	  		JsonObject workProductObj = object.get("WorkProduct").getAsJsonObject();
           	  		if(null != workProductObj) {
           	  			workProduct = retrieveUserStory(restApi, workProductObj.get("_ref").getAsString());
           	  		}
           	  	}
    	        String type = object.get("Type").getAsString();
    	        String priority = object.get("Priority").getAsString();
    	        String owner = "";
    	        if(null != object.get("Owner") && !object.get("Owner").isJsonNull()) {
           	  		JsonObject ownerObj = object.get("Owner").getAsJsonObject();
           	  		owner = ownerObj.get("_refObjectName").getAsString();
           	  	}
    	        String method = object.get("Method").getAsString();
    	        String lastVerdict = "";
    	        if(null != object.get("LastVerdict") && !object.get("LastVerdict").isJsonNull()) {
    	        	lastVerdict = object.get("LastVerdict").getAsString();
    	        }
    	        String lastBuild = "";
        		if(null != object.get("LastBuild") && !object.get("LastBuild").isJsonNull()) {
        			lastBuild = object.get("LastBuild").getAsString();
        		}
    	        String lastRun = "";
    	        if(null != object.get("LastRun") && !object.get("LastRun").isJsonNull()) {
    	        	lastRun = object.get("LastRun").getAsString();
    	        }
    	        String creationDate = object.get("CreationDate").getAsString();
    	        String description = object.get("Description").getAsString();
    	        
    	        String validationExpectedResult = "";
    	        if(null != object.get("ValidationExpectedResult") && !object.get("ValidationExpectedResult").isJsonNull()) {
       	  			if(null != object.get("ValidationExpectedResult") && !object.get("ValidationExpectedResult").isJsonNull()){
           				validationExpectedResult = object.get("ValidationExpectedResult").getAsString();
            		}
           	  	}
    	        String validationInput = "";
    	        if(null != object.get("ValidationInput") && !object.get("ValidationInput").isJsonNull()) {
       	  			if(null != object.get("ValidationInput") && !object.get("ValidationInput").isJsonNull()){
       	  				validationInput = object.get("ValidationInput").getAsString();
            		}
           	  	}
    	        String discussionCount = "";
    	        if(null != object.get("Discussion") && !object.get("Discussion").isJsonNull()) {
           	  		JsonArray discussionObj = object.get("Discussion").getAsJsonArray();
           	  		discussionCount = discussionObj.size()+"";
    	        }
    	        String howmanyminutesdoesthistaketorun = "";
    	        if(null != object.get("Howmanyminutesdoesthistaketorun") && !object.get("Howmanyminutesdoesthistaketorun").isJsonNull()) {
    	        	howmanyminutesdoesthistaketorun = object.get("Howmanyminutesdoesthistaketorun").getAsInt()+"";
    	        }
    	        String notes = "";
    	        if(null != object.get("Notes") && !object.get("Notes").isJsonNull()) {
    	        	notes = object.get("Notes").getAsString();
    	        }
    	        String objectId = "";
    	        if(null != object.get("ObjectID") && !object.get("ObjectID").isJsonNull()) {
    	        	objectId = object.get("ObjectID").getAsString();
    	        }
    	        String objective = "";
    	        if(null != object.get("Objective") && !object.get("Objective").isJsonNull()) {
    	        	objective = object.get("Objective").getAsString();
    	        }
    	        String lastUpdateDate = "";
    	        if(null != object.get("LastUpdateDate") && !object.get("LastUpdateDate").isJsonNull()) {
    	        	lastUpdateDate = object.get("LastUpdateDate").getAsString();
    	        }
    	        String postConditions = "";
    	        if(null != object.get("PostConditions") && !object.get("PostConditions").isJsonNull()) {
    	        	postConditions = object.get("PostConditions").getAsString();
    	        }
    	        String preConditions = "";
    	        if(null != object.get("PreConditions") && !object.get("PreConditions").isJsonNull()) {
    	        	preConditions = object.get("PreConditions").getAsString();
    	        }
    	        String project = "";
    	        if(null != object.get("Project") && !object.get("Project").isJsonNull()) {
           	  		JsonObject projectObj = object.get("Project").getAsJsonObject();
           	  		if(null != projectObj) {
           	  			project = projectObj.get("_refObjectName").getAsString();
           	  		}
           	  	}
    	        String risk = "";
    	        if(null != object.get("Risk") && !object.get("Risk").isJsonNull()) {
    	        	risk = object.get("Risk").getAsString();
    	        }
    	        
    	        String tags = "";
    	        if(null != object.get("Tags") && !object.get("Tags").isJsonNull()) {
    	        	JsonArray tagsArray = object.get("Tags").getAsJsonArray();
    	        	for(int t=0; t<tagsArray.size(); t++) {
    	        		JsonObject tagObj = tagsArray.get(t).getAsJsonObject();
    	        		if(tags.isEmpty()) {
    	        			tags = tagObj.get("_refObjectName").getAsString();
    	        		} else {
    	        			tags = tags +", "+tagObj.get("_refObjectName").getAsString();
    	        		}
    	        	}
    	        }
    	        
    	        String testFolder = "";
    	        if(null != object.get("TestFolder") && !object.get("TestFolder").isJsonNull()) {
           	  		JsonObject testFolderObj = object.get("TestFolder").getAsJsonObject();
           	  		if(null != testFolderObj) {
           	  			testFolder = retrieveTestFolder(restApi, testFolderObj.get("_ref").getAsString());
           	  		}
           	  	}
    	        
    	        String isthisaCandidateforAutomation = "";
    	        if(null != object.get("IsthisaCandidateforAutomation") && !object.get("IsthisaCandidateforAutomation").isJsonNull()) {
    	        	isthisaCandidateforAutomation = Boolean.toString(object.get("IsthisaCandidateforAutomation").getAsBoolean());
    	        }
    	        
    	        String gherkinLanguage = "";
    	        if(null != object.get("GherkinLanguage") && !object.get("GherkinLanguage").isJsonNull()) {
    	        	gherkinLanguage = object.get("GherkinLanguage").getAsString();
    	        }
    	        
    	        String displayColor = "";
    	        if(null != object.get("DisplayColor") && !object.get("DisplayColor").isJsonNull()) {
    	        	displayColor = object.get("DisplayColor").getAsString();
    	        }
    	        
    	        testCase.setFormattedID(formattedId);
    	        testCase.setName(name);
    	        testCase.setWorkProduct(workProduct);
    	        testCase.setType(type);
    	        testCase.setPriority(priority);
    	        testCase.setOwner(owner);
    	        testCase.setMethod(method);
    	        testCase.setLastVerdict(lastVerdict);
    	        testCase.setLastBuild(lastBuild);
    	        testCase.setLastRun(lastRun);
    	        testCase.setCreationDate(creationDate);
    	        testCase.setDescription(Jsoup.parse(description).text());
    	        testCase.setActiveDefects("");
    	        testCase.setDiscussionCount(discussionCount);
    	        testCase.setHowManyMinutesDoesThisTakeToRun(howmanyminutesdoesthistaketorun);
    	        testCase.setLastUpdateDate(lastUpdateDate);
    	        testCase.setNotes(Jsoup.parse(notes).text());
    	        testCase.setObjectID(objectId);
    	        testCase.setObjective(Jsoup.parse(objective).text());
    	        testCase.setPostConditions(Jsoup.parse(postConditions).text());
    	        testCase.setPreConditions(Jsoup.parse(preConditions).text());
    	        testCase.setProject(project);
    	        testCase.setRisk(risk);
    	        testCase.setTags(tags);
    	        testCase.setTestFolder(testFolder);
    	        testCase.setValidationExpectedResult(Jsoup.parse(validationExpectedResult).text());
    	        testCase.setValidationInput(Jsoup.parse(validationInput).text());
    	        testCase.setIsThisACandidateForAutomation(isthisaCandidateforAutomation);
    	        testCase.setGherkinLanguage(Jsoup.parse(gherkinLanguage).text());
    	        testCase.setDisplayColor(displayColor);
    	        testCases.add(testCase);
        	}
        	
        	HSSFWorkbook workbook = new HSSFWorkbook();
        	HSSFSheet sheet = workbook.createSheet();
        	HSSFFont font = workbook.createFont();
    		font.setFontName("Trebuchet MS");
    		HSSFCellStyle style = workbook.createCellStyle();
            style.setFont(font);
            style.setWrapText(true);
            style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            style.setBorderTop(HSSFCellStyle.BORDER_THIN);
            style.setBorderRight(HSSFCellStyle.BORDER_THIN);
            style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            style.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
            int rowNum = 1;
            
        	for(TestCaseDTO testCase: testCases) {
        		HSSFRow row = sheet.createRow(rowNum);
    			
    			short cellNum = 0;
    			
    			HSSFCell cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getFormattedID());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getName());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getWorkProduct());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getType());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getPriority());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getOwner());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getMethod());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getLastVerdict());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getLastBuild());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getLastRun());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getCreationDate());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getDescription());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getActiveDefects());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getDiscussionCount());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getHowManyMinutesDoesThisTakeToRun());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getLastUpdateDate());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getNotes());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getObjectID());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getObjective());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getPostConditions());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getPreConditions());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getProject());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getRisk());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getTags());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getTestFolder());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getValidationInput());
    			cellNum++;
    			
    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getValidationExpectedResult());
    			cellNum++;

    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getIsThisACandidateForAutomation());
    			cellNum++;

    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getGherkinLanguage());
    			cellNum++;

    			cell = row.createCell(cellNum);
    			cell.setCellStyle(style);
    			cell.setCellValue(testCase.getDisplayColor());
    			cellNum++;
    			
    			rowNum++;
        	}
        	SimpleDateFormat formatter = new SimpleDateFormat("dd-mm-yy");
        	Date today = new Date();
        	String dateStr = formatter.format(today);
        	System.out.println(new File("retrieve-testcases"+dateStr+".xls").getAbsolutePath());
        	FileOutputStream outputStream = new FileOutputStream("retrieve-testcases"+dateStr+".xls");
            workbook.write(outputStream);
        }
    }
    
    private static String retrieveUserStory(RallyRestApi restApi, String userStory) throws IOException, ParseException {
    	QueryRequest storyRequest = new QueryRequest("HierarchicalRequirement");
        storyRequest.setLimit(1000);
        storyRequest.setScopedDown(true);
        storyRequest.setFetch(new Fetch("FormattedID","Name"));
        userStory = userStory.substring(userStory.lastIndexOf("/")+1);
        userStory = userStory.replace(".js", "");
        QueryFilter queryFilter = new QueryFilter("ObjectID", "=", userStory);
        storyRequest.setQueryFilter(queryFilter);
        String wsapiVersion = "1.43";
        restApi.setWsapiVersion(wsapiVersion);
        QueryResponse testSetQueryResponse = restApi.query(storyRequest);
        if(testSetQueryResponse.getResults().size() > 0) {
        	JsonObject testSetJsonObject = testSetQueryResponse.getResults().get(0).getAsJsonObject();
        	return testSetJsonObject.get("FormattedID").getAsString()+": "+testSetJsonObject.get("Name").getAsString();
        }
        return "";
    }
    
    private static String retrieveTestFolder(RallyRestApi restApi, String testFolder) throws IOException, ParseException {
    	QueryRequest storyRequest = new QueryRequest("TestFolder");
        storyRequest.setLimit(1000);
        storyRequest.setScopedDown(true);
        storyRequest.setFetch(new Fetch("FormattedID","Name"));
        testFolder = testFolder.substring(testFolder.lastIndexOf("/")+1);
        testFolder = testFolder.replace(".js", "");
        QueryFilter queryFilter = new QueryFilter("ObjectID", "=", testFolder);
        storyRequest.setQueryFilter(queryFilter);
        String wsapiVersion = "1.43";
        restApi.setWsapiVersion(wsapiVersion);
        QueryResponse testSetQueryResponse = restApi.query(storyRequest);
        if(testSetQueryResponse.getResults().size() > 0) {
        	JsonObject testSetJsonObject = testSetQueryResponse.getResults().get(0).getAsJsonObject();
        	return testSetJsonObject.get("FormattedID").getAsString()+": "+testSetJsonObject.get("Name").getAsString();
        }
        return "";
    }
    
    private static RallyRestApi loginRally() throws URISyntaxException {
    	String rallyURL = "https://rally1.rallydev.com";
     	String myUserName = "mohammed.saquib@pearson.com";
     	String myUserPassword = "Rally@123";
     	return new RallyRestApi(new URI(rallyURL), myUserName, myUserPassword);
    }
}
