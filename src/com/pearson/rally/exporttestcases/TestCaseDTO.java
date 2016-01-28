package com.pearson.rally.exporttestcases;

import java.lang.reflect.Field;

public class TestCaseDTO {

	private String formattedID;
	private String name;
	private String workProduct;
	private String type;
	private String priority;
	private String owner;
	private String method;
	private String lastVerdict;
	private String lastBuild;
	private String lastRun;
	private String creationDate;
	private String description;
	private String activeDefects;
	private String discussionCount;
	private String howManyMinutesDoesThisTakeToRun;
	private String lastUpdateDate;
	private String notes;
	private String objectID;
	private String objective;
	private String postConditions;
	private String preConditions;
	private String project;
	private String risk;
	private String tags;
	private String testFolder;
	private String validationExpectedResult;
	private String validationInput;
	private String isThisACandidateForAutomation;
	private String gherkinLanguage;
	private String displayColor;
	
	public String getFormattedID() {
		return formattedID;
	}
	public void setFormattedID(String formattedID) {
		this.formattedID = formattedID;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getWorkProduct() {
		return workProduct;
	}
	public void setWorkProduct(String workProduct) {
		this.workProduct = workProduct;
	}
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public String getPriority() {
		return priority;
	}
	public void setPriority(String priority) {
		this.priority = priority;
	}
	public String getOwner() {
		return owner;
	}
	public void setOwner(String owner) {
		this.owner = owner;
	}
	public String getMethod() {
		return method;
	}
	public void setMethod(String method) {
		this.method = method;
	}
	public String getLastVerdict() {
		return lastVerdict;
	}
	public void setLastVerdict(String lastVerdict) {
		this.lastVerdict = lastVerdict;
	}
	public String getLastBuild() {
		return lastBuild;
	}
	public void setLastBuild(String lastBuild) {
		this.lastBuild = lastBuild;
	}
	public String getLastRun() {
		return lastRun;
	}
	public void setLastRun(String lastRun) {
		this.lastRun = lastRun;
	}
	public String getCreationDate() {
		return creationDate;
	}
	public void setCreationDate(String creationDate) {
		this.creationDate = creationDate;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public String getActiveDefects() {
		return activeDefects;
	}
	public void setActiveDefects(String activeDefects) {
		this.activeDefects = activeDefects;
	}
	public String getDiscussionCount() {
		return discussionCount;
	}
	public void setDiscussionCount(String discussionCount) {
		this.discussionCount = discussionCount;
	}
	public String getHowManyMinutesDoesThisTakeToRun() {
		return howManyMinutesDoesThisTakeToRun;
	}
	public void setHowManyMinutesDoesThisTakeToRun(
			String howManyMinutesDoesThisTakeToRun) {
		this.howManyMinutesDoesThisTakeToRun = howManyMinutesDoesThisTakeToRun;
	}
	public String getLastUpdateDate() {
		return lastUpdateDate;
	}
	public void setLastUpdateDate(String lastUpdateDate) {
		this.lastUpdateDate = lastUpdateDate;
	}
	public String getNotes() {
		return notes;
	}
	public void setNotes(String notes) {
		this.notes = notes;
	}
	public String getObjectID() {
		return objectID;
	}
	public void setObjectID(String objectID) {
		this.objectID = objectID;
	}
	public String getObjective() {
		return objective;
	}
	public void setObjective(String objective) {
		this.objective = objective;
	}
	public String getPostConditions() {
		return postConditions;
	}
	public void setPostConditions(String postConditions) {
		this.postConditions = postConditions;
	}
	public String getPreConditions() {
		return preConditions;
	}
	public void setPreConditions(String preConditions) {
		this.preConditions = preConditions;
	}
	public String getProject() {
		return project;
	}
	public void setProject(String project) {
		this.project = project;
	}
	public String getRisk() {
		return risk;
	}
	public void setRisk(String risk) {
		this.risk = risk;
	}
	public String getTags() {
		return tags;
	}
	public void setTags(String tags) {
		this.tags = tags;
	}
	public String getTestFolder() {
		return testFolder;
	}
	public void setTestFolder(String testFolder) {
		this.testFolder = testFolder;
	}
	public String getValidationExpectedResult() {
		return validationExpectedResult;
	}
	public void setValidationExpectedResult(String validationExpectedResult) {
		this.validationExpectedResult = validationExpectedResult;
	}
	public String getValidationInput() {
		return validationInput;
	}
	public void setValidationInput(String validationInput) {
		this.validationInput = validationInput;
	}
	public String getIsThisACandidateForAutomation() {
		return isThisACandidateForAutomation;
	}
	public void setIsThisACandidateForAutomation(
			String isThisACandidateForAutomation) {
		this.isThisACandidateForAutomation = isThisACandidateForAutomation;
	}
	public String getGherkinLanguage() {
		return gherkinLanguage;
	}
	public void setGherkinLanguage(String gherkinLanguage) {
		this.gherkinLanguage = gherkinLanguage;
	}
	public String getDisplayColor() {
		return displayColor;
	}
	public void setDisplayColor(String displayColor) {
		this.displayColor = displayColor;
	}
	
/*	public String toString() {
		  StringBuilder result = new StringBuilder();
		  String newLine = System.getProperty("line.separator");

		  result.append( this.getClass().getName() );
		  result.append( " Object {" );
		  result.append(newLine);

		  Field[] fields = this.getClass().getDeclaredFields();
		  for ( Field field : fields  ) {
		    result.append("  ");
		    try {
		      result.append( field.getName() );
		      result.append(": ");
		      result.append( field.get(this) );
		    } catch ( IllegalAccessException ex ) {
		      System.out.println(ex);
		    }
		    result.append(newLine);
		  }
		  result.append("}");

		  return result.toString();
		}
 	*/
}
