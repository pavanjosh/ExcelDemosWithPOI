package model;

public class Invoice {

	private String firstName;
	private String LastName;
	
	private String fullName;
	
	private Double finalRate;
	
	private Double finalTotalWorkedHours;

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	public String getLastName() {
		return LastName;
	}

	public void setLastName(String lastName) {
		LastName = lastName;
	}

	public String getFullName() {
		return fullName;
	}

	public void setFullName(String fullName) {
		this.fullName = fullName;
	}

	public Double getFinalRate() {
		return finalRate;
	}

	public void setFinalRate(Double finalRate) {
		this.finalRate = finalRate;
	}

	public Double getFinalTotalWorkedHours() {
		return finalTotalWorkedHours;
	}

	public void setFinalTotalWorkedHours(Double finalTotalWorkedHours) {
		this.finalTotalWorkedHours = finalTotalWorkedHours;
	}

	}
