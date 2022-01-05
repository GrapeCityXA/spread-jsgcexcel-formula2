package com.grapecity.bean;

import java.util.Map;

public class Calc {
	String currentSheet;
	Map<String,Object> dirtyCells;
	
	public String getCurrentSheet() {
		return currentSheet;
	}
	public void setCurrentSheet(String currentSheet) {
		this.currentSheet = currentSheet;
	}
	public Map<String, Object> getDirtyCells() {
		return dirtyCells;
	}
	public void setDirtyCells(Map<String, Object> dirtyCells) {
		this.dirtyCells = dirtyCells;
	}
	
	
}
