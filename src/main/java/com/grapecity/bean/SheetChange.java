package com.grapecity.bean;

import java.util.List;
import java.util.Map;

public class SheetChange {
	String oldSheetName;
	String newSheetName;
	List<Map<String,Object>> dirtyCells;
	public String getOldSheetName() {
		return oldSheetName;
	}
	public void setOldSheetName(String oldSheetName) {
		this.oldSheetName = oldSheetName;
	}
	public String getNewSheetName() {
		return newSheetName;
	}
	public void setNewSheetName(String newSheetName) {
		this.newSheetName = newSheetName;
	}
	public List<Map<String, Object>> getDirtyCells() {
		return dirtyCells;
	}
	public void setDirtyCells(List<Map<String, Object>> dirtyCells) {
		this.dirtyCells = dirtyCells;
	}
	
}
