package com.grapecity.controller;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.grapecity.bean.Calc;
import com.grapecity.bean.Sheet;
import com.grapecity.bean.SheetChange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Visibility;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxOpenOptions;
import com.grapecity.util.GZip;

@Controller
public class SpreadController {
	public static Workbook workbook;
	static {
		workbook = new Workbook();
    }
	

	
	@RequestMapping(value = "/initWorkBook", method = RequestMethod.POST)
	@ResponseBody
	public Map<String,Object> initWorkBook() {
		Map<String,Object> returnMap = new HashMap<String,Object>();
		XlsxOpenOptions options = new XlsxOpenOptions();
		workbook.open(SpreadController.class.getClassLoader().getResourceAsStream("Excel/test2.xlsx"),options);
		//workbook.open(SpreadController.class.getClassLoader().getResourceAsStream("Excel/Wicked.xlsx"),options);
		workbook.setEnableCalculation(false);
		List<Sheet> sheetNameList = new ArrayList<Sheet>();
		for(int i=0;i<workbook.getWorksheets().getCount();i++) {
			Sheet sheet = new Sheet();
			IWorksheet worksheet = workbook.getWorksheets().get(i);
			sheet.setName(worksheet.getName());
			if(Visibility.Hidden.equals(worksheet.getVisible())) {
				sheet.setVisiable(false);
			}else {
				sheet.setVisiable(true);
			}
			sheetNameList.add(sheet);
		}
		returnMap.put("sheetNames", sheetNameList);
		
		IWorksheet activeSheet = workbook.getActiveSheet();
		String activeSheetName = activeSheet.getName();
		returnMap.put("activeSheetName", activeSheetName);
		
		String result = activeSheet.toJson();
		result = GZip.compress(result);
		returnMap.put("sheetJSON", result);

		return returnMap;
	}
	
	@RequestMapping(value = "/getSheet", method = RequestMethod.POST)
	@ResponseBody
	public String getSheet(@RequestBody SheetChange sheetChange) {
		IWorksheet oldSheet = workbook.getWorksheets().get(sheetChange.getOldSheetName());
		if(sheetChange.getDirtyCells().size()>0) {
			for(int i=0;i<sheetChange.getDirtyCells().size();i++) {
				Map<String,Object> dirtyCell = sheetChange.getDirtyCells().get(i);
				int row = (int) dirtyCell.get("row");
				int col = (int) dirtyCell.get("col");
				oldSheet.getRange(row, col).setValue(dirtyCell.get("newValue"));
			}
			workbook.setEnableCalculation(true);
		}
		
		IWorksheet newSheet = workbook.getWorksheets().get(sheetChange.getNewSheetName());
		String result = null;
		if(newSheet!=null) {
			result = newSheet.toJson();
			result = GZip.compress(result);	
		}
		return result;
	}
	
	@RequestMapping(value = "/getCalcResult", method = RequestMethod.POST)
	@ResponseBody
	public String getCalcResult(@RequestBody Calc calc) {
		IWorksheet currentSheet = workbook.getWorksheets().get(calc.getCurrentSheet());
		Map<String,Object> map = calc.getDirtyCells();
		for (Map.Entry<String, Object> entry : map.entrySet()) {
			String key = entry.getKey();
			IWorksheet dirtySheet = workbook.getWorksheets().get(key);
			List<Map<String,Object>> sheetDirtys = (List<Map<String, Object>>) entry.getValue();
			for(int i=0;i<sheetDirtys.size();i++) {
				Map<String,Object> dirtyCell = sheetDirtys.get(i);
				int row = (int) dirtyCell.get("row");
				int col = (int) dirtyCell.get("col");
				dirtySheet.getRange(row, col).setValue(dirtyCell.get("newValue"));
			}
		}
		workbook.setEnableCalculation(true);
		String result = null;
		if(currentSheet!=null) {
			result = currentSheet.toJson();
			result = GZip.compress(result);	
		}
		return result;
	}
	
	/*public static void main(String[] args) {
		Workbook workbook = new Workbook();
		XlsxOpenOptions options = new XlsxOpenOptions();
		workbook.open(SpreadJsgcExcelFormulaApplication.class.getClassLoader().getResourceAsStream("Excel/Wicked.xlsx"),options);
		workbook.setEnableCalculation(false);
		List<String> sheetNameList = new ArrayList()<String>;
		for(int i=0;i<workbook.getWorksheets().getCount();i++) {
			String sheetName = workbook.getWorksheets().get(i).getName();
			workbook.getNames();
			sheetNameList.add(sheetName);
		}
		
	}*/

}
