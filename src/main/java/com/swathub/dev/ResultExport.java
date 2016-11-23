package com.swathub.dev;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.http.auth.AuthScope;
import org.apache.http.auth.UsernamePasswordCredentials;
import org.apache.http.client.CredentialsProvider;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.impl.client.BasicCredentialsProvider;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.hssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import javax.imageio.IIOException;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ResultExport {
	private static Map<String, String> valueMap = new HashMap<String, String>();
	static {
		valueMap.put("type.flow", "Flow");
		valueMap.put("type.sop", "System Operation");
		valueMap.put("type.pop", "Page Operation");
	}

	private static String apiGet(URIBuilder url, String user, String pass, JSONObject proxy) throws Exception {
		CredentialsProvider credsProvider = new BasicCredentialsProvider();
		credsProvider.setCredentials(
				new AuthScope(url.getHost(), url.getPort()),
				new UsernamePasswordCredentials(user, pass));
		CloseableHttpClient httpclient = HttpClients.custom()
				.setDefaultCredentialsProvider(credsProvider)
				.build();

		String result = null;
		try {
			HttpGet httpget = new HttpGet(url.build());
			CloseableHttpResponse response = httpclient.execute(httpget);
			try {
				result = EntityUtils.toString(response.getEntity());
			} finally {
				response.close();
			}
		} finally {
			httpclient.close();
		}
		return result;
	}

	private static int fetchSteps(JSONArray steps, HSSFWorkbook workbook, HSSFCreationHelper creationHelper,
								  HSSFSheet sheet, int rowCnt, HashMap<String, JSONObject> validResults, String[] platforms) throws Exception {
		HSSFRow row;
		HSSFFont boldFont = workbook.createFont();
		boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);

		HSSFCellStyle titleStyle = workbook.createCellStyle();
		titleStyle.setFont(boldFont);

		HSSFCellStyle tableHeader = workbook.createCellStyle();
		tableHeader.setFont(boldFont);
		tableHeader.setAlignment(CellStyle.ALIGN_CENTER);
		tableHeader.setBorderBottom(CellStyle.BORDER_THIN);
		tableHeader.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		tableHeader.setBorderLeft(CellStyle.BORDER_THIN);
		tableHeader.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		tableHeader.setBorderRight(CellStyle.BORDER_THIN);
		tableHeader.setRightBorderColor(IndexedColors.BLACK.getIndex());
		tableHeader.setBorderTop(CellStyle.BORDER_THIN);
		tableHeader.setTopBorderColor(IndexedColors.BLACK.getIndex());

		HSSFCellStyle tableCell =  workbook.createCellStyle();
		tableCell.setBorderBottom(CellStyle.BORDER_THIN);
		tableCell.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		tableCell.setBorderLeft(CellStyle.BORDER_THIN);
		tableCell.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		tableCell.setBorderRight(CellStyle.BORDER_THIN);
		tableCell.setRightBorderColor(IndexedColors.BLACK.getIndex());
		tableCell.setBorderTop(CellStyle.BORDER_THIN);
		tableCell.setTopBorderColor(IndexedColors.BLACK.getIndex());

		HSSFCellStyle cellDisabled =  workbook.createCellStyle();
		cellDisabled.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		cellDisabled.setBorderBottom(CellStyle.BORDER_THIN);
		cellDisabled.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		cellDisabled.setBorderLeft(CellStyle.BORDER_THIN);
		cellDisabled.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		cellDisabled.setBorderRight(CellStyle.BORDER_THIN);
		cellDisabled.setRightBorderColor(IndexedColors.BLACK.getIndex());
		cellDisabled.setBorderTop(CellStyle.BORDER_THIN);
		cellDisabled.setTopBorderColor(IndexedColors.BLACK.getIndex());

		HSSFCellStyle lineDisabled =  workbook.createCellStyle();
		lineDisabled.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());

		for (int i = 0; i < steps.length(); i++) {
			JSONObject step = steps.getJSONObject(i);
			if (step.getString("seqNo").equals("")) {
				row = sheet.createRow(rowCnt++);
				row.createCell(0).setCellValue("Test Case");
				row.getCell(0).setCellStyle(titleStyle);
				row.createCell(1).setCellValue(step.getString("stepTitle"));

				row = sheet.createRow(rowCnt++);
				row.createCell(0).setCellValue("Parameter");
				row.getCell(0).setCellStyle(tableHeader);
				row.createCell(1).setCellValue("Name");
				row.getCell(1).setCellStyle(tableHeader);
				row.createCell(2).setCellValue("Value");
				row.getCell(2).setCellStyle(tableHeader);

				for (int j = 0; j < step.getJSONArray("paramData").length(); j++) {
					JSONObject item = step.getJSONArray("paramData").getJSONObject(j);
					row = sheet.createRow(rowCnt++);
					if (item.getBoolean("runtimeEnabled")) {
						row.createCell(0).setCellValue(item.getString("name"));
						row.getCell(0).setCellStyle(tableCell);
						row.createCell(1).setCellValue(item.getString("variable"));
						row.getCell(1).setCellStyle(tableCell);
						row.createCell(2).setCellValue(item.getString("value"));
						row.getCell(2).setCellStyle(tableCell);
					} else {
						row.createCell(0).setCellValue(item.getString("name"));
						row.getCell(0).setCellStyle(cellDisabled);
						row.createCell(1).setCellValue(item.getString("variable"));
						row.getCell(1).setCellStyle(cellDisabled);
						row.createCell(2).setCellValue(item.getString("value"));
						row.getCell(2).setCellStyle(cellDisabled);
					}
				}
				rowCnt++;
			} else {
				String title = step.getString("stepTitle");
				if (step.has("typeName") && !step.isNull("typeName")) {
					title = title + "(" + step.getString("typeName") + ")";
				}
				row = sheet.createRow(rowCnt++);
				if (step.getBoolean("executed")) {
					row.createCell(0).setCellValue(step.getString("seqNo"));
					row.createCell(1).setCellValue("Name");
					row.getCell(1).setCellStyle(titleStyle);
					row.createCell(2).setCellValue(title);
				} else {
					row.createCell(0).setCellValue(step.getString("seqNo"));
					row.createCell(1).setCellValue("Name");
					row.getCell(1).setCellStyle(lineDisabled);
					row.createCell(2).setCellValue(title);
				}

				row = sheet.createRow(rowCnt++);
				row.createCell(1).setCellValue("Type");
				row.getCell(1).setCellStyle(titleStyle);
				row.createCell(2).setCellValue(valueMap.get("type." + step.getString("type")));

				List<JSONObject> paramData = new ArrayList<JSONObject>();
				JSONObject comment = null;
				for (int j = 0; j < step.getJSONArray("paramData").length(); j++) {
					JSONObject item = step.getJSONArray("paramData").getJSONObject(j);
					if (item.getString("code").equals("comment")) {
						comment = item;
					} else {
						paramData.add(item);
					}
				}

				if (comment != null) {
					row = sheet.createRow(rowCnt++);
					row.createCell(1).setCellValue("Comment");
					row.getCell(1).setCellStyle(titleStyle);
					row.createCell(2).setCellValue(comment.getString("value"));
				}

				if (!step.isNull("evidences") && step.getJSONObject("evidences").has("url")) {
					row = sheet.createRow(rowCnt++);
					row.createCell(1).setCellValue("URL");
					row.getCell(1).setCellStyle(titleStyle);
					row.createCell(2).setCellValue(step.getJSONObject("evidences").getString("url"));
				}

				row = sheet.createRow(rowCnt++);
				row.createCell(1).setCellValue("Parameter");
				row.getCell(1).setCellStyle(tableHeader);
				row.createCell(2).setCellValue("Name");
				row.getCell(2).setCellStyle(tableHeader);
				row.createCell(3).setCellValue("Value");
				row.getCell(3).setCellStyle(tableHeader);

				for (JSONObject item : paramData) {
					row = sheet.createRow(rowCnt++);
					if (!item.isNull("runtimeEnabled") && item.getBoolean("runtimeEnabled")) {
						row.createCell(1).setCellValue(item.getString("name"));
						row.getCell(1).setCellStyle(tableCell);
						row.createCell(2).setCellValue(item.getString("variable"));
						row.getCell(2).setCellStyle(tableCell);
						row.createCell(3).setCellValue(item.getString("value"));
						row.getCell(3).setCellStyle(tableCell);
					} else {
						row.createCell(1).setCellValue(item.getString("name"));
						row.getCell(1).setCellStyle(cellDisabled);
						row.createCell(2).setCellValue(item.getString("variable"));
						row.getCell(2).setCellStyle(cellDisabled);
						row.createCell(3).setCellValue(item.getString("value"));
						row.getCell(3).setCellStyle(cellDisabled);
					}
				}

				if (!step.isNull("evidences") && step.getJSONObject("evidences").has("screenshots")) {
					row = sheet.createRow(rowCnt++);
					JSONArray screenshots = step.getJSONObject("evidences").getJSONArray("screenshots");
					for (int j = 0; j < screenshots.length(); j++) {
						int rowCntTemp = 0, colCnt = 1, size = 1;
						String screenshot = screenshots.getString(j).replace("_s.png", ".png");
						for (String platform : platforms) {
							JSONObject validResult = validResults.get(platform);
							if (validResult == null) continue;

							row.createCell(colCnt).setCellValue(validResult.getString("execPlatform"));
							URL imageUrl = new URL(validResult.getString("baseURL").concat(screenshot));
							BufferedImage image;
							try {
								image = ImageIO.read(imageUrl);
							} catch (IIOException e) {
								System.out.println("Image URL may not exist:" + imageUrl.toString());
								break;
							}
							ByteArrayOutputStream baos = new ByteArrayOutputStream();
							ImageIO.write(image, "png", baos);

							try {
								int pictureIdx = workbook.addPicture(baos.toByteArray(), Workbook.PICTURE_TYPE_PNG);
								HSSFPatriarch drawing = sheet.createDrawingPatriarch();

								HSSFClientAnchor anchor = creationHelper.createClientAnchor();
								anchor.setCol1(colCnt);
								anchor.setRow1(rowCnt);
								HSSFPicture picture = drawing.createPicture(anchor, pictureIdx);
								picture.resize();
								colCnt = picture.getPreferredSize().getCol2() + 1;
								if (picture.getPreferredSize().getRow2() > rowCntTemp) {
									rowCntTemp = picture.getPreferredSize().getRow2();
								}
							} catch (IllegalArgumentException e) {
								System.out.println("Image export error:" + imageUrl.toString());
							}

							if (size == 6) {
								if (rowCntTemp > rowCnt) rowCnt = rowCntTemp;
								row = sheet.createRow(rowCnt++);
								colCnt = 1;
								size = 1;
							} else {
								size++;
							}
						}
						if (rowCntTemp > rowCnt) rowCnt = rowCntTemp + 1;
						row = sheet.createRow(rowCnt++);
					}
				}

				rowCnt++;
			}

			rowCnt = fetchSteps(step.getJSONArray("steps"), workbook, creationHelper, sheet, rowCnt, validResults, platforms);
		}

		return rowCnt;
	}

	public static void main(String[] args) throws Exception{
		if (args.length != 2) {
			System.out.println("Usage: java -jar ResultExport.jar <config file> <target path>");
			return;
		}

		File configFile = new File(args[0]);
		if (!configFile.exists() || configFile.isDirectory()) {
			System.out.println("Config file is not exist.");
			return;
		}

		File targetFolder = new File(args[1]);
		if (!targetFolder.exists() && !targetFolder.mkdirs()) {
			System.out.println("Create target folder error.");
			return;
		}

		JSONObject config = new JSONObject(FileUtils.readFileToString(configFile, "UTF-8"));
		URIBuilder casesUrl = new URIBuilder(config.getString("serverUrl"));
		casesUrl.setPath("/swathub/api/" + config.getString("workspaceOwner") + "/" +
				config.getString("workspaceName") + "/sets/" + config.getString("setID") + "/scenarios");
		casesUrl.addParameter("tags", config.getString("tags"));

System.out.println("casesUrl: " + casesUrl.toString());

		String apiResult = apiGet(casesUrl, config.getString("username"), config.getString("apiKey"), null);
		if (apiResult == null || ("").equals(apiResult)) {
			System.out.println("Config file is not correct.");
			return;
		}
		JSONArray scenarios = new JSONArray(apiResult);

		for (int i = 0; i < scenarios.length(); i++) {
			JSONObject scenario = scenarios.getJSONObject(i);
			JSONArray testcases = scenario.getJSONArray("testcases");

			for (int j = 0; j < testcases.length(); j++) {
				JSONObject testcase = testcases.getJSONObject(j);
				System.out.println("Start creating xls file.");
				JSONArray results = testcase.getJSONArray("results");
				if (results.length() == 0) {
					System.out.println("No result for this test case, file will not be created.");
					System.out.println("");
					continue;
				}

				// get latest result base url for platforms
				String[] status = new String[config.getJSONArray("status").length()];
				for (int k = 0; k < config.getJSONArray("status").length(); k++) {
					status[k] = config.getJSONArray("status").getString(k);
				}

				String[] platforms = new String[config.getJSONArray("platforms").length()];
				for (int k = 0; k < config.getJSONArray("platforms").length(); k++) {
					platforms[k] = config.getJSONArray("platforms").getString(k);
				}

				String[] platformList = new String[config.getJSONArray("platforms").length()];
				for (int k = 0; k < config.getJSONArray("platforms").length(); k++) {
					platformList[k] = config.getJSONArray("platforms").getString(k);
				}

				HashMap<String, JSONObject> validResults = new HashMap<String, JSONObject>();
				JSONObject validResult = null;
				for (int k = 0; k < results.length(); k++) {
					JSONObject result = results.getJSONObject(k);
					if (ArrayUtils.contains(platformList, result.getString("execPlatform")) && ArrayUtils.contains(status, result.getString("status"))) {
						if (validResult == null) validResult = result;
						validResults.put(result.getString("execPlatform"), result);
						platformList = ArrayUtils.removeElement(platformList, result.getString("execPlatform"));
						if (platformList.length == 0) {
							break;
						}
					}
				}
				if (validResult == null) {
					System.out.println("No valid result for all platforms, file will not be created.");
					System.out.println("");
					continue;
				}
				if (platformList.length > 0) {
					for (String platform : platformList) {
						System.out.println("No valid result for this platform:" + platform);
					}
				}

				// get result object
				URIBuilder resultUrl = new URIBuilder(config.getString("serverUrl"));
				resultUrl.setPath("/swathub/api/" + config.getString("workspaceOwner") + "/" +
						config.getString("workspaceName") + "/results/" + validResult.getInt("id"));
				String strResult = apiGet(resultUrl, config.getString("username"), config.getString("apiKey"), null);
				if (strResult == null || ("").equals(strResult)) {
					System.out.println("Result not exists, file will not be created.");
					System.out.println("");
					continue;
				}
				JSONObject caseResult = new JSONObject(strResult);

				// create result sheet
				HSSFWorkbook workbook = new HSSFWorkbook();
				HSSFCreationHelper creationHelper = workbook.getCreationHelper();

				String safeName = WorkbookUtil.createSafeSheetName(testcase.getString("name"));
				HSSFSheet resultSheet = workbook.createSheet(safeName);

				// set basic information
				int rowCnt = 0;

				resultSheet.setColumnWidth(0, 4500);
				resultSheet.setColumnWidth(1, 4500);
				resultSheet.setColumnWidth(2, 4500);
				resultSheet.setColumnWidth(3, 4500);

				HSSFFont boldFont = workbook.createFont();
				boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
				HSSFCellStyle titleStyle = workbook.createCellStyle();
				titleStyle.setFont(boldFont);

				HSSFRow row = resultSheet.createRow(rowCnt);
				row.createCell(0).setCellValue("Scenario");
				row.getCell(0).setCellStyle(titleStyle);
				row.createCell(1).setCellValue(scenario.getString("name"));
				rowCnt = rowCnt + 2;

				// create result sheet
				fetchSteps(caseResult.getJSONArray("result"), workbook, creationHelper, resultSheet, rowCnt, validResults, platforms);

				ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
				workbook.write(outputStream);

				FileUtils.writeByteArrayToFile(new File(targetFolder, scenario.getString("name") +
						"_" + testcase.getString("name") + ".xls"), outputStream.toByteArray());
				outputStream.close();
				System.out.println(scenario.getString("name") + "_" + testcase.getString("name") + ".xls is created.");
				System.out.println("");
			}
		}
	}
}
