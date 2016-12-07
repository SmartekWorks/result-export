package com.swathub.dev;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
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
import org.apache.poi.hssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import javax.imageio.IIOException;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ResultExport {
	private static Map<String, String> valueMap = new HashMap<String, String>();
	static {
		valueMap.put("en.type.flow", "Flow");
		valueMap.put("en.type.sop", "System Operation");
		valueMap.put("en.type.pop", "Page Operation");
		valueMap.put("en.status.finished", "Done");
		valueMap.put("en.status.failed", "Error");
		valueMap.put("en.status.ok", "OK");
		valueMap.put("en.status.ng", "NG");
		valueMap.put("ja.type.flow", "ワークフロー");
		valueMap.put("ja.type.sop", "システムオペレーション");
		valueMap.put("ja.type.pop", "画面オペレーション");
		valueMap.put("ja.status.finished", "完了");
		valueMap.put("ja.status.failed", "エラー");
		valueMap.put("ja.status.ok", "OK");
		valueMap.put("ja.status.ng", "NG");
		valueMap.put("zh_cn.type.flow", "业务流程");
		valueMap.put("zh_cn.type.sop", "系统操作");
		valueMap.put("zh_cn.type.pop", "页面操作");
		valueMap.put("zh_cn.status.finished", "完成");
		valueMap.put("zh_cn.status.failed", "错误");
		valueMap.put("zh_cn.status.ok", "OK");
		valueMap.put("zh_cn.status.ng", "NG");
		valueMap.put("en.scenario", "Scenario");
		valueMap.put("en.result", "Test Result");
		valueMap.put("en.execNode", "Execution Node");
		valueMap.put("en.execPlatform", "Execution Platform");
		valueMap.put("en.testServer", "Test Server");
		valueMap.put("en.apiServer", "API Server");
		valueMap.put("en.result.execTime", "Execution Time");
		valueMap.put("en.result.error", "Error");
		valueMap.put("en.result.initBy", "Initiated By");
		valueMap.put("en.result.verifyBy", "Last Verified By");
		valueMap.put("en.case", "Test Case");
		valueMap.put("en.param", "Parameter");
		valueMap.put("en.variable", "Variable");
		valueMap.put("en.value", "Value");
		valueMap.put("en.name", "Name");
		valueMap.put("en.component.type", "Type");
		valueMap.put("en.comment", "Comment");
		valueMap.put("en.url", "URL");
		valueMap.put("ja.scenario", "シナリオ");
		valueMap.put("ja.result", "テスト結果");
		valueMap.put("ja.execNode", "実行ノード");
		valueMap.put("ja.execPlatform", "実行プラットフォーム");
		valueMap.put("ja.testServer", "テストサーバー");
		valueMap.put("ja.apiServer", "APIサーバー");
		valueMap.put("ja.result.execTime", "実行時間");
		valueMap.put("ja.result.error", "エラー");
		valueMap.put("ja.result.initBy", "実行者");
		valueMap.put("ja.result.verifyBy", "最終確認者");
		valueMap.put("ja.case", "テストケース");
		valueMap.put("ja.param", "パラメータ");
		valueMap.put("ja.variable", "変数");
		valueMap.put("ja.value", "値");
		valueMap.put("ja.name", "名前");
		valueMap.put("ja.component.type", "タイプ");
		valueMap.put("ja.comment", "コメント");
		valueMap.put("ja.url", "URL");
		valueMap.put("zh_cn.scenario", "测试流程");
		valueMap.put("zh_cn.result", "执行结果");
		valueMap.put("zh_cn.execNode", "执行节点");
		valueMap.put("zh_cn.execPlatform", "执行平台");
		valueMap.put("zh_cn.testServer", "测试服务器");
		valueMap.put("zh_cn.apiServer", "API服务器");
		valueMap.put("zh_cn.result.execTime", "执行时间");
		valueMap.put("zh_cn.result.error", "执行错误");
		valueMap.put("zh_cn.result.initBy", "执行发起者");
		valueMap.put("zh_cn.result.verifyBy", "最后验证者");
		valueMap.put("zh_cn.case", "测试用例");
		valueMap.put("zh_cn.param", "参数");
		valueMap.put("zh_cn.variable", "变量");
		valueMap.put("zh_cn.value", "值");
		valueMap.put("zh_cn.name", "名称");
		valueMap.put("zh_cn.component.type", "操作类型");
		valueMap.put("zh_cn.comment", "注释");
		valueMap.put("zh_cn.url", "URL");
	}
	private static String ROOT_PATH = "/api/";

	private static String apiGet(URIBuilder url, String user, String pass) throws Exception {
		CredentialsProvider credsProvider = new BasicCredentialsProvider();
		credsProvider.setCredentials(
				new AuthScope(url.getHost(), url.getPort()),
				new UsernamePasswordCredentials(user, pass));
		CloseableHttpClient httpclient = HttpClients.custom()
				.setDefaultCredentialsProvider(credsProvider)
				.build();

		String result;
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

	private static int fetchExcelSteps(JSONArray steps, HSSFWorkbook workbook, HSSFCreationHelper creationHelper,
								  HSSFSheet sheet, int rowCnt, JSONObject summary, String locale) throws Exception {
		HSSFRow row;
		HSSFFont boldFont = workbook.createFont();
		boldFont.setBold(true);
		HSSFFont errorFont = workbook.createFont();
		errorFont.setBold(true);
		errorFont.setColor(IndexedColors.RED.getIndex());

		HSSFCellStyle titleStyle = workbook.createCellStyle();
		titleStyle.setFont(boldFont);

		HSSFCellStyle errorStyle = workbook.createCellStyle();
		errorStyle.setFont(errorFont);

		HSSFCellStyle tableHeader = workbook.createCellStyle();
		tableHeader.setFont(boldFont);
		tableHeader.setAlignment(HorizontalAlignment.CENTER);
		tableHeader.setBorderBottom(BorderStyle.THIN);
		tableHeader.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		tableHeader.setBorderLeft(BorderStyle.THIN);
		tableHeader.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		tableHeader.setBorderRight(BorderStyle.THIN);
		tableHeader.setRightBorderColor(IndexedColors.BLACK.getIndex());
		tableHeader.setBorderTop(BorderStyle.THIN);
		tableHeader.setTopBorderColor(IndexedColors.BLACK.getIndex());

		HSSFCellStyle tableCell =  workbook.createCellStyle();
		tableCell.setBorderBottom(BorderStyle.THIN);
		tableCell.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		tableCell.setBorderLeft(BorderStyle.THIN);
		tableCell.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		tableCell.setBorderRight(BorderStyle.THIN);
		tableCell.setRightBorderColor(IndexedColors.BLACK.getIndex());
		tableCell.setBorderTop(BorderStyle.THIN);
		tableCell.setTopBorderColor(IndexedColors.BLACK.getIndex());

		HSSFCellStyle cellDisabled =  workbook.createCellStyle();
		cellDisabled.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		cellDisabled.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellDisabled.setBorderBottom(BorderStyle.THIN);
		cellDisabled.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		cellDisabled.setBorderLeft(BorderStyle.THIN);
		cellDisabled.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		cellDisabled.setBorderRight(BorderStyle.THIN);
		cellDisabled.setRightBorderColor(IndexedColors.BLACK.getIndex());
		cellDisabled.setBorderTop(BorderStyle.THIN);
		cellDisabled.setTopBorderColor(IndexedColors.BLACK.getIndex());

		HSSFCellStyle lineDisabled =  workbook.createCellStyle();
		lineDisabled.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		lineDisabled.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		for (int i = 0; i < steps.length(); i++) {
			JSONObject step = steps.getJSONObject(i);
			if (step.getString("seqNo").equals("")) {
				row = sheet.createRow(rowCnt++);
				row.createCell(0).setCellValue(valueMap.get(locale + ".case"));
				row.getCell(0).setCellStyle(titleStyle);
				row.createCell(1).setCellValue(step.getString("stepTitle"));

				row = sheet.createRow(rowCnt++);
				row.createCell(0).setCellValue(valueMap.get(locale + ".param"));
				row.getCell(0).setCellStyle(tableHeader);
				row.createCell(1).setCellValue(valueMap.get(locale + ".variable"));
				row.getCell(1).setCellStyle(tableHeader);
				row.createCell(2).setCellValue(valueMap.get(locale + ".value"));
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
					row.createCell(1).setCellValue(valueMap.get(locale + ".name"));
					row.getCell(1).setCellStyle(titleStyle);
					row.createCell(2).setCellValue(title);
				} else {
					row.createCell(0).setCellValue(step.getString("seqNo"));
					row.createCell(1).setCellValue(valueMap.get(locale + ".name"));
					row.getCell(0).setCellStyle(lineDisabled);
					row.createCell(2).setCellValue(title);
				}

				row = sheet.createRow(rowCnt++);
				row.createCell(1).setCellValue(valueMap.get(locale + ".component.type"));
				row.getCell(1).setCellStyle(titleStyle);
				row.createCell(2).setCellValue(valueMap.get(locale + ".type." + step.getString("type")));

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
					row.createCell(1).setCellValue(valueMap.get(locale + ".comment"));
					row.getCell(1).setCellStyle(titleStyle);
					row.createCell(2).setCellValue(comment.getString("value"));
				}

				if (!step.isNull("evidences") && step.getJSONObject("evidences").has("url")) {
					row = sheet.createRow(rowCnt++);
					row.createCell(1).setCellValue(valueMap.get(locale + ".url"));
					row.getCell(1).setCellStyle(titleStyle);
					row.createCell(2).setCellValue(step.getJSONObject("evidences").getString("url"));
				}

				if (!step.isNull("error")) {
					row = sheet.createRow(rowCnt++);
					row.createCell(1).setCellValue(valueMap.get(locale + ".result.error"));
					row.getCell(1).setCellStyle(titleStyle);
					row.createCell(2).setCellValue(step.getString("error"));
					row.getCell(2).setCellStyle(errorStyle);
				}

				row = sheet.createRow(rowCnt++);
				row.createCell(1).setCellValue(valueMap.get(locale + ".param"));
				row.getCell(1).setCellStyle(tableHeader);
				row.createCell(2).setCellValue(valueMap.get(locale + ".variable"));
				row.getCell(2).setCellStyle(tableHeader);
				row.createCell(3).setCellValue(valueMap.get(locale + ".value"));
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
					JSONArray screenshots = step.getJSONObject("evidences").getJSONArray("screenshots");
					for (int j = 0; j < screenshots.length(); j++) {
						String screenshot = screenshots.getString(j).replace("_s.png", ".png");

						URL imageUrl = new URL(summary.getString("baseURL").concat(screenshot));
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
							anchor.setCol1(1);
							anchor.setRow1(rowCnt);
							HSSFPicture picture = drawing.createPicture(anchor, pictureIdx);
							picture.resize();
							rowCnt = picture.getPreferredSize().getRow2();
						} catch (IllegalArgumentException e) {
							System.out.println("Image export error:" + imageUrl.toString());
						}
					}
				}

				rowCnt++;
			}

			rowCnt = fetchExcelSteps(step.getJSONArray("steps"), workbook, creationHelper, sheet, rowCnt, summary, locale);
		}

		return rowCnt;
	}

	private static void fetchHtmlSteps(JSONArray steps, ZipArchiveOutputStream archive, JSONObject summary) {
		for (int i = 0; i < steps.length(); i++) {
			JSONObject step = steps.getJSONObject(i);
			if (!step.isNull("evidences") && step.getJSONObject("evidences").has("html")) {
				try {
					URL htmlURL = new URL(summary.getString("baseURL") + step.getJSONObject("evidences").getString("html"));
					BufferedReader htmlInStream = new BufferedReader(new InputStreamReader(htmlURL.openStream()));

					archive.putArchiveEntry(new ZipArchiveEntry(step.getString("seqNo") + ".html"));
					IOUtils.copy(htmlInStream, archive);
					archive.closeArchiveEntry();

					htmlInStream.close();
				} catch (IOException e) {
					// continue
				}
			}
			fetchHtmlSteps(step.getJSONArray("steps"), archive, summary);
		}
	}

	private static void fetchDiagSteps(JSONArray steps, ZipArchiveOutputStream archive, JSONObject summary, JSONObject config) {
		for (int i = 0; i < steps.length(); i++) {
			JSONObject step = steps.getJSONObject(i);
			if (!step.isNull("evidences")) {
				if (step.getJSONObject("evidences").has("html")) {
					String htmlFile = step.getJSONObject("evidences").getString("html");
					if (!"".equals(htmlFile)) {
						try {
							URL htmlURL = new URL(summary.getString("baseURL") + htmlFile);
							BufferedReader htmlInStream = new BufferedReader(new InputStreamReader(htmlURL.openStream()));

							archive.putArchiveEntry(new ZipArchiveEntry("evidences/" + htmlFile));
							IOUtils.copy(htmlInStream, archive);
							archive.closeArchiveEntry();

							htmlInStream.close();
						} catch (IOException e) {
							// continue
						}
					}
				}
				if (step.getJSONObject("evidences").has("log")) {
					String logFile = step.getJSONObject("evidences").getString("log");
					if (!"".equals(logFile)) {
						try {
							URL logURL = new URL(summary.getString("baseURL") + step.getString("seqNo").replace("-", "/") + "/" + logFile);
							BufferedReader logInStream = new BufferedReader(new InputStreamReader(logURL.openStream()));

							archive.putArchiveEntry(new ZipArchiveEntry("evidences/" + step.getString("seqNo").replace("-", "/") + "/" + logFile));
							IOUtils.copy(logInStream, archive);
							archive.closeArchiveEntry();

							logInStream.close();
						} catch (IOException e) {
							// continue
						}
					}
				}
				if (step.getJSONObject("evidences").has("console")) {
					String consoleFile = step.getJSONObject("evidences").getString("console");
					if (!"".equals(consoleFile)) {
						try {
							URL consoleURL = new URL(summary.getString("baseURL") + step.getString("seqNo").replace("-", "/") + "/" + consoleFile);
							BufferedReader consoleInStream = new BufferedReader(new InputStreamReader(consoleURL.openStream()));

							archive.putArchiveEntry(new ZipArchiveEntry("evidences/" + step.getString("seqNo").replace("-", "/") + "/" + consoleFile));
							IOUtils.copy(consoleInStream, archive);
							archive.closeArchiveEntry();

							consoleInStream.close();
						} catch (IOException e) {
							// continue
						}
					}
				}
				if (step.getJSONObject("evidences").has("console")) {
					JSONArray screenshots = step.getJSONObject("evidences").getJSONArray("screenshots");
					for (int j = 0; j < screenshots.length(); j++) {
						String screenshot = screenshots.getString(j).replace("_s.png", ".png");
						try {
							URL imageUrl = new URL(summary.getString("baseURL").concat(screenshot));
							BufferedImage image;
							try {
								image = ImageIO.read(imageUrl);
							} catch (IIOException e) {
								System.out.println("Image URL may not exist:" + imageUrl.toString());
								continue;
							}
							ByteArrayOutputStream baos = new ByteArrayOutputStream();
							ImageIO.write(image, "png", baos);

							archive.putArchiveEntry(new ZipArchiveEntry("evidences/" + screenshot));
							baos.writeTo(archive);
							archive.closeArchiveEntry();
						} catch (IOException e) {
							// continue
						}
					}
				}
			}
			if (!step.isNull("error") && "pop".equals(step.getString("type"))) {
				try {
					URIBuilder apiURL = new URIBuilder(config.getString("serverUrl"));
					apiURL.setPath(ROOT_PATH + config.getString("workspaceOwner") + "/" +
							config.getString("workspaceName") + "/pages/" + step.getString("pageCode"));
					String apiResult = apiGet(apiURL, config.getString("username"), config.getString("apiKey"));
					JSONObject page = new JSONObject(apiResult);

					archive.putArchiveEntry(new ZipArchiveEntry(page.getString("name") + ".shtm"));
					ByteArrayInputStream shtmBIS = new ByteArrayInputStream(Base64.decodeBase64(page.getString("data").getBytes()));
					IOUtils.copy(shtmBIS, archive);
					archive.closeArchiveEntry();
					shtmBIS.close();

					archive.putArchiveEntry(new ZipArchiveEntry(page.getString("name") + ".rule"));
					ByteArrayInputStream ruleBIS = new ByteArrayInputStream(page.getString("rule").getBytes());
					IOUtils.copy(ruleBIS, archive);
					archive.closeArchiveEntry();
					ruleBIS.close();
				} catch (Exception e) {
					// continue
				}
			}
			fetchDiagSteps(step.getJSONArray("steps"), archive, summary, config);
		}
	}

	public static void main(String[] args) throws Exception {

		if (args.length != 3) {
			System.out.println("Usage: java -jar ResultExport.jar <config file> <format> <target path>");
			return;
		}

		File configFile = new File(args[0]);
		if (!configFile.exists() || configFile.isDirectory()) {
			System.out.println("Config file is not exist.");
			return;
		}

		String format = args[1];
		if (!"excel".equals(format) && !"html".equals(format) && !"diag".equals(format)) {
			System.out.println("Format must be one of the following: excel, html, diag");
			return;
		}

		JSONObject config = new JSONObject(FileUtils.readFileToString(configFile, "UTF-8"));
		String locale = config.getString("locale");

		URIBuilder apiURL = new URIBuilder(config.getString("serverUrl"));
		apiURL.setPath(ROOT_PATH + config.getString("workspaceOwner") + "/" +
				config.getString("workspaceName") + "/results/" + args[2]);
		apiURL.addParameter("lang", locale);

		String apiResult = apiGet(apiURL, config.getString("username"), config.getString("apiKey"));
		if (apiResult == null || ("").equals(apiResult)) {
			System.out.println("Config file is not correct.");
			return;
		}

		JSONObject result = new JSONObject(apiResult);
		JSONObject summary = result.getJSONObject("summary");
		if ("excel".equals(format)) {
			// create result sheet
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFCreationHelper creationHelper = workbook.getCreationHelper();

			HSSFSheet resultSheet = workbook.createSheet("Result");

			// set basic information
			int rowCnt = 1;

			resultSheet.setColumnWidth(0, 4500);
			resultSheet.setColumnWidth(1, 4500);
			resultSheet.setColumnWidth(2, 4500);
			resultSheet.setColumnWidth(3, 4500);

			HSSFFont boldFont = workbook.createFont();
			boldFont.setBold(true);
			HSSFFont errorFont = workbook.createFont();
			errorFont.setBold(true);
			errorFont.setColor(IndexedColors.RED.getIndex());

			HSSFCellStyle titleStyle = workbook.createCellStyle();
			titleStyle.setFont(boldFont);

			HSSFCellStyle errorStyle = workbook.createCellStyle();
			errorStyle.setFont(errorFont);

			HSSFRow row = resultSheet.createRow(rowCnt);
			row.createCell(0).setCellValue(valueMap.get(locale + ".scenario"));
			row.getCell(0).setCellStyle(titleStyle);
			row.createCell(1).setCellValue(summary.getString("scenarioName"));

			row = resultSheet.createRow(rowCnt++);
			row.createCell(0).setCellValue(valueMap.get(locale + ".result"));
			row.getCell(0).setCellStyle(titleStyle);
			row.createCell(1).setCellValue(valueMap.get(locale + ".status." + summary.getString("status")));

			row = resultSheet.createRow(rowCnt++);
			row.createCell(0).setCellValue(valueMap.get(locale + ".execNode"));
			row.getCell(0).setCellStyle(titleStyle);
			row.createCell(1).setCellValue(summary.getString("execNode"));

			row = resultSheet.createRow(rowCnt++);
			row.createCell(0).setCellValue(valueMap.get(locale + ".execPlatform"));
			row.getCell(0).setCellStyle(titleStyle);
			row.createCell(1).setCellValue(summary.getString("execPlatform"));

			row = resultSheet.createRow(rowCnt++);
			row.createCell(0).setCellValue(valueMap.get(locale + ".testServer"));
			row.getCell(0).setCellStyle(titleStyle);
			row.createCell(1).setCellValue(summary.getString("testServer"));

			row = resultSheet.createRow(rowCnt++);
			row.createCell(0).setCellValue(valueMap.get(locale + ".apiServer"));
			row.getCell(0).setCellStyle(titleStyle);
			row.createCell(1).setCellValue(summary.getString("apiServer"));

			row = resultSheet.createRow(rowCnt++);
			row.createCell(0).setCellValue(valueMap.get(locale + ".result.execTime"));
			row.getCell(0).setCellStyle(titleStyle);
			row.createCell(1).setCellValue(summary.getDouble("duration") + "s (" +
					summary.getString("timeStart") + "~" + summary.getString("timeEnd") + ")");

			row = resultSheet.createRow(rowCnt++);
			if ("NA".equals(summary.getString("error"))) {
				row.createCell(0).setCellValue(valueMap.get(locale + ".result.error"));
				row.getCell(0).setCellStyle(titleStyle);
				row.createCell(1).setCellValue(summary.getString("error"));
			} else {
				row.createCell(0).setCellValue(valueMap.get(locale + ".result.error"));
				row.getCell(0).setCellStyle(titleStyle);
				row.createCell(1).setCellValue(summary.getString("error"));
				row.getCell(1).setCellStyle(errorStyle);
			}

			row = resultSheet.createRow(rowCnt++);
			row.createCell(0).setCellValue(valueMap.get(locale + ".result.initBy"));
			row.getCell(0).setCellStyle(titleStyle);
			row.createCell(1).setCellValue(summary.getString("initBy") + "(" + summary.getString("initTime") + ")");

			row = resultSheet.createRow(rowCnt++);
			if ("NA".equals(summary.getString("verifyBy"))) {
				row.createCell(0).setCellValue(valueMap.get(locale + ".result.verifyBy"));
				row.getCell(0).setCellStyle(titleStyle);
				row.createCell(1).setCellValue(summary.getString("verifyBy"));
			} else {
				row.createCell(0).setCellValue(valueMap.get(locale + ".result.verifyBy"));
				row.getCell(0).setCellStyle(titleStyle);
				row.createCell(1).setCellValue(summary.getString("verifyBy") + "(" + summary.getString("verifyTime") + ")");
			}

			rowCnt ++;

			fetchExcelSteps(result.getJSONArray("result"), workbook, creationHelper, resultSheet, rowCnt, summary, locale);

			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			workbook.write(outputStream);

			FileUtils.writeByteArrayToFile(new File(summary.getString("scenarioName") +
					"-" + summary.getString("caseName") + "-" + args[2] + ".xls"), outputStream.toByteArray());
			outputStream.close();
			System.out.println("Excel file is created.");
		} else if ("html".equals(format)) {
			ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
			ZipArchiveOutputStream archive = new ZipArchiveOutputStream(byteOut);

			fetchHtmlSteps(result.getJSONArray("result"), archive, summary);

			archive.finish();
			archive.flush();
			archive.close();

			byteOut.flush();
			FileOutputStream fop = new FileOutputStream(summary.getString("scenarioName") +
					"-" + summary.getString("caseName") + "-" + args[2] + ".zip");
			byteOut.writeTo(fop);
			byteOut.close();
			fop.close();
			System.out.println("Html zip file is created.");
		} else if ("diag".equals(format)) {
			ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
			ZipArchiveOutputStream archive = new ZipArchiveOutputStream(byteOut);

			fetchDiagSteps(result.getJSONArray("result"), archive, summary, config);

			archive.putArchiveEntry(new ZipArchiveEntry("result.json"));
			ByteArrayInputStream resultBIS = new ByteArrayInputStream(apiResult.getBytes());
			IOUtils.copy(resultBIS, archive);
			archive.closeArchiveEntry();
			resultBIS.close();

			archive.finish();
			archive.flush();
			archive.close();
			byteOut.flush();

			FileOutputStream fop = new FileOutputStream("diag-" + summary.getString("scenarioName") +
					"-" + summary.getString("caseName") + "-" + args[2] + ".zip");
			byteOut.writeTo(fop);
			byteOut.close();
			fop.close();
			System.out.println("Diagnosis zip file is created.");
		} else {
			System.out.println("Format must be one of the following: excel, html, diag");
		}
	}

}
