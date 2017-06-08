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
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

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

	private static String lastPageCode = null;

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

	private static String fetchHtmlSteps(JSONArray steps, ZipArchiveOutputStream archive, String html, JSONObject summary, String locale) throws Exception {
		for (int i = 0; i < steps.length(); i++) {
			JSONObject step = steps.getJSONObject(i);
			if (step.getString("seqNo").equals("")) {
				html += "<div class=\"pure-g\"><div class=\"pure-u-1\">&nbsp;</div>";

				html += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".case") + "：</div></div>";
				html += "<div class=\"pure-u-20-24\">" + step.getString("stepTitle") + "</div>";

				html += "<div class=\"pure-u-1 title first\"><table class=\"pure-table\"><thead><tr>";
				html += "<th>" + valueMap.get(locale + ".param") + "</th>";
				html += "<th>" + valueMap.get(locale + ".variable") + "</th>";
				html += "<th>" + valueMap.get(locale + ".value") + "</th>";
				html += "</tr></thead>";

				html += "<tbody>";
				for (int j = 0; j < step.getJSONArray("paramData").length(); j++) {
					JSONObject item = step.getJSONArray("paramData").getJSONObject(j);
					html += "<tr>";
					if (item.getBoolean("runtimeEnabled")) {
						html += "<td>" + item.getString("name") + "</td>";
						html += "<td>" + item.getString("variable") + "</td>";
						html += "<td>" + item.getString("value") + "</td>";
					} else {
						html += "<td style=\"background-color: grey;\">" + item.getString("name") + "</td>";
						html += "<td style=\"background-color: grey;\">" + item.getString("variable") + "</td>";
						html += "<td style=\"background-color: grey;\">" + item.getString("value") + "</td>";
					}
					html += "</tr>";
				}
				html += "</tbody></table></div></div>";
			} else {
				String title = step.getString("stepTitle");
				if (step.has("typeName") && !step.isNull("typeName")) {
					title = title + "(" + step.getString("typeName") + ")";
				}
				html += "<div class=\"pure-g\"><div class=\"pure-u-1\">&nbsp;</div>";
				if (step.getBoolean("executed")) {
					html += "<div class=\"pure-u-2-24\"><div class=\"first\">" + step.getString("seqNo") + "</div></div>";
				} else {
					html += "<div class=\"pure-u-2-24\" style=\"background-color: grey;\">><div class=\"first\">" + step.getString("seqNo") + "</div></div>";
				}
				html += "<div class=\"pure-u-3-24\"><div class=\"title\">" + valueMap.get(locale + ".name") + "：</div></div>";
				html += "<div class=\"pure-u-19-24\">" + title + "</div>";

				html += "<div class=\"pure-u-2-24\">&nbsp;</div>";
				html += "<div class=\"pure-u-3-24\"><div class=\"title\">" + valueMap.get(locale + ".component.type") + "：</div></div>";
				html += "<div class=\"pure-u-19-24\">" + valueMap.get(locale + ".type." + step.getString("type")) + "</div>";

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
					html += "<div class=\"pure-u-2-24\">&nbsp;</div>";
					html += "<div class=\"pure-u-3-24\"><div class=\"title\">" + valueMap.get(locale + ".comment") + "：</div></div>";
					html += "<div class=\"pure-u-19-24\">" + comment.getString("value") + "</div>";
				}

				if (!step.isNull("evidences") && step.getJSONObject("evidences").has("url")) {
					html += "<div class=\"pure-u-2-24\">&nbsp;</div>";
					html += "<div class=\"pure-u-3-24\"><div class=\"title\">" + valueMap.get(locale + ".url") + "：</div></div>";
					html += "<div class=\"pure-u-19-24\">" + step.getJSONObject("evidences").getString("url") + "</div>";
				}

				if (!step.isNull("error")) {
					html += "<div class=\"pure-u-2-24\">&nbsp;</div>";
					html += "<div class=\"pure-u-3-24\"><div class=\"title\">" + valueMap.get(locale + ".result.error") + "：</div></div>";
					html += "<div class=\"pure-u-19-24\" style=\"color: red;\">" + step.getString("error") + "</div>";
				}

				html += "<div class=\"pure-u-2-24\">&nbsp;</div>";
				html += "<div class=\"pure-u-22-24\"><table class=\"pure-table\"><thead>";
				html += "<th>" + valueMap.get(locale + ".param") + "</th>";
				html += "<th>" + valueMap.get(locale + ".variable") + "</th>";
				html += "<th>" + valueMap.get(locale + ".value") + "</th>";
				html += "</tr></thead>";

				for (JSONObject item : paramData) {
					html += "<tr>";
					if (!item.isNull("runtimeEnabled") && item.getBoolean("runtimeEnabled")) {
						html += "<td>" + item.getString("name") + "</td>";
						html += "<td>" + item.getString("variable") + "</td>";
						html += "<td>" + item.getString("value") + "</td>";
					} else {
						html += "<td style=\"background-color: grey;\">" + item.getString("name") + "</td>";
						html += "<td style=\"background-color: grey;\">" + item.getString("variable") + "</td>";
						html += "<td style=\"background-color: grey;\">" + item.getString("value") + "</td>";
					}
					html += "</tr>";
				}
				html += "</tbody></table></div>";

				if (!step.isNull("evidences") && step.getJSONObject("evidences").has("screenshots")) {
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

							html += "<div class=\"pure-u-2-24\">&nbsp;</div>";
							html += "<div class=\"pure-u-20-24\"><img class=\"pure-img\" src=\"./evidences/" + screenshot + "\"></div>";
							html += "<div class=\"pure-u-2-24\">&nbsp;</div>";
						} catch (IOException e) {
							// continue
						}
					}
				}

				html += "</div>";
			}

			html = fetchHtmlSteps(step.getJSONArray("steps"), archive, html, summary, locale);
		}

		return html;
	}

	private static void fetchSourceSteps(JSONArray steps, ZipArchiveOutputStream archive, JSONObject summary) {
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
			fetchSourceSteps(step.getJSONArray("steps"), archive, summary);
		}
	}

	private static void fetchDiagSteps(JSONArray steps, ZipArchiveOutputStream archive, JSONObject summary, JSONObject config) {
		for (int i = 0; i < steps.length(); i++) {
			JSONObject step = steps.getJSONObject(i);
			if (step.getBoolean("executed") && "pop".equals(step.getString("type"))) {
				if (step.isNull("pageCode")) {
					lastPageCode = null;
				} else {
					lastPageCode = step.getString("pageCode");
				}
			}
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
				if (step.getJSONObject("evidences").has("screenshots")) {
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
			if (!step.isNull("error") && lastPageCode != null) {
				try {
					URIBuilder apiURL = new URIBuilder(config.getString("serverUrl"));
					apiURL.setPath(ROOT_PATH + config.getString("workspaceOwner") + "/" +
							config.getString("workspaceName") + "/pages/" + lastPageCode);
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

	private static void fetchRawSteps(JSONArray steps, ZipArchiveOutputStream archive, JSONObject summary) {
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
				if (step.getJSONObject("evidences").has("screenshots")) {
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
				if (step.getJSONObject("evidences").has("files")) {
					JSONArray files = step.getJSONObject("evidences").getJSONArray("files");
					for (int j = 0; j < files.length(); j++) {
						String file = files.getString(j);
						try {
							URL fileUrl = new URL(summary.getString("baseURL").concat(file));
							BufferedReader fileInStream = new BufferedReader(new InputStreamReader(fileUrl.openStream()));

							archive.putArchiveEntry(new ZipArchiveEntry("evidences/" + file));
							IOUtils.copy(fileInStream, archive);
							archive.closeArchiveEntry();

							fileInStream.close();
						} catch (IOException e) {
							// continue
						}
					}
				}
			}
			fetchRawSteps(step.getJSONArray("steps"), archive, summary);
		}
	}

	private static void createFile(String format, JSONObject config, String resultId) throws Exception{
		String locale = config.getString("locale");

		URIBuilder apiURL = new URIBuilder(config.getString("serverUrl"));
		apiURL.setPath(ROOT_PATH + config.getString("workspaceOwner") + "/" +
				config.getString("workspaceName") + "/results/" + resultId);
		apiURL.addParameter("lang", locale);

		String apiResult = apiGet(apiURL, config.getString("username"), config.getString("apiKey"));
		if (apiResult == null || ("").equals(apiResult)) {
			System.out.println("Config file is not correct.");
			return;
		}

		JSONObject result = new JSONObject(apiResult);
		JSONObject summary = result.getJSONObject("summary");
		if ("raw".equals(format)) {
			ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
			ZipArchiveOutputStream archive = new ZipArchiveOutputStream(byteOut);

			archive.putArchiveEntry(new ZipArchiveEntry("result.json"));
			ByteArrayInputStream bytesInStream = new ByteArrayInputStream(apiResult.getBytes());
			IOUtils.copy(bytesInStream, archive);
			bytesInStream.close();
			archive.closeArchiveEntry();

			fetchRawSteps(result.getJSONArray("result"), archive, summary);

			archive.finish();
			archive.flush();
			archive.close();

			byteOut.flush();
			FileOutputStream fop = new FileOutputStream("raw-" + summary.getString("scenarioName") +
					"-" + summary.getString("caseName") + "-" + resultId + ".zip");
			byteOut.writeTo(fop);
			byteOut.close();
			fop.close();
			System.out.println("Raw file is created. Result ID:" + resultId);
		}
		else if ("excel".equals(format)) {
			// create result sheet
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFCreationHelper creationHelper = workbook.getCreationHelper();

			HSSFSheet resultSheet = workbook.createSheet("Result");

			// set basic information
			int rowCnt = 0;

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

			HSSFRow row = resultSheet.createRow(rowCnt++);
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
					"-" + summary.getString("caseName") + "-" + resultId + ".xls"), outputStream.toByteArray());
			outputStream.close();
			System.out.println("Excel file is created. Result ID:" + resultId);
		} else if ("html".equals(format)) {
			ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
			ZipArchiveOutputStream archive = new ZipArchiveOutputStream(byteOut);

			String execInfo = "<div class=\"pure-g\"><div class=\"pure-u-1\">&nbsp;</div>";
			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".scenario") + "：</div></div>";
			execInfo += "<div class=\"pure-u-20-24\">" + summary.getString("scenarioName") + "</div>";

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".result") + "：</div></div>";
			execInfo += "<div class=\"pure-u-20-24\">" + valueMap.get(locale + ".status." + summary.getString("status")) + "</div>";

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".execNode") + "：</div></div>";
			execInfo += "<div class=\"pure-u-20-24\">" + summary.getString("execNode") + "</div>";

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".execPlatform") + "：</div></div>";
			execInfo += "<div class=\"pure-u-20-24\">" + summary.getString("execPlatform") + "</div>";

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".testServer") + "：</div></div>";
			execInfo += "<div class=\"pure-u-20-24\">" + summary.getString("testServer") + "</div>";

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".apiServer") + "：</div></div>";
			execInfo += "<div class=\"pure-u-20-24\">" + summary.getString("apiServer") + "</div>";

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".result.execTime") + "：</div></div>";
			execInfo += "<div class=\"pure-u-20-24\">" + summary.getDouble("duration") + "s (" +
					summary.getString("timeStart") + "~" + summary.getString("timeEnd") + ")</div>";

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".result.error") + "：</div></div>";
			if ("NA".equals(summary.getString("error"))) {
				execInfo += "<div class=\"pure-u-20-24\">NA</div>";
			} else {
				execInfo += "<div class=\"pure-u-20-24\" style=\"color: red;\">" + summary.getString("error") + "</div>";
			}

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".result.initBy") + "：</div></div>";
			execInfo += "<div class=\"pure-u-20-24\">" + summary.getString("initBy") + "(" + summary.getString("initTime") + ")</div>";

			execInfo += "<div class=\"pure-u-4-24\"><div class=\"title first\">" + valueMap.get(locale + ".result.verifyBy") + "：</div></div>";
			if ("NA".equals(summary.getString("verifyBy"))) {
				execInfo += "<div class=\"pure-u-20-24\">NA</div>";
			} else {
				execInfo += "<div class=\"pure-u-20-24\">" + summary.getString("verifyBy") + "(" + summary.getString("verifyTime") + ")</div>";
			}

			String stepInfo = fetchHtmlSteps(result.getJSONArray("result"), archive, "", summary, locale);

			ClassLoader classloader = Thread.currentThread().getContextClassLoader();
			String template = IOUtils.toString(classloader.getResourceAsStream("template.html"));
			template = template.replace("_exec-info_", execInfo);
			template = template.replace("_step-info_", stepInfo);

			archive.putArchiveEntry(new ZipArchiveEntry("index.html"));
			ByteArrayInputStream bytesInStream = new ByteArrayInputStream(template.getBytes());
			IOUtils.copy(bytesInStream, archive);
			bytesInStream.close();
			archive.closeArchiveEntry();

			archive.putArchiveEntry(new ZipArchiveEntry("css/pure-min.css"));
			InputStream inStream = classloader.getResourceAsStream("pure-min.css");
			IOUtils.copy(inStream, archive);
			inStream.close();
			archive.closeArchiveEntry();

			archive.flush();
			archive.close();

			byteOut.flush();
			FileOutputStream fop = new FileOutputStream("html-" + summary.getString("scenarioName") +
					"-" + summary.getString("caseName") + "-" + resultId + ".zip");
			byteOut.writeTo(fop);
			byteOut.close();
			fop.close();
			System.out.println("Html zip file is created. Result ID:" + resultId);
		} else if ("source".equals(format)) {
			ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
			ZipArchiveOutputStream archive = new ZipArchiveOutputStream(byteOut);

			fetchSourceSteps(result.getJSONArray("result"), archive, summary);

			archive.finish();
			archive.flush();
			archive.close();

			byteOut.flush();
			FileOutputStream fop = new FileOutputStream("source-" + summary.getString("scenarioName") +
					"-" + summary.getString("caseName") + "-" + resultId + ".zip");
			byteOut.writeTo(fop);
			byteOut.close();
			fop.close();
			System.out.println("Page sources file is created. Result ID:" + resultId);
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
					"-" + summary.getString("caseName") + "-" + resultId + ".zip");
			byteOut.writeTo(fop);
			byteOut.close();
			fop.close();
			System.out.println("Diagnosis zip file is created. Result ID:" + resultId);
		}
	}

	public static void main(String[] args) throws Exception {

		if (args.length != 3) {
			System.out.println("Usage: java -jar ResultExport.jar <config file> <format> <target file>");
			return;
		}

		File configFile = new File(args[0]);
		if (!configFile.exists() || configFile.isDirectory()) {
			System.out.println("Config file is not exist.");
			return;
		}

		String format = args[1];
		if (!"excel".equals(format) && !"html".equals(format) && !"diag".equals(format) && !"raw".equals(format)) {
			System.out.println("Format must be one of the following: excel, html, diag, raw");
			return;
		}

		File targetFile = new File(args[2]);
		if (!targetFile.exists() || targetFile.isDirectory()) {
			System.out.println("Target file is not exist.");
			return;
		}

		JSONObject config = new JSONObject(FileUtils.readFileToString(configFile, "UTF-8"));
		JSONObject target = new JSONObject(FileUtils.readFileToString(targetFile, "UTF-8"));
		JSONObject filters = target.getJSONObject("filters");
		ArrayList<String> validResults = new ArrayList<String>();

		if (target.has("ids") && target.getJSONArray("ids").length() > 0) {
			JSONArray ids = target.getJSONArray("ids");
			for (int i = 0; i < ids.length(); i++) {
				validResults.add(ids.get(i).toString());
			}
		} else {
			URIBuilder casesUrl = new URIBuilder(config.getString("serverUrl"));
			casesUrl.setPath(ROOT_PATH + config.getString("workspaceOwner") + "/" +
					config.getString("workspaceName") + "/sets/" + filters.getString("setID") + "/scenarios");
			casesUrl.addParameter("tags", filters.getString("tags"));

			String apiResult = apiGet(casesUrl, config.getString("username"), config.getString("apiKey"));
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
					JSONArray results = testcase.getJSONArray("results");

					String targetStatus = filters.getString("status");

					String targetPlatform = filters.getString("platform");

					Date beforeDate = new Date();
					String beforeDateString = filters.getString("beforeDate");
					if (!("").equals(beforeDateString)) {
						DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
						dateFormat.setTimeZone(TimeZone.getTimeZone("Asia/Tokyo"));
						beforeDate = dateFormat.parse(beforeDateString);
					}

					for (int k = 0; k < results.length(); k++) {
						JSONObject result = results.getJSONObject(k);

						String status = result.getString("status");

						String platform = result.getString("execPlatform");

						String timeStartString = result.getString("timeStart");
						DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
						dateFormat.setTimeZone(TimeZone.getTimeZone("Asia/Tokyo"));
						Date timeStart = dateFormat.parse(timeStartString);

						if (timeStart.getTime() - beforeDate.getTime() > 0) {
							continue;
						}

						if (("".equals(targetPlatform) || targetPlatform.equals(platform)) && ("".equals(targetStatus) || targetStatus.equals(status))) {
							validResults.add(String.valueOf(result.getInt("id")));
							break;
						}
					}
				}
			}
		}

		for (String result : validResults) {
			createFile(format, config, result);
		}
	}
}
