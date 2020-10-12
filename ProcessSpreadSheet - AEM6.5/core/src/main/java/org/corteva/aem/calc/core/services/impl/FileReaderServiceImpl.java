package org.corteva.aem.calc.core.services.impl;

import java.io.File;
import java.io.IOException;
import java.util.Objects;

import javax.jcr.Node;
import javax.jcr.RepositoryException;
import javax.jcr.Session;

import org.apache.jackrabbit.commons.JcrUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.osgi.service.component.annotations.Component;
import org.corteva.aem.calc.core.services.FileReaderService;
import org.corteva.aem.calc.core.constants.AppConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Component(service = FileReaderService.class, immediate = true)
public class FileReaderServiceImpl implements FileReaderService {

	// Default logger
	private final Logger log = LoggerFactory.getLogger(this.getClass());

	/**
	 * Overridden method to read the uploaded excel file
	 */
	@Override
	public void readExcel(Session session, String filePath) {

		try {

			// Creating a workbook from the excel file
			XSSFWorkbook workbook = new XSSFWorkbook(new File(filePath));

			//workbook.setMissingCellPolicy(MissingCellPolicy.RETURN_BLANK_AS_NULL);

			// Getting the number of sheets
			int numberOfSheets = workbook.getNumberOfSheets();

			for (int i = 0; i < numberOfSheets; i++) {

				// Getting the sheet at the ith position
				XSSFSheet sheet = workbook.getSheetAt(i);

				log.info("Reading sheet: {}", sheet.getSheetName());

				if (sheet.getSheetName().equals(AppConstants.SHEET_SELECT_SAVINGS)) {
					
					log.info("Inside Select Savings Sheet");

					int tradeCode = 0, productColumnIndex = 1, progromCodeColumnIndex = 2, soldAsUOM = 3,
							selectSavingsEarningRate = 4, requriesPrepay = 5;

					for (Row myrow : sheet) {

						// Skip 1st Row
						if (myrow.getRowNum() != 0) {

							for (Cell cell : myrow) {

								String nodePath = AppConstants.INITIAL_STRUCTURE + "select-savings/";
								String programCode = "", product = "";

								cell = myrow.getCell(progromCodeColumnIndex);

								if (Objects.nonNull(cell)) {
									programCode = readCellData(cell);
									log.info("programCode: " + programCode);
									createNode(nodePath, session, programCode);
								}

								cell = myrow.getCell(productColumnIndex);

								if (Objects.nonNull(cell)) {
									product = readCellData(cell);
									log.info("product: " + product);
									createNode(nodePath + programCode, session, product);
								}

								cell = myrow.getCell(requriesPrepay);

								if (Objects.nonNull(cell)) {
									String prePay = readCellData(cell);
									createProperties(nodePath + programCode + "/" + product, session,
											AppConstants.SELECT_SAVINGS_REQUIRES_PREPAY, prePay);
								}

								cell = myrow.getCell(selectSavingsEarningRate);

								if (Objects.nonNull(cell)) {
									String earningsRate = readCellData(cell);
									createProperties(nodePath + programCode + "/" + product, session,
											AppConstants.SELECT_SAVINGS_EARNINGS_RATE, earningsRate);
								}

								cell = myrow.getCell(tradeCode);

								if (Objects.nonNull(cell)) {
									String tradeCodeValue = readCellData(cell);
									createProperties(nodePath + programCode + "/" + product, session,
											AppConstants.SELECT_SAVINGS_TRADE_CODE, tradeCodeValue);
								}

								cell = myrow.getCell(soldAsUOM);

								if (Objects.nonNull(cell)) {
									String soldAsUOMValue = readCellData(cell);
									createProperties(nodePath + programCode + "/" + product, session,
											AppConstants.SELECT_SAVINGS_SOLD_AS_UOM, soldAsUOMValue);
								}

							}

						}
					}
				}

				if (sheet.getSheetName().equals(AppConstants.SHEET_ENLIST_AHEAD)) {
					
					log.info("Inside Enlist Ahead Sheet");

					int product = 0, programCode = 1, programMatchRate = 2, applicationUOM = 3, soldAsUOM = 4,
							coversionFactor = 5, payRate = 6;

					for (Row myrow : sheet) {

						// Skip 1st Row
						if (myrow.getRowNum() != 0) {

							for (Cell cell : myrow) {

								String nodePath = AppConstants.INITIAL_STRUCTURE + "enlist-ahead/";
								String programCodeValue = "", productValue = "";
								
								cell = myrow.getCell(programCode);
								if (Objects.nonNull(cell)) {
									programCodeValue = readCellData(cell);
									log.info("programCode: " + programCodeValue);
									createNode(nodePath, session, programCodeValue);
								}
								
								cell = myrow.getCell(product);
								
								if (Objects.nonNull(cell)) {
									productValue = readCellData(cell);
									log.info("product: " + productValue);
									createNode(nodePath + programCodeValue, session, productValue);
								}
								
								cell = myrow.getCell(programMatchRate);
								
								if (Objects.nonNull(cell)) {
									String programMatchRateValue = readCellData(cell);
									createProperties(nodePath + programCodeValue + "/" + productValue, session,
											AppConstants.ENLIST_AHEAD_PROGRAM_MATCH_RATE, programMatchRateValue);
								}
								
								cell = myrow.getCell(applicationUOM);
								
								if (Objects.nonNull(cell)) {
									String applicationUOMValue = readCellData(cell);
									createProperties(nodePath + programCodeValue + "/" + productValue, session,
											AppConstants.ENLIST_AHEAD_APPLICATION_UOM, applicationUOMValue);
								}

								cell = myrow.getCell(soldAsUOM);

								if (Objects.nonNull(cell)) {
									String soldAsUOMValue = readCellData(cell);
									createProperties(nodePath + programCodeValue + "/" + productValue, session,
											AppConstants.ENLIST_AHEAD_SOLD_AS_UOM, soldAsUOMValue);
								}

								cell = myrow.getCell(coversionFactor);

								if (Objects.nonNull(cell)) {
									String coversionFactorValue = readCellData(cell);
									createProperties(nodePath + programCodeValue + "/" + productValue, session,
											AppConstants.ENLIST_AHEAD_CONVERSION_FACTOR, coversionFactorValue);
								}

								cell = myrow.getCell(payRate);

								if (Objects.nonNull(cell)) {
									String payRateValue = readCellData(cell);
									createProperties(nodePath + programCodeValue + "/" + productValue, session,
											AppConstants.ENLIST_AHEAD_PAY_RATE, payRateValue);
								}

							}

						}
					}

				}
			}
			workbook.close();

		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			log.error("Error Occured in ReadExcel: " + e.getMessage());
		} catch (Exception f) {
			log.error("Error in ReadExcel: " + f.getMessage());
		}

	}

	private void createProperties(String pathName, Session session, String name, String value) {
		try {
			Node baseNode = session.getNode(pathName);
			if (baseNode != null) {
				baseNode.setProperty(name, value);
				log.info("Properties created at: " + baseNode + name + value);
			}
		} catch (RepositoryException e) {
			log.error("RepositoryException in createProperties: " + e.getMessage());
		}

	}

	private void createNode(String pathName, Session session, String readCellData) {
		try {
			Node baseNode = session.getNode(pathName);
			if (baseNode != null) {
				baseNode = JcrUtils.getOrCreateByPath(baseNode, readCellData, false, AppConstants.NT_UNSTRUCTURED,
						AppConstants.NT_UNSTRUCTURED, true);
				log.info("Node created at: " + baseNode + readCellData);
			}
		} catch (RepositoryException e) {
			log.error("RepositoryException in createNode: " + e.getMessage());
		}

	}

	private String readCellData(Cell cell) {
		String value = null;
		try {
			if (cell != null) {
				switch (cell.getCellTypeEnum()) {
				case BLANK:
					break;
				case STRING:
					value = cell.getStringCellValue();
					break;
				case NUMERIC:
					value = String.valueOf(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					value = String.valueOf(cell.getBooleanCellValue());
					break;
				default:
					break;
				}
			}
		} catch (Exception e) {
			log.error("Exception in readCellData: " + e.getMessage());
		}
		return value;
	}

}
