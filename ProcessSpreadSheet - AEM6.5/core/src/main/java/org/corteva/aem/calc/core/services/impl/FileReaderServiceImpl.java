package org.corteva.aem.calc.core.services.impl;

import java.io.File;
import java.io.IOException;

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

			// Getting the number of sheets
			int numberOfSheets = workbook.getNumberOfSheets();

			for (int i = 0; i < numberOfSheets; i++) {
				
				// Getting the sheet at the ith position
				XSSFSheet sheet = workbook.getSheetAt(i);

				log.info("Reading sheet: {}", sheet.getSheetName());

				if (sheet.getSheetName().equals(AppConstants.SHEET_SELECT_SAVINGS)) {
					
					int tradeCode = 0, productColumnIndex = 1, progromCodeColumnIndex = 2, soldAsUOM = 3,
							selectSavingsEarningRate = 4, requriesPrepay = 5;

					for (Row myrow : sheet) {
						
						//Skip 1st Row
						if (myrow.getRowNum() != 0) {
							
							for (Cell cell : myrow) {

								cell = myrow.getCell(progromCodeColumnIndex);
								String programCode = readCellData(cell);
								log.info("programCode: " + programCode);
								createNode(AppConstants.INITIAL_STRUCTURE, session, programCode);
								cell = myrow.getCell(productColumnIndex);
								String product = readCellData(cell);
								log.info("product: " + product);
								createNode(AppConstants.INITIAL_STRUCTURE + programCode, session, product);
								cell = myrow.getCell(requriesPrepay);
								String prePay = readCellData(cell);
								createProperties(AppConstants.INITIAL_STRUCTURE + programCode + "/" + product, session, AppConstants.SELECT_SAVINGS_REQUIRES_PREPAY, prePay);
								cell = myrow.getCell(selectSavingsEarningRate);
								String earningsRate = readCellData(cell);
								createProperties(AppConstants.INITIAL_STRUCTURE + programCode + "/" + product, session, AppConstants.SELECT_SAVINGS_EARNINGS_RATE, earningsRate);
								cell = myrow.getCell(tradeCode);
								String tradeCodeValue = readCellData(cell);
								createProperties(AppConstants.INITIAL_STRUCTURE + programCode + "/" + product, session, AppConstants.SELECT_SAVINGS_TRADE_CODE, tradeCodeValue);
								cell = myrow.getCell(soldAsUOM);
								String soldAsUOMValue = readCellData(cell);
								createProperties(AppConstants.INITIAL_STRUCTURE + programCode + "/" + product, session, AppConstants.SELECT_SAVINGS_SOLD_AS_UOM, soldAsUOMValue);
								

							}

						}
					}
				}
			}
			workbook.close();

		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			log.error(e.getMessage(), e);
		}

	}

	private void createProperties(String pathName, Session session, String name, String value) {
		try {
			Node baseNode = session.getNode(pathName);
			if (baseNode != null) {
				baseNode.setProperty(name, value);
			}
		} catch (RepositoryException e) {
			log.error("RepositoryException in createProperties: "+e.getMessage());
		}
		
	}

	private void createNode(String pathName, Session session, String readCellData) {
		try {
			Node baseNode = session.getNode(pathName);
			if(baseNode != null) {
				baseNode = JcrUtils.getOrCreateByPath(baseNode, readCellData, false, AppConstants.NT_UNSTRUCTURED, AppConstants.NT_UNSTRUCTURED, true);
			}
		} catch (RepositoryException e) {
			log.error("RepositoryException in createNode: "+e.getMessage());
		}

	}

	private String readCellData(Cell cell) {
		String value = null;
		try {
			switch (cell.getCellTypeEnum()) {
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
		} catch (Exception e) {
			log.error("Exception in readCellData: "+e.getMessage());
		}
		return value;
	}

}
