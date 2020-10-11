package org.corteva.aem.calc.core.services;

import javax.jcr.Session;

public interface FileReaderService {
	
	void readExcel(Session session, String filePath);
}
