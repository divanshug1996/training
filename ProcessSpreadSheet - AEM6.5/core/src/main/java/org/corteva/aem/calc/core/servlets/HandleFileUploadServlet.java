package org.corteva.aem.calc.core.servlets;

import org.corteva.aem.calc.core.constants.AppConstants;

import java.io.File;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.Map;

import javax.servlet.Servlet;
import javax.jcr.Session;

import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.io.FileUtils;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.request.RequestParameter;
import org.apache.sling.api.servlets.HttpConstants;
import org.apache.sling.api.servlets.SlingAllMethodsServlet;
import org.osgi.framework.Constants;
import org.osgi.service.component.annotations.Component;
import org.osgi.service.component.annotations.Reference;
import org.corteva.aem.calc.core.services.FileReaderService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Component(service = Servlet.class, property = { Constants.SERVICE_DESCRIPTION + "= Handle File Upload Servlet",
		"sling.servlet.methods=" + HttpConstants.METHOD_POST,
		"sling.servlet.paths=" + "/bin/processspreadsheet/fileupload" })
public class HandleFileUploadServlet extends SlingAllMethodsServlet {

	// Generated serial version UID
	private static final long serialVersionUID = -6272122496444152824L;

	// Default logger
	private final Logger log = LoggerFactory.getLogger(this.getClass());

	// Path of the temporary file
	private String tempFilePath;

	// PrintWriter instance to set response
	private PrintWriter printWriter;
	
	// Injecting reference of the FileReaderService
	@Reference
	private FileReaderService fileReaderService;

	@Override
	protected void doPost(SlingHttpServletRequest request, SlingHttpServletResponse response) {

		log.info("Invoking HandleFileUploadServlet...");

		try {

			// Check if the file is multi-part
			final boolean isMultipart = ServletFileUpload.isMultipartContent(request);

			// Setting the temporary file path - This path will be on the server from
			// where the AEM is running
			tempFilePath = System.getProperty("user.dir");

			// Getting the writer instance from the response object
			printWriter = response.getWriter();
			
			String createdFilePath = AppConstants.EMPTY_STRING;
			
			// Temporary file
			File file = null;

			if (isMultipart) {

				// Getting the request parameters from the request object
				Map<String, RequestParameter[]> parameters = request.getRequestParameterMap();

				// Getting the request parameters from the entry set
				for (final Map.Entry<String, RequestParameter[]> pairs : parameters.entrySet()) {

					// Getting the value of request parameter - first element only
					RequestParameter parameter = pairs.getValue()[0];

					// Checking if the posted value is a file or JCR path
					final boolean isFormField = parameter.isFormField();

					if (!isFormField) {
						// Getting the input stream object
						InputStream inputStream = parameter.getInputStream();

						// Creating a temporary file
						file = File.createTempFile("sample", ".xlsx", new File(tempFilePath));

						// Writing contents from input stream to the temporary file
						FileUtils.copyInputStreamToFile(inputStream, file);
						
						createdFilePath = file.getAbsolutePath();
					}
				}
				printWriter.println("File uploaded successfully");
				
				log.debug("Created File Path: "+createdFilePath);
				
				//final ResourceResolver resolver = request.getResourceResolver();
				
				final Session session = request.getResourceResolver().adaptTo(Session.class);
				
				fileReaderService.readExcel(session, createdFilePath);
				
				log.info("Records have been read from the file");
				//printWriter.println("Records have been read from the file");
				
				// Deleting the temporary file
				file.delete();
				
				//printWriter.println("File has been processed successfully");
				
			}
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			//printWriter.println(e.getMessage());
		} finally {
			if (printWriter != null) {
				printWriter.close();
			}
		}
	}

}
