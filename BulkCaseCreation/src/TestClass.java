import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map.Entry;

import javax.security.auth.Subject;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.filenet.api.collection.ContentElementList;
import com.filenet.api.collection.DocumentSet;
import com.filenet.api.collection.FolderSet;
import com.filenet.api.constants.AutoClassify;
import com.filenet.api.constants.AutoUniqueName;
import com.filenet.api.constants.CheckinType;
import com.filenet.api.constants.ClassNames;
import com.filenet.api.constants.DefineSecurityParentage;
import com.filenet.api.constants.PropertyNames;
import com.filenet.api.constants.RefreshMode;
import com.filenet.api.constants.ReservationType;
import com.filenet.api.core.Connection;
import com.filenet.api.core.ContentTransfer;
import com.filenet.api.core.Document;
import com.filenet.api.core.Domain;
import com.filenet.api.core.Factory;
import com.filenet.api.core.Folder;
import com.filenet.api.core.ObjectStore;
import com.filenet.api.core.ReferentialContainmentRelationship;
import com.filenet.api.property.FilterElement;
import com.filenet.api.property.Properties;
import com.filenet.api.property.PropertyFilter;
import com.filenet.api.util.UserContext;
import com.ibm.casemgmt.api.Case;
import com.ibm.casemgmt.api.CaseType;
import com.ibm.casemgmt.api.constants.ModificationIntent;
import com.ibm.casemgmt.api.context.CaseMgmtContext;
import com.ibm.casemgmt.api.context.SimpleP8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleVWSessionCache;
import com.ibm.casemgmt.api.objectref.ObjectStoreReference;

public class TestClass {
	String uri = "http://ibmbaw:9080/wsi/FNCEWS40MTOM";
	String username = "dadmin";
	String password = "dadmin";
	String folderName = "C://Templates/";
	String TOS = "tos";
	UserContext old = null;
	CaseMgmtContext oldCmc = null;
	int columnCount = 1;

	public void fetchDocument() {

		try {
			Connection conn = Factory.Connection.getConnection(uri);
			Subject subject = UserContext.createSubject(conn, username, password, "FileNetP8WSI");
			UserContext.get().pushSubject(subject);

			Domain domain = Factory.Domain.fetchInstance(conn, null, null);
			System.out.println("Domain: " + domain.get_Name());
			System.out.println("Connection to Content Platform Engine successful");
			ObjectStore targetOS = (ObjectStore) domain.fetchObject(ClassNames.OBJECT_STORE, TOS, null);
			System.out.println("Object Store =" + targetOS.get_DisplayName());

			SimpleVWSessionCache vwSessCache = new SimpleVWSessionCache();
			CaseMgmtContext cmc = new CaseMgmtContext(vwSessCache, new SimpleP8ConnectionCache());
			oldCmc = CaseMgmtContext.set(cmc);

			PropertyFilter pf = new PropertyFilter();
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_SIZE, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_ELEMENTS, null));
			String folderPath = "/Bulk Case Creation";
			Folder myFolder = Factory.Folder.fetchInstance(targetOS, folderPath, null);
			DocumentSet myLoanDocs = myFolder.get_ContainedDocuments();
			Iterator itr = myLoanDocs.iterator();
			while (itr.hasNext()) {
				Document doc = (Document) itr.next();
				doc.fetchProperties(pf);
				SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
				String docCheckInDate = formatter.format(doc.get_DateCheckedIn());
				String todayDate = formatter.format(new Date());
				if (docCheckInDate.equals(todayDate)) {
					createBulkCase(targetOS, doc);
				} else {
					System.out.println("No Templates Available, Please upload template and try again..!!");
				}
			}

		} catch (Exception e) {
			System.out.println(e);
		} finally {
			if (oldCmc != null) {
				CaseMgmtContext.set(oldCmc);
			}

			if (old != null) {
				UserContext.set(old);
			}
		}

	}

	public void createBulkCase(ObjectStore targetOS, Document doc) throws IOException {
		ContentElementList docContentList = doc.get_ContentElements();
		Iterator iter = docContentList.iterator();
		while (iter.hasNext()) {
			ContentTransfer ct = (ContentTransfer) iter.next();
			InputStream stream = ct.accessContentStream();
			int rowLastCell = 0;
			HashMap<Integer, String> headers = new HashMap<Integer, String>();
			HashMap<Integer, HashMap<String, Object>> caseProperties = new HashMap<Integer, HashMap<String, Object>>();
			ObjectStoreReference targetOsRef = new ObjectStoreReference(targetOS);
			CaseType caseType = CaseType.fetchInstance(targetOsRef, doc.get_Name());
			XSSFWorkbook workbook = new XSSFWorkbook(stream);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			String headerValue;
			int rowNum = 0;
			if (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				int colNum = 0;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					headerValue = cell.getStringCellValue();
					if (headerValue.contains("*")) {
						if (headerValue.contains("datetime")) {
							headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
							headerValue += "dateField";
						} else {
							headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
						}
					}
					if (headerValue.contains("datetime")) {
						headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
						headerValue += "dateField";
					} else {
						headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
					}
					headers.put(colNum++, headerValue);
				}
				rowLastCell = row.getLastCellNum();
				Cell cell1 = row.createCell(rowLastCell, Cell.CELL_TYPE_STRING);
				if (row.getRowNum() == 0) {
					cell1.setCellValue("Status");
				}
			}
			while (rowIterator.hasNext()) {
				HashMap<String, Object> rowValue = new HashMap<String, Object>();
				Row row = rowIterator.next();
				int colNum = 0;

				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
					try {
						if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
							colNum++;
						} else {
							if (headers.get(colNum).contains("dateField")) {
								String symName = headers.get(colNum).replace("dateField", "");
								if (HSSFDateUtil.isCellDateFormatted(cell)) {
									Date date = cell.getDateCellValue();
									rowValue.put(symName, date);
									colNum++;
								}
							} else {
								rowValue.put(headers.get(colNum), getCharValue(cell));
								colNum++;
							}
						}
					} catch (Exception e) {
						System.out.println(e);
						e.printStackTrace();
					}
				}
				caseProperties.put(++rowNum, rowValue);

			}
			System.out.println(caseProperties);
			Iterator<Entry<Integer, HashMap<String, Object>>> caseProperty = caseProperties.entrySet().iterator();
			while (caseProperty.hasNext()) {
				try {
					Case pendingCase = null;
					String caseId = "";
					Entry<Integer, HashMap<String, Object>> propertyPair = caseProperty.next();
					System.out.println("RowNumber :   " + propertyPair.getKey());
					pendingCase = Case.createPendingInstance(caseType);
					Iterator<Entry<String, Object>> propertyValues = (propertyPair.getValue()).entrySet().iterator();
					while (propertyValues.hasNext()) {
						Entry<String, Object> propertyValuesPair = propertyValues.next();
						pendingCase.getProperties().putObjectValue(propertyValuesPair.getKey(),
								propertyValuesPair.getValue());
						propertyValues.remove();
					}
					pendingCase.save(RefreshMode.REFRESH, null, ModificationIntent.MODIFY);
					caseId = pendingCase.getId().toString();
					System.out.println("Case_ID: " + caseId);
				} catch (Exception e) {
					System.out.println(e);
					e.printStackTrace();
				}
			}

		}

	}

	private static Object getCharValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();

		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return null;
	}
}
