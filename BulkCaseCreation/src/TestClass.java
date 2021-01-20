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
		int caseCount = 0;
		Iterator iter = docContentList.iterator();
		while (iter.hasNext()) {
			ContentTransfer ct = (ContentTransfer) iter.next();
			InputStream stream = ct.accessContentStream();
			int rowLastCell = 0;
			HashMap<Integer, String> headers = new HashMap<Integer, String>();
			HashMap<String, String> propDescMap = new HashMap<String, String>();
			HashMap<Integer, HashMap<String, Object>> caseProperties = new HashMap<Integer, HashMap<String, Object>>();
			ObjectStoreReference targetOsRef = new ObjectStoreReference(targetOS);
			CaseType caseType = CaseType.fetchInstance(targetOsRef, doc.get_Name());
			XSSFWorkbook workbook = new XSSFWorkbook(stream);
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFSheet sheet1 = workbook.getSheetAt(1);
			Iterator<Row> rowIterator = sheet.iterator();
			Iterator<Row> rowIterator1 = sheet1.iterator();
			while (rowIterator1.hasNext()) {
				Row row = rowIterator1.next();
				if (row.getRowNum() > 0) {
					String key = null, value = null;
					key = row.getCell(0).getStringCellValue();
					value = row.getCell(1).getStringCellValue();
					if (key != null && value != null) {
						propDescMap.put(key, value);
					}
				}
			}
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
				int validRowCount = 0;
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
									rowValue.put(propDescMap.get(symName), date);
									colNum++;
								}
							} else {
								rowValue.put(propDescMap.get(headers.get(colNum)), getCharValue(cell));
								colNum++;
							}
							validRowCount += 1;
						}
					} catch (Exception e) {
						System.out.println(e);
						e.printStackTrace();
					}

				}
				if (validRowCount != 0) {
					caseProperties.put(++rowNum, rowValue);
				} else {
					break;
				}
			}
			int rowNum1 = 1;
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
					Row row = sheet.getRow(rowNum1++);
					Cell cell1 = row.createCell(row.getLastCellNum());
					if (!caseId.isEmpty()) {
						caseCount += 1;
						System.out.println("CaseCount: " + caseCount);
						cell1.setCellValue("Success");
					} else {
						cell1.setCellValue("Failure");
					}
				} catch (Exception e) {
					System.out.println(e);
					e.printStackTrace();
				}
			}
			InputStream is = null;
			ByteArrayOutputStream bos = null;
			try {
				bos = new ByteArrayOutputStream();
				workbook.write(bos);
				byte[] barray = bos.toByteArray();
				is = new ByteArrayInputStream(barray);
				String docTitle = doc.get_Name();
				FolderSet folderSet = doc.get_FoldersFiledIn();
				Folder folder = null;
				Iterator<Folder> folderSetIterator = folderSet.iterator();
				if (folderSetIterator.hasNext()) {
					folder = folderSetIterator.next();
				}
				String folderPath = folder.get_PathName();
				folderPath += " Response";
				Folder responseFolder = Factory.Folder.fetchInstance(targetOS, folderPath, null);
				updateDocument(targetOS, is, doc, docTitle, responseFolder);

			} catch (Exception e) {
				System.out.println(e);
				e.printStackTrace();
			} finally {
				if (bos != null) {
					bos.close();
				}
				if (is != null) {
					is.close();
				}
				if (stream != null) {
					stream.close();
				}
			}
		}
	}

	private void updateDocument(ObjectStore os, InputStream is, Document doc, String docTitle, Folder responseFolder) {
		// TODO Auto-generated method stub
		String docClassName = doc.getClassName() + "Response";
		Document updateDoc = Factory.Document.createInstance(os, docClassName);
		ContentElementList contentList = Factory.ContentElement.createList();
		ContentTransfer contentTransfer = Factory.ContentTransfer.createInstance();
		contentTransfer.setCaptureSource(is);
		contentTransfer.set_RetrievalName(docTitle + ".xlsx");
		contentTransfer.set_ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		contentList.add(contentTransfer);

		updateDoc.set_ContentElements(contentList);
		updateDoc.checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION);
		Properties p = updateDoc.getProperties();
		p.putValue("DocumentTitle", docTitle);
		updateDoc.setUpdateSequenceNumber(null);
		updateDoc.save(RefreshMode.REFRESH);
		ReferentialContainmentRelationship rc = responseFolder.file(updateDoc, AutoUniqueName.AUTO_UNIQUE, docTitle,
				DefineSecurityParentage.DO_NOT_DEFINE_SECURITY_PARENTAGE);
		rc.save(RefreshMode.REFRESH);
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
