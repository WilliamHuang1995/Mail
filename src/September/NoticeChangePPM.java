package September;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.INode;
import com.agile.api.IProgram;
import com.agile.api.ITableDesc;
import com.agile.api.IUser;
import com.agile.api.UserConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventDirtyCell;
import com.agile.px.IEventInfo;
import com.agile.px.IUpdateEventInfo;
import com.anselm.plm.util.AUtil;
import com.anselm.plm.utilobj.Ini;
import com.anselm.plm.utilobj.LogIt; 

//this event action is used to detect changes in the title block
public class NoticeChangePPM implements IEventAction {
	static HSSFSheet spreadsheet;
	static HSSFWorkbook workbook;
	static int rowcount = 0;
	static final String URL = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&classid=18022&objid=";

	@Override
	// This code looks for an existing item under the category Parts and throws
	// and error if it exists.
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo event) {
		LogIt log = new LogIt();
		Ini ini = new Ini();
		String str = ini.getValue("Settings", "role");
		log.log(ini.getValue("Settings", "role"));
		ArrayList<String> array = new ArrayList<String>(Arrays.asList(str.split("\\s*,\\s*")));
		log.log(array);

		IUpdateEventInfo info = (IUpdateEventInfo) event;

		// initialize filename
		Date today = new Date();
		SimpleDateFormat format = new SimpleDateFormat("MMdd");
		String filename = format.format(today) + ".xlsx";
		String fileLoc = ini.getValue("File Location", "FILEPATH");

		File temp = new File(fileLoc + filename);
		boolean exists = temp.exists();
		System.out.println(exists);

		try {
			rowcount = initializeWorkbook(exists, fileLoc, filename);
		} catch (Exception e) {
			log.log("error");
		}
		int index = 1;
		Map<String, Object[]> empinfo = new TreeMap<String, Object[]>();
		if (!exists) {
			empinfo.put(index+"", new Object[] { "User Role", "Name", "Email", "Project Name","Project Link" });
		}
	
		// get admin session
		session = AUtil.getAgileSession(ini, "AgileAP");
		try {
			IEventDirtyCell[] cells = info.getCells();
			IProgram program = (IProgram) session.getObject(info.getDataObject().getAgileClass(),
					info.getDataObject().getName());
			for (IEventDirtyCell dirtyCell : cells) {
				if (array.contains(dirtyCell.getAttribute().getName())) {
					String attributeName = dirtyCell.getAttribute().getName();
					ICell cleanCell = program.getCell(dirtyCell.getAttributeId());
					
					
					//gets the objID integer
					String objID = program.getId().toString();
					ArrayList<String> idList = new ArrayList<String>(Arrays.asList(objID.split("\\s*=\\s*")));
					objID = idList.get(1).substring(0,idList.get(1).indexOf(" "));
					
					
					
					ArrayList<String> oldValue = new ArrayList(
							Arrays.asList(cleanCell.getValue().toString().split("\\s*;\\s*")));
					ArrayList<String> newValue = new ArrayList<String>(
							Arrays.asList(dirtyCell.getValue().toString().split("\\s*;\\s*")));
					newValue.removeAll(oldValue);
					log.log(newValue);
					Iterator<String> it = newValue.iterator();
					// loop thru new ppl
					while (it.hasNext()) {
						String nextUser = it.next();
						String userID = nextUser.substring(nextUser.lastIndexOf("(") + 1, nextUser.lastIndexOf(")"));
						IUser user = (IUser) session.getObject(IUser.OBJECT_TYPE, userID);
						String userEmail = (String) user.getValue(UserConstants.ATT_GENERAL_INFO_EMAIL);
						String userName = (String) user.getValue(UserConstants.ATT_GENERAL_INFO_FIRST_NAME);
						log.log(userName);
						log.log(userEmail);
						log.log(attributeName + " " + userName + " " + userEmail + " " + program.getName());
						empinfo.put(++index + "",
								new Object[] { attributeName, userName, userEmail, program.getName(),ini.getValue("Server Info", "url")+ URL + objID});
					}
				}
			}

			Set<String> keyid = empinfo.keySet();
			
			for (String key : keyid) {
				HSSFRow row = spreadsheet.createRow(rowcount++);
				Object[] objectArr = empinfo.get(key);
				int cellid = 0;
				for (Object obj : objectArr) {
					Cell cell = row.createCell(cellid++);
					cell.setCellValue((String) obj);
				}
			}

			// Write the workbook in file system
			try {
				FileOutputStream out = new FileOutputStream(new File(fileLoc + filename));
				workbook.write(out);
				out.close();
				log.log("close");
				System.out.println("Write Successful");
			} catch (IOException e) {
			}

		}
		// file:///C:/Users/user/Downloads/933_SDK_Samples/sdk/documentation/html/interfacecom_1_1agile_1_1px_1_1IUpdateTitleBlockEventInfo.html
		// file:///C:/Users/user/Downloads/933_SDK_Samples/sdk/documentation/html/interfacecom_1_1agile_1_1px_1_1IEventDirtyCell.html

		catch (

		APIException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return new EventActionResult(event, new ActionResult(ActionResult.NORESULT, null));
	}

	private int initializeWorkbook(boolean exists, String fileLoc, String filename) throws Exception {
		int rowcount=0;
		if (exists) {
			FileInputStream fis = new FileInputStream(new File(fileLoc + filename));
			workbook = new HSSFWorkbook(fis);
			spreadsheet = workbook.getSheetAt(0);
			// used to append to the current file
			rowcount = spreadsheet.getLastRowNum()+1;
			System.out.println("ROWCOUNT: "+rowcount);
			fis.close();

		} else {
			// Create blank workbook
			workbook = new HSSFWorkbook();
			// Create a blank sheet
			spreadsheet = workbook.createSheet(" Employee Info ");
			// Create row object
			rowcount = 0;
		}
		return rowcount;
	}

	// http://anselm-demoplm:7001/Agile/PLMServlet?fromPCClient=true&module=ActivityHandler&requestUrl=module%3DActivityHandler%26opcode%3DdisplayObject%26classid%3D18022%26objid%3D12422%26tabid%3D0%26

}
