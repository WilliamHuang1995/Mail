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

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.INode;
import com.agile.api.IProgram;
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

/**
 * <h1>NoticeChangePPM</h1> The NoticeChangePPM program extends IEventAction and
 * adds PLM users to an excel file. This excel file will be used as a reference
 * to send out email notifying users of these new changes This custom PX should
 * be used to detect the 'Save As' event else it would not work.
 *
 * @author William Huang
 * @version 1.0
 * @since 2016-09-26
 */
public class NoticeChangePPM implements IEventAction {
	static HSSFSheet spreadsheet;
	static HSSFWorkbook workbook;
	static int rowcount = 0;
	static final String URL = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&classid=18022&objid=";

	/**********************************************************************
	 * 內部變數 log : 自定義方便debug的log物件 ini : 自定義方便讀取config的程式
	 **********************************************************************/
	private LogIt log = null;
	private Ini ini = null;
	private String filename;
	private String fileLocation;
	private IAgileSession session;
	ArrayList<String> roleList;
	private String currentUser;

	/*
	 * Constructor
	 */
	public NoticeChangePPM() {

		ini = new Ini();
		log = new LogIt("Role Change");

		// Initialize filepath
		Date today = new Date();
		SimpleDateFormat format = new SimpleDateFormat("MMdd");
		filename = format.format(today) + ".xlsx";
		fileLocation = ini.getValue("File Location", "FILE_PATH");
		try {
			log.setLogFile(ini.getValue("File Location", "LOG_FILE_PATH") + format.format(today) + ".log");

		} catch (Exception e) {

			e.printStackTrace();
		}

		// Initialize role list
		String roleFromConfig = ini.getValue("Settings", "role");
		roleList = new ArrayList<String>(Arrays.asList(roleFromConfig.split("\\s*,\\s*")));
		
		//Initialize current user
		currentUser = "Unknown";

	}

	private boolean checkExist() {
		return (new File(fileLocation + filename)).exists();
	}

	@Override
	// This code looks for an existing item under the category Parts and throws
	// and error if it exists.
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo event) {

		try {
			// The event is an update event
			IUpdateEventInfo info = (IUpdateEventInfo) event;

			// Get current user
			IUser sessionUser = session.getCurrentUser();
			currentUser = (String) sessionUser.getValue(UserConstants.ATT_GENERAL_INFO_FIRST_NAME);
			
			// get admin session
			session = AUtil.getAgileSession(ini, "AgileAP");

			// to be frank, the starting index does not matter.
			int mappingIndex = 1;
			Map<String, Object[]> empinfo = new TreeMap<String, Object[]>();
			if (!checkExist()) {
				empinfo.put(mappingIndex + "",
						new Object[] { "User Role", "Name", "user.id", "Email", "Project Name", "Project Link", "Assigned By" });
			}

			IEventDirtyCell[] cells = info.getCells();
			IProgram program = (IProgram) session.getObject(info.getDataObject().getAgileClass(),
					info.getDataObject().getName());
			for (IEventDirtyCell dirtyCell : cells) {
				if (roleList.contains(dirtyCell.getAttribute().getName())) {
					String attributeName = dirtyCell.getAttribute().getName();
					ICell cleanCell = program.getCell(dirtyCell.getAttributeId());

					// gets the objID integer
					String objID = program.getId().toString();
					ArrayList<String> idList = new ArrayList<String>(Arrays.asList(objID.split("\\s*=\\s*")));
					objID = idList.get(1).substring(0, idList.get(1).indexOf(" "));

					ArrayList<String> oldValue = new ArrayList<String>(
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
						//String dataType = cleanCell.getDataType() == 4 ? "Single List" : "Multi List";
						log.log(userName);
						log.log(userEmail);
						log.log(attributeName + " " + userName + " " + userEmail + " " + program.getName());
						empinfo.put(++mappingIndex + "", new Object[] { attributeName, userName, userID, userEmail,
								program.getName(), ini.getValue("Server Info", "url") + URL + objID, currentUser });
					}
				}
			}

			Set<String> keyid = empinfo.keySet();
			if (keyid.size() == 0) {
				log.log("No changes to roles made, exiting");
				log.close();
				return new EventActionResult(event, new ActionResult(ActionResult.NORESULT, null));

			}

			// initialize workbook
			rowcount = initializeWorkbook(checkExist(), fileLocation, filename);
			for (String key : keyid) {
				HSSFRow row = spreadsheet.createRow(rowcount++);
				Object[] objectArr = empinfo.get(key);
				int cellid = 0;
				for (Object obj : objectArr) {
					Cell cell = row.createCell(cellid++);
					cell.setCellValue((String) obj);
				}
			}
			// initialize workbook as well

			// Write the workbook in file system
			try {
				FileOutputStream out = new FileOutputStream(new File(fileLocation + filename));
				workbook.write(out);
				out.close();
				log.log("close");
				log.log("Write Successful");
			} catch (IOException e) {
			}

		}

		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		log.close();
		return new EventActionResult(event, new ActionResult(ActionResult.NORESULT, null));
	}

	private int initializeWorkbook(boolean exists, String fileLoc, String filename) throws Exception {
		int rowcount = 0;
		if (exists) {
			FileInputStream fis = new FileInputStream(new File(fileLoc + filename));
			workbook = new HSSFWorkbook(fis);
			spreadsheet = workbook.getSheetAt(0);
			// used to append to the current file
			rowcount = spreadsheet.getLastRowNum() + 1;
			log.log("ROWCOUNT: " + rowcount);
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
