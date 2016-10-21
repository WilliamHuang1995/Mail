package September;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Iterator;

import com.agile.api.APIException;
import com.agile.api.IAgileList;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IProgram;
import com.agile.api.ProgramConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.EventConstants;
import com.agile.px.ICreateEventInfo;
import com.agile.px.IEventAction;
import com.agile.px.IEventDirtyCell;
import com.agile.px.IEventDirtyRow;
import com.agile.px.IEventDirtyTable;
import com.agile.px.IEventInfo;
import com.agile.px.IObjectEventInfo;
import com.agile.px.IUpdateEventInfo;
import com.agile.px.IUpdateTableEventInfo;
import com.anselm.plm.util.AUtil;
import com.anselm.plm.utilobj.Ini;

//this event action is used to detect changes in the title block
public class NoticeChangePPM implements IEventAction {

	public static String URL = "192.168.13.106";
	public static String USERNAME = "admin";
	public static String PASSWORD = "agile933";

	@Override
	// This code looks for an existing item under the category Parts and throws
	// and error if it exists.
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo event) {

		Ini ini = new Ini();
		ArrayList<String> array = new ArrayList<String>(Arrays.asList("EE", "PM", "ME", "BIOS", "DQA", "TEST"));

		IUpdateEventInfo info = (IUpdateEventInfo) event;

		try {
			IEventDirtyCell[] cells = info.getCells();
			IProgram program = (IProgram) session.getObject(info.getDataObject().getAgileClass(),
					info.getDataObject().getName());
			for (IEventDirtyCell cell : cells) {
				if (array.contains(cell.getAttribute().getName())) {
					String attributeName = cell.getAttribute().getName();
					System.out.println(attributeName);
					System.out.println(cell.getValue());

					//TODO find a way to filter out what is added
					//TODO find a way to distinguish things from being added and removed
					// gets the previous value
					ICell name = program.getCell(cell.getAttributeId());
					System.out.println("Original val: " + name.getValue());
					System.out.println("Dirty val: " + cell.getValue());
				}
			}
			// file:///C:/Users/user/Downloads/933_SDK_Samples/sdk/documentation/html/interfacecom_1_1agile_1_1px_1_1IUpdateTitleBlockEventInfo.html
			// file:///C:/Users/user/Downloads/933_SDK_Samples/sdk/documentation/html/interfacecom_1_1agile_1_1px_1_1IEventDirtyCell.html

			session = AUtil.getAgileSession(ini, "AgileAP");
			System.out.println(session.getCurrentUser());

			System.out.println(program.getName());
		} catch (APIException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return new EventActionResult(event, new ActionResult(ActionResult.NORESULT, null));
	}

	private static String[] getlist(ICell cell) throws APIException {

		IAgileList cl = (IAgileList) cell.getValue();// 取得PARTS子類別

		IAgileList[] selected = cl.getSelection();
		if (selected == null) {
			return null;
		}
		String value[] = new String[selected.length];

		for (int i = 0; i < selected.length; i++) {
			value[i] = selected[i].getValue().toString();
		}

		return value;
	}
	/*
	 * Original val: Administrator, admin (admin);Gabriele Braga, Arrow
	 * Electronics (1ae001.01);Julie Fozard, Channel One Limited (1co001.01)
	 * 
	 * Dirty val: Administrator, admin (admin);Gabriele Braga, Arrow Electronics
	 * (1ae001.01);Julie Fozard, Channel One Limited (1co001.01);Pang, Philip,
	 * Avnet Electronics Marketing Australia (1ae002.01) PM
	 */

	// http://anselm-demoplm:7001/Agile/PLMServlet?fromPCClient=true&module=ActivityHandler&requestUrl=module%3DActivityHandler%26opcode%3DdisplayObject%26classid%3D18022%26objid%3D12422%26tabid%3D0%26

}
