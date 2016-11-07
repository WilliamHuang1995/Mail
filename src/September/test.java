package September;

import java.text.SimpleDateFormat;
import java.util.Date;

import com.agile.api.IAgileSession;
import com.agile.api.IProgram;
import com.agile.api.ProgramConstants;
import com.anselm.plm.utilobj.Ini;
import com.anselm.plm.utilobj.LogIt;

import util.WebClient;

public class test {
	public static void main(String[] args) throws Exception{
		Ini ini = new Ini();
		LogIt log = new LogIt();
		String username = ini.getValue("AgileAP", "username");
		String password = ini.getValue("AgileAP", "password");
		String connectString = ini.getValue("AgileAP", "url");
		IAgileSession session = WebClient.getAgileSession(username, password, connectString);
		IProgram project = (IProgram) session.getObject(IProgram.OBJECT_TYPE, "PROJECT0000004");
		Date startDate = (Date) project.getValue(ProgramConstants.ATT_GENERAL_INFO_SCHEDULE_START_DATE);
		Date endDate = (Date) project.getValue(ProgramConstants.ATT_GENERAL_INFO_SCHEDULE_END_DATE);
		SimpleDateFormat format = new SimpleDateFormat("MM/dd");
		log.log("Start: "+ format.format(startDate)+ " Finish: "+format.format(endDate));
		Date actualStartdate = (Date) project.getValue(ProgramConstants.ATT_GENERAL_INFO_ACTUAL_START_DATE);
		Date today = new Date();
		log.log("today: "+format.format(today)+" actual start date: "+format.format(actualStartdate));
		
	}
}
