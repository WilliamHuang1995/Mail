package September;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.agile.api.IAgileSession;
import com.agile.px.ICustomTable;
import com.anselm.plm.util.AUtil;
import com.anselm.plm.utilobj.Ini;
import com.anselm.plm.utilobj.LogIt;

public class SignOffBlackList implements ICustomTable{

	public Collection<HashMap<String, String>> getTable(IAgileSession session, Map map) throws Exception {

		List<HashMap<String,String>> result = new ArrayList<HashMap<String,String>>();
		Ini ini = new Ini("C:/Users/user/Desktop/Anselm/Config.ini");
		LogIt log = new LogIt("");
		Connection conA = null;
		
		try{
			conA = AUtil.getDbConn(ini, "AgileDB");
			String sql = 
					"select "
					+ "round((sysdate-w.last_upd)-2*FLOOR((sysdate-w.last_upd)/7)-DECODE(SIGN(TO_CHAR(sysdate,'D')-TO_CHAR(w.last_upd,'D')),-1,2,0)+DECODE(TO_CHAR(w.last_upd,'D'),7,1,0)-DECODE(TO_CHAR(sysdate,'D'),7,1,0),2) as WORKDAYS "
					+ ", round((sysdate-w.last_upd),2)  DAYS 			"
					+ ", s.user_name_assigned   		USER_NAME 		"
					+ ", usr.loginid 					USER_ACCOUNT	"
					+ ", n3.description 				CHANGE_TYPE 	"
					+ ", c.change_number 				CHANGE_NUMBER 	"
					+ ", s.last_upd 					LAST_UPD 		"
					+ ", n1.description 				STATUS_FROM		"
					+ ", n2.description 				STATUS_TO 		"
					+ ", c.description 					CHANGE_DESC 	"
					+ ", s.id, s.change_id, s.process_id, s.user_assigned "
					+ "from signoff s, change c, workflow_process w, nodetable n1, nodetable n2, agileuser usr, nodetable n3 "+
					"where "+
					"  s.signoff_status=0 and c.delete_flag is null and s.change_id=c.id "+
					"  and c.process_id=s.process_id and w.id=c.process_id and c.subclass=n3.id "+
					"and w.state=n1.id and w.next_state=n2.id and usr.id=s.user_assigned "+
					"order by days desc, s.last_upd desc";
			ResultSet rs = conA.createStatement().executeQuery(sql);
			ResultSetMetaData rsmd = rs.getMetaData();
			int numCols = rsmd.getColumnCount();
			for(int i=1;i<=numCols;i++)log.log(rsmd.getColumnName(i));
			int count = 0;
			while(rs.next()){
				Map<String,String> datarow = new HashMap<String,String>();
				for(int i=1;i<=numCols;i++)datarow.put(rsmd.getColumnName(i), rs.getString(i));
				result.add((HashMap<String, String>) datarow);
				if(count>1000)break;
			}
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			try{conA.close();log.log("Agile DB Closed.");}catch(Exception e){}
		}
		
		
		return result;
	}

}

