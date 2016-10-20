package util;

import java.util.HashMap;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.IAgileSession;
import com.agile.api.IUser;

/*
 * This file mimics common behaviors you can do on the front end.
 */
public class WebClient {
	public static IAgileSession getAgileSession(String username, String password, String connectString) throws APIException {
		HashMap<Integer, String> params = new HashMap<Integer, String>();
		params.put(AgileSessionFactory.USERNAME, username);
		params.put(AgileSessionFactory.PASSWORD, password);
		AgileSessionFactory m_factory = AgileSessionFactory.getInstance(connectString);
		return m_factory.createSession(params);
	}

	public static String getCurrentUser(IAgileSession session) throws APIException {
		IUser currentUser = session.getCurrentUser();
		String userName = currentUser.getName();
		
		return userName;
	}


}
