package util;

import java.util.HashMap;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.IAgileSession;

public class WebClient {
	public static IAgileSession login(String username, String password, String connectString) throws APIException {
		HashMap params = new HashMap();
		params.put(AgileSessionFactory.USERNAME, username);
		params.put(AgileSessionFactory.PASSWORD, password);
		AgileSessionFactory m_factory = AgileSessionFactory.getInstance(connectString);
		return m_factory.createSession(params);
	}
}
