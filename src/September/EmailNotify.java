package September;

import java.util.HashMap;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.ChangeConstants;
import com.agile.api.IAdmin;
import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IStatus;
import com.agile.api.IUser;
import com.agile.api.UserConstants;

public class EmailNotify extends ServerInfo {
	IAgileSession m_session;
	IAdmin m_admin;
	AgileSessionFactory m_factory;

	public EmailNotify() {
	}

	public static void main(String[] args) {
		try {
			EmailNotify en = new EmailNotify();
			en.run();
		} catch (Exception ex) {
			System.out.println(ex);
		}
	}

	private void run() throws Exception {
		this.m_session = login(username, password, connectString);
		if (m_session != null) {
			System.out.println("Successfully logged in.");
		}
		
	}

	/*
	 * Creating a session and logging in This is sample code provided by Oracle
	 */
	private IAgileSession login(String username, String password, String connectString) throws APIException {
		
		// Create the params variable to hold login parameters
		HashMap params = new HashMap();

		// Put username and password values into params
		params.put(AgileSessionFactory.USERNAME, username);
		params.put(AgileSessionFactory.PASSWORD, password);

		// Get an Agile server instance. ("agileserver" is the name of the Agile
		// proxy server,
		// and "virtualPath" is the name of the virtual path used for the Agile
		// system.)
		m_factory = AgileSessionFactory.getInstance(connectString);

		// Create the Agile PLM session and log in
		return m_factory.createSession(params);

	}
}
