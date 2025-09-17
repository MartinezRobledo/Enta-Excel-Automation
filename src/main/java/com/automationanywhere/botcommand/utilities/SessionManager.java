package com.automationanywhere.botcommand.utilities;

import java.util.concurrent.ConcurrentHashMap;
import java.util.Map;

public class SessionManager {
    private static final Map<String, Session> sessions = new ConcurrentHashMap<>();

    public static void addSession(String sessionId, Session session) {
        sessions.put(sessionId, session);
    }

    public static Session getSession(String sessionId) {
        return sessions.get(sessionId);
    }

    public static void removeSession(String sessionId) {
        Session session = sessions.remove(sessionId);
        if (session != null && session.excelApp != null) {
            session.excelApp.invoke("Quit");
        }
    }

    public static Map<String, Session> getSessions() {
        return sessions;
    }

}
