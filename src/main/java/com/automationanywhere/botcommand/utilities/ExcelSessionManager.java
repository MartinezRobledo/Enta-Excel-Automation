package com.automationanywhere.botcommand.utilities;

import java.util.concurrent.ConcurrentHashMap;
import java.util.Map;

public class ExcelSessionManager {
    private static final Map<String, ExcelSession> sessions = new ConcurrentHashMap<>();

    public static void addSession(String sessionId, ExcelSession session) {
        sessions.put(sessionId, session);
    }

    public static ExcelSession getSession(String sessionId) {
        return sessions.get(sessionId);
    }

    public static void removeSession(String sessionId) {
        ExcelSession session = sessions.remove(sessionId);
        if (session != null && session.excelApp != null) {
            session.excelApp.invoke("Quit");
        }
    }

    public static Map<String, ExcelSession> getSessions() {
        return sessions;
    }

}
