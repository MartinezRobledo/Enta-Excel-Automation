package com.automationanywhere.botcommand.utilities;

import com.automationanywhere.toolchain.runtime.session.CloseableSessionObject;

public class ExcelSession implements CloseableSessionObject {

    private final String sessionId;
    private final Session session;

    public ExcelSession(String sessionId, Session session) {
        this.sessionId = sessionId;
        this.session = session;
    }

    public String getSessionId() {
        return sessionId;
    }

    public Session getSession() {
        return session;
    }

    @Override
    public void close() {
        SessionManager.removeSession(sessionId);
    }

    @Override
    public boolean isClosed() {
        return SessionManager.getSession(sessionId) == null;
    }
}
