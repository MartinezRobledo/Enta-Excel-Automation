package com.automationanywhere.botcommand.utilities;

import com.automationanywhere.toolchain.runtime.session.CloseableSessionObject;

import java.io.Serializable;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Wrapper que viaja en el SessionValue de AA y marca el libro asociado.
 * Debe implementar CloseableSessionObject para ser aceptado por SessionValue.Builder.withSessionObject(...)
 */
public class ExcelSession implements CloseableSessionObject, Serializable {
    private static final long serialVersionUID = 1L;

    private final String sessionId;
    private final String workbookKey; // clave normalizada del libro abierto en esta variable
    private transient Session session; // objeto COM no serializable

    // Estado de cierre requerido por CloseableSessionObject
    private final AtomicBoolean closed = new AtomicBoolean(false);

    public ExcelSession(String sessionId, Session session, String workbookKey) {
        this.sessionId = sessionId;
        this.session = session;
        this.workbookKey = workbookKey;
    }

    public String getSessionId() {
        return sessionId;
    }

    public Session getSession() {
        return session;
    }

    public String getWorkbookKey() {
        return workbookKey;
    }

    /**
     * Importante: No cerramos Excel ni libros acá.
     * Solo marcamos el wrapper como "cerrado". El cierre real lo hacen las acciones (CloseWorkbook / CloseAllSessions).
     */
    @Override
    public void close() {
        closed.set(true);
    }

    @Override
    public boolean isClosed() {
        return closed.get();
    }
}