package com.automationanywhere.botcommand.utilities;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

import java.sql.Connection;
import java.util.HashMap;
import java.util.Map;

public class Session {
    // Instancia compartida de Excel.Application
    public ActiveXComponent excelApp;

    // Libros abiertos en ESTA instancia de Excel (clave: workbookKey normalizado)
    public Map<String, Dispatch> openWorkbooks;

    // Conexiones OLE DB asociadas (si aplica)
    public Map<String, Connection> oleDbConnections;

    // Flag original (respetada)
    public Boolean global; // nueva flag

    public Session(ActiveXComponent app) {
        this.excelApp = app;
        this.openWorkbooks = new HashMap<>();
        this.oleDbConnections = new HashMap<>();
        this.global = false;
    }
}