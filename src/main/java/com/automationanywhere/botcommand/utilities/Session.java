package com.automationanywhere.botcommand.utilities;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

import java.sql.Connection;
import java.util.HashMap;
import java.util.Map;

public class Session {
    public ActiveXComponent excelApp;
    public Map<String, Dispatch> openWorkbooks;
    public Map<String, Connection> oleDbConnections;
    public Boolean global; // nueva flag

    public Session(ActiveXComponent app) {
        this.excelApp = app;
        this.openWorkbooks = new HashMap<>();
        this.oleDbConnections = new HashMap<>();
        this.global = false;
    }
}
