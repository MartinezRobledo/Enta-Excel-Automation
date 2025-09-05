package com.automationanywhere.botcommand.utilities;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

import java.util.HashMap;
import java.util.Map;

public class ExcelSession {
    public ActiveXComponent excelApp; // Excel.Application
    public Map<String, Dispatch> openWorkbooks; // path -> Workbook

    public ExcelSession(ActiveXComponent app) {
        this.excelApp = app;
        this.openWorkbooks = new HashMap<>();
    }
}
