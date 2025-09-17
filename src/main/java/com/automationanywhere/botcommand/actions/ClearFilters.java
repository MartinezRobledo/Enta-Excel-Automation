package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;

@BotCommand
@CommandPkg(
        label = "Clear Filters",
        name = "clearFilters",
        description = "Removes all filters from a sheet",
        icon = "excel.svg"
)
public class ClearFilters {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionId,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Workbook path")
            @NotEmpty
            String workbookName,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            String selectSheetBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NotEmpty
            Double sheetIndex
    ) {

        Session session = SessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found: " + sessionId);

        Dispatch wb = session.openWorkbooks.get(workbookName);
        if (wb == null)
            throw new BotCommandException("Workbook not open: " + workbookName);

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet = "index".equalsIgnoreCase(selectSheetBy)
                ? Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch()
                : Dispatch.call(sheets, "Item", sheetName).toDispatch();

        try {
            // Desactivar filtros
            Dispatch.put(sheet, "AutoFilterMode", false);
        } catch (Exception e) {
            throw new BotCommandException("Error clearing filters: " + e.getMessage());
        }
    }
}
