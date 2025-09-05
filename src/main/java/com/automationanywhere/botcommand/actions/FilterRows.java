package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.ExcelSessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListAddButtonLabel;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListEmptyLabel;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListEntryUnique;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListLabel;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@BotCommand
@CommandPkg(
        label = "Filter Rows",
        name = "filterRows",
        description = "Filtra filas de un sheet según criterios",
        icon = "excel.svg"
)
public class FilterRows {

    @Idx(index = "5.3", type = AttributeType.TEXT, name = "Column")
    @Pkg(label = "Column", default_value_type = DataType.STRING)
    @NotEmpty
    private String entryColumn;

    @Idx(index = "5.4", type = AttributeType.TEXT, name = "Criteria")
    @Pkg(label = "Criteria", default_value_type = DataType.STRING)
    @NotEmpty
    private String entryCriteria;

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionId,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Workbook Path")
            @NotEmpty
            String sourceWorkbookName,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            @NotEmpty
            String originSheetName,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            @NotEmpty
            Double originSheetIndex,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Por letra (A,B,C...)", value = "letter")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "Por nombre de encabezado", value = "header"))
            })
            @Pkg(label = "Referencia de columna", default_value = "letter", default_value_type = DataType.STRING)
            String referenceMode,

            @Idx(index = "4.2.1", type = AttributeType.CHECKBOX)
            @Pkg(label = "Aplicar TRIM a encabezados (solo si usa nombre)", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean trimHeaders,

            @Idx(index = "5", type = AttributeType.ENTRYLIST, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(title = "Column", label = "Column")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(title = "Criteria", label = "Criteria"))
            })
            @Pkg(label = "Provide filter entries")
            @EntryListLabel(value = "Provide entry")
            //Button label which displays the entry form
            @EntryListAddButtonLabel(value = "Add entry")
            //Unique rule for the column, this value is the column TITLE.
            @EntryListEntryUnique(value = "NAME")
            //Message to dispaly in the table when no entries are present.
            @EntryListEmptyLabel(value = "No parameters added")
            List<Value> entryList
    ) {

        if (entryList == null || entryList.isEmpty()) {
            throw new BotCommandException("No filter entries provided. Please add at least one entry.");
        }

        // Convertir entryList a Map<columna, valores>
        Map<String, List<String>> criteriaMap = new HashMap<>();
        int entryIndex = 1;

        for (Value v : entryList) {
            @SuppressWarnings("unchecked")
            Map<String, Object> row = (Map<String, Object>) v.get();

            String colKey = row.get("Column") != null ? row.get("Column").toString().trim() :
                    row.get("column") != null ? row.get("column").toString().trim() : "";

            String criteria = row.get("Criteria") != null ? row.get("Criteria").toString().trim() :
                    row.get("criteria") != null ? row.get("criteria").toString().trim() : "";

            if (colKey.isEmpty()) {
                throw new BotCommandException("Entry " + entryIndex + ": Column value cannot be empty.");
            }
            if (criteria.isEmpty()) {
                throw new BotCommandException("Entry " + entryIndex + ": Criteria value cannot be empty.");
            }

            criteriaMap.computeIfAbsent(colKey, k -> new ArrayList<>()).add(criteria);
            entryIndex++;
        }


        ExcelSession session = ExcelSessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found: " + sessionId);

        Dispatch wb = session.openWorkbooks.get(sourceWorkbookName);
        if (wb == null)
            throw new BotCommandException("Workbook not open: " + sourceWorkbookName);

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet = "index".equalsIgnoreCase(selectSheetBy)
                ? Dispatch.call(sheets, "Item", originSheetIndex.intValue()).toDispatch()
                : Dispatch.call(sheets, "Item", originSheetName).toDispatch();

        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        int firstRow = Dispatch.get(usedRange, "Row").getInt();
        int totalCols = Dispatch.get(Dispatch.get(usedRange, "Columns").toDispatch(), "Count").getInt();

        // Map headers si referenceMode = header
        Map<String, Integer> headerToIndex = new HashMap<>();
        if ("header".equalsIgnoreCase(referenceMode)) {
            for (int c = 1; c <= totalCols; c++) {
                Dispatch cell = Dispatch.call(sheet, "Cells", firstRow, c).toDispatch();
                String val = safeVariantToString(Dispatch.get(cell, "Value"));
                if (trimHeaders && val != null) val = val.trim();
                headerToIndex.put(val.toLowerCase(), c);

            }
        }

        // Aplicar AutoFilter para cada entrada
        for (Value v : entryList) {
            @SuppressWarnings("unchecked")
            Map<String, Object> row = (Map<String, Object>) v.get();
            String colKey = row.get("Column") != null ? row.get("Column").toString() : "";
            String criteria = row.get("Criteria") != null ? row.get("Criteria").toString() : "";

            String keyNorm = colKey.trim().toLowerCase();
            int colIndex = "letter".equalsIgnoreCase(referenceMode)
                    ? excelColumnLetterToNumber(colKey)
                    : headerToIndex.getOrDefault(keyNorm, -1);

            if (colIndex <= 0) {
                throw new BotCommandException("Invalid column: " + colKey + ". Index: " + keyNorm + ". Letter: " + colIndex);
            }

            Dispatch range = Dispatch.get(sheet, "UsedRange").toDispatch();
            Dispatch.call(range, "AutoFilter", colIndex, criteria);
        }
    }

    private static int excelColumnLetterToNumber(String col) {
        int res = 0;
        for (int i = 0; i < col.length(); i++) {
            res = res * 26 + (col.charAt(i) - 'A' + 1);
        }
        return res;
    }

    private static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return o != null ? o.toString() : "";
    }
}
