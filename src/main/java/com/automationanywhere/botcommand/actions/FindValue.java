package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

@BotCommand
@CommandPkg(
        label = "Find",
        name = "findValue",
        description = "Busca un valor en la hoja y devuelve la dirección de la celda de la primera coincidencia",
        icon = "excel.svg",
        return_type = DataType.STRING,
        return_label = "Cell Address",
        return_required = true
)
public class FindValue {

    @Execute
    public Value action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NotEmpty
            Double sheetIndex,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Standard", value = "standard")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Regex", value = "regex"))
            })
            @Pkg(label = "Search Mode", default_value = "standard", default_value_type = DataType.STRING)
            String searchMode,

            // --- STANDARD MODE ---
            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Value to search")
            String valueToSearch,

            @Idx(index = "3.1.2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Case sensitive", default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean caseSensitive,

            @Idx(index = "3.1.3", type = AttributeType.CHECKBOX)
            @Pkg(label = "Use wildcard (*, ?)", default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean useWildcard,

            @Idx(index = "3.1.4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1.4.1", pkg = @Pkg(label = "By Rows", value = "rows")),
                    @Idx.Option(index = "3.1.4.2", pkg = @Pkg(label = "By Columns", value = "columns"))
            })
            @Pkg(label = "Search order", default_value = "rows", default_value_type = DataType.STRING)
            String searchOrder,

            // --- REGEX MODE ---
            @Idx(index = "3.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Regex Pattern")
            String regexPattern,

            // --- END UI ---
            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Absolute address", default_value_type = DataType.BOOLEAN, default_value = "true")
            Boolean absolute
    ) {

        Session session = excelSession.getSession();
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found o closed");

        if (session.openWorkbooks.isEmpty())
            throw new BotCommandException("No workbook is open in this session.");

        Dispatch wb = session.openWorkbooks.values().iterator().next();

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet;
        if ("index".equalsIgnoreCase(selectSheetBy)) {
            sheet = Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch();
        } else {
            sheet = Dispatch.call(sheets, "Item", sheetName).toDispatch();
        }

        if ("regex".equalsIgnoreCase(searchMode)) {
            return new StringValue(regexSearch(sheet, regexPattern));
        } else {
            return new StringValue(standardSearch(sheet, valueToSearch, caseSensitive, useWildcard, searchOrder, absolute));
        }

    }

    private String standardSearch(Dispatch sheet, String value, Boolean caseSensitive, Boolean useWildcard, String searchOrder, Boolean absolute) {
        if (value == null || value.trim().isEmpty()) {
            throw new BotCommandException("Search value cannot be empty.");
        }

        Dispatch cells = Dispatch.get(sheet, "Cells").toDispatch();

        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        Dispatch firstCell = Dispatch.call(usedRange, "Cells", 1, 1).toDispatch();

        Dispatch found;
        try {
            found = Dispatch.call(cells, "Find",
                    new Variant(value),                                   // What
                    firstCell,                                           // After
                    new Variant(-4123),                                  // LookIn: xlValues
                    new Variant(useWildcard != null && useWildcard ? 2 : 1), // LookAt
                    new Variant("columns".equalsIgnoreCase(searchOrder) ? 2 : 1), // SearchOrder
                    new Variant(1),                                      // SearchDirection
                    new Variant(caseSensitive != null && caseSensitive) // MatchCase
            ).toDispatch();

        } catch (Exception e) {
            throw new BotCommandException("Excel Find call failed: " + e.getMessage(), e);
        }

        if (found == null || found.m_pDispatch == 0) {
            return "";
        }

        if(absolute)
            return Dispatch.get(found, "Address").getString();
        else {
            // GetAddress(rowAbsolute, columnAbsolute, referenceStyle)
            String cellAddress = Dispatch.call(found, "Address",
                    false,  // RowAbsolute
                    false,  // ColumnAbsolute
                    1       // xlA1 style
            ).getString();
            return  cellAddress;
        }
    }

    private String regexSearch(Dispatch sheet, String pattern) {
        if (pattern == null || pattern.isEmpty()) throw new BotCommandException("Regex pattern required.");
        Pattern regex = Pattern.compile(pattern);

        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        int firstRow = Dispatch.get(usedRange, "Row").getInt();
        int totalRows = Dispatch.get(Dispatch.get(usedRange, "Rows").toDispatch(), "Count").getInt();
        int totalCols = Dispatch.get(Dispatch.get(usedRange, "Columns").toDispatch(), "Count").getInt();

        for (int r = firstRow; r < firstRow + totalRows; r++) {
            for (int c = 1; c <= totalCols; c++) {
                Dispatch cell = Dispatch.call(sheet, "Cells", r, c).toDispatch();
                String val = safeVariantToString(Dispatch.get(cell, "Value"));
                if (val != null) {
                    Matcher m = regex.matcher(val);
                    if (m.find()) {
                        return Dispatch.get(cell, "Address").getString();
                    }
                }
            }
        }
        return "";
    }

    private static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return null;
        Object o = v.toJavaObject();
        return o != null ? o.toString() : null;
    }

}
