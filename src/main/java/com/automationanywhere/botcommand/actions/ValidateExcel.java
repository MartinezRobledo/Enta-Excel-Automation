package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListAddButtonLabel;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListEmptyLabel;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListEntryUnique;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListLabel;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.data.Value;
import com.jacob.com.Dispatch;

import java.util.*;
import java.util.stream.Collectors;


@BotCommand
@CommandPkg(
        label = "Validate Excel",
        name = "validateExcel",
        description = "Validates workbook sheets and headers according to the EntryList configuration. Returns empty string if valid, else error message.",
        icon = "excel.svg",
        return_type = DataType.STRING,
        return_label = "Validation Message",
        return_required = true
)
public class ValidateExcel {

    @Idx(index = "5.3", type = AttributeType.TEXT, name = "SheetValues")
    @Pkg(label = "Sheet", default_value_type = DataType.STRING)
    @NotEmpty
    private String entrySheetValues;

    @Idx(index = "5.4", type = AttributeType.TEXT, name = "Headers")
    @Pkg(label = "Headers", default_value_type = DataType.STRING)
    @NotEmpty
    private String entryHeadersValues;

    @Idx(index = "6.3", type = AttributeType.TEXT, name = "SheetRanges")
    @Pkg(label = "Sheet", default_value_type = DataType.STRING)
    @NotEmpty
    private String entrySheetRanges;

    @Idx(index = "6.4", type = AttributeType.TEXT, name = "Range")
    @Pkg(label = "Range", default_value_type = DataType.STRING)
    @NotEmpty
    private String entryHeadersRange;

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.NUMBER)
            @Pkg(label = "Expected number of sheets (optional)")
            Double expectedSheetCount,

            @Idx(index = "3", type = AttributeType.LIST)
            @Pkg(label = "Expected sheet names (optional)")
            List<Object> expectedSheetNames,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Por index (1,2,3...)", value = "index")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "Por nombre de hoja", value = "sheet"))
            })
            @Pkg(label = "Referencia de hoja", default_value = "index", default_value_type = DataType.STRING)
            String referenceMode,

            @Idx(index = "5", type = AttributeType.ENTRYLIST, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(title = "SheetValues", label = "Sheet")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(title = "Headers", label = "Headers (comma-separated)"))
            })
            @Pkg(label = "Provide sheets and expected headers")
            @EntryListLabel(value = "Provide entry")
            @EntryListAddButtonLabel(value = "Add entry")
            @EntryListEntryUnique(value = "SheetValues")
            @EntryListEmptyLabel(value = "No sheets added")
            List<Value> sheetHeadersEntryList,

            @Idx(index = "6", type = AttributeType.ENTRYLIST, options = {
                    @Idx.Option(index = "6.1", pkg = @Pkg(title = "SheetRanges", label = "Sheet")),
                    @Idx.Option(index = "6.2", pkg = @Pkg(title = "Range", label = "Range of headers"))
            })
            @Pkg(label = "Provide sheets and header ranges")
            @EntryListLabel(value = "Provide entry")
            @EntryListAddButtonLabel(value = "Add entry")
            @EntryListEntryUnique(value = "SheetRanges")
            @EntryListEmptyLabel(value = "No sheets added")
            List<Value> sheetRangesEntryList,

            @Idx(index = "7", type = AttributeType.CHECKBOX)
            @Pkg(label = "Headers are discontinuous", description = "Check if there are empty columns between headers",
                    default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean allowDiscontinuousHeaders
    ) {
        // Este String se retorna al final. Si queda vacío => válido.
        String errorMessage = "";

        Session session = excelSession.getSession();
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found o closed");

        if (session.openWorkbooks.isEmpty())
            throw new BotCommandException("No workbook is open in this session.");

        Dispatch wb = session.openWorkbooks.values().iterator().next();

        try {
            Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
            int sheetCount = Dispatch.get(sheets, "Count").getInt();

            if (expectedSheetCount != null && expectedSheetCount.intValue() != sheetCount) {
                errorMessage = "Expected number of sheets: " + expectedSheetCount.intValue()
                        + " but found: " + sheetCount;
                return new StringValue(errorMessage);
            }

            if (expectedSheetNames != null && !expectedSheetNames.isEmpty()) {
                for (Object o : expectedSheetNames) {
                    if (o == null) continue;
                    String expectedName = o.toString();
                    boolean found = false;
                    for (int i = 1; i <= sheetCount; i++) {
                        Dispatch s = Dispatch.call(sheets, "Item", i).toDispatch();
                        String sheetName = Dispatch.get(s, "Name").toString();
                        if (expectedName.equalsIgnoreCase(sheetName)) {
                            found = true;
                            break;
                        }
                    }
                    if (!found) {
                        errorMessage = "Expected sheet name not found: " + expectedName;
                        return new StringValue(errorMessage);
                    }
                }
            }

            // parsear sheetHeadersMap
            Map<String, List<String>> sheetHeadersMap = new HashMap<>();
            if (sheetHeadersEntryList != null) {
                for (Value entryValue : sheetHeadersEntryList) {
                    @SuppressWarnings("unchecked")
                    Map<String, Object> entryMap = (Map<String, Object>) entryValue.get();
                    String sheetKey = entryMap.get("SheetValues").toString();
                    String headersCsv = entryMap.get("Headers").toString();
                    List<String> headersList = Arrays.stream(headersCsv.split(","))
                            .map(String::trim)
                            .filter(s -> !s.isEmpty())
                            .collect(Collectors.toList());
                    sheetHeadersMap.put(sheetKey, headersList);
                }
            }

            // sheetRanges map
            Map<String, String> sheetRangesMap = new HashMap<>();
            if (sheetRangesEntryList != null) {
                for (Value entryValue : sheetRangesEntryList) {
                    @SuppressWarnings("unchecked")
                    Map<String, Object> entryMap = (Map<String, Object>) entryValue.get();
                    String sheetKey = entryMap.get("SheetRanges").toString();
                    String range = entryMap.get("Range").toString();
                    sheetRangesMap.put(sheetKey, range);
                }
            }

            // validar cada sheet
            for (Map.Entry<String, List<String>> e : sheetHeadersMap.entrySet()) {
                String sheetKey = e.getKey();
                List<String> expectedHeaders = e.getValue();

                Dispatch sheet = null;
                try {
                    if ("sheet".equalsIgnoreCase(referenceMode)) {
                        sheet = Dispatch.call(sheets, "Item", sheetKey.trim()).toDispatch();
                    } else {
                        int index = Integer.parseInt(sheetKey);
                        sheet = Dispatch.call(sheets, "Item", index).toDispatch();
                    }
                } catch (Exception ex) {
                    errorMessage = "Sheet not found or invalid reference: " + sheetKey;
                    return new StringValue(errorMessage);
                }

                String headerRange = sheetRangesMap.get(sheetKey);
                if (headerRange == null || headerRange.isEmpty()) {
                    errorMessage = "No header range defined for sheet: " + sheetKey;
                    return new StringValue(errorMessage);
                }

                Dispatch range = Dispatch.call(sheet, "Range", headerRange).toDispatch();
                int lastRow = Dispatch.get(Dispatch.get(range, "Rows").toDispatch(), "Count").getInt();
                int lastCol = Dispatch.get(Dispatch.get(range, "Columns").toDispatch(), "Count").getInt();

                boolean headersFound = false;
                for (int r = 1; r <= lastRow; r++) {
                    List<String> rowValues = new ArrayList<>();
                    for (int c = 1; c <= lastCol; c++) {
                        Object val = Dispatch.get(
                                Dispatch.call(range, "Cells", r, c).toDispatch(),
                                "Value"
                        ).toJavaObject();
                        String cellValue = val != null ? val.toString().trim() : "";
                        if (allowDiscontinuousHeaders) {
                            if (!cellValue.isEmpty()) rowValues.add(cellValue);
                        } else {
                            rowValues.add(cellValue);
                        }
                    }

                    // Compara los headers esperados con los valores de la fila
                    if (rowValues.size() >= expectedHeaders.size()) {
                        boolean match = true;
                        for (int i = 0, j = 0; i < expectedHeaders.size() && j < rowValues.size(); i++, j++) {
                            if (!expectedHeaders.get(i).equals(rowValues.get(j))) {
                                errorMessage += errorMessage + "Header esperado: " + expectedHeaders.get(i) +
                                        " no coincide con el header obtenido: " + rowValues.get(j) + ".";
                                match = false;
                                break;
                            }
                        }
                        if (match) {
                            headersFound = true;
                            break;
                        }
                    }
                }

                if (!headersFound) {
                    errorMessage = errorMessage + " Headers mismatch or not found in sheet: " + sheetKey;
                    return new StringValue(errorMessage);
                }
            }

            // si llegamos hasta acá sin retornar, todo ok
            return new StringValue("");
        } catch (Exception ex) {
            errorMessage = "Error validating Excel: " + ex.getMessage();
            return new StringValue(errorMessage);
        }
    }
}
