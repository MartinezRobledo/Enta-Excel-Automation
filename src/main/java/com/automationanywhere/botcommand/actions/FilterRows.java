package com.automationanywhere.botcommand.actions;

import java.util.*;
import java.util.stream.Collectors;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
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
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;

// Asumimos los imports/annotaciones del SDK de tu bot framework (Automation Anywhere Package SDK)
// import com.automationanywhere.commands.annotations.*;
// import com.automationanywhere.core.*;
// etc.

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
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String originSheetName,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
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

            @Idx(index = "4.2.2", type = AttributeType.TEXT)
            @Pkg(label = "Headers range (e.g., C9:BM9)")
            @NotEmpty
            String headersRange,

            @Idx(index = "5", type = AttributeType.ENTRYLIST, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(title = "Column", label = "Column")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(title = "Criteria", label = "Criteria"))
            })
            @Pkg(label = "Provide filter entries criteria o multiple criterias delimited by ;")
            @EntryListLabel(value = "Provide entry")
            @EntryListAddButtonLabel(value = "Add entry")
            //@EntryListEntryUnique(value = "Column")
            @EntryListEmptyLabel(value = "No parameters added")
            List<Value> entryList
    ) {

        if (entryList == null || entryList.isEmpty()) {
            throw new BotCommandException("No filter entries provided. Please add at least one entry.");
        }

        // Normalizar/validar y agrupar criterios por columna (permite OR en una misma columna)
        Map<String, List<String>> criteriaMap = new LinkedHashMap<>();
        int entryIndex = 1;
        for (Value v : entryList) {
            @SuppressWarnings("unchecked")
            Map<String, Object> row = (Map<String, Object>) v.get();

            String colKey = row.get("Column") != null ? row.get("Column").toString().trim()
                    : row.get("column") != null ? row.get("column").toString().trim() : "";
            String criteria = row.get("Criteria") != null ? row.get("Criteria").toString().trim()
                    : row.get("criteria") != null ? row.get("criteria").toString().trim() : "";

            if (colKey.isEmpty()) {
                throw new BotCommandException("Entry " + entryIndex + ": Column value cannot be empty.");
            }
            if (criteria.isEmpty()) {
                throw new BotCommandException("Entry " + entryIndex + ": Criteria value cannot be empty.");
            }

            // Dividir el string de criterios por ';' y agregar cada uno por separado
            String[] criteriaParts = criteria.split(";");
            for (String c : criteriaParts) {
                String trimmed = c.trim();
                if (!trimmed.isEmpty()) {
                    criteriaMap.computeIfAbsent(colKey, k -> new ArrayList<>()).add(trimmed);
                }
            }

            entryIndex++;
        }

        // ---------- Obtener sesión/Excel/Workbook/Sheet ----------
        Session session = SessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found: " + sessionId);

        Dispatch wb = session.openWorkbooks.get(sourceWorkbookName);
        if (wb == null)
            throw new BotCommandException("Workbook not open: " + sourceWorkbookName);

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet = "index".equalsIgnoreCase(selectSheetBy)
                ? Dispatch.call(sheets, "Item", originSheetIndex.intValue()).toDispatch()
                : Dispatch.call(sheets, "Item", originSheetName).toDispatch();

        // ---------- Parsear rango de encabezados ----------
        HeaderRange hr = parseHeaderRange(headersRange); // valida formato y misma fila
        int headerRow = hr.row;
        int headerStartCol = hr.startCol;
        int headerEndCol = hr.endCol;
        if (headerEndCol < headerStartCol) {
            throw new BotCommandException("Headers range inválido: la columna final es menor a la inicial.");
        }

        // ---------- Última fila usada ----------
        int lastRow = Math.max(headerRow, getLastRow(sheet));

        // ---------- Detectar si hay tabla que incluya el headerRow ----------
        Dispatch dataRange = null;
        Dispatch listObjects = Dispatch.get(sheet, "ListObjects").toDispatch();
        int tableCount = Dispatch.get(listObjects, "Count").getInt();

        if (tableCount > 0) {
            for (int i = 1; i <= tableCount; i++) {
                Dispatch table = Dispatch.call(listObjects, "Item", i).toDispatch();
                Dispatch tableRange = Dispatch.get(table, "Range").toDispatch();

                int tblFirstRow = Dispatch.get(tableRange, "Row").getInt();
                int tblLastRow = tblFirstRow +
                        Dispatch.get(Dispatch.get(tableRange, "Rows").toDispatch(), "Count").getInt() - 1;

                if (headerRow >= tblFirstRow && headerRow <= tblLastRow) {
                    dataRange = tableRange; // usamos toda la tabla
                    break;
                }
            }
        }

        // Si no encontramos tabla, usamos rango normal
        if (dataRange == null) {
            Dispatch topLeft = Dispatch.call(sheet, "Cells", headerRow, headerStartCol).toDispatch();
            Dispatch bottomRight = Dispatch.call(sheet, "Cells", lastRow, headerEndCol).toDispatch();
            dataRange = Dispatch.call(sheet, "Range", topLeft, bottomRight).toDispatch();
        }

        // ---------- Map headers si referenceMode = header ----------
        Map<String, Integer> headerToIndex = new HashMap<>();
        if ("header".equalsIgnoreCase(referenceMode)) {
            boolean doTrim = (trimHeaders == null) ? true : trimHeaders;
            int field = 1;
            for (int col = headerStartCol; col <= headerEndCol; col++, field++) {
                Dispatch cell = Dispatch.call(sheet, "Cells", headerRow, col).toDispatch();
                String val = safeVariantToString(Dispatch.get(cell, "Value"));
                if (doTrim && val != null) val = val.trim();
                if (val != null && !val.isEmpty()) {
                    headerToIndex.put(val.toLowerCase(), field); // 1-based relativo al rango
                }
            }
        }

        // ---------- Aplicar AutoFilter ----------
        for (Map.Entry<String, List<String>> e : criteriaMap.entrySet()) {
            String colKeyOriginal = e.getKey();
            String keyNorm = colKeyOriginal.trim().toLowerCase();

            int fieldIndex;
            if ("letter".equalsIgnoreCase(referenceMode)) {
                int absCol = excelColumnLetterToNumber(colKeyOriginal);
                if (absCol < headerStartCol || absCol > headerEndCol) {
                    throw new BotCommandException(
                            "Column letter '" + colKeyOriginal + "' (col " + absCol + ") está fuera del headers range [" +
                                    columnNumberToLetter(headerStartCol) + ":" + columnNumberToLetter(headerEndCol) + "]."
                    );
                }
                fieldIndex = (absCol - headerStartCol) + 1; // 1-based relativo al rango
            } else { // header
                fieldIndex = headerToIndex.getOrDefault(keyNorm, -1);
                if (fieldIndex <= 0) {
                    String available = String.join(", ",
                            headerToIndex.keySet().stream().sorted().collect(Collectors.toList()));
                    throw new BotCommandException(
                            "Header '" + colKeyOriginal + "' no encontrado en la fila " + headerRow +
                                    ". Headers disponibles en el rango: [" + available + "]."
                    );
                }
            }

            List<String> criteriaList = e.getValue();
            if (criteriaList.size() == 1) {
                Dispatch.call(dataRange, "AutoFilter",
                        new Variant(fieldIndex),
                        new Variant(criteriaList.get(0)));
            } else {
                Variant[] variants = new Variant[criteriaList.size()];
                for (int i = 0; i < criteriaList.size(); i++) {
                    variants[i] = new Variant(criteriaList.get(i));
                }
                Dispatch.callN(dataRange, "AutoFilter",
                        new Object[]{ new Variant(fieldIndex), variants, new Variant(7) });
            }
        }

    }

    // ---------- Helpers ----------

    private static class HeaderRange {
        int row;
        int startCol;
        int endCol;
        HeaderRange(int row, int startCol, int endCol) {
            this.row = row; this.startCol = startCol; this.endCol = endCol;
        }
    }

    private static HeaderRange parseHeaderRange(String range) {
        if (range == null || range.trim().isEmpty())
            throw new BotCommandException("Headers range no puede estar vacío.");

        String r = range.trim().toUpperCase(Locale.ROOT).replace("$", "");
        String[] parts = r.split(":");
        if (parts.length == 1) {
            CellRef c = parseCellRef(parts[0]);
            return new HeaderRange(c.row, c.col, c.col);
        } else if (parts.length == 2) {
            CellRef a = parseCellRef(parts[0]);
            CellRef b = parseCellRef(parts[1]);
            if (a.row != b.row) {
                throw new BotCommandException("Headers range debe estar en una única fila (ej.: C9:BM9).");
            }
            int start = Math.min(a.col, b.col);
            int end   = Math.max(a.col, b.col);
            return new HeaderRange(a.row, start, end);
        } else {
            throw new BotCommandException("Formato inválido para Headers range. Use ej.: C9 o C9:BM9.");
        }
    }

    private static class CellRef {
        int row; int col;
        CellRef(int row, int col) { this.row = row; this.col = col; }
    }

    private static CellRef parseCellRef(String addr) {
        if (addr == null || addr.isEmpty())
            throw new BotCommandException("Referencia de celda vacía en el headers range.");

        String s = addr.trim().toUpperCase(Locale.ROOT);
        int i = 0, n = s.length();

        StringBuilder colSb = new StringBuilder();
        while (i < n && s.charAt(i) >= 'A' && s.charAt(i) <= 'Z') { colSb.append(s.charAt(i++)); }

        StringBuilder rowSb = new StringBuilder();
        while (i < n && Character.isDigit(s.charAt(i))) { rowSb.append(s.charAt(i++)); }

        if (colSb.length() == 0 || rowSb.length() == 0 || i != n) {
            throw new BotCommandException("Dirección inválida: '" + addr + "'. Ej.: C9 o C9:BM9");
        }

        int col = excelColumnLetterToNumber(colSb.toString());
        int row = Integer.parseInt(rowSb.toString());
        return new CellRef(row, col);
    }

    private static int getLastRow(Dispatch sheet) {
        // xlCellTypeLastCell = 11
        try {
            Dispatch cells = Dispatch.get(sheet, "Cells").toDispatch();
            Dispatch lastCell = Dispatch.call(cells, "SpecialCells", new Variant(11)).toDispatch();
            return Dispatch.get(lastCell, "Row").getInt();
        } catch (Exception ex) {
            // Fallback a UsedRange si no hay SpecialCells disponibles
            try {
                Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
                int firstRow = Dispatch.get(usedRange, "Row").getInt();
                int totalRows = Dispatch.get(Dispatch.get(usedRange, "Rows").toDispatch(), "Count").getInt();
                return firstRow + totalRows - 1;
            } catch (Exception ignored) {
                return 1;
            }
        }
    }

    private static int excelColumnLetterToNumber(String col) {
        if (col == null) return -1;
        String s = col.trim().toUpperCase(Locale.ROOT);
        int res = 0;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch < 'A' || ch > 'Z') {
                throw new BotCommandException("Invalid column letter: '" + col + "'");
            }
            res = res * 26 + (ch - 'A' + 1);
        }
        return res;
    }

    private static String columnNumberToLetter(int col) {
        StringBuilder sb = new StringBuilder();
        int n = col;
        while (n > 0) {
            int rem = (n - 1) % 26;
            sb.insert(0, (char)('A' + rem));
            n = (n - 1) / 26;
        }
        return sb.toString();
    }

    private static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return o != null ? o.toString() : "";
    }

}
