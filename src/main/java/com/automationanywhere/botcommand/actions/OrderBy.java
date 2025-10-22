package com.automationanywhere.botcommand.actions;

import java.util.*;
import java.util.stream.Collectors;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListAddButtonLabel;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListEmptyLabel;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListLabel;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Order By (Sort rows)",
        name = "orderBy",
        description = "Ordena filas completas por múltiples columnas (como Excel: Primero por… Luego por…)",
        icon = "excel.svg"
)
public class OrderBy {
    @Idx(index = "6.3", type = AttributeType.TEXT, name = "Column")
    @Pkg(label = "Column", default_value_type = DataType.STRING)
    @NotEmpty
    private String entryColumn;

    @Idx(index = "6.4", type = AttributeType.TEXT, name = "Order")
    @Pkg(label = "Order", default_value_type = DataType.STRING)
    @NotEmpty
    private String entryCriteria;

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            Double sheetIndex,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Por letra (A,B,C...)", value = "letter")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Por nombre de encabezado", value = "header"))
            })
            @Pkg(label = "Referencia de columna", default_value = "letter", default_value_type = DataType.STRING)
            String referenceMode,

            @Idx(index = "3.2.1", type = AttributeType.CHECKBOX)
            @Pkg(label = "Aplicar TRIM a encabezados (solo si usa nombre)", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean trimHeaders,

            @Idx(index = "3.2.2", type = AttributeType.TEXT)
            @Pkg(label = "Headers range (e.g., C9:BM9)")
            @NotEmpty
            String headersRange,

            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Match case", default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean matchCase,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Rows (Top-to-bottom)", value = "rows")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Columns (Left-to-right)", value = "columns"))
            })
            @Pkg(label = "Orientation", default_value = "rows", default_value_type = DataType.STRING)
            String orientation,

            @Idx(index = "6", type = AttributeType.ENTRYLIST, options = {
                    @Idx.Option(index = "6.1", pkg = @Pkg(title = "Column", label = "Column")),
                    @Idx.Option(index = "6.2", pkg = @Pkg(title = "Order", label = "Order (asc|desc)"))
            })
            @Pkg(label = "Sort keys (en orden de prioridad)")
            @EntryListLabel(value = "Add sort key")
            @EntryListAddButtonLabel(value = "Add key")
            @EntryListEmptyLabel(value = "No keys added")
            List<Value> sortEntries
    ) {
        if (sortEntries == null || sortEntries.isEmpty()) {
            throw new BotCommandException("Agregá al menos una clave de ordenación.");
        }

        // === Sesión / Workbook / Sheet ===
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet = resolveSheet(sheets, selectSheetBy, sheetName, sheetIndex);

        // === Parse header range y mapa de encabezados ===
        HeaderRange hr = parseHeaderRange(headersRange); // valida formato y misma fila
        int headerRow = hr.row;
        int headerStartCol = hr.startCol;
        int headerEndCol = hr.endCol;
        if (headerEndCol < headerStartCol) {
            throw new BotCommandException("Headers range inválido: la columna final es menor a la inicial.");
        }

        boolean doTrimHeaders = (trimHeaders == null) ? true : trimHeaders;
        Map<String, Integer> headerToIndex = new HashMap<>();
        if ("header".equalsIgnoreCase(referenceMode)) {
            int field = 1;
            for (int col = headerStartCol; col <= headerEndCol; col++, field++) {
                Dispatch cell = Dispatch.call(sheet, "Cells", headerRow, col).toDispatch();
                String val = safeVariantToString(Dispatch.get(cell, "Value"));
                if (doTrimHeaders && val != null) val = val.trim();
                if (val != null && !val.isEmpty()) {
                    headerToIndex.put(val.toLowerCase(Locale.ROOT), field); // 1-based relativo al rango
                }
            }
            if (headerToIndex.isEmpty()) {
                throw new BotCommandException("No se detectaron encabezados en " + headersRange + ".");
            }
        }

        // Última fila usada
        int lastRow = Math.max(headerRow, getLastRow(sheet));
        if (lastRow <= headerRow) {
            // No hay datos para ordenar
            return;
        }

        // === Detectar si hay una tabla que incluya el headerRow ===
        Dispatch dataRange = detectDataRange(sheet, headerRow, headerStartCol, headerEndCol, lastRow);

        // === Configurar Sort ===
        Dispatch sort = Dispatch.get(sheet, "Sort").toDispatch();
        Dispatch sortFields = Dispatch.get(sort, "SortFields").toDispatch();
        // Limpiar sort previos
        try { Dispatch.call(sortFields, "Clear"); } catch (Exception ignored) {}

        // Normalizar entradas
        List<SortKey> keys = normalizeSortEntries(sortEntries);

        // Agregar SortFields en orden
        for (int i = 0; i < keys.size(); i++) {
            SortKey k = keys.get(i);
            int absCol;
            if ("letter".equalsIgnoreCase(referenceMode)) {
                absCol = excelColumnLetterToNumber(k.column);
                if (absCol < headerStartCol || absCol > headerEndCol) {
                    throw new BotCommandException(
                            "Column letter '" + k.column + "' (col " + absCol + ") está fuera del headers range [" +
                                    columnNumberToLetter(headerStartCol) + ":" + columnNumberToLetter(headerEndCol) + "].");
                }
            } else { // header
                int fieldIndex = headerToIndex.getOrDefault(k.column.trim().toLowerCase(Locale.ROOT), -1);
                if (fieldIndex <= 0) {
                    String available = String.join(", ",
                            headerToIndex.keySet().stream().sorted().collect(Collectors.toList()));
                    throw new BotCommandException(
                            "Header '" + k.column + "' no encontrado en la fila " + headerRow +
                                    ". Headers disponibles: [" + available + "].");
                }
                // Convertir fieldIndex (1-based dentro del rango) a columna absoluta
                absCol = headerStartCol + fieldIndex - 1;
            }

            // keyRange: desde headerRow hasta lastRow en la columna seleccionada (incluye header; Header=xlYes)
            Dispatch topLeft = Dispatch.call(sheet, "Cells", headerRow, absCol).toDispatch();
            Dispatch bottomRight = Dispatch.call(sheet, "Cells", lastRow, absCol).toDispatch();
            Dispatch keyRange = Dispatch.call(sheet, "Range", topLeft, bottomRight).toDispatch();

            // Add devuelve un SortField, luego seteamos su "Order": 1=asc, 2=desc
            Dispatch sortField = Dispatch.call(sortFields, "Add", keyRange).toDispatch();
            int orderVal = "desc".equalsIgnoreCase(k.order) ? 2 : 1;
            try { Dispatch.put(sortField, "Order", new Variant(orderVal)); } catch (Exception ignored) {}
        }

        // Rango a ordenar = dataRange (incluye header)
        Dispatch.call(sort, "SetRange", dataRange);

        // Header: 1 = xlYes (0=Guess, 2=No)
        try { Dispatch.put(sort, "Header", new Variant(1)); } catch (Exception ignored) {}

        // MatchCase
        boolean mc = (matchCase != null && matchCase);
        try { Dispatch.put(sort, "MatchCase", new Variant(mc)); } catch (Exception ignored) {}

        // Orientation: 1 = xlTopToBottom (rows), 2 = xlLeftToRight (columns)
        int orient = "columns".equalsIgnoreCase(orientation) ? 2 : 1;
        try { Dispatch.put(sort, "Orientation", new Variant(orient)); } catch (Exception ignored) {}

        // Aplicar
        Dispatch.call(sort, "Apply");
    }

    // ===================== Helpers =====================

    private static Dispatch resolveSheet(Dispatch sheets, String selectSheetBy, String sheetName, Double sheetIndex) {
        int count = Dispatch.get(sheets, "Count").getInt();
        if ("index".equalsIgnoreCase(selectSheetBy)) {
            if (sheetIndex == null) throw new BotCommandException("Sheet Index es requerido cuando 'Select sheet by' = index.");
            if (sheetIndex.intValue() < 1 || sheetIndex.intValue() > count)
                throw new BotCommandException("Sheet Index fuera de rango (1.." + count + ").");
            return Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch();
        } else {
            if (sheetName == null || sheetName.trim().isEmpty()) {
                throw new BotCommandException("Sheet Name es requerido cuando 'Select sheet by' = name.");
            }
            return Dispatch.call(sheets, "Item", sheetName).toDispatch();
        }
    }

    private static Dispatch detectDataRange(Dispatch sheet, int headerRow, int headerStartCol, int headerEndCol, int lastRow) {
        Dispatch listObjects = Dispatch.get(sheet, "ListObjects").toDispatch();
        try {
            int tableCount = Dispatch.get(listObjects, "Count").getInt();
            if (tableCount > 0) {
                for (int i = 1; i <= tableCount; i++) {
                    Dispatch table = Dispatch.call(listObjects, "Item", i).toDispatch();
                    Dispatch tableRange = Dispatch.get(table, "Range").toDispatch();
                    int tblFirstRow = Dispatch.get(tableRange, "Row").getInt();
                    int tblLastRow = tblFirstRow +
                            Dispatch.get(Dispatch.get(tableRange, "Rows").toDispatch(), "Count").getInt() - 1;
                    if (headerRow >= tblFirstRow && headerRow <= tblLastRow) {
                        return tableRange; // usar toda la tabla
                    }
                }
            }
        } catch (Exception ignored) {}

        // Si no hay tabla, usar rango rectangular basado en headersRange y última fila
        Dispatch topLeft = Dispatch.call(sheet, "Cells", headerRow, headerStartCol).toDispatch();
        Dispatch bottomRight = Dispatch.call(sheet, "Cells", lastRow, headerEndCol).toDispatch();
        return Dispatch.call(sheet, "Range", topLeft, bottomRight).toDispatch();
    }

    // === Parseo de headersRange (mismo estilo que FilterRows) ===
    private static class HeaderRange {
        int row; int startCol; int endCol;
        HeaderRange(int row, int startCol, int endCol) { this.row = row; this.startCol = startCol; this.endCol = endCol; }
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
            int end = Math.max(a.col, b.col);
            return new HeaderRange(a.row, start, end);
        } else {
            throw new BotCommandException("Formato inválido para Headers range. Use ej.: C9 o C9:BM9.");
        }
    }
    private static class CellRef { int row; int col; CellRef(int row, int col){ this.row=row; this.col=col; } }
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

    private static List<SortKey> normalizeSortEntries(List<Value> sortEntries) {
        List<SortKey> keys = new ArrayList<>();
        int idx = 1;
        for (Value v : sortEntries) {
            @SuppressWarnings("unchecked")
            Map<String, Object> row = (Map<String, Object>) v.get();
            String col = getStr(row, "Column", true, idx);
            String order = getStr(row, "Order", true, idx);
            order = order.trim().toLowerCase(Locale.ROOT);
            if (!order.equals("asc") && !order.equals("desc")) {
                throw new BotCommandException("Entrada " + idx + ": 'Order' debe ser 'asc' o 'desc'.");
            }
            keys.add(new SortKey(col, order));
            idx++;
        }
        return keys;
    }

    private static String getStr(Map<String, Object> row, String key, boolean required, int idx) {
        Object o = row.get(key) != null ? row.get(key) : row.get(key.toLowerCase(Locale.ROOT));
        String s = o == null ? "" : o.toString().trim();
        if (required && s.isEmpty()) {
            throw new BotCommandException("Entrada " + idx + ": '" + key + "' no puede estar vacío.");
        }
        return s;
    }

    private static class SortKey {
        String column; String order;
        SortKey(String column, String order){ this.column=column; this.order=order; }
    }
}