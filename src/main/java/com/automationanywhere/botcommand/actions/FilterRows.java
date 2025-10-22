package com.automationanywhere.botcommand.actions;

import java.util.*;
import java.util.stream.Collectors;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListAddButtonLabel;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListEmptyLabel;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListEntryUnique;
import com.automationanywhere.commandsdk.annotations.rules.EntryList.EntryListLabel;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
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
            @SelectModes
            String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String originSheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NotEmpty
            Double originSheetIndex,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Por letra (A,B,C...)", value = "letter")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Por nombre de encabezado", value = "header"))
            })
            @Pkg(label = "Referencia de columna", default_value = "letter", default_value_type = DataType.STRING)
            String referenceMode,

            @Idx(index = "3.2.1", type = AttributeType.CHECKBOX)
            @Pkg(label = "Aplicar TRIM a encabezados (solo si usa nombre)", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean trimHeaders,

            @Idx(index = "4", type = AttributeType.TEXT)
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
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

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

            // --- Obtener columna y fieldIndex (igual que tu código base) ---
            String colKeyOriginal = e.getKey();
            String keyNorm = colKeyOriginal.trim().toLowerCase();

            int fieldIndex;
            int absCol; // columna absoluta 1-based en la hoja

            if ("letter".equalsIgnoreCase(referenceMode)) {
                absCol = excelColumnLetterToNumber(colKeyOriginal);
                if (absCol < headerStartCol || absCol > headerEndCol) {
                    throw new BotCommandException(
                            "Column letter '" + colKeyOriginal + "' (col " + absCol + ") está fuera del headers range [" +
                                    columnNumberToLetter(headerStartCol) + ":" + columnNumberToLetter(headerEndCol) + "]."
                    );
                }
                fieldIndex = (absCol - headerStartCol) + 1; // 1-based relativo al rango
            } else { // referenceMode = "header"
                fieldIndex = headerToIndex.getOrDefault(keyNorm, -1);
                if (fieldIndex <= 0) {
                    String available = String.join(", ",
                            headerToIndex.keySet().stream().sorted().collect(Collectors.toList()));
                    throw new BotCommandException(
                            "Header '" + colKeyOriginal + "' no encontrado en la fila " + headerRow +
                                    ". Headers disponibles en el rango: [" + available + "]."
                    );
                }
                absCol = headerStartCol + fieldIndex - 1; // convertir fieldIndex relativo a columna absoluta
            }

            // --- Lista de criterios para esta columna ---
            List<String> criteriaList = e.getValue();

            // --- AND dentro de un único string de criterio ---
            if (criteriaList.size() == 1 && containsAnd(criteriaList.get(0))) {
                String andExpr = criteriaList.get(0);

                // Parsear condiciones (tu parse debe llenar op, rhs, rhsNum)
                List<SimpleCondition> conditions = parseAndConditions(andExpr);
                if (conditions.isEmpty()) {
                    throw new BotCommandException("No se detectaron condiciones válidas en: " + andExpr);
                }

                // 1) EXACTAMENTE 2 condiciones -> usar AND nativo de AutoFilter
                if (conditions.size() == 2) {
                    String crit1 = toExcelCriteriaString(conditions.get(0)); // ej "<>0"
                    String crit2 = toExcelCriteriaString(conditions.get(1)); // ej "<>'Impuestos'"

                    // xlAnd = 1
                    Dispatch.callN(
                            dataRange,
                            "AutoFilter",
                            new Object[]{
                                    new Variant(fieldIndex),
                                    new Variant(crit1),
                                    new Variant(1),      // xlAnd
                                    new Variant(crit2)
                            }
                    );
                    continue;
                }

                // 2) 3 o más condiciones -> AND sintético (materializa valores que cumplen TODAS)
                if (conditions.size() >= 3) {
                    LinkedHashSet<String> displayKeys = new LinkedHashSet<>();
                    for (int r = headerRow + 1; r <= lastRow; r++) {
                        Dispatch cell = Dispatch.call(sheet, "Cells", r, absCol).toDispatch();

                        // Evaluar condiciones con VALUE (mantiene tipos reales)
                        Variant vVal = Dispatch.get(cell, "Value");
                        Object valObj = (vVal == null || vVal.isNull()) ? null : vVal.toJavaObject();
                        if (matchesAll(conditions, valObj)) {
                            // Para filtrar por lista, usar el TEXTO mostrado (lo que ve el usuario)
                            String shown = safeVariantToString(Dispatch.get(cell, "Text"));
                            displayKeys.add(shown);
                        }
                    }

                    if (displayKeys.isEmpty()) {
                        // Fuerza "sin resultados"
                        Dispatch.call(dataRange, "AutoFilter",
                                new Variant(fieldIndex), new Variant("#__NO_MATCH__"));
                    } else {
                        // xlFilterValues = 7, Criteria1 debe ser SafeArray de VARIANT (strings visibles)
                        Variant vArray = buildVariantArrayFromStrings(displayKeys);
                        Dispatch.callN(
                                dataRange,
                                "AutoFilter",
                                new Object[]{ new Variant(fieldIndex), vArray, new Variant(7) }
                        );
                    }
                    continue;
                }

                // 3) Fallback improbable (solo 1 condición detectada)
                String only = toExcelCriteriaString(conditions.get(0));
                Dispatch.call(dataRange, "AutoFilter",
                        new Variant(fieldIndex), new Variant(only));
                continue;
            }

            // --- LÓGICA EXISTENTE ---
            if (criteriaList.size() == 1) {
                // Un único criterio simple (ej. "<>0" o "<>Impuestos")
                Dispatch.call(dataRange, "AutoFilter",
                        new Variant(fieldIndex),
                        new Variant(criteriaList.get(0)));
            } else {
                // Varios criterios -> OR (multi-select en Excel)
                Variant[] variants = new Variant[criteriaList.size()];
                for (int i = 0; i < criteriaList.size(); i++) {
                    variants[i] = new Variant(criteriaList.get(i));
                }
                // xlFilterValues = 7
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

    // --- Helpers para AND sintético ---

    private static class SimpleCondition {
        String op;      // =, <>, >, >=, <, <=
        String rhs;     // literal como texto (sin comillas)
        Double rhsNum;  // si es numérico, valor parseado; si no, null
    }

    private static List<SimpleCondition> parseAndConditions(String expr) {
        // Admite && como separador de AND. También tolera " AND " convirtiéndolo a && si querés extender.
        String normalized = expr.replaceAll("\\s+AND\\s+", "&&");
        String[] parts = normalized.split("\\&\\&");
        List<SimpleCondition> out = new ArrayList<>();
        for (String raw : parts) {
            String s = raw.trim();
            if (s.isEmpty()) continue;
            out.add(parseCondition(s));
        }
        return out;
    }

    private static SimpleCondition parseCondition(String s) {
        // Soporta operadores: >=, <=, <>, =, >, <
        String[] ops = new String[] {">=", "<=", "<>", "=", ">", "<"};
        String op = null;
        int pos = -1;
        for (String o : ops) {
            pos = indexOfOp(s, o);
            if (pos >= 0) { op = o; break; }
        }
        if (op == null) {
            throw new BotCommandException("Condición inválida (se esperaba =, <>, >, >=, <, <=): " + s);
        }
        String rhs = s.substring(pos + op.length()).trim();

        // Quitar comillas simples o dobles si envuelven el literal ('...' o "...")
        if ((rhs.startsWith("'") && rhs.endsWith("'")) || (rhs.startsWith("\"") && rhs.endsWith("\""))) {
            rhs = rhs.substring(1, rhs.length() - 1);
        }

        SimpleCondition c = new SimpleCondition();
        c.op = op;
        c.rhs = rhs;
        c.rhsNum = toDoubleOrNull(rhs);
        return c;
    }

    private static int indexOfOp(String s, String op) {
        // Busca el operador respetando espacios arbitrarios (ej.: "<>   'Impuestos'")
        String trim = s.trim();
        // Búsqueda directa; si querés más robustez, podés normalizar espacios.
        return trim.indexOf(op);
    }

    private static Double toDoubleOrNull(String s) {
        if (s == null) return null;
        String t = s.trim();
        if (t.isEmpty()) return null;
        try {
            // tolera coma decimal convirtiéndola a punto
            return Double.valueOf(t.replace(',', '.'));
        } catch (NumberFormatException ex) {
            return null;
        }
    }

    private static boolean matchesAll(List<SimpleCondition> conditions, Object valObj) {
        for (SimpleCondition c : conditions) {
            if (!matches(c, valObj)) return false;
        }
        return true;
    }

    private static boolean matches(SimpleCondition c, Object valObj) {
        String cellStr = (valObj == null) ? "" : valObj.toString();
        Double cellNum = toDoubleOrNull(cellStr);

        switch (c.op) {
            case "<>":
                if (c.rhsNum != null && cellNum != null) return Double.compare(cellNum, c.rhsNum) != 0;
                return !cellStr.equals(c.rhs);
            case "=":
                if (c.rhsNum != null && cellNum != null) return Double.compare(cellNum, c.rhsNum) == 0;
                return cellStr.equals(c.rhs);
            case ">":
                if (c.rhsNum != null && cellNum != null) return cellNum > c.rhsNum;
                return cellStr.compareTo(c.rhs) > 0;
            case ">=":
                if (c.rhsNum != null && cellNum != null) return cellNum >= c.rhsNum;
                return cellStr.compareTo(c.rhs) >= 0;
            case "<":
                if (c.rhsNum != null && cellNum != null) return cellNum < c.rhsNum;
                return cellStr.compareTo(c.rhs) < 0;
            case "<=":
                if (c.rhsNum != null && cellNum != null) return cellNum <= c.rhsNum;
                return cellStr.compareTo(c.rhs) <= 0;
            default:
                throw new BotCommandException("Operador no soportado: " + c.op);
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

    // --- Helpers de parsing y evaluación ---

    private static boolean containsAnd(String s) {
        if (s == null) return false;
        String t = s.replace("&amp;", "&"); // por si vino HTML-encoded
        return t.contains("&&") || t.matches(".*\\sAND\\s.*") || t.contains(";");
    }


    private static Double tryToNumber(Object o) {
        if (o == null) return null;
        if (o instanceof Number) return ((Number) o).doubleValue();
        String s = o.toString().trim().replace(",", "."); // opcional: manejar locales
        if (s.isEmpty()) return null;
        try { return Double.parseDouble(s); } catch (Exception e) { return null; }
    }

    // --- Utilidad para armar Variant conservando tipo ---
    private static final class AutomationUtils {
        static Variant toVariantKeepingType(Object o) {
            if (o == null) return new Variant("");
            if (o instanceof Boolean) return new Variant((Boolean) o);
            if (o instanceof Integer) return new Variant((Integer) o);
            if (o instanceof Long)    return new Variant((Long) o);
            if (o instanceof Double)  return new Variant((Double) o);
            if (o instanceof Float)   return new Variant(((Float) o).doubleValue());
            // Si te llegan Date (COM DATE), podrías mapearlas según tu experiencia con JACOB.
            // Por ahora, lo pasamos como String si no es tipo básico:
            return new Variant(o.toString());
        }
    }

    // Convierte una SimpleCondition a la cadena de criterio que espera Excel.
    // Reglas: rhsNum != null -> número sin comillas; rhs textual -> comillas simples si no están.
    private static String toExcelCriteriaString(SimpleCondition c) {
        String op = c.op;
        if (c.rhsNum != null) {
            return op + c.rhsNum;  // números sin comillas
        }
        String s = (c.rhs == null) ? "" : c.rhs;
        boolean quoted = (s.startsWith("'") && s.endsWith("'")) || (s.startsWith("\"") && s.endsWith("\""));
        return op + (quoted ? s : "'" + s + "'");
    }

    // Crea un SafeArray(VT_VARIANT) a partir de una colección de strings (display values)
    // y lo envuelve en Variant, listo para pasar a AutoFilter con xlFilterValues.
    private static Variant buildVariantArrayFromStrings(Collection<String> items) {
        SafeArray sa = new SafeArray(Variant.VariantVariant, items.size());
        int i = 0;
        for (String s : items) {
            sa.setVariant(i++, new Variant(s == null ? "" : s));
        }
        return new Variant(sa);
    }

}
