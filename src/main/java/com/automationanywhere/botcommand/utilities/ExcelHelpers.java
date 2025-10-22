package com.automationanywhere.botcommand.utilities;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.*;

public class ExcelHelpers {


    // --- Límites reales de Excel (XFD = 16384; última fila = 1_048_576) ---
    public static final int EXCEL_MAX_ROWS = 1_048_576;
    public static final int EXCEL_MAX_COLS = 16_384;


    public static final int xlUp       = -4162;
    public static final int xlCalculationAutomatic = -4105;
    public static final int xlCalculationManual    = -4135;


    // ------------------------------------------------------------
    // A1Ref: referencia A1 con soporte de absolutos ($)
    // ------------------------------------------------------------
    public static final class A1Ref {
        public final int row;      // 1-based
        public final int col;      // 1-based
        public final boolean absRow;
        public final boolean absCol;

        public A1Ref(int row, int col, boolean absRow, boolean absCol) {
            if (row < 1 || row > EXCEL_MAX_ROWS)
                throw new BotCommandException("Row out of range for Excel: " + row);
            if (col < 1 || col > EXCEL_MAX_COLS)
                throw new BotCommandException("Column out of range for Excel: " + col);
            this.row = row;
            this.col = col;
            this.absRow = absRow;
            this.absCol = absCol;
        }

        @Override public String toString() { return ExcelHelpers.buildA1(this); }

        @Override public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof A1Ref)) return false;
            A1Ref that = (A1Ref) o;
            return row == that.row && col == that.col && absRow == that.absRow && absCol == that.absCol;
        }
        @Override public int hashCode() {
            int result = row;
            result = 31 * result + col;
            result = 31 * result + (absRow ? 1 : 0);
            result = 31 * result + (absCol ? 1 : 0);
            return result;
        }
    }


    // ------------------------------------------------------------
    // parseA1: "$C$9" | "C$9" | "$C9" | "C9"  -> A1Ref (1-based)
    // ------------------------------------------------------------
    public static A1Ref parseA1(String a1) {
        if (a1 == null) throw new BotCommandException("A1 reference is null.");
        String s = a1.trim().toUpperCase();
        if (s.isEmpty()) throw new BotCommandException("A1 reference is empty.");

        int i = 0;
        boolean absCol = false, absRow = false;

        // Optional $ before column
        if (i < s.length() && s.charAt(i) == '$') { absCol = true; i++; }

        // Column letters (A..Z)
        int colStart = i;
        while (i < s.length()) {
            char ch = s.charAt(i);
            if (ch >= 'A' && ch <= 'Z') i++; else break;
        }
        if (i == colStart) {
            throw new BotCommandException("Invalid A1: missing column letters in '" + a1 + "'.");
        }
        String colLetters = s.substring(colStart, i);

        // Optional $ before row
        if (i < s.length() && s.charAt(i) == '$') { absRow = true; i++; }

        // Row digits
        int rowStart = i;
        while (i < s.length()) {
            char ch = s.charAt(i);
            if (ch >= '0' && ch <= '9') i++; else break;
        }
        if (i == rowStart) {
            throw new BotCommandException("Invalid A1: missing row digits in '" + a1 + "'.");
        }
        String rowDigits = s.substring(rowStart, i);

        // No trailing garbage
        if (i != s.length()) {
            throw new BotCommandException("Invalid A1: trailing characters in '" + a1 + "'.");
        }

        int col = excelColumnLetterToNumber(colLetters);
        int row;
        try {
            row = Integer.parseInt(rowDigits);
        } catch (NumberFormatException nfe) {
            throw new BotCommandException("Invalid A1 row number in '" + a1 + "'.");
        }

        return new A1Ref(row, col, absRow, absCol);
    }

    // ------------------------------------------------------------
    // buildA1: arma "A1" con $ según flags de A1Ref
    // ------------------------------------------------------------
    public static String buildA1(A1Ref r) {
        String colLetters = numberToColumnLetter(r.col);
        StringBuilder sb = new StringBuilder();
        if (r.absCol) sb.append('$');
        sb.append(colLetters);
        if (r.absRow) sb.append('$');
        sb.append(r.row);
        return sb.toString();
    }

    /**
     * Última fila (índice 1-based) con datos o fórmula en la columna (por índice).
     * Devuelve 0 si la columna está completamente vacía.
     *
     * Robusto: NO usa UsedRange. Sube con End(xlUp) desde la última fila de Excel.
     */
    public static int getLastDataRowInColumn(Dispatch sheet, int columnIndex) {
        if (sheet == null || sheet.m_pDispatch == 0) return 0;
        if (columnIndex < 1 || columnIndex > EXCEL_MAX_COLS)
            throw new BotCommandException("Column out of range: " + columnIndex);

        // Ir al fondo de la hoja en esa columna y subir con End(xlUp)
        Dispatch bottom = Dispatch.call(sheet, "Cells", EXCEL_MAX_ROWS, columnIndex).toDispatch();
        Dispatch lastInCol = Dispatch.call(bottom, "End", new Variant(xlUp)).toDispatch();
        int row = Dispatch.get(lastInCol, "Row").getInt();

        // Si cayó en fila 1, verificar si realmente hay algo en (1, col)
        Dispatch cell = Dispatch.call(sheet, "Cells", row, columnIndex).toDispatch();
        boolean hasFormula = false;
        try { hasFormula = Dispatch.get(cell, "HasFormula").getBoolean(); } catch (Exception ignore) {}
        Variant v = Dispatch.get(cell, "Value");
        boolean emptyValue = (v == null || v.isNull() || v.toString().trim().isEmpty());

        // Columna vacía → 0
        if (row == 1 && emptyValue && !hasFormula) return 0;

        return row;
    }

    /**
     * ÍNDICE de la última fila con datos (constantes o fórmulas) en la columna dada por letra (A, B, ..., AA, ...).
     * Devuelve 0 si la columna está completamente vacía.
     */
    public static int getLastDataRowInColumn(Dispatch sheet, String columnLetter) {
        if (columnLetter == null || columnLetter.trim().isEmpty()) return 0;
        int col = colLetterToIndex(columnLetter.trim());
        return getLastDataRowInColumn(sheet, col);
    }


    public static int getNumberOfRows(Dispatch sheet) {
        if (sheet == null || sheet.m_pDispatch == 0) return -1;

        final int xlCellTypeVisible = 12;

        // Tomamos UsedRange como base
        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        if (usedRange == null || usedRange.m_pDispatch == 0) return -1;

        // Tomar SOLO la primera columna del UsedRange (evita problemas con rangos no contiguos)
        Dispatch firstCol = Dispatch.call(usedRange, "Columns", new Variant(1)).toDispatch();

        // Filas visibles en esa columna
        Dispatch visibleCol = Dispatch.call(firstCol, "SpecialCells", new Variant(xlCellTypeVisible)).toDispatch();

        // Contar celdas visibles (equivale a filas visibles)
        int count = Dispatch.get(visibleCol, "Count").getInt();

        // Si hay AutoFilter, restar 1 (encabezado)
        Dispatch autoFilter = Dispatch.get(sheet, "AutoFilter").toDispatch();
        if (autoFilter != null && autoFilter.m_pDispatch != 0 && count > 0) {
            count -= 1;
        }

        return count;
    }


    /** Intenta obtener la fila de header desde AutoFilter, o sino desde la primera tabla, o sino el inicio de UsedRange. */
    private static int getHeaderRow(Dispatch sheet) {
        try {
            Dispatch autoFilter = Dispatch.get(sheet, "AutoFilter").toDispatch();
            Dispatch rng = Dispatch.get(autoFilter, "Range").toDispatch();
            return Dispatch.get(rng, "Row").getInt();
        } catch (Exception ignore) {}

        try {
            Dispatch listObjects = Dispatch.get(sheet, "ListObjects").toDispatch();
            int count = Dispatch.get(listObjects, "Count").getInt();
            int min = Integer.MAX_VALUE;
            for (int i = 1; i <= count; i++) {
                Dispatch lo = Dispatch.call(listObjects, "Item", i).toDispatch();
                Dispatch hdr = Dispatch.get(lo, "HeaderRowRange").toDispatch();
                int r = Dispatch.get(hdr, "Row").getInt();
                if (r < min) min = r;
            }
            if (min != Integer.MAX_VALUE) return min;
        } catch (Exception ignore) {}

        try {
            Dispatch used = Dispatch.get(sheet, "UsedRange").toDispatch();
            return Dispatch.get(used, "Row").getInt();
        } catch (Exception ignore) {}

        return 0;
    }

    /** Devuelve true si el header está dentro de las celdas visibles actuales. */
    private static boolean headerIsVisible(Dispatch sheet, Dispatch used, Dispatch visible, Dispatch app) {
        int headerRow = getHeaderRow(sheet);
        if (headerRow <= 0) return false;
        try {
            Dispatch headerRowRange = Dispatch.call(sheet, "Rows", headerRow).toDispatch();
            Dispatch visHeader = Dispatch.call(app, "Intersect", visible, headerRowRange).toDispatch();
            return visHeader != null && visHeader.m_pDispatch != 0;
        } catch (Exception e) {
            return false;
        }
    }


    /** ÍNDICE de la última fila con datos reales (constantes o fórmulas). 0 si no hay datos. */
    public static int getLastDataRow(Dispatch sheet) {
        Dispatch used = Dispatch.get(sheet, "UsedRange").toDispatch();
        if (used == null || used.m_pDispatch == 0) return 0;

        int usedFirstRow = Dispatch.get(used, "Row").getInt();
        int usedFirstCol = Dispatch.get(used, "Column").getInt();
        int usedRows     = Dispatch.get(Dispatch.get(used, "Rows").toDispatch(), "Count").getInt();
        int usedCols     = Dispatch.get(Dispatch.get(used, "Columns").toDispatch(), "Count").getInt();
        if (usedRows <= 0 || usedCols <= 0) return 0;

        int lastPossibleRow = usedFirstRow + usedRows - 1;
        int lastDataRow = 0;

        // Escanear cada columna del UsedRange y tomar el máximo End(xlUp).Row
        for (int c = usedFirstCol; c <= usedFirstCol + usedCols - 1; c++) {
            Dispatch bottom = Dispatch.call(sheet, "Cells", lastPossibleRow, c).toDispatch();
            Dispatch lastInCol = Dispatch.call(bottom, "End", new Variant(xlUp)).toDispatch();
            int rowInCol = Dispatch.get(lastInCol, "Row").getInt();
            if (rowInCol > lastDataRow) lastDataRow = rowInCol;
        }
        // Si el sheet está vacío, Excel puede devolver usedFirstRow aun sin datos "reales"
        return (lastDataRow < usedFirstRow) ? 0 : lastDataRow;
    }

    /**
     * Obtiene el número de columnas con datos en una hoja de Excel
     * @param sheet Dispatch de la hoja
     * @return número de columnas
     */
    public static int getLastColumn(Dispatch sheet) {
        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        Dispatch colsRange = Dispatch.get(usedRange, "Columns").toDispatch();
        return Dispatch.get(colsRange, "Count").getInt();
    }


    /**
     * Convierte letra de columna a índice (A=1, B=2, ...)
     */
    public static int colLetterToIndex(String col) {
        col = col.toUpperCase();
        int index = 0;
        for (int i = 0; i < col.length(); i++) {
            index = index * 26 + (col.charAt(i) - 'A' + 1);
        }
        return index;
    }


    public static String numberToColumnLetter(int col) {
        StringBuilder sb = new StringBuilder();
        while (col > 0) {
            int rem = (col - 1) % 26;
            sb.insert(0, (char) ('A' + rem));
            col = (col - 1) / 26;
        }
        return sb.toString();
    }

    // Convierte letras de columna → número (A=1, B=2, ...)
    private static int colLetterToNumber(String col) {
        int res = 0;
        for (int i = 0; i < col.length(); i++) {
            res = res * 26 + (col.charAt(i) - 'A' + 1);
        }
        return res;
    }

    // Convierte número → letras (1=A, 2=B, ...)
    private static String colNumberToLetter(int num) {
        StringBuilder sb = new StringBuilder();
        while (num > 0) {
            int rem = (num - 1) % 26;
            sb.insert(0, (char) ('A' + rem));
            num = (num - 1) / 26;
        }
        return sb.toString();
    }

    // Divide un rango en sub-rangos excluyendo columnas ignoradas
    public static List<String> splitRangeByIgnoredColumns(String fullRange, List<String> ignoreCols) {
        List<String> result = new ArrayList<>();

        // Ej: fullRange = "B3:G40"
        String[] parts = fullRange.split(":");
        if (parts.length != 2) return Collections.singletonList(fullRange);

        String startCell = parts[0].toUpperCase();
        String endCell = parts[1].toUpperCase();

        int startCol = colLetterToNumber(startCell.replaceAll("\\d", ""));
        int startRow = Integer.parseInt(startCell.replaceAll("\\D", ""));
        int endCol = colLetterToNumber(endCell.replaceAll("\\d", ""));
        int endRow = Integer.parseInt(endCell.replaceAll("\\D", ""));

        // Pasar columnas a ignorar a números
        Set<Integer> ignoreSet = new HashSet<>();
        for (String col : ignoreCols) {
            if (col != null && !col.trim().isEmpty()) {
                ignoreSet.add(colLetterToNumber(col.trim().toUpperCase()));
            }
        }

        int col = startCol;
        while (col <= endCol) {
            // Saltar columnas ignoradas
            while (col <= endCol && ignoreSet.contains(col)) {
                col++;
            }
            if (col > endCol) break;

            int blockStart = col;
            // Avanzar hasta la siguiente ignorada o final
            while (col <= endCol && !ignoreSet.contains(col)) {
                col++;
            }
            int blockEnd = col - 1;

            // Armar sub-rango
            String subRange = colNumberToLetter(blockStart) + startRow + ":" +
                    colNumberToLetter(blockEnd) + endRow;
            result.add(subRange);
        }

        return result;
    }

    public static int excelColumnLetterToNumber(String col) {
        int res = 0; col = col.toUpperCase();
        for (int i = 0; i < col.length(); i++) res = res * 26 + (col.charAt(i) - 'A' + 1);
        return res;
    }
    public static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return (o != null) ? o.toString() : "";
    }
    public static boolean getBool(Dispatch app, String prop) {
        try { return Dispatch.get(app, prop).getBoolean(); } catch (Exception e) { return true; }
    }
    public static int getInt(Dispatch app, String prop) {
        try { return Dispatch.get(app, prop).getInt(); } catch (Exception e) { return xlCalculationAutomatic; }
    }
    public static void putBool(Dispatch app, String prop, boolean v) {
        try { Dispatch.put(app, prop, v); } catch (Exception ignore) {}
    }
    public static void putInt(Dispatch app, String prop, int v) {
        try { Dispatch.put(app, prop, new Variant(v)); } catch (Exception ignore) {}
    }

    //Convertir Header Name en Column Index
    public static int headerNameToColumnIndex(Dispatch sheet, String columnName, int firstRow, int colsCnt){
        if (columnName == null || columnName.isEmpty())
            throw new BotCommandException("Column header not provided.");
        int colIndex = -1;
        String target = columnName.trim();
        for (int c = 1; c <= colsCnt; c++) {
            Dispatch hdrCell = Dispatch.call(sheet, "Cells", firstRow, c).toDispatch();
            String hdr = safeVariantToString(Dispatch.get(hdrCell, "Value"));
            if (hdr != null && hdr.trim().equalsIgnoreCase(target)) { colIndex = c; break; }
        }
        if (colIndex == -1) throw new BotCommandException("Header not found: " + target);
        return colIndex;
    }

    /** Incrementa la COLUMNA de una dirección A1, preservando los '$' si existen.
     *  Ej: incrementColumnInA1("B3", 2) -> "D3"
     *      incrementColumnInA1("$B$3", 2) -> "$D$3"
     */
    public static String incrementColumnInA1(String address, int step) {
        if (step < 0) throw new BotCommandException("Step (columns) must be >= 0.");

        A1Ref r = parseA1(address);
        long newCol = (long) r.col + step;                 // r.col ya es int (1-based)
        if (newCol < 1 || newCol > EXCEL_MAX_COLS)
            throw new BotCommandException("Column exceeds Excel limit (XFD / 16384).");

        A1Ref r2 = new A1Ref(r.row, (int) newCol, r.absRow, r.absCol); // nueva ref
        return buildA1(r2);
    }

    /** Incrementa la FILA de una dirección A1, preservando los '$' si existen.
     *  Ej: incrementRowInA1("A4", 4) -> "A8"
     *      incrementRowInA1("$A$4", 4) -> "$A$8"
     */
    public static String incrementRowInA1(String address, int step) {
        if (step < 0) throw new BotCommandException("Step (rows) must be >= 0.");

        A1Ref r = parseA1(address);
        long newRow = (long) r.row + step; // usar long para prevenir overflow intermedio

        if (newRow < 1 || newRow > EXCEL_MAX_ROWS) {
            throw new BotCommandException("Row exceeds Excel limit (1,048,576).");
        }

        // >>> Crear NUEVA instancia (A1Ref es inmutable)
        A1Ref r2 = new A1Ref((int) newRow, r.col, r.absRow, r.absCol);
        return buildA1(r2);
    }

    // Puedes pegarlo dentro de ConvertColumnToNumber (como private static)
// o mover a ExcelHelpers si preferís.
    public static int countTextNumbersInRange(Dispatch sheet, int startRow, int endRow, int colIndex) {
        if (endRow < startRow) return 0;

        // 1) Construir el rango [startRow..endRow] en colIndex
        Dispatch start = Dispatch.call(sheet, "Cells", startRow, colIndex).toDispatch();
        Dispatch end   = Dispatch.call(sheet, "Cells", endRow,   colIndex).toDispatch();
        Dispatch rng   = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]{ start, end }, new int[1]).toDispatch();

        // 2) Obtener Application y la Address A1 del rango (absoluta)
        Dispatch app   = Dispatch.get(sheet, "Application").toDispatch();
        String addr    = Dispatch.get(rng, "Address").toString(); // ej: $B$2:$B$4098

        // 3) Intento 1: Evaluate con funciones en inglés y separador coma
        //    Cuenta celdas que SON TEXTO y que al aplicar VALUE() resultan NUMÉRICAS.
        String f1 = "SUMPRODUCT(--ISTEXT(" + addr + "),--ISNUMBER(VALUE(" + addr + ")))";

        try {
            // Evaluate devuelve Variant; redondeamos a int
            double d = Dispatch.call(app, "Evaluate", f1).getDouble();
            return (int)Math.round(d);
        } catch (Exception ignore) {
            // 4) Intento 2: misma fórmula con separador ';' (locales que no aceptan coma)
            String f2 = "SUMPRODUCT(--ISTEXT(" + addr + ");--ISNUMBER(VALUE(" + addr + ")))";
            try {
                double d = Dispatch.call(app, "Evaluate", f2).getDouble();
                return (int)Math.round(d);
            } catch (Exception ignore2) {
                // 5) Fallback conservador: no distingue “texto numérico” de “texto no numérico”.
                //    Aun así, sirve para detectar conversión incompleta a grandes rasgos.
                Dispatch wf = Dispatch.get(app, "WorksheetFunction").toDispatch();
                int nonEmpty = (int)Math.round(Dispatch.call(wf, "CountA", rng).getDouble());
                int numeric  = (int)Math.round(Dispatch.call(wf, "Count",  rng).getDouble());
                int textish  = Math.max(0, nonEmpty - numeric);
                return textish;
            }
        }
    }

}