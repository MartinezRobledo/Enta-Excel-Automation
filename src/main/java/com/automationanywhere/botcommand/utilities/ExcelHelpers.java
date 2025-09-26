package com.automationanywhere.botcommand.utilities;

import com.automationanywhere.botcommand.data.Value;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.*;

public class ExcelHelpers {


    private static final int xlUp       = -4162;
    private static final int xlToLeft   = -4159;

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

    /** ÍNDICE de la última columna con datos reales. 0 si no hay datos. */
    public static int getLastDataColumn(Dispatch sheet) {
        Dispatch used = Dispatch.get(sheet, "UsedRange").toDispatch();
        if (used == null || used.m_pDispatch == 0) return 0;

        int usedFirstRow = Dispatch.get(used, "Row").getInt();
        int usedFirstCol = Dispatch.get(used, "Column").getInt();
        int usedRows     = Dispatch.get(Dispatch.get(used, "Rows").toDispatch(), "Count").getInt();
        int usedCols     = Dispatch.get(Dispatch.get(used, "Columns").toDispatch(), "Count").getInt();
        if (usedRows <= 0 || usedCols <= 0) return 0;

        int lastPossibleCol = usedFirstCol + usedCols - 1;
        int lastDataCol = 0;

        // Escanear cada fila del UsedRange y tomar el máximo End(xlToLeft).Column
        for (int r = usedFirstRow; r <= usedFirstRow + usedRows - 1; r++) {
            Dispatch right = Dispatch.call(sheet, "Cells", r, lastPossibleCol).toDispatch();
            Dispatch lastInRow = Dispatch.call(right, "End", new Variant(xlToLeft)).toDispatch();
            int colInRow = Dispatch.get(lastInRow, "Column").getInt();
            if (colInRow > lastDataCol) lastDataCol = colInRow;
        }
        return (lastDataCol < usedFirstCol) ? 0 : lastDataCol;
    }

    /** CANTIDAD de filas con datos (desde la primera fila del UsedRange hasta la última fila con datos). */
    public static int getDataRowCount(Dispatch sheet) {
        Dispatch used = Dispatch.get(sheet, "UsedRange").toDispatch();
        if (used == null || used.m_pDispatch == 0) return 0;
        int usedFirstRow = Dispatch.get(used, "Row").getInt();
        int lastDataRow = getLastDataRow(sheet);
        if (lastDataRow == 0 || lastDataRow < usedFirstRow) return 0;
        return lastDataRow - usedFirstRow + 1;
    }

    /** (Opcional) CANTIDAD de filas con datos desde una fila de header (incluyéndola o no). */
    public static int getDataRowCountFromHeader(Dispatch sheet, int headerRow, boolean includeHeader) {
        int last = getLastDataRow(sheet);
        if (last == 0 || last < headerRow) return 0;
        return includeHeader ? (last - headerRow + 1) : (last - headerRow);
    }


    /**
     * Obtiene el número de filas con datos en una hoja de Excel
     * @param sheet Dispatch de la hoja
     * @return número de filas
     */
    public static int getLastRow(Dispatch sheet) {
        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        Dispatch rowsRange = Dispatch.get(usedRange, "Rows").toDispatch();
        return Dispatch.get(rowsRange, "Count").getInt();
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
     * Filtra filas de un sheet por una o más columnas y criterios múltiples.
     *
     * @param sheet        Dispatch de la hoja
     * @param columns      Lista de columnas a filtrar (puede ser letra A-Z o nombre de header)
     * @param criteriaMap  Mapa: columna -> lista de valores aceptados
     * @return Listado de filas filtradas como listas de strings
     */
    public static List<List<String>> filterRows(Dispatch sheet, List<String> columns, Map<String, List<String>> criteriaMap) {
        List<List<String>> result = new ArrayList<>();

        // Obtener UsedRange
        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        int rowCount = getLastRow(sheet);
        int colCount = getLastColumn(sheet);

        // Mapear nombres de headers a índices
        Map<String, Integer> headerIndexMap = new HashMap<>();
        Dispatch headerRow = Dispatch.call(usedRange, "Rows", 1).toDispatch();
        for (int c = 1; c <= colCount; c++) {
            Dispatch cell = Dispatch.call(headerRow, "Cells", 1, c).toDispatch();
            String header = Dispatch.get(cell, "Value").toString();
            headerIndexMap.put(header.trim(), c);
        }

        // Determinar índices de columnas a filtrar
        List<Integer> filterIndices = new ArrayList<>();
        for (String col : columns) {
            if (headerIndexMap.containsKey(col)) {
                filterIndices.add(headerIndexMap.get(col));
            } else {
                // Asumimos que es letra de columna
                int index = colLetterToIndex(col);
                if (index <= colCount) {
                    filterIndices.add(index);
                }
            }
        }

        // Recorrer filas (desde fila 2, porque fila 1 son headers)
        for (int r = 2; r <= rowCount; r++) {
            boolean match = true;
            for (int i = 0; i < filterIndices.size(); i++) {
                int colIndex = filterIndices.get(i);
                Dispatch cell = Dispatch.call(usedRange, "Cells", r, colIndex).toDispatch();
                String value = Dispatch.get(cell, "Value").toString();

                List<String> allowed = criteriaMap.get(columns.get(i));
                if (allowed != null && !allowed.contains(value)) {
                    match = false;
                    break;
                }
            }
            if (match) {
                // Guardar toda la fila
                List<String> row = new ArrayList<>();
                for (int c = 1; c <= colCount; c++) {
                    Dispatch cell = Dispatch.call(usedRange, "Cells", r, c).toDispatch();
                    Object val = Dispatch.get(cell, "Value").toJavaObject();
                    row.add(val != null ? val.toString() : "");
                }
                result.add(row);
            }
        }

        return result;
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

    public static Map<String, List<String>> parseFilterCriteria(List<Value> entryList) {
        Map<String, List<String>> map = new HashMap<>();
        if (entryList == null) return map;

        for (Value v : entryList) {
            String json = v.toString(); // cada Value viene como "Column:Criteria"
            String[] parts = json.split(":", 2);
            if (parts.length == 2) {
                String key = parts[0].trim();
                String[] values = parts[1].split(";");
                List<String> list = new ArrayList<>();
                for (String val : values) list.add(val.trim());
                map.put(key, list);
            }
        }
        return map;
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

}



