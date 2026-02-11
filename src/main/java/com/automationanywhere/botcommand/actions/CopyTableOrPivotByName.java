package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Copy Table/Pivot (full / values / rowLabels / columnLabels)",
        name = "copyTableOrPivotByName",
        description = "Copia como valores una Tabla o Pivot por nombre. Permite copiar todo, solo datos, solo etiquetas de filas o solo etiquetas de columnas.",
        icon = "excel.svg",
        return_type = DataType.STRING,
        return_required = true,
        return_label = "Destino ocupado (A1)",
        return_description = "Dirección A1 final ocupada en el destino"
)
public class CopyTableOrPivotByName {

    @Execute
    public Value action(
            // ===== ORIGEN =====
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session (origen)")
            @NotEmpty @com.automationanywhere.commandsdk.annotations.rules.SessionObject ExcelSession srcSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Tabla (ListObject)", value = "table")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Tabla Dinámica (PivotTable)", value = "pivot"))
            })
            @Pkg(label = "Tipo de objeto", default_value = "table", default_value_type = DataType.STRING)
            String objectType,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Nombre de la Tabla o Pivot")
            @NotEmpty String objectName,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Sheet origen (opcional): acelera la búsqueda")
            String sourceSheetOpt,

            // ===== DESTINO =====
            @Idx(index = "5", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session (destino)")
            @NotEmpty @com.automationanywhere.commandsdk.annotations.rules.SessionObject ExcelSession dstSession,

            @Idx(index = "6", type = AttributeType.TEXT)
            @Pkg(label = "Sheet destino")
            @NotEmpty String destSheetName,

            @Idx(index = "7", type = AttributeType.TEXT)
            @Pkg(label = "Celda destino (top-left), ej.: AN4190")
            @NotEmpty String destTopLeft,

            // ===== CONTENIDO =====
            @Idx(index = "8", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "8.1", pkg = @Pkg(label = "Full (completo)", value = "full")),
                    @Idx.Option(index = "8.2", pkg = @Pkg(label = "Values (solo datos)", value = "values")),
                    @Idx.Option(index = "8.3", pkg = @Pkg(label = "RowLabels (solo etiquetas de filas - Pivot)", value = "rowlabels")),
                    @Idx.Option(index = "8.4", pkg = @Pkg(label = "ColumnLabels (solo etiquetas de columnas - Pivot)", value = "columnlabels"))
            })
            @Pkg(label = "Contenido a copiar", default_value = "full", default_value_type = DataType.STRING)
            String contentMode,

            // ===== OPCIONES: TABLA =====
            @Idx(index = "9", type = AttributeType.CHECKBOX)
            @Pkg(label = "Si es Tabla: incluir header", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean includeTableHeader,

            @Idx(index = "10", type = AttributeType.CHECKBOX)
            @Pkg(label = "Si es Tabla: incluir TotalsRow", default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean includeTableTotals,

            // ===== OPCIONES: PIVOT =====
            @Idx(index = "11", type = AttributeType.CHECKBOX)
            @Pkg(label = "Pivot: incluir field headers (encabezados de campos)", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean pivotIncludeFieldHeaders,

            @Idx(index = "12", type = AttributeType.CHECKBOX)
            @Pkg(label = "Pivot (Values): excluir Grand Totals", default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean pivotExcludeGrandTotals,

            // ===== OPCIONES de RETORNO =====
            @Idx(index = "13", type = AttributeType.CHECKBOX)
            @Pkg(label = "Retornar con nombre de hoja", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean returnWithSheetName,

            @Idx(index = "14", type = AttributeType.CHECKBOX)
            @Pkg(label = "Retornar rango fijo ($)", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean returnAbsoluteRange
    ) {
        // --- Workbooks ---
        Session sSrc = ExcelObjects.requireSession(srcSession);
        Dispatch wbSrc = ExcelObjects.requireWorkbook(sSrc, srcSession);

        Session sDst = ExcelObjects.requireSession(dstSession);
        Dispatch wbDst = ExcelObjects.requireWorkbook(sDst, dstSession);

        boolean isTable = !"pivot".equalsIgnoreCase(objectType);
        if (objectName == null || objectName.trim().isEmpty()) {
            throw new BotCommandException("Debe indicar el nombre de la Tabla o Pivot.");
        }
        String mode = (contentMode == null) ? "full" : contentMode.trim().toLowerCase();

        // --- Resolver RANGO ORIGEN ---
        Dispatch srcRange;
        if (isTable) {
            // Modos permitidos en Tabla: full / values
            if ("rowlabels".equals(mode) || "columnlabels".equals(mode)) {
                throw new BotCommandException("rowLabels/columnLabels aplican solo a Pivot. En Tabla usá 'full' o 'values'.");
            }
            IncludeTableParts parts = resolveTableRange(
                    wbSrc, sourceSheetOpt, objectName,
                    Boolean.TRUE.equals(includeTableHeader),
                    Boolean.TRUE.equals(includeTableTotals),
                    mode
            );
            srcRange = parts.range;
        } else {
            // Pivot
            if ("full".equals(mode)) {
                srcRange = findPivotTableRange2(wbSrc, sourceSheetOpt, objectName); // TableRange2
            } else if ("values".equals(mode)) {
                Dispatch pvt = findPivotObject(wbSrc, sourceSheetOpt, objectName);
                Variant v = Dispatch.get(pvt, "DataBodyRange");
                if (v == null || v.isNull()) {
                    throw new BotCommandException("La Pivot no tiene área de datos (¿sin campos de valores?).");
                }
                srcRange = v.toDispatch();
                if (Boolean.TRUE.equals(pivotExcludeGrandTotals)) {
                    srcRange = shrinkPivotGrandTotalsIfAny(pvt, srcRange);
                }
            } else if ("rowlabels".equals(mode)) {
                srcRange = buildPivotRowLabelsRange(wbSrc, sourceSheetOpt, objectName,
                        Boolean.TRUE.equals(pivotIncludeFieldHeaders),
                        Boolean.TRUE.equals(pivotExcludeGrandTotals));

            } else if ("columnlabels".equals(mode)) {
                srcRange = buildPivotColumnLabelsRange(wbSrc, sourceSheetOpt, objectName,
                        Boolean.TRUE.equals(pivotIncludeFieldHeaders),
                        Boolean.TRUE.equals(pivotExcludeGrandTotals));
            } else {
                throw new BotCommandException("Contenido inválido: " + contentMode);
            }
        }

        int totalRows = count(srcRange, "Rows");
        int totalCols = count(srcRange, "Columns");
        if (totalRows <= 0 || totalCols <= 0) {
            throw new BotCommandException("El rango origen está vacío.");
        }

        // --- Destino ---
        Dispatch dstSheet = requireSheetByName(wbDst, destSheetName);
        Dispatch topLeft  = Dispatch.call(dstSheet, "Range", destTopLeft.trim()).toDispatch();
        Dispatch dstArea  = Dispatch.callN(topLeft, "Resize",
                new Variant[]{ new Variant(totalRows), new Variant(totalCols) }).toDispatch();

        // --- Copiar valores (sin portapapeles) ---
        Dispatch.put(dstArea, "Value2", Dispatch.get(srcRange, "Value2"));

        // --- A1 final ocupado (parametrizable) ---
        boolean includeSheet = (returnWithSheetName == null) ? true : returnWithSheetName.booleanValue();
        boolean absolute     = (returnAbsoluteRange == null) ? true : returnAbsoluteRange.booleanValue();

        // Address(RowAbsolute, ColumnAbsolute [, ReferenceStyle] [, External] [, RelativeTo])
        String a1 = Dispatch.callN(
                dstArea,
                "Address",
                new Variant[]{ new Variant(absolute), new Variant(absolute) }
        ).getString();

        if (includeSheet) {
            String sheetName = Dispatch.get(dstSheet, "Name").getString();
            // Excel: comillas simples alrededor del nombre; duplicar si contiene ' dentro
            String quotedSheet = "'" + sheetName.replace("'", "''") + "'";
            return new StringValue(quotedSheet + "!" + a1);
        } else {
            return new StringValue(a1);
        }

    }

    // ================== HELPERS ==================

    // -------- Sheet / Count / Coords --------
    private static Dispatch requireSheetByName(Dispatch wb, String sheetName) {
        try {
            Dispatch sheets = Dispatch.get(wb, "Worksheets").toDispatch();
            return Dispatch.call(sheets, "Item", sheetName).toDispatch();
        } catch (Exception e) {
            throw new BotCommandException("No existe la hoja '" + sheetName + "'.");
        }
    }
    private static Dispatch getSheetOrNull(Dispatch wb, String sheetName) {
        try {
            Dispatch sheets = Dispatch.get(wb, "Worksheets").toDispatch();
            return Dispatch.call(sheets, "Item", sheetName).toDispatch();
        } catch (Exception e) { return null; }
    }
    private static int count(Dispatch range, String rowsOrColumns) {
        return Dispatch.get(Dispatch.get(range, rowsOrColumns).toDispatch(), "Count").getInt();
    }
    private static int getRow(Dispatch range)    { return Dispatch.get(range, "Row").getInt(); }
    private static int getColumn(Dispatch range) { return Dispatch.get(range, "Column").getInt(); }

    // -------- PIVOT --------
    private static Dispatch findPivotObject(Dispatch wb, String sheetFilter, String pivotName) {
        Dispatch sh = (sheetFilter == null || sheetFilter.trim().isEmpty())
                ? null : getSheetOrNull(wb, sheetFilter.trim());
        if (sh != null) {
            return getPivotObjectOnSheet(sh, pivotName);
        }
        Dispatch sheets = Dispatch.get(wb, "Worksheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();
        for (int i = 1; i <= count; i++) {
            Dispatch ws = Dispatch.call(sheets, "Item", i).toDispatch();
            Dispatch p = tryGetPivotObjectOnSheet(ws, pivotName);
            if (p != null) return p;
        }
        throw new BotCommandException("No se encontró la PivotTable '" + pivotName + "' en el libro.");
    }
    private static Dispatch getPivotObjectOnSheet(Dispatch sheet, String pivotName) {
        Dispatch p = tryGetPivotObjectOnSheet(sheet, pivotName);
        if (p == null)
            throw new BotCommandException("No existe la PivotTable '" + pivotName + "' en la hoja: "
                    + Dispatch.get(sheet, "Name").getString());
        return p;
    }
    private static Dispatch tryGetPivotObjectOnSheet(Dispatch sheet, String pivotName) {
        try {
            Dispatch pvtTables = Dispatch.get(sheet, "PivotTables").toDispatch();
            return Dispatch.call(pvtTables, "Item", pivotName).toDispatch();
        } catch (Exception ignore) { return null; }
    }
    private static Dispatch findPivotTableRange2(Dispatch wb, String sheetFilter, String pivotName) {
        Dispatch p = findPivotObject(wb, sheetFilter, pivotName);
        return Dispatch.get(p, "TableRange2").toDispatch(); // toda la Pivot
    }

    // Solo datos: recorta última fila/col si hay GrandTotals
    // Reemplazar el helper original por ESTE
    private static Dispatch shrinkPivotGrandTotalsIfAny(Dispatch pvt, Dispatch dataArea) {
        // Flags GT
        boolean rowGrand = getBoolProp(pvt, "RowGrand");
        boolean colGrand = getBoolProp(pvt, "ColumnGrand");

        // Tamaño actual del DataBodyRange
        int rows = count(dataArea, "Rows");
        int cols = count(dataArea, "Columns");

        // Estructura de la Pivot
        int rowFields    = safeCount(pvt, "RowFields");
        int colFields    = safeCount(pvt, "ColumnFields");
        int dataFields   = safeCount(pvt, "DataFields");
        int dataOrient   = getDataFieldOrientation(pvt); // 1=xlRowField, 2=xlColumnField, 0=desconocido

        // -----------------------------------------------------
        // 1) ColGrand: recortar la ÚLTIMA COLUMNA solo si
        //    - hay Grand Total de columnas activo,
        //    - NO hay ColumnFields (caso simple),
        //    - los Valores (Σ) están en COLUMNAS,
        //    - y el ancho del DBR es MAYOR a #DataFields (=> hay una col extra que es GT).
        // -----------------------------------------------------
        if (colGrand && cols > 1) {
            boolean safeToShrinkCols = false;
            if (colFields == 0 && dataOrient == 2 /* xlColumnField */ && dataFields >= 1) {
                // En este layout, DBR width esperado = dataFields (+1 si hay GT).
                if (cols > dataFields) {
                    safeToShrinkCols = true;
                }
            }
            if (safeToShrinkCols) {
                dataArea = Dispatch.callN(
                        dataArea, "Resize",
                        new Variant[]{ new Variant(rows), new Variant(cols - 1) }
                ).toDispatch();
                cols = cols - 1; // mantener "cols" coherente si más abajo hicieras otro recorte
            }
        }

        // -----------------------------------------------------
        // 2) RowGrand: recortar la ÚLTIMA FILA solo si
        //    - hay Grand Total de filas activo,
        //    - hay RowFields (si no, carece de sentido),
        //    - y además hay ColumnFields (layout típico con GT de filas)
        //      o bien los Valores están en FILAS con >1 medida (caso menos común).
        //  Nota: súper conservador para NO cortar datos reales en layouts complejos.
        // -----------------------------------------------------
        if (rowGrand && rows > 1) {
            boolean safeToShrinkRows = false;
            if (rowFields > 0) {
                if (colFields > 0 || (dataOrient == 1 /* xlRowField */ && dataFields > 1)) {
                    safeToShrinkRows = true;
                }
            }
            if (safeToShrinkRows) {
                dataArea = Dispatch.callN(
                        dataArea, "Resize",
                        new Variant[]{ new Variant(rows - 1), new Variant(cols) }
                ).toDispatch();
                // rows--; // no hace falta continuar usando "rows" después
            }
        }

        return dataArea;
    }

// ===== Helpers internos del helper =====

    private static boolean getBoolProp(Dispatch obj, String prop) {
        try { return Dispatch.get(obj, prop).getBoolean(); }
        catch (Exception e) { return false; }
    }

    private static int safeCount(Dispatch pvt, String collectionName) {
        try {
            Dispatch coll = Dispatch.get(pvt, collectionName).toDispatch();
            return Dispatch.get(coll, "Count").getInt();
        } catch (Exception e) {
            return 0;
        }
    }

    private static int getDataFieldOrientation(Dispatch pvt) {
        try {
            Dispatch dpf = Dispatch.get(pvt, "DataPivotField").toDispatch();
            // 1 = xlRowField, 2 = xlColumnField, 0 = xlHidden (u otro)
            return Dispatch.get(dpf, "Orientation").getInt();
        } catch (Exception e) {
            return 0;
        }
    }

    // Construye un rango contiguo con SOLO las etiquetas de FILAS de la Pivot
    private static Dispatch buildPivotRowLabelsRange(Dispatch wb, String sheetFilter, String pivotName, boolean includeFieldHeaders, boolean excludeGrandTotals) {
        Dispatch pvt = findPivotObject(wb, sheetFilter, pivotName);
        Variant vData = Dispatch.get(pvt, "DataBodyRange");
        if (vData == null || vData.isNull())
            throw new BotCommandException("La Pivot no tiene área de datos (¿sin campos de valores?).");
        Dispatch dbr = vData.toDispatch();

        Dispatch rowRange = Dispatch.get(pvt, "RowRange").toDispatch(); // área de etiquetas de filas
        Dispatch ws = Dispatch.get(rowRange, "Worksheet").toDispatch();


        boolean rowGrand = false;
        try { rowGrand = Dispatch.get(pvt, "RowGrand").getBoolean(); } catch (Exception ignore) {}
        int dbrRows = count(dbr, "Rows");
        if (excludeGrandTotals && rowGrand && dbrRows > 1) {
            dbrRows = dbrRows - 1;  // recortamos la última fila (Total general)
        }


        // Alinear verticalmente con DataBodyRange (ya ajustado arriba si aplica)
        int startRow = includeFieldHeaders ? getRow(rowRange) : getRow(dbr);
        int rows = includeFieldHeaders
                ? (getRow(dbr) + dbrRows - getRow(rowRange))   // top de headers hasta el último dato (sin GT)
                : dbrRows;

        int startCol = getColumn(rowRange);
        int cols = count(rowRange, "Columns");

        // Construir rango
        Dispatch tl = Dispatch.call(ws, "Cells", startRow, startCol).toDispatch();
        Dispatch br = Dispatch.call(ws, "Cells", startRow + rows - 1, startCol + cols - 1).toDispatch();
        return Dispatch.call(ws, "Range", tl, br).toDispatch();

    }

    // Construye un rango contiguo con SOLO las etiquetas de COLUMNAS de la Pivot
    private static Dispatch buildPivotColumnLabelsRange(Dispatch wb, String sheetFilter, String pivotName,
                                                        boolean includeFieldHeaders, boolean excludeGrandTotals) {
        Dispatch pvt = findPivotObject(wb, sheetFilter, pivotName);
        Variant vData = Dispatch.get(pvt, "DataBodyRange");
        if (vData == null || vData.isNull())
            throw new BotCommandException("La Pivot no tiene área de datos (¿sin campos de valores?).");
        Dispatch dbr = vData.toDispatch();

        Dispatch colRange = Dispatch.get(pvt, "ColumnRange").toDispatch(); // área de etiquetas de columnas
        Dispatch ws = Dispatch.get(colRange, "Worksheet").toDispatch();

        // Alinear horizontalmente con DataBodyRange (excluye headers izquierdos si aplica)
        int startCol = getColumn(dbr);
        int cols = count(dbr, "Columns");

        // Excluir Grand Total de columnas si corresponde
        boolean colGrand = false;
        try { colGrand = Dispatch.get(pvt, "ColumnGrand").getBoolean(); } catch (Exception ignore) {}
        if (excludeGrandTotals && colGrand && cols > 1) cols = cols - 1;

        // Altura: todas las filas de encabezados o solo la última (leaf)
        int startRow, rows;
        if (includeFieldHeaders) {
            startRow = getRow(colRange);
            rows = count(colRange, "Rows");
        } else {
            startRow = getRow(colRange) + count(colRange, "Rows") - 1; // última fila de encabezados
            rows = 1;
        }

        Dispatch tl = Dispatch.call(ws, "Cells", startRow, startCol).toDispatch();
        Dispatch br = Dispatch.call(ws, "Cells", startRow + rows - 1, startCol + cols - 1).toDispatch();
        return Dispatch.call(ws, "Range", tl, br).toDispatch();
    }

    // -------- TABLAS (ListObject) --------

    private static class IncludeTableParts {
        Dispatch range;
    }

    private static IncludeTableParts resolveTableRange(
            Dispatch wb, String sheetFilter, String tableName,
            boolean includeHeader, boolean includeTotals, String mode
    ) {
        IncludeTableParts out = new IncludeTableParts();

        Dispatch sheet = (sheetFilter == null || sheetFilter.trim().isEmpty())
                ? null : getSheetOrNull(wb, sheetFilter.trim());
        Dispatch lo = (sheet != null) ? getListObjectOnSheet(sheet, tableName)
                : findListObjectOnAnySheet(wb, tableName);

        if ("values".equals(mode)) {
            Variant v = Dispatch.get(lo, "DataBodyRange");
            if (v == null || v.isNull()) throw new BotCommandException("La Tabla no tiene filas de datos.");
            out.range = v.toDispatch();
            return out;
        }
        // mode = full
        if (includeHeader && includeTotals) {
            out.range = Dispatch.get(lo, "Range").toDispatch();
            return out;
        }
        if (includeHeader && !includeTotals) {
            Dispatch hdr = Dispatch.get(lo, "HeaderRowRange").toDispatch();
            Variant vdb = Dispatch.get(lo, "DataBodyRange");
            if (vdb == null || vdb.isNull()) { out.range = hdr; return out; }
            Dispatch dbr = vdb.toDispatch();

            Dispatch ws = Dispatch.get(hdr, "Worksheet").toDispatch();
            int tlRow = getRow(hdr), tlCol = getColumn(hdr);
            int brRow = getRow(dbr) + count(dbr, "Rows") - 1;
            int brCol = getColumn(dbr) + count(dbr, "Columns") - 1;
            Dispatch tl = Dispatch.call(ws, "Cells", tlRow, tlCol).toDispatch();
            Dispatch br = Dispatch.call(ws, "Cells", brRow, brCol).toDispatch();
            out.range = Dispatch.call(ws, "Range", tl, br).toDispatch();
            return out;
        }
        if (!includeHeader && !includeTotals) {
            Variant vdb = Dispatch.get(lo, "DataBodyRange");
            if (vdb == null || vdb.isNull()) throw new BotCommandException("No hay datos (DataBodyRange vacío).");
            out.range = vdb.toDispatch();
            return out;
        } else {
            // !includeHeader && includeTotals => Data + Totals (si existen ambos)
            Variant vdb = Dispatch.get(lo, "DataBodyRange");
            Variant vt  = Dispatch.get(lo, "TotalsRowRange");
            if ((vdb == null || vdb.isNull()) && (vt == null || vt.isNull())) {
                throw new BotCommandException("Tabla sin datos ni TotalsRow.");
            } else if (vdb == null || vdb.isNull()) {
                out.range = vt.toDispatch();
                return out;
            } else if (vt == null || vt.isNull()) {
                out.range = vdb.toDispatch();
                return out;
            } else {
                Dispatch dbr = vdb.toDispatch();
                Dispatch tot = vt.toDispatch();
                Dispatch ws  = Dispatch.get(dbr, "Worksheet").toDispatch();
                int tlRow = getRow(dbr), tlCol = getColumn(dbr);
                int brRow = getRow(tot) + count(tot, "Rows") - 1;
                int brCol = getColumn(dbr) + count(dbr, "Columns") - 1;
                Dispatch tl = Dispatch.call(ws, "Cells", tlRow, tlCol).toDispatch();
                Dispatch br = Dispatch.call(ws, "Cells", brRow, brCol).toDispatch();
                out.range = Dispatch.call(ws, "Range", tl, br).toDispatch();
                return out;
            }
        }
    }

    private static Dispatch getListObjectOnSheet(Dispatch sheet, String tableName) {
        try {
            Dispatch los = Dispatch.get(sheet, "ListObjects").toDispatch();
            return Dispatch.call(los, "Item", tableName).toDispatch();
        } catch (Exception e) {
            throw new BotCommandException("No existe la Tabla '" + tableName + "' en la hoja: "
                    + Dispatch.get(sheet, "Name").getString());
        }
    }
    private static Dispatch findListObjectOnAnySheet(Dispatch wb, String tableName) {
        Dispatch sheets = Dispatch.get(wb, "Worksheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();
        for (int i = 1; i <= count; i++) {
            Dispatch ws = Dispatch.call(sheets, "Item", i).toDispatch();
            try {
                Dispatch los = Dispatch.get(ws, "ListObjects").toDispatch();
                return Dispatch.call(los, "Item", tableName).toDispatch();
            } catch (Exception ignore) { }
        }
        throw new BotCommandException("No se encontró la Tabla '" + tableName + "' en el libro.");
    }
}