package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.TableValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.data.model.Schema;
import com.automationanywhere.botcommand.data.model.table.Table;
import com.automationanywhere.botcommand.data.model.table.Row;
import com.automationanywhere.botcommand.exception.BotCommandException;

import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.*;

@BotCommand
@CommandPkg(
        label = "Get Table (visible rows)",
        name = "getTable",
        description = "Devuelve una tabla (AA) a partir de un rango de headers, leyendo solo filas visibles.",
        icon = "excel.svg",
        return_type = DataType.TABLE,
        return_required = true,
        return_label = "Tabla resultante"
)
public class GetTable {

    // Constantes Excel/COM
    private static final int xlCellTypeVisible = 12; // SpecialCells
    private static final int xlDown = -4121;         // Range.End(xlDown)

    @Execute
    public TableValue action(
            // ==== Sesión / Hoja ====
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @SessionObject @NotEmpty ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            Double sheetIndex,

            // ==== Rango de headers ====
            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Header range (A1 notation)", description = "Ej.: A1:D1 (una sola fila con los encabezados)")
            @NotEmpty String headerRangeA1,

            // ==== Opción: discontinuidad ====
            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "La tabla puede ser discontinua (permitir filas vacías intermedias)",
                    default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean allowDiscontinuous
    ) {

        // === Workbook / Sheet ===
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet = resolveSheet(wb, selectSheetBy, sheetName, sheetIndex);

        try {
            // --- Obtener el rango de headers ---
            Dispatch headerRange = Dispatch.call(sheet, "Range", headerRangeA1.trim()).toDispatch();
            if (headerRange == null || headerRange.m_pDispatch == 0) {
                throw new BotCommandException("El rango de headers no es válido: " + headerRangeA1);
            }
            int headerRows = getCount(Dispatch.get(headerRange, "Rows").toDispatch());
            int headerCols = getCount(Dispatch.get(headerRange, "Columns").toDispatch());
            if (headerRows != 1 || headerCols < 1) {
                throw new BotCommandException("El rango de headers debe ser una única fila con una o más columnas. Recibido: " + headerRangeA1);
            }
            int headerFirstRow = Dispatch.get(headerRange, "Row").getInt();
            int headerFirstCol = Dispatch.get(headerRange, "Column").getInt();
            int headerLastCol  = headerFirstCol + headerCols - 1;

            // --- Leer los headers (texto por columna) ---
            List<String> headers = new ArrayList<>(headerCols);
            for (int j = 1; j <= headerCols; j++) {
                Dispatch cell = Dispatch.call(headerRange, "Cells", 1, j).toDispatch();
                String h = safeStr(Dispatch.get(cell, "Value").toJavaObject());
                headers.add(h);
            }

            // --- Calcular primera fila de datos y última fila de la tabla ---
            int firstDataRow = headerFirstRow + 1;
            if (firstDataRow > ExcelHelpers.EXCEL_MAX_ROWS) {
                return new TableValue(buildTableWithSchema(headers)); // fuera de rango
            }

            int lastRow;
            if (Boolean.TRUE.equals(allowDiscontinuous)) {
                // Modo discontiguo: última fila real con datos entre todas las columnas
                int maxLast = 0;
                for (int col = headerFirstCol; col <= headerLastCol; col++) {
                    int lr = ExcelHelpers.getLastDataRowInColumn(sheet, col);
                    if (lr > maxLast) maxLast = lr;
                }
                lastRow = Math.max(firstDataRow, maxLast);
            } else {
                // Modo contiguo: usar columna guía = primera del rango
                int guideCol = headerFirstCol;
                Dispatch guideStart = cell(sheet, firstDataRow, guideCol);
                Dispatch below = cell(sheet, firstDataRow + 1, guideCol);
                if (isEmpty(guideStart)) {
                    // no hay datos en la primera fila de datos -> intentar mirar la siguiente visible (igual leeremos visibles más abajo)
                    lastRow = firstDataRow - 1; // sin datos
                } else if (!isEmpty(below)) {
                    Dispatch lastInBlock = Dispatch.call(guideStart, "End", new Variant(xlDown)).toDispatch();
                    lastRow = Dispatch.get(lastInBlock, "Row").getInt();
                } else {
                    lastRow = firstDataRow; // un solo registro
                }
            }

            if (lastRow < firstDataRow) {
                // no hay datos
                return new TableValue(buildTableWithSchema(headers));
            }

            // --- Rango de datos completo (sin header) ---
            Dispatch dataStart = cell(sheet, firstDataRow, headerFirstCol);
            Dispatch dataEnd   = cell(sheet, lastRow,      headerLastCol);
            Dispatch dataRange = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]{dataStart, dataEnd}, new int[1]).toDispatch();

            // --- Solo filas visibles ---
            Dispatch visible;
            try {
                visible = Dispatch.call(Dispatch.get(dataRange, "Cells").toDispatch(),
                        "SpecialCells", new Variant(xlCellTypeVisible)).toDispatch();
            } catch (Exception noVisible) {
                // Si no hay visibles, retornar tabla vacía con schema
                return new TableValue(buildTableWithSchema(headers));
            }

            // --- Construir TABLE (schema = headers; filas = visibles sin la fila de header) ---
            Table table = buildTableWithSchema(headers);
            List<Row> rows = ensureRows(table);

            Dispatch areas = Dispatch.get(visible, "Areas").toDispatch();
            int aCount = getCount(areas);

            for (int a = 1; a <= aCount; a++) {
                Dispatch area = Dispatch.call(areas, "Item", a).toDispatch();
                Dispatch areaRows = Dispatch.get(area, "Rows").toDispatch();
                int rCount = getCount(areaRows);
                for (int i = 1; i <= rCount; i++) {
                    Dispatch r = Dispatch.call(areaRows, "Item", i).toDispatch();

                    // Leer fila completa de la tabla
                    Row row = new Row();
                    List<Value> rv = new ArrayList<>(headerCols);
                    boolean allEmpty = true;

                    for (int j = 1; j <= headerCols; j++) {
                        Dispatch c = Dispatch.call(r, "Cells", 1, j).toDispatch();
                        String v = safeStr(Dispatch.get(c, "Value").toJavaObject());
                        if (!v.isEmpty()) allEmpty = false;
                        rv.add(new StringValue(v));
                    }

                    // Saltar filas completamente vacías (para que la primera fila sea la primera con datos)
                    if (!allEmpty) {
                        row.setValues(rv);
                        rows.add(row);
                    }
                }
            }

            return new TableValue(table);

        } catch (Exception e) {
            throw new BotCommandException("GetTable: " + e.getMessage(), e);
        }
    }

    // ============= Helpers =============

    private static Dispatch resolveSheet(Dispatch wb, String by, String name, Double index) {
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();
        if ("index".equalsIgnoreCase(by)) {
            if (index == null) throw new BotCommandException("Sheet Index es requerido cuando se selecciona por índice.");
            int i = index.intValue();
            if (i < 1 || i > count) throw new BotCommandException("Sheet Index fuera de rango (1.." + count + ").");
            return Dispatch.call(sheets, "Item", i).toDispatch();
        } else {
            if (name == null || name.trim().isEmpty())
                throw new BotCommandException("Sheet Name es requerido cuando se selecciona por nombre.");
            try {
                return Dispatch.call(sheets, "Item", name.trim()).toDispatch();
            } catch (Exception e) {
                throw new BotCommandException("No existe la hoja '" + name + "'.");
            }
        }
    }

    private static Dispatch cell(Dispatch sheet, int row, int col) {
        return Dispatch.call(sheet, "Cells", row, col).toDispatch();
    }

    private static int getCount(Dispatch collection) {
        return Dispatch.get(collection, "Count").getInt();
    }

    private static boolean isEmpty(Dispatch cell) {
        try {
            Variant v = Dispatch.get(cell, "Value2");
            if (v == null || v.isNull()) return true;
            Object o = v.toJavaObject();
            if (o == null) return true;
            String s = o.toString();
            return (s == null || s.trim().isEmpty());
        } catch (Exception e) {
            return true;
        }
    }

    private static String safeStr(Object o) {
        if (o == null) return "";
        String s = o.toString();
        return (s == null) ? "" : s.trim();
    }

    private static Table buildTableWithSchema(List<String> headers) {
        Table t = new Table();
        List<Schema> schema = new ArrayList<>(headers.size());
        for (String h : headers) {
            Schema sc = new Schema();
            sc.setName(h);
            schema.add(sc);
        }
        t.setSchema(schema);
        ensureRows(t);
        return t;
    }

    private static List<Row> ensureRows(Table t) {
        List<Row> rows = t.getRows();
        if (rows == null) {
            rows = new ArrayList<>();
            t.setRows(rows);
        }
        return rows;
    }
}