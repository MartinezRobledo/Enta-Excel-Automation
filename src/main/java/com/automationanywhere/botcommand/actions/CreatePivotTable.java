package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.jacob.com.SafeArray;

import java.util.ArrayList;
import java.util.List;


/**
 * CreatePivotTable - SOLO PLAN B (PivotTableWizard)
 * - Mismo libro (pide solo la session).
 * - Hoja origen, hoja destino y celda destino.
 * - Headers en la fila 1 (simplificación).
 * - NUEVO: Rango de datos A1 opcional (sourceRangeA1). Si no se indica, auto A1:última fila/col usada.
 * - Row fields: lista (por header o letra).
 * - Value fields: lista (por header o letra) -> SUM y 'Valores' como columnas.
 * - Nombre de la PT: requerido; falla si ya existe.
 */
@BotCommand
@CommandPkg(
        label = "Create Pivot Table",
        name = "createPivotTable",
        description = "Crea una Tabla Dinámica (PivotTableWizard) en el mismo libro. Rows por lista; Values sumados en columnas. Admite rango A1 opcional.",
        icon = "excel.svg"
)
public class CreatePivotTable {
    // --- Excel constants ---
    private static final int xlDatabase = 1;
    private static final int xlRowField = 1;
    private static final int xlColumnField = 2;
    private static final int xlSum = -4157;

    @Execute
    public void action(
            // 1) Session
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

            // 2) Hoja origen
            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Hoja de origen (datos)", default_value_type = DataType.STRING)
            @NotEmpty String sourceSheetName,

            // 3) Hoja destino
            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Hoja destino (Tabla Dinámica)", default_value_type = DataType.STRING)
            @NotEmpty String destSheetName,

            // 4) Celda destino (top-left)
            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Celda destino (top-left), ej.: AN4190")
            @NotEmpty String destTopLeft,

            // 5) Modo de referencia
            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Por header", value = "header")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Por letra (A,B,...)", value = "letter"))
            })
            @Pkg(label = "Las columnas se indican por", default_value = "header", default_value_type = DataType.STRING)
            @SelectModes String referenceMode,

            // 6) Row fields (lista)
            @Idx(index = "6", type = AttributeType.LIST)
            @Pkg(label = "Campos de FILA (headers o letras según el modo)")
            @NotEmpty List<Value> rowFields,

            // 7) Value fields (lista)
            @Idx(index = "7", type = AttributeType.LIST)
            @Pkg(label = "Campos a SUMAR (headers o letras según el modo)")
            @NotEmpty List<Value> valueFields,

            // 8) Nombre de la PT
            @Idx(index = "8", type = AttributeType.TEXT)
            @Pkg(label = "Nombre de la Tabla Dinámica")
            @NotEmpty String pivotName,

            // 9) NUEVO: Rango de datos (A1) opcional
            @Idx(index = "9", type = AttributeType.TEXT)
            @Pkg(label = "Rango de datos (A1) - opcional, ej.: A1:BI4182", default_value_type = DataType.STRING)
            String sourceRangeA1
    ) {
        // --- Workbook & Sheets ---
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch src = requireSheetByName(wb, sourceSheetName);
        Dispatch dst = requireSheetByName(wb, destSheetName);

        // --- Listas saneadas ---
        List<String> rows = toStrings(rowFields);
        List<String> vals = toStrings(valueFields);
        if (rows.isEmpty()) throw new BotCommandException("Agregá al menos un campo de FILA.");
        if (vals.isEmpty()) throw new BotCommandException("Agregá al menos un campo a SUMAR.");

        String ptName = pivotName.trim();
        if (ptName.isEmpty()) throw new BotCommandException("El nombre de la tabla dinámica no puede ser vacío.");

        // --- Determinar el rango fuente ---
        final int headerRow = 1; // headers en fila 1
        int lastRow = ExcelHelpers.getLastDataRow(src);
        int lastCol = ExcelHelpers.getLastColumn(src);
        if (lastRow < headerRow) throw new BotCommandException("No hay datos debajo de la fila de encabezados.");

        int firstColForHeaders = 1;    // por defecto A
        int lastColForHeaders  = lastCol;

        Dispatch srcRange;
        if (sourceRangeA1 != null && !sourceRangeA1.trim().isEmpty()) {
            RangeBounds rb = parseA1RangeBounds(sourceRangeA1.trim());
            if (rb.firstRow != 1) {
                throw new BotCommandException("El rango debe comenzar en la fila 1 (encabezados en fila 1).");
            }
            srcRange = Dispatch.call(src, "Range", sourceRangeA1.trim()).toDispatch();
            // ajustar límites para validar encabezados
            firstColForHeaders = rb.firstCol;
            lastColForHeaders  = rb.lastCol;
            // opcional: alinear lastRow a lo provisto (para logs/consistencia)
            lastRow = rb.lastRow;
        } else {
            // Comportamiento actual: desde A1 a última col/fila usada
            String lastColLetter = ExcelHelpers.numberToColumnLetter(lastCol);
            String autoA1 = "A" + headerRow + ":" + lastColLetter + lastRow;
            srcRange = Dispatch.call(src, "Range", autoA1).toDispatch();
        }

        Dispatch destRange = Dispatch.call(dst, "Range", destTopLeft.trim()).toDispatch();

        // --- Validaciones: headers no vacíos y existencia de campos pedidos ---
        validateNoEmptyHeaders(src, headerRow, firstColForHeaders, lastColForHeaders, true);
        List<String> rowNames = resolveFieldNames(src, referenceMode, rows, headerRow, true);
        List<String> valueNames = resolveFieldNames(src, referenceMode, vals, headerRow, true);

        // --- Duplicado de nombre de PT en hoja destino ---
        Dispatch pvtTables = Dispatch.get(dst, "PivotTables").toDispatch();
        if (pivotNameExists(pvtTables, ptName)) {
            throw new BotCommandException("Ya existe una Tabla Dinámica llamada '" + ptName + "' en la hoja destino.");
        }

        // --- SOLO PLAN B: PivotTableWizard ---
        Dispatch pvt;
        try {
            Dispatch.callN(
                    dst, "PivotTableWizard",
                    new Variant[]{
                            new Variant(xlDatabase),
                            new Variant(srcRange),
                            new Variant(destRange),
                            new Variant(ptName)
                    }
            );
            pvt = Dispatch.call(pvtTables, "Item", ptName).toDispatch();
        } catch (Exception e) {
            throw new BotCommandException(
                    "No se pudo crear la Tabla Dinámica (Plan B). Detalle: " + safeMsg(e), e
            );
        }

        // --- Configurar Row fields ---
        int pos = 1;
        for (String fieldName : rowNames) {
            Dispatch pf = Dispatch.call(pvt, "PivotFields", fieldName).toDispatch();
            Dispatch.put(pf, "Orientation", new Variant(xlRowField));
            Dispatch.put(pf, "Position", new Variant(pos++));
            disableSubtotals(pf);
            Dispatch.put(pf, "RepeatLabels", new Variant(true)); // <-- NUEVO
        }
        // Layout tabular para toda la tabla
        //Dispatch.put(pvt, "RowAxisLayout", new Variant(1)); // xlTabularRow



        // --- Configurar Data fields (SUM) ---
        for (String fieldName : valueNames) {
            Dispatch pf = Dispatch.call(pvt, "PivotFields", fieldName).toDispatch();
            String caption = "Sum " + fieldName;
            Dispatch.callN(pvt, "AddDataField",
                    new Variant[]{ new Variant(pf), new Variant(caption), new Variant(xlSum) });
        }

        // --- 'Valores' como columnas ---
        try {
            Dispatch dataField = Dispatch.get(pvt, "DataPivotField").toDispatch();
            Dispatch.put(dataField, "Orientation", new Variant(xlColumnField));
        } catch (Exception ignore) {
            // Si solo hay un DataField, algunas versiones no exponen DataPivotField. Se ignora.
        }
    }

    // ===== Helpers =====

    private static Dispatch requireSheetByName(Dispatch wb, String name) {
        try {
            Dispatch sheets = Dispatch.get(wb, "Worksheets").toDispatch();
            return Dispatch.call(sheets, "Item", name).toDispatch();
        } catch (Exception e) {
            throw new BotCommandException("No existe la hoja '" + name + "' en el libro activo.");
        }
    }

    private static List<String> toStrings(List<Value> list) {
        List<String> out = new ArrayList<String>();
        if (list == null) return out;
        for (Value v : list) {
            Object o = (v == null ? null : v.get());
            String s = (o == null) ? "" : o.toString().trim();
            if (!s.isEmpty()) out.add(s);
        }
        return out;
    }

    private static List<String> resolveFieldNames(
            Dispatch sheet,
            String mode,
            List<String> entries,
            int headerRow,
            boolean trim
    ) {
        List<String> out = new ArrayList<String>();
        boolean doTrim = trim;
        for (String s : entries) {
            if ("letter".equalsIgnoreCase(mode)) {
                int col = ExcelHelpers.excelColumnLetterToNumber(s);
                Dispatch cell = Dispatch.call(sheet, "Cells", headerRow, col).toDispatch();
                Variant v = Dispatch.get(cell, "Value2");
                String header = (v == null || v.isNull()) ? "" : v.toString();
                if (doTrim) header = header.trim();
                if (header.isEmpty()) {
                    throw new BotCommandException(
                            "El encabezado en " + s + headerRow + " está vacío. La PT usa captions de encabezado."
                    );
                }
                out.add(header);
            } else {
                out.add(doTrim ? s.trim() : s);
            }
        }
        return out;
    }

    // Valida que no existan encabezados vacíos en la fila 1 entre firstCol y lastCol
    private static void validateNoEmptyHeaders(Dispatch sheet, int headerRow, int firstCol, int lastCol, boolean trim) {
        for (int col = firstCol; col <= lastCol; col++) {
            Dispatch cell = Dispatch.call(sheet, "Cells", headerRow, col).toDispatch();
            Variant v = Dispatch.get(cell, "Value2");
            String header = (v == null || v.isNull()) ? "" : v.toString();
            if (trim) header = header.trim();
            if (header.isEmpty()) {
                String colLetter = ExcelHelpers.numberToColumnLetter(col);
                throw new BotCommandException(
                        "Encabezado vacío en " + colLetter + headerRow +
                                ". Excel requiere encabezados no vacíos para todas las columnas del rango."
                );
            }
        }
    }

    private static boolean pivotNameExists(Dispatch pvtTables, String name) {
        try {
            Dispatch.call(pvtTables, "Item", name).toDispatch();
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private static String safeMsg(Exception e) {
        String m = (e == null ? "" : e.getMessage());
        return (m == null ? "" : m);
    }

    // --- NUEVO: bounds simples para un rango A1 ---
    private static class RangeBounds {
        int firstCol, lastCol, firstRow, lastRow;
    }

    // --- NUEVO: parsea "A1:BI4182" -> col/row inicial y final ---
    private static RangeBounds parseA1RangeBounds(String a1) {
        String[] parts = a1.split(":");
        String start = parts[0].trim();
        String end   = (parts.length > 1 ? parts[1].trim() : start);

        String sColLetters = start.replaceAll("\\d", "");
        String eColLetters = end.replaceAll("\\d", "");
        String sRowDigits  = start.replaceAll("\\D", "");
        String eRowDigits  = end.replaceAll("\\D", "");

        if (sColLetters.isEmpty() || eColLetters.isEmpty())
            throw new BotCommandException("Rango A1 inválido: " + a1);

        RangeBounds rb = new RangeBounds();
        rb.firstCol = ExcelHelpers.excelColumnLetterToNumber(sColLetters);
        rb.lastCol  = ExcelHelpers.excelColumnLetterToNumber(eColLetters);
        rb.firstRow = sRowDigits.isEmpty() ? 1 : Integer.parseInt(sRowDigits);
        rb.lastRow  = eRowDigits.isEmpty() ? rb.firstRow : Integer.parseInt(eRowDigits);

        if (rb.firstCol <= 0 || rb.lastCol <= 0 || rb.firstCol > rb.lastCol || rb.firstRow <= 0 || rb.lastRow < rb.firstRow) {
            throw new BotCommandException("Rango A1 inválido: " + a1);
        }
        return rb;
    }

    // Desactiva todos los subtotales de un PivotField (12 funciones: Sum, Count, etc.)
    private static void disableSubtotals(Dispatch pivotField) {
        // Excel espera un array de 12 booleans; todos en false = sin subtotales
        SafeArray sa = new SafeArray(Variant.VariantBoolean, 12);
        for (int i = 0; i < 12; i++) {
            sa.setBoolean(i, false);
        }
        Variant arr = new Variant();
        arr.putSafeArray(sa);
        Dispatch.put(pivotField, "Subtotals", arr);
    }

}
