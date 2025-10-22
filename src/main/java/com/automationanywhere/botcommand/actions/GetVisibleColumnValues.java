package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

@BotCommand
@CommandPkg(
        label = "Get Visible Column Values",
        name = "getVisibleColumnValues",
        description = "Devuelve los valores visibles de una columna según el filtro aplicado",
        icon = "excel.svg",
        return_label = "Values",
        return_type = DataType.LIST,
        return_required = true
)
public class GetVisibleColumnValues {

    @Execute
    public Value action(
            // 1) Sesión de Excel
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

            // 2) Selección de hoja por nombre o índice
            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index="2.1", pkg=@Pkg(label="Name",  value="name")),
                    @Idx.Option(index="2.2", pkg=@Pkg(label="Index", value="index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NumberInteger @GreaterThanEqualTo("1") @NotEmpty Double sheetIndex,

            // 3) Selección de columna por encabezado o por letra
            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Header", value = "header")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Letter", value = "letter"))
            })
            @Pkg(label = "Select Column By", default_value = "letter", default_value_type = DataType.STRING)
            @SelectModes String selectColumnBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Header Name")
            @NotEmpty String columnName,

            @Idx(index = "3.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Letter (A, B, ...)")
            @NotEmpty String columnLetter,

            // 4) Fila inicial
            @Idx(index = "4", type = AttributeType.NUMBER)
            @Pkg(label = "Start Row", default_value = "2", default_value_type = DataType.NUMBER)
            @NumberInteger @GreaterThanEqualTo("1") @NotEmpty Double startRowInput,

            // 5) Mostrar texto formateado (Range.Text) en lugar del valor subyacente (Range.Value)
            @Idx(index = "5", type = AttributeType.CHECKBOX)
            @Pkg(label = "Return displayed text (formatted)", default_value = "false", default_value_type = DataType.BOOLEAN)
            @NotEmpty Boolean returnDisplayedText
    ) {

        // --- Validaciones básicas ---
        boolean byName  = "name".equalsIgnoreCase(selectSheetBy);
        boolean byIndex = "index".equalsIgnoreCase(selectSheetBy);
        if (!byName && !byIndex) {
            throw new BotCommandException("Select sheet by debe ser 'name' o 'index'.");
        }
        if (byName && (sheetName == null || sheetName.trim().isEmpty())) {
            throw new BotCommandException("Sheet Name no puede estar vacío.");
        }
        if (byIndex && (sheetIndex == null || sheetIndex < 1)) {
            throw new BotCommandException("Sheet Index debe ser >= 1.");
        }

        boolean byHeader = "header".equalsIgnoreCase(selectColumnBy);
        boolean byLetter = "letter".equalsIgnoreCase(selectColumnBy);
        if (!byHeader && !byLetter) {
            throw new BotCommandException("Select Column By debe ser 'header' o 'letter'.");
        }
        if (byHeader && (columnName == null || columnName.trim().isEmpty())) {
            throw new BotCommandException("Column Header Name no puede estar vacío.");
        }
        if (byLetter && (columnLetter == null || columnLetter.trim().isEmpty())) {
            throw new BotCommandException("Column Letter no puede estar vacía.");
        }

        int startRow = startRowInput.intValue();
        if (startRow < 1) startRow = 1;

        // --- Obtener sesión/Excel/hoja ---
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb     = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet  = byIndex
                ? Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch()
                : Dispatch.call(sheets, "Item", sheetName).toDispatch();

        // --- Última fila usada ---
        int lastRow = getLastRow(sheet);
        if (lastRow < startRow) {
            return emptyListValue();
        }

        // --- Resolver columna absoluta ---
        int absCol;
        if (byLetter) {
            absCol = excelColumnLetterToNumber(columnLetter);
            if (absCol < 1) throw new BotCommandException("Column Letter inválida: '" + columnLetter + "'");
        } else {
            // Por convención: si Start Row = 2, el header está en 1
            int headerRow = (startRow > 1) ? (startRow - 1) : 1;
            absCol = findColumnByHeader(sheet, headerRow, columnName);
            if (absCol <= 0) {
                throw new BotCommandException("Encabezado '" + columnName + "' no encontrado en fila " + headerRow + ".");
            }
        }

        // --- Determinar el ÁMBITO del filtro (AutoFilter.Range o Tabla) ---
        Dispatch scopeRange = null;        // rango global (puede incluir header si es AutoFilter.Range)
        Integer scopeHeaderRow = null;     // fila de header si aplica AutoFilter.Range

        // 1) AutoFilter activo y la columna cae dentro del rango del filtro
        try {
            Variant vAF = Dispatch.get(sheet, "AutoFilter");
            if (vAF != null && !vAF.isNull()) {
                Dispatch autoFilter = vAF.toDispatch();
                Dispatch afRange    = Dispatch.get(autoFilter, "Range").toDispatch();
                int frFirstCol = Dispatch.get(afRange, "Column").getInt();
                int frCols     = Dispatch.get(Dispatch.get(afRange, "Columns").toDispatch(), "Count").getInt();
                if (absCol >= frFirstCol && absCol <= (frFirstCol + frCols - 1)) {
                    scopeRange     = afRange;
                    scopeHeaderRow = Dispatch.get(afRange, "Row").getInt();
                }
            }
        } catch (Exception ignore) { /* no AF o no accesible */ }

        // 2) Si no hay AutoFilter aplicable, buscar Tabla (ListObject) que contenga la columna
        if (scopeRange == null) {
            try {
                Dispatch listObjects = Dispatch.get(sheet, "ListObjects").toDispatch();
                int tableCount = Dispatch.get(listObjects, "Count").getInt();
                for (int i = 1; i <= tableCount; i++) {
                    Dispatch table      = Dispatch.call(listObjects, "Item", i).toDispatch();
                    Dispatch tableRange = Dispatch.get(table, "Range").toDispatch();
                    int tblFirstCol = Dispatch.get(tableRange, "Column").getInt();
                    int tblCols     = Dispatch.get(Dispatch.get(tableRange, "Columns").toDispatch(), "Count").getInt();
                    if (absCol >= tblFirstCol && absCol <= (tblFirstCol + tblCols - 1)) {
                        Variant vBody = Dispatch.get(table, "DataBodyRange"); // solo datos
                        if (vBody != null && !vBody.isNull()) {
                            scopeRange = vBody.toDispatch();
                            scopeHeaderRow = null;
                        } else {
                            return emptyListValue(); // tabla sin filas
                        }
                        break;
                    }
                }
            } catch (Exception ignore) { /* sin ListObjects */ }
        }

        List<Value> out = new ArrayList<>();

        if (scopeRange != null) {
            // --- Caso con scope (AutoFilter o Tabla): iterar Areas de visibles ---
            Dispatch visibleRange;
            try {
                // xlCellTypeVisible = 12 (aplicar sobre el scope completo)
                visibleRange = Dispatch.call(scopeRange, "SpecialCells", new Variant(12)).toDispatch();
            } catch (Exception ex) {
                return emptyListValue();
            }

            // Determinar última fila con datos en la columna
            int lastRowCol = getLastUsedRowInColumn(sheet, absCol);
            int effectiveLastRow = Math.min(lastRow, lastRowCol);
            if (effectiveLastRow < startRow) {
                return emptyListValue();
            }

            Dispatch areas = Dispatch.get(visibleRange, "Areas").toDispatch();
            int areaCount = Dispatch.get(areas, "Count").getInt();

            for (int a = 1; a <= areaCount; a++) {
                Dispatch area = Dispatch.call(areas, "Item", a).toDispatch();
                Dispatch aRows = Dispatch.get(area, "Rows").toDispatch();
                int aRowCount = Dispatch.get(aRows, "Count").getInt();

                for (int r = 1; r <= aRowCount; r++) {
                    Dispatch rowRange = Dispatch.call(aRows, "Item", r).toDispatch();
                    int rowNum = Dispatch.get(rowRange, "Row").getInt();

                    // Excluir header si aplica AutoFilter
                    if (scopeHeaderRow != null && rowNum == scopeHeaderRow.intValue()) continue;

                    // Respetar la fila de inicio
                    if (rowNum < startRow) continue;

                    // Limitar por última fila con datos en la columna
                    if (rowNum > effectiveLastRow) continue;

                    Dispatch cell = Dispatch.call(sheet, "Cells", rowNum, absCol).toDispatch();
                    Variant v = Dispatch.get(cell, returnDisplayedText ? "Text" : "Value");
                    out.add(new StringValue(safeVariantToString(v)));
                }
            }
        } else {
            // --- Fallback: sin AF/Tabla -> usar el rectángulo de la COLUMNA desde startRow hasta la última fila usada en la columna ---
            int lastRowCol = getLastUsedRowInColumn(sheet, absCol);
            int effectiveLastRow = Math.min(lastRow, lastRowCol);
            if (effectiveLastRow < startRow) {
                return emptyListValue();
            }

            Dispatch topLeft     = Dispatch.call(sheet, "Cells", startRow,         absCol).toDispatch();
            Dispatch bottomRight = Dispatch.call(sheet, "Cells", effectiveLastRow, absCol).toDispatch();
            Dispatch colRange    = Dispatch.call(sheet, "Range", topLeft, bottomRight).toDispatch();

            // 1) Intento principal: SpecialCells(xlCellTypeVisible)
            try {
                Dispatch visibleRange = Dispatch.call(colRange, "SpecialCells", new Variant(12)).toDispatch();
                Dispatch cells = Dispatch.get(visibleRange, "Cells").toDispatch();
                int count = Dispatch.get(cells, "Count").getInt();
                for (int i = 1; i <= count; i++) {
                    Dispatch cell = Dispatch.call(cells, "Item", i).toDispatch();
                    Variant v = Dispatch.get(cell, returnDisplayedText ? "Text" : "Value");
                    out.add(new StringValue(safeVariantToString(v)));
                }
            } catch (Exception ex) {
                // 2) Fallback robusto: recorrer del startRow al effectiveLastRow y tomar solo filas no ocultas
                for (int row = startRow; row <= effectiveLastRow; row++) {
                    Dispatch rowObj = Dispatch.call(sheet, "Rows", row).toDispatch();
                    boolean hidden = false;
                    try {
                        hidden = Dispatch.get(rowObj, "Hidden").getBoolean();
                    } catch (Exception ignore) {
                        // Si no podemos leer Hidden, asumimos visible
                    }
                    if (hidden) continue;

                    Dispatch cell = Dispatch.call(sheet, "Cells", row, absCol).toDispatch();
                    Variant v = Dispatch.get(cell, returnDisplayedText ? "Text" : "Value");
                    out.add(new StringValue(safeVariantToString(v)));
                }
            }
        }

        // Devolver ListValue
        trimTrailingEmpty(out);
        ListValue lv = new ListValue();
        lv.set(out);
        return lv;
    }

    // ------------------------- Helpers -------------------------

    private static ListValue emptyListValue() {
        ListValue lv = new ListValue();
        lv.set(new ArrayList<>());
        return lv;
    }

    private static int findColumnByHeader(Dispatch sheet, int headerRow, String headerText) {
        if (headerText == null) return -1;
        String target = headerText.trim().toLowerCase(Locale.ROOT);

        // Usar UsedRange para acotar
        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        int firstCol = Dispatch.get(usedRange, "Column").getInt();
        int colCount = Dispatch.get(Dispatch.get(usedRange, "Columns").toDispatch(), "Count").getInt();
        int lastCol = firstCol + colCount - 1;

        for (int c = firstCol; c <= lastCol; c++) {
            Dispatch cell = Dispatch.call(sheet, "Cells", headerRow, c).toDispatch();
            String val = safeVariantToString(Dispatch.get(cell, "Value"))
                    .trim().toLowerCase(Locale.ROOT);
            if (!val.isEmpty() && val.equals(target)) {
                return c;
            }
        }
        return -1;
    }

    private static int getLastRow(Dispatch sheet) {
        // xlCellTypeLastCell = 11
        try {
            Dispatch cells    = Dispatch.get(sheet, "Cells").toDispatch();
            Dispatch lastCell = Dispatch.call(cells, "SpecialCells", new Variant(11)).toDispatch();
            return Dispatch.get(lastCell, "Row").getInt();
        } catch (Exception ex) {
            // Fallback a UsedRange
            try {
                Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
                int firstRow  = Dispatch.get(usedRange, "Row").getInt();
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

    private static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return o != null ? o.toString() : "";
    }

    private static int getLastUsedRowInColumn(Dispatch sheet, int absCol) {
        // xlUp = -4162
        final int XL_UP = -4162;
        try {
            Dispatch rows = Dispatch.get(sheet, "Rows").toDispatch();
            int rowCount = Dispatch.get(rows, "Count").getInt(); // normalmente 1,048,576 en xlsx
            Dispatch bottomCell = Dispatch.call(sheet, "Cells", rowCount, absCol).toDispatch();
            Dispatch lastCell = Dispatch.call(bottomCell, "End", new Variant(XL_UP)).toDispatch();
            return Dispatch.get(lastCell, "Row").getInt();
        } catch (Exception ex) {
            // Fallback: UsedRange
            try {
                Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
                int firstRow = Dispatch.get(usedRange, "Row").getInt();
                int totalRows = Dispatch.get(Dispatch.get(usedRange, "Rows").toDispatch(), "Count").getInt();
                return firstRow + totalRows - 1;
            } catch (Exception ignore) {
                return 1;
            }
        }
    }
    // helper
    private static void trimTrailingEmpty(List<Value> list) {
        int i = list.size() - 1;
        while (i >= 0) {
            Value v = list.get(i);
            if (v instanceof StringValue && ((StringValue) v).get().trim().isEmpty()) {
                list.remove(i--);
            } else {
                break;
            }
        }
    }

}