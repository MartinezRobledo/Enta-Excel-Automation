package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "AutoFill",
        name = "autoFill",
        description = "Rellena como el doble‑click del fill handle: copia hacia abajo hasta el final del bloque contiguo adyacente.",
        icon = "excel.svg",
        return_type = DataType.STRING,
        return_required = true,
        return_label = "Destino ocupado (A1)",
        return_description = "Dirección A1 final del rango destino"
)
public class AutoFill {

    // Constantes COM (Excel)
    private static final int xlDown = -4121;       // XlDirection.xlDown
    private static final int xlFillDefault = 0;    // XlAutoFillType.xlFillDefault

    @Execute
    public Value action(
            // ===== Sesión / Hoja =====
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

            // ===== Origen / Preferencia de guía =====
            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Start cell (A1), ej.: D2")
            @NotEmpty
            String startA1,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Auto (Left then Right)", value = "auto")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "Left", value = "left")),
                    @Idx.Option(index = "4.3", pkg = @Pkg(label = "Right", value = "right"))
            })
            @Pkg(label = "Guide preference", default_value = "auto", default_value_type = DataType.STRING)
            String guidePreference,

            // (Reservado para futuro) Fill type; por ahora siempre xlFillDefault
            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Default (let Excel decide)", value = "default"))
            })
            @Pkg(label = "Fill type", default_value = "default", default_value_type = DataType.STRING)
            String fillType,

            @Idx(index = "6", type = AttributeType.CHECKBOX)
            @Pkg(label = "Ignorar filas vacías en la guía (expandir sobre saltos)",
                    default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean spanGaps
    ) {
        // === Workbook / Sheet ===
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet = resolveSheet(wb, selectSheetBy, sheetName, sheetIndex);

        // Origen
        String cellRef = (startA1 == null ? "" : startA1.trim());
        if (cellRef.isEmpty()) throw new BotCommandException("Start cell no puede estar vacío.");
        Dispatch start = Dispatch.call(sheet, "Range", cellRef).toDispatch();
        int startRow = Dispatch.get(start, "Row").getInt();
        int startCol = Dispatch.get(start, "Column").getInt();

        // Determinar columna guía
        Integer guideCol = pickGuideColumn(sheet, startRow, startCol, guidePreference);

        if (guideCol == null) {
            // No hay guía útil (ambas adyacentes vacías al nivel de startRow)
            // => comportamiento Excel: doble‑click no hace nada
            return new StringValue(addressOf(start));
        }

        Dispatch guideCell = cell(sheet, startRow, guideCol);
        // Si la guía en la fila de inicio está vacía, no hay referencia para extender
        if (isEmpty(guideCell)) {
            return new StringValue(addressOf(start));
        }

        // Calcular la última fila destino
        int lastRow = startRow;

        if (Boolean.TRUE.equals(spanGaps)) {
            // NUEVO: ignorar vacíos intermedios => ir hasta la última fila con datos en la guía
            int lastData = ExcelHelpers.getLastDataRowInColumn(sheet, guideCol);
            if (lastData > startRow) lastRow = lastData;
        } else {
            // Comportamiento clásico (contiguo): si debajo hay algo, usamos End(xlDown), sino nos quedamos
            Dispatch below = cell(sheet, startRow + 1, guideCol);
            if (!isEmpty(below)) {
                Dispatch lastInBlock = Dispatch.call(guideCell, "End", new Variant(xlDown)).toDispatch();
                lastRow = Dispatch.get(lastInBlock, "Row").getInt();
            }
        }

        if (lastRow <= startRow) {
            // Nada que rellenar
            return new StringValue(addressOf(start));
        }

        // Destination = desde start hasta la fila final en la MISMA columna del origen
        Dispatch destBottomRight = cell(sheet, lastRow, startCol);
        Dispatch destination = Dispatch.call(sheet, "Range", start, destBottomRight).toDispatch();

        // AutoFill (dejo que Excel elija la estrategia: xlFillDefault)
        Dispatch.callN(start, "AutoFill", new Object[] { destination, new Variant(xlFillDefault) });

        // Devolver la dirección del rango destino
        String addr = addressOf(destination);
        String shName = Dispatch.get(Dispatch.get(destination, "Worksheet").toDispatch(), "Name").getString();
        return new StringValue(shName + "!" + addr);
    }

    // ================= Helpers =================

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

    private static Integer pickGuideColumn(Dispatch sheet, int row, int startCol, String pref) {
        String p = (pref == null ? "auto" : pref.trim().toLowerCase());
        boolean canLeft = startCol > 1;
        boolean leftHasValue = canLeft && !isEmpty(cell(sheet, row, startCol - 1));
        boolean rightHasValue = !isEmpty(cell(sheet, row, startCol + 1));

        if ("left".equals(p)) {
            return leftHasValue ? startCol - 1 : null;
        } else if ("right".equals(p)) {
            return rightHasValue ? startCol + 1 : null;
        } else {
            // auto: prioriza izquierda, sino derecha
            if (leftHasValue) return startCol - 1;
            if (rightHasValue) return startCol + 1;
            return null;
        }
    }

    private static Dispatch cell(Dispatch sheet, int row, int col) {
        return Dispatch.call(sheet, "Cells", row, col).toDispatch();
    }

    private static boolean isEmpty(Dispatch cell) {
        try {
            Variant v = Dispatch.get(cell, "Value2");
            if (v == null || v.isNull()) return true;
            Object o = v.toJavaObject();
            if (o == null) return true;
            String s = o.toString();
            return s == null || s.trim().isEmpty();
        } catch (Exception e) {
            return true;
        }
    }

    private static String addressOf(Dispatch range) {
        // Address(RowAbsolute=false, ColumnAbsolute=false)
        Variant addr = Dispatch.callN(range, "Address",
                new Variant[] { new Variant(false), new Variant(false) });
        return addr.getString();
    }
}