package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

import java.util.List;
import java.util.stream.Collectors;

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.*; // getBool, getInt, putBool, putInt, xlCalculationManual, etc.

/**
 * Inserta una LISTA de valores o fórmulas en fila o columna comenzando en una celda inicial.
 * Ejemplos:
 *  - startCell=B3 + direction=row    -> B3, C3, D3, ...
 *  - startCell=D8 + direction=column -> D8, D9, D10, ...
 *
 * Basado en el patrón robusto de InsertValue (manejo de sesión, optimizaciones y fallback de COM thread).
 */
@BotCommand
@CommandPkg(
        label = "Set Values",
        name = "setValues",
        description = "Inserta una lista de valores o fórmulas en fila o columna comenzando desde una celda inicial",
        icon = "excel.svg"
)
public class SetValues {

    @Execute
    public void action(
            // 1) Sesión
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

            // 2) Selección de hoja (igual patrón que tu InsertValue)
            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name",  value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NumberInteger @GreaterThanEqualTo("1") @NotEmpty Double sheetIndex,

            // 3) Lista de valores / fórmulas
            @Idx(index = "3", type = AttributeType.LIST)
            @Pkg(label = "Valores / Fórmulas (lista)")
            @NotEmpty List<Value> values,

            // 4) ¿Se interpretan como fórmula?
            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "¿Es Fórmula?", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean isFormula,

            // 5) Dirección y celda inicial
            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Fila",    value = "row")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Columna", value = "column"))
            })
            @Pkg(label = "Dirección", default_value = "row", default_value_type = DataType.STRING)
            String direction,

            @Idx(index = "6", type = AttributeType.TEXT)
            @Pkg(label = "Celda inicial (ej. B3)")
            @NotEmpty String startCell
    ) {
        try {
            run(excelSession, selectSheetBy, sheetName, sheetIndex, values, isFormula, direction, startCell);
        } catch (Exception first) {
            // Fallback robusto por si el hilo no tenía COM inicializado
            try {
                ComThread.InitSTA();
                run(excelSession, selectSheetBy, sheetName, sheetIndex, values, isFormula, direction, startCell);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("InsertValues failed: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private void run(
            ExcelSession excelSession,
            String selectSheetBy, String sheetName, Double sheetIndex,
            List<Value> Valvalues, Boolean isFormula,
            String direction, String startCell
    ) {

        List<String> values = Valvalues.stream()
                .map(v -> {
                    if (v instanceof StringValue) {
                        return ((StringValue) v).get();
                    } else {
                        throw new BotCommandException("Todos los elementos deben ser texto.");
                    }
                })
                .collect(Collectors.toList());

        if (values == null || values.isEmpty()) {
            throw new BotCommandException("La lista de valores/fórmulas no puede estar vacía.");
        }
        if (startCell == null || startCell.trim().isEmpty()) {
            throw new BotCommandException("La celda inicial es obligatoria (ej. B3).");
        }
        if (!"row".equalsIgnoreCase(direction) && !"column".equalsIgnoreCase(direction)) {
            throw new BotCommandException("Dirección inválida. Usa 'row' o 'column'.");
        }

        // 1) Re-attach a Excel en este hilo (mismo patrón que tu InsertValue)
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb     = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet  = ExcelObjects.requireSheet(wb, selectSheetBy, sheetName, sheetIndex);
        Dispatch app    = Dispatch.get(wb, "Application").toDispatch();

        // 2) Guardar estado y optimizar (igual patrón que InsertValue)
        boolean prevUpd = getBool(app, "ScreenUpdating");
        boolean prevEvt = getBool(app, "EnableEvents");
        boolean prevAlr = getBool(app, "DisplayAlerts");
        int     prevCalc= getInt (app, "Calculation");

        putBool(app, "ScreenUpdating", false);
        putBool(app, "EnableEvents",   false);
        putBool(app, "DisplayAlerts",  false);
        putInt (app, "Calculation",    xlCalculationManual);

        try {
            // 3) Obtener fila/col de la celda inicial vía Range (evitamos parse manual A1)
            Dispatch start = Dispatch.call(sheet, "Range", startCell).toDispatch();
            int startRow = getInt(start, "Row");
            int startCol = getInt(start, "Column");

            // 4) Insertar secuencialmente (robusto).
            //    Nota: Se puede optimizar a escritura por bloque con SafeArray 2D,
            //    pero este enfoque ya rinde muy bien con calc/events/screen off.
            boolean asFormula = Boolean.TRUE.equals(isFormula);
            final boolean byRow = "row".equalsIgnoreCase(direction);

            for (int i = 0; i < values.size(); i++) {
                int r = byRow ? startRow : startRow + i;
                int c = byRow ? startCol + i : startCol;

                Dispatch cell = Dispatch.call(sheet, "Cells", r, c).toDispatch();
                String v = values.get(i);

                if (asFormula) {
                    // Debe incluir '=' si corresponde (no alteramos el string)
                    Dispatch.put(cell, "Formula", v);
                } else {
                    Dispatch.put(cell, "Value2", v);
                }
            }

        } finally {
            // 5) Restaurar estado de Excel (mismo patrón que InsertValue)
            putInt (app, "Calculation",   prevCalc);
            putBool(app, "DisplayAlerts", prevAlr);
            putBool(app, "EnableEvents",  prevEvt);
            putBool(app, "ScreenUpdating",prevUpd);
        }
    }
}