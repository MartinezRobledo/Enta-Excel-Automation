package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
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
        label = "Has Formula Errors",
        name = "hasFormulaErrors",
        description = "Devuelve TRUE si existen errores en fórmulas (y opcionalmente en constantes) en el alcance indicado.",
        icon = "excel.svg",
        return_type = DataType.BOOLEAN,
        return_required = true,
        return_label = "¿Hay errores?",
        return_description = "TRUE si existe al menos un error; FALSE en caso contrario."
)
public class HasFormulaErrors {

    // --- Constantes COM (Excel) ---
    private static final int xlCellTypeFormulas  = -4123; // FIX
    private static final int xlCellTypeConstants = 2;     // para opcional “constantes con error”
    private static final int xlErrors            = 16;

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

            // ===== Alcance =====
            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Entire sheet", value = "sheet")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Range (A1)", value = "range")),
                    @Idx.Option(index = "3.3", pkg = @Pkg(label = "Column letter", value = "column"))
            })
            @Pkg(label = "Scope", default_value = "sheet", default_value_type = DataType.STRING)
            String scope,

            @Idx(index = "3.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Range A1 (ej.: C2:K500)")
            String a1Range,

            @Idx(index = "3.3.1", type = AttributeType.TEXT)
            @Pkg(label = "Column letter (ej.: D)")
            String columnLetter,

            // ===== Opcional: incluir errores en constantes (no solo fórmulas) =====
            @Idx(index = "4", type = AttributeType.BOOLEAN)
            @Pkg(label = "Include constant errors (not only formulas)?",
                    default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean includeConstantErrors
    ) {
        // --- Workbook / Sheet ---
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet = resolveSheet(wb, selectSheetBy, sheetName, sheetIndex);

        // --- Determinar rango base según scope ---
        Dispatch baseRange;
        String s = (scope == null ? "sheet" : scope.trim().toLowerCase());
        switch (s) {
            case "range":
                if (a1Range == null || a1Range.trim().isEmpty())
                    throw new BotCommandException("Debe indicar Range A1 cuando Scope = range.");
                baseRange = Dispatch.call(sheet, "Range", a1Range.trim()).toDispatch();
                break;
            case "column":
                String col = (columnLetter == null ? "" : columnLetter.trim());
                if (col.isEmpty())
                    throw new BotCommandException("Debe indicar Column letter cuando Scope = column.");
                baseRange = Dispatch.call(sheet, "Range", col + ":" + col).toDispatch();
                break;
            case "sheet":
            default:
                baseRange = Dispatch.get(sheet, "Cells").toDispatch();
                break;
        }

        // --- Acotar a UsedRange para performance (cuando aplica) ---
        Dispatch app = Dispatch.get(sheet, "Application").toDispatch();
        Dispatch used = null;
        try {
            used = Dispatch.get(sheet, "UsedRange").toDispatch();
        } catch (Exception ignore) {}

        Dispatch effectiveRange = baseRange;
        if (used != null && used.m_pDispatch != 0 && (s.equals("sheet") || s.equals("column"))) {
            try {
                Dispatch intersect = Dispatch.call(app, "Intersect", baseRange, used).toDispatch();
                if (intersect != null && intersect.m_pDispatch != 0) {
                    effectiveRange = intersect;
                }
            } catch (Exception ignore) {
                // Si Intersect falla, seguimos con baseRange
            }
        }

        // --- Buscar errores en fórmulas ---
        boolean hasFormulaErrors = hasSpecialCells(effectiveRange, xlCellTypeFormulas, xlErrors);

        // --- (Opcional) Buscar también errores en constantes ---
        boolean hasConstErrors = false;
        if (Boolean.TRUE.equals(includeConstantErrors)) {
            hasConstErrors = hasSpecialCells(effectiveRange, xlCellTypeConstants, xlErrors);
        }

        return new BooleanValue(hasFormulaErrors || hasConstErrors);
    }

    // ================= Helpers =================
    private static Dispatch resolveSheet(Dispatch wb, String selectSheetBy, String name, Double index) {
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();
        if ("index".equalsIgnoreCase(selectSheetBy)) {
            if (index == null) throw new BotCommandException("Sheet Index es requerido cuando 'Select sheet by' = index.");
            int i = index.intValue();
            if (i < 1 || i > count) throw new BotCommandException("Sheet Index fuera de rango (1.." + count + ").");
            return Dispatch.call(sheets, "Item", i).toDispatch();
        } else {
            if (name == null || name.trim().isEmpty())
                throw new BotCommandException("Sheet Name es requerido cuando 'Select sheet by' = name.");
            try {
                return Dispatch.call(sheets, "Item", name.trim()).toDispatch();
            } catch (Exception e) {
                throw new BotCommandException("No existe la hoja '" + name + "'.");
            }
        }
    }

    /** Devuelve true si hay coincidencias en SpecialCells(type, value); Excel lanza excepción cuando no hay. */
    private static boolean hasSpecialCells(Dispatch range, int type, int value) {
        try {
            Dispatch res = Dispatch.callN(range, "SpecialCells",
                    new Object[]{ new Variant(type), new Variant(value) }).toDispatch();
            return res != null && res.m_pDispatch != 0;
        } catch (Exception noMatches) {
            return false; // típico: “no cells found”
        }
    }
}