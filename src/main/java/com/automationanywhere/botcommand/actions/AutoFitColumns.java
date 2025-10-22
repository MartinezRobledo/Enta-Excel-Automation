package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
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

@BotCommand
@CommandPkg(
        label = "AutoFit Columns",
        name = "autoFitColumns",
        description = "Ajusta el ancho de las columnas para que el contenido sea visible (evita #######).",
        icon = "excel.svg",
        return_label = "Success",
        return_type = DataType.BOOLEAN,
        return_required = true
)
public class AutoFitColumns {

    @Execute
    public Value action(
            // 1) Sesión de Excel
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

            // 2) Selección de hoja por nombre o índice
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

            // 3) Ámbito a aplicar
            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "UsedRange",   value = "usedrange")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Entire Sheet", value = "entiresheet"))
            })
            @Pkg(label = "Scope", default_value = "usedrange", default_value_type = DataType.STRING)
            @SelectModes String scope,

            // 4) Considerar sólo celdas visibles (útil con filtros)
            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Consider only visible cells (filters)", default_value = "false", default_value_type = DataType.BOOLEAN)
            @NotEmpty Boolean onlyVisibleCells
    ) {

        // Validaciones básicas
        boolean byName = "name".equalsIgnoreCase(selectSheetBy);
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
        boolean scopeUsedRange   = "usedrange".equalsIgnoreCase(scope);
        boolean scopeEntireSheet = "entiresheet".equalsIgnoreCase(scope);
        if (!scopeUsedRange && !scopeEntireSheet) {
            throw new BotCommandException("Scope debe ser 'usedrange' o 'entiresheet'.");
        }

        // Obtener sesión/libro/hoja
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet = byIndex
                ? Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch()
                : Dispatch.call(sheets, "Item", sheetName).toDispatch();

        try {
            if (scopeUsedRange) {
                // Usar UsedRange para performance
                Variant vUR = Dispatch.get(sheet, "UsedRange");
                if (vUR == null || vUR.isNull()) {
                    return new BooleanValue(true); // nada que ajustar
                }
                Dispatch usedRange = vUR.toDispatch();

                Dispatch targetRange = usedRange;
                if (Boolean.TRUE.equals(onlyVisibleCells)) {
                    // xlCellTypeVisible = 12
                    try {
                        targetRange = Dispatch.call(usedRange, "SpecialCells", new Variant(12)).toDispatch();
                    } catch (Exception ignore) {
                        // Si no hay celdas visibles (caso extremo), no hacemos nada
                        return new BooleanValue(true);
                    }
                }

                Dispatch columns = Dispatch.get(targetRange, "Columns").toDispatch();
                Dispatch.call(columns, "AutoFit");

            } else {
                // Toda la hoja (puede ser lento en libros grandes)
                Dispatch sheetColumns = Dispatch.get(sheet, "Columns").toDispatch();

                if (Boolean.TRUE.equals(onlyVisibleCells)) {
                    // Intersectar Columns con las filas visibles de UsedRange para acotar el cálculo
                    Variant vUR = Dispatch.get(sheet, "UsedRange");
                    if (vUR != null && !vUR.isNull()) {
                        Dispatch usedRange = vUR.toDispatch();
                        Dispatch app = Dispatch.get(wb, "Application").toDispatch();
                        // Tomar sólo visibles dentro del UsedRange
                        Dispatch visible;
                        try {
                            visible = Dispatch.call(usedRange, "SpecialCells", new Variant(12)).toDispatch();
                            Variant vIntersect = Dispatch.call(app, "Intersect", sheetColumns, visible);
                            if (vIntersect != null && !vIntersect.isNull()) {
                                Dispatch colVis = Dispatch.get(vIntersect.toDispatch(), "Columns").toDispatch();
                                Dispatch.call(colVis, "AutoFit");
                                return new BooleanValue(true);
                            }
                        } catch (Exception ignored) {
                            // si falla SpecialCells, seguimos con Columns completo
                        }
                    }
                }

                // Fallback/General: AutoFit sobre todas las columnas
                Dispatch.call(sheetColumns, "AutoFit");
            }

            return new BooleanValue(true);

        } catch (BotCommandException ex) {
            throw ex;
        } catch (Exception ex) {
            throw new BotCommandException("No se pudo ajustar el ancho de columnas: " + ex.getMessage(), ex);
        }
    }
}