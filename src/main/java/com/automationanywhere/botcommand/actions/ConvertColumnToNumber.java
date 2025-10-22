package com.automationanywhere.botcommand.actions;

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
import com.jacob.com.Variant;

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.*;

@BotCommand
@CommandPkg(
        label = "Convert Column as Number",
        name  = "convertColumnToNumber",
        description = "Convierte números pegados como texto en NUMÉRICOS sobre una columna",
        icon = "excel.svg"
)
public class ConvertColumnToNumber {

    // Constantes Excel que usamos (evitamos dependencias externas)
    private static final int xlDelimited = 1;          // Range.TextToColumns DataType
    private static final int xlCalculationManual = -4135;

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

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

            @Idx(index = "4", type = AttributeType.NUMBER)
            @Pkg(label = "Start Row", default_value = "2", default_value_type = DataType.NUMBER)
            @NumberInteger @GreaterThanEqualTo("1") @NotEmpty Double startRowInput
    ) {
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb     = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet  = ExcelObjects.requireSheet(wb, selectSheetBy, sheetName, sheetIndex);
        Dispatch app    = Dispatch.get(wb, "Application").toDispatch();

        // Guardar estado y optimizar
        boolean prevUpd = getBool(app, "ScreenUpdating");
        boolean prevEvt = getBool(app, "EnableEvents");
        boolean prevAlr = getBool(app, "DisplayAlerts");
        int     prevCal = getInt (app, "Calculation");

        putBool(app, "ScreenUpdating", false);
        putBool(app, "EnableEvents",  false);
        putBool(app, "DisplayAlerts", false);
        putInt (app, "Calculation",   xlCalculationManual);

        try {
            // Activar hoja y chequear protección
            Dispatch.call(sheet, "Activate"); // simplifica contexto
            if (getBool(sheet, "ProtectContents")) {
                throw new BotCommandException("La hoja está protegida (ProtectContents=true). No se puede convertir la columna.");
            } // [7](https://exceloffthegrid.com/vba-code-worksheet-protection/)

            // Resolver colIndex
            int colIndex;
            Dispatch used = Dispatch.get(sheet, "UsedRange").toDispatch();
            int firstRow  = Dispatch.get(used, "Row").getInt();
            int rowsCnt   = Dispatch.get(Dispatch.get(used, "Rows").toDispatch(), "Count").getInt();
            int colsCnt   = Dispatch.get(Dispatch.get(used, "Columns").toDispatch(), "Count").getInt();

            if ("letter".equalsIgnoreCase(selectColumnBy)) {
                if (columnLetter == null || columnLetter.isEmpty())
                    throw new BotCommandException("Column letter not provided.");
                colIndex = excelColumnLetterToNumber(columnLetter);
            } else {
                String target = columnName.trim();
                colIndex = ExcelHelpers.headerNameToColumnIndex(sheet, target, firstRow, colsCnt);
            }

            int startRow     = startRowInput.intValue();
            int usedLastRow  = (rowsCnt > 0) ? (firstRow + rowsCnt - 1) : startRow;
            int lastDataRow  = ExcelHelpers.getLastDataRowInColumn(sheet, colIndex);
            int lastRowIndex = (lastDataRow > 0) ? Math.max(lastDataRow, startRow)
                    : Math.max(usedLastRow, startRow);
            int rowSize      = Math.max(1, lastRowIndex - startRow + 1);

            // Armar rango como Range(start,end) (evitamos Resize)
            Dispatch start = Dispatch.call(sheet, "Cells", startRow, colIndex).toDispatch();
            Dispatch end   = Dispatch.call(sheet, "Cells", lastRowIndex, colIndex).toDispatch();
            Dispatch rng   = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]{ start, end }, new int[1]).toDispatch();
            String addr    = Dispatch.get(rng, "Address").toString();

            // ===== Estrategia 1: TextToColumns (General) =====
            boolean converted = false;
            try {
                // TextToColumns con DataType = xlDelimited y defaults (coerciona a General)
                // Todos los argumentos son opcionales; al no especificar delimitadores, Excel
                // reinterpreta los valores numéricos (General) respetando separadores del sistema. [1](https://learn.microsoft.com/en-us/office/vba/api/Excel.Range.TextToColumns)
                Dispatch.callN(rng, "TextToColumns", new Variant[] {
                        // Destination (omitimos => mismo lugar)
                        Variant.VT_MISSING, // 1
                        new Variant(xlDelimited), // DataType
                        Variant.VT_MISSING, Variant.VT_MISSING, // TextQualifier, ConsecutiveDelimiter
                        Variant.VT_MISSING, Variant.VT_MISSING, Variant.VT_MISSING, Variant.VT_MISSING, // Tab;Semicolon;Comma;Space
                        Variant.VT_MISSING, Variant.VT_MISSING, // Other, OtherChar
                        Variant.VT_MISSING, // FieldInfo (default->General)
                        Variant.VT_MISSING, Variant.VT_MISSING, // DecimalSeparator, ThousandsSeparator
                        Variant.VT_MISSING  // TrailingMinusNumbers
                });
                converted = true;
            } catch (Exception ignore) {
                // Seguimos al fallback
            }

            // ===== Fallback: fórmula VALUE + pegar valores =====
            if (!converted) {
                // Poner fórmula (R1C1) que convierte texto->número y luego reemplazar por valores
                Dispatch.put(rng, "FormulaR1C1", "=VALUE(RC[0])"); // convierte cada celda a número [3](https://www.exceldemy.com/convert-text-to-number-excel-vba/)
                // Sustituir fórmula por valores (equivalente a “pegar valores”)
                Variant vals = Dispatch.get(rng, "Value2");  // leer valores numéricos resultantes [5](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)
                Dispatch.put(rng, "Value2", vals);           // escribir de vuelta (valores estáticos)
            }

        } finally {
            // Restaurar estado Excel
            putInt (app, "Calculation",   prevCal);
            putBool(app, "DisplayAlerts", prevAlr);
            putBool(app, "EnableEvents",  prevEvt);
            putBool(app, "ScreenUpdating",prevUpd);
        }
    }
}