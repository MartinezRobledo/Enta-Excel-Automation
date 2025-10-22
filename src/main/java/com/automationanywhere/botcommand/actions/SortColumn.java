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

import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.*;
// headerNameToColumnIndex, excelColumnLetterToNumber, getLastDataRowInColumn

/**
 * Ordena una columna ASC/DESC:
 * - Selección de columna ORIGEN por Header o Letter.
 * - Indicar si la columna tiene header (xlYes/xlNo).
 * - Orden: Asc / Desc.
 *
 * Nota: Ordena SOLO ESA COLUMNA (no expande a toda la tabla).
 * Si más adelante querés una variante que "mantenga filas alineadas"
 * expandiendo a todo el bloque (tabla), lo armamos sin tocar la UI.
 */
@BotCommand
@CommandPkg(
        label = "Sort Column",
        name = "sortColumn",
        description = "Ordena una columna (Header/Letter) en ascendente o descendente.",
        icon = "excel.svg"
)
public class SortColumn {

    // --- Constantes Excel / VBA ---
    // XlSortOrder
    private static final int xlAscending  = 1;  // Ascendente
    private static final int xlDescending = 2;  // Descendente
    // XlYesNoGuess
    private static final int xlGuess = 0;
    private static final int xlYes   = 1;
    private static final int xlNo    = 2;
    // Calculation
    private static final int xlCalculationAutomatic = -4105;
    private static final int xlCalculationManual    = -4135;

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @SessionObject @NotEmpty ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name",  value = "name")),
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

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Header", value = "header")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Letter", value = "letter"))
            })
            @Pkg(label = "Select Column By", default_value = "letter", default_value_type = DataType.STRING)
            @SelectModes String selectColumnBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Header Name")
            String columnHeader,

            @Idx(index = "3.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Letter (A, B, ...)")
            String columnLetter,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "No",  value = "no"))
            })
            @Pkg(label = "Column has header?", default_value = "no", default_value_type = DataType.STRING)
            @SelectModes String hasHeader,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Ascending (A-Z / Small-Large)", value = "asc")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Descending (Z-A / Large-Small)", value = "desc"))
            })
            @Pkg(label = "Sort order", default_value = "asc", default_value_type = DataType.STRING)
            @SelectModes String sortOrder
    ) {
        int order1 = "desc".equalsIgnoreCase(sortOrder) ? xlDescending : xlAscending; // por defecto asc
        int headerFlag = "yes".equalsIgnoreCase(hasHeader) ? xlYes : xlNo;

        // 1) Excel, workbook, hoja
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb     = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet  = ExcelObjects.requireSheet(wb, selectSheetBy, sheetName, sheetIndex);

        // 2) UsedRange: headerRow y colsCnt
        Dispatch used = Dispatch.get(sheet, "UsedRange").toDispatch();
        if (used == null || used.m_pDispatch == 0) return;

        int headerRow = Dispatch.get(used, "Row").getInt();                                      // primera fila del UsedRange
        int colsCnt   = Dispatch.get(Dispatch.get(used, "Columns").toDispatch(), "Count").getInt();

        // 3) Resolver colIndex
        int colIndex;
        if ("letter".equalsIgnoreCase(selectColumnBy)) {
            if (columnLetter == null || columnLetter.trim().isEmpty())
                throw new BotCommandException("Column letter not provided.");
            colIndex = excelColumnLetterToNumber(columnLetter.trim());
        } else {
            if (columnHeader == null || columnHeader.trim().isEmpty())
                throw new BotCommandException("Column header name not provided.");
            colIndex = headerNameToColumnIndex(sheet, columnHeader.trim(), headerRow, colsCnt);
            if (colIndex <= 0) throw new BotCommandException("Header not found: " + columnHeader);
        }

        // 4) Última fila con datos en ESA columna (helper tuyo)
        int lastDataRow = getLastDataRowInColumn(sheet, colIndex);
        if (lastDataRow <= 0) return;

        int startRow = ("yes".equalsIgnoreCase(hasHeader)) ? headerRow : (headerRow + 1);
        if (lastDataRow < startRow) return;

        // Armar rango: Range(Cells(startRow, col), Cells(lastDataRow, col))
        Dispatch start = Dispatch.call(sheet, "Cells", startRow, colIndex).toDispatch();
        Dispatch end   = Dispatch.call(sheet, "Cells", lastDataRow, colIndex).toDispatch();
        Dispatch rng   = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]{ start, end }, new int[1]).toDispatch();

        // Key1: la celda “clave” dentro del rango. Para Header=xlYes, Excel acepta que sea el header.
        Dispatch key1 = start;

        // 5) Optimización Excel
        Dispatch app = Dispatch.get(wb, "Application").toDispatch();
        boolean prevUpd = true, prevEvt = true, prevAlr = true; int prevCalc = xlCalculationAutomatic;

        try {
            prevUpd = getBool(app, "ScreenUpdating");
            prevEvt = getBool(app, "EnableEvents");
            prevAlr = getBool(app, "DisplayAlerts");
            prevCalc= getInt (app, "Calculation");

            putBool(app, "ScreenUpdating", false);
            putBool(app, "EnableEvents",   false);
            putBool(app, "DisplayAlerts",  false);
            putInt (app, "Calculation",    xlCalculationManual);

            // 6) Ordenar: Range.Sort(Key1, Order1, ..., Header)
            // Usamos sólo los argumentos esenciales: Key1, Order1, Header.
            // (El resto usa defaults; por ejemplo, orientación por filas/top-to-bottom).
            Dispatch.callN(rng, "Sort", new Variant[] {
                    new Variant(key1),               // Key1
                    new Variant(order1),             // Order1 (xlAscending/xlDescending)
                    Variant.VT_MISSING,              // Key2
                    Variant.VT_MISSING,              // Type
                    Variant.VT_MISSING,              // Order2
                    Variant.VT_MISSING,              // Key3
                    Variant.VT_MISSING,              // Order3
                    new Variant(headerFlag),         // Header (xlYes/xlNo/xlGuess)
                    Variant.VT_MISSING,              // OrderCustom
                    Variant.VT_MISSING,              // MatchCase
                    Variant.VT_MISSING,              // Orientation (default TopToBottom)
                    Variant.VT_MISSING,              // SortMethod
                    Variant.VT_MISSING,              // DataOption1
                    Variant.VT_MISSING,              // DataOption2
                    Variant.VT_MISSING               // DataOption3
            });

        } finally {
            // Restaurar estado Excel
            putInt (app, "Calculation",   prevCalc);
            putBool(app, "DisplayAlerts", prevAlr);
            putBool(app, "EnableEvents",  prevEvt);
            putBool(app, "ScreenUpdating",prevUpd);
        }
    }

    // === Helpers internos (idénticos estilo a CopyTableOptimized) ===
    private static boolean getBool(Dispatch app, String prop) {
        try { return Dispatch.get(app, prop).getBoolean(); } catch (Exception e) { return true; }
    }
    private static int getInt(Dispatch app, String prop) {
        try { return Dispatch.get(app, prop).getInt(); } catch (Exception e) { return xlCalculationAutomatic; }
    }
    private static void putBool(Dispatch app, String prop, boolean v) {
        try { Dispatch.put(app, prop, v); } catch (Exception ignore) {}
    }
    private static void putInt(Dispatch app, String prop, int v) {
        try { Dispatch.put(app, prop, new Variant(v)); } catch (Exception ignore) {}
    }
}