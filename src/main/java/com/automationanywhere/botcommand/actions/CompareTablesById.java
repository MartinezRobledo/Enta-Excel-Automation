package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.data.impl.DictionaryValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.*;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.*;

@BotCommand
@CommandPkg(
        label = "Compare Tables by ID Column",
        name = "compareExcelRangesHash",
        description = "Compares ranges from two Excel workbooks using column ID",
        icon = "excel.svg",
        return_type = DataType.DICTIONARY,
        return_label = "Comparison result",
        return_required = true
)
public class CompareTablesById {

    // Constantes Excel
    private static final int XL_CELL_TYPE_VISIBLE = 12; // xlCellTypeVisible


    @Execute
    public DictionaryValue action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession sourceExcelSession,

            @Idx(index = "2", type = AttributeType.SESSION)
            @Pkg(label = "Destination Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession destExcelSession,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectOriginSheetBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            @NotEmpty
            String originSheetName,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double originSheetIndex,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select destination sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectDestSheetBy,

            @Idx(index = "4.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Destination Sheet Name")
            @NotEmpty
            String destSheetName,

            @Idx(index = "4.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Destination Sheet Index (1-based)")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double destSheetIndex,

            @Idx(index = "5", type = AttributeType.TEXT)
            @Pkg(label = "ID Column Table 1")
            @NotEmpty String idCol1,

            @Idx(index = "6", type = AttributeType.TEXT)
            @Pkg(label = "ID Column Table 2")
            @NotEmpty String idCol2
    ) {
        Map<String, Value> result = new HashMap<>();

        try {
            // ==== Sesión y workbooks ====
            Session sourceSession = sourceExcelSession.getSession();
            if (sourceSession == null || sourceSession.excelApp == null)
                throw new BotCommandException("Source Session not found o closed");

            if (sourceSession.openWorkbooks.isEmpty())
                throw new BotCommandException("Source workbook: No workbook is open in this session.");

            Dispatch wb1 = sourceSession.openWorkbooks.values().iterator().next();

            Session destSession = destExcelSession.getSession();
            if (destSession == null || destSession.excelApp == null)
                throw new BotCommandException("Destination Session not found o closed");

            if (sourceSession.openWorkbooks.isEmpty())
                throw new BotCommandException("Destination workbook: No workbook is open in this session.");

            Dispatch wb2 = destSession.openWorkbooks.values().iterator().next();

            // ==== Hojas ====
            Dispatch sheet1 = selectOriginSheetBy.equals("name")
                    ? Dispatch.call(Dispatch.get(wb1, "Sheets").toDispatch(), "Item", originSheetName).toDispatch()
                    : Dispatch.call(Dispatch.get(wb1, "Sheets").toDispatch(), "Item", originSheetIndex.intValue()).toDispatch();

            Dispatch sheet2 = selectDestSheetBy.equals("name")
                    ? Dispatch.call(Dispatch.get(wb2, "Sheets").toDispatch(), "Item", destSheetName).toDispatch()
                    : Dispatch.call(Dispatch.get(wb2, "Sheets").toDispatch(), "Item", destSheetIndex.intValue()).toDispatch();

            // ==== Usar UsedRange de cada hoja ====
            Dispatch usedRange1 = Dispatch.get(sheet1, "UsedRange").toDispatch();
            Dispatch usedRange2 = Dispatch.get(sheet2, "UsedRange").toDispatch();

            // Validar que hay contenido
            int usedRows1 = getCount(Dispatch.get(usedRange1, "Rows").toDispatch());
            int usedCols1 = getCount(Dispatch.get(usedRange1, "Columns").toDispatch());
            int usedRows2 = getCount(Dispatch.get(usedRange2, "Rows").toDispatch());
            int usedCols2 = getCount(Dispatch.get(usedRange2, "Columns").toDispatch());


            if (usedRows1 < 1 || usedCols1 < 1 || usedRows2 < 1 || usedCols2 < 1) {
                result.put("equal", new BooleanValue(false));
                result.put("message", new StringValue("Alguna hoja no tiene datos (UsedRange vacío)."));
                return new DictionaryValue(result);
            }

            // ==== Hallar índice de columna por header (en la PRIMERA fila del UsedRange) ====
            int idColIndex1 = getColumnIndexByHeader(sheet1, usedRange1, idCol1);
            int idColIndex2 = getColumnIndexByHeader(sheet2, usedRange2, idCol2);

            // ==== Excluir header: recortar UsedRange a "dataRange" (desde fila 2 del UsedRange) ====
            if (usedRows1 <= 1 || usedRows2 <= 1) {
                result.put("equal", new BooleanValue(false));
                result.put("message", new StringValue("No hay filas de datos (solo header) en alguna tabla."));
                return new DictionaryValue(result);
            }
            Dispatch dataRange1 = resizeWithoutHeader(usedRange1, usedRows1, usedCols1);
            Dispatch dataRange2 = resizeWithoutHeader(usedRange2, usedRows2, usedCols2);

            // ==== Obtener solo filas visibles (respetando filtros/ocultas) ====
            Dispatch visibleRows1 = getVisibleRows(dataRange1);
            Dispatch visibleRows2 = getVisibleRows(dataRange2);

            if (visibleRows1 == null || visibleRows2 == null) {
                result.put("equal", new BooleanValue(false));
                result.put("message", new StringValue("No hay filas visibles en al menos una tabla."));
                return new DictionaryValue(result);
            }

            // ==== Contar frecuencias del ID en filas visibles ====
            Map<String, Integer> freq1 = countIdFrequencies(visibleRows1, idColIndex1);
            Map<String, Integer> freq2 = countIdFrequencies(visibleRows2, idColIndex2);

            boolean equal = freq1.equals(freq2);

            result.put("equal", new BooleanValue(equal));
            if (!equal) {

                // Después de calcular freq1 y freq2
                Set<String> allKeys = new HashSet<>();
                allKeys.addAll(freq1.keySet());
                allKeys.addAll(freq2.keySet());

                List<String> faltanEnTabla2 = new ArrayList<>();
                List<String> faltanEnTabla1 = new ArrayList<>();

                for (String key : allKeys) {
                    int c1 = freq1.getOrDefault(key, 0);
                    int c2 = freq2.getOrDefault(key, 0);
                    if (c1 > c2) faltanEnTabla2.add(key);
                    if (c2 > c1) faltanEnTabla1.add(key);
                }

                result.put("faltanEnTabla1", new StringValue(faltanEnTabla1.toString()));
                result.put("faltanEnTabla2", new StringValue(faltanEnTabla2.toString()));
                result.put("message", new StringValue("Las frecuencias de ID difieren entre filas visibles."));
            } else {
                result.put("message", new StringValue("Las frecuencias de ID coinciden en filas visibles."));
            }

            return new DictionaryValue(result);

        } catch (Exception e) {
            // Encapsular en BotCommandException para A360
            throw new BotCommandException("Error comparando tablas: " + e.getMessage(), e);
        }
    }

    // === Helpers ===

    /** Devuelve Columns/Rows.Count de una colección (Dispatch). */
    private int getCount(Dispatch collection) {
        return Dispatch.get(collection, "Count").getInt(); // ✓ acá sí es int
    }

    /** Normaliza Variant a String segura (simple y suficiente para comparar IDs). */
    private String v2str(Variant v) {
        if (v == null || v.isNull()) return "";
        String s = v.toString();
        return (s != null) ? s.trim() : "";
    }

    /** Obtiene índice de columna (1-based) buscando header en la primera fila del UsedRange. */
    private int getColumnIndexByHeader(Dispatch sheet, Dispatch usedRange, String headerText) {
        Dispatch columns = Dispatch.get(usedRange, "Columns").toDispatch();
        int colCount = getCount(columns);

        // Primera fila del UsedRange (no necesariamente fila 1 del sheet)
        Dispatch rowsInUR = Dispatch.get(usedRange, "Rows").toDispatch();
        Dispatch headerRow = Dispatch.call(rowsInUR, "Item", 1).toDispatch();

        for (int j = 1; j <= colCount; j++) {
            Dispatch cell = Dispatch.call(headerRow, "Cells", 1, j).toDispatch();
            Variant v = Dispatch.get(cell, "Value");
            String headerVal = v2str(v);
            if (headerVal.equalsIgnoreCase(headerText)) return j;
        }
        throw new BotCommandException("Header '" + headerText + "' no encontrado en la primera fila del UsedRange.");
    }

    /** Crea un rango sin la fila de header: Offset(1,0).Resize(rows-1, cols). */
    private Dispatch resizeWithoutHeader(Dispatch usedRange, int usedRows, int usedCols) {
        Dispatch offset = Dispatch.call(usedRange, "Offset", 1, 0).toDispatch(); // desde 2da fila del UR
        // Resize(RowSize, ColumnSize)
        return Dispatch.call(offset, "Resize", usedRows - 1, usedCols).toDispatch();
    }

    /** Devuelve un Range compuesto de las filas visibles dentro de dataRange (excluye header). */
    private Dispatch getVisibleRows(Dispatch dataRange) {
        Dispatch dataRows = Dispatch.get(dataRange, "Rows").toDispatch();
        try {
            return Dispatch.call(dataRows, "SpecialCells", new Variant(XL_CELL_TYPE_VISIBLE)).toDispatch();
        } catch (Exception ex) {
            // No hay visibles -> null
            return null;
        }
    }

    /**
     * Cuenta frecuencias de la columna ID en filas visibles.
     * Recorre Areas para soportar rangos discontinuos (típico de filtros).
     */
    private Map<String, Integer> countIdFrequencies(Dispatch visibleRows, int idColIndex) {
        Map<String, Integer> freq = new HashMap<>();

        // Manejar rangos discontinuos (Areas)
        Dispatch areas = Dispatch.get(visibleRows, "Areas").toDispatch();
        int areasCount = getCount(areas);


        for (int a = 1; a <= areasCount; a++) {
            Dispatch area = Dispatch.call(areas, "Item", a).toDispatch();
            Dispatch areaRows = Dispatch.get(area, "Rows").toDispatch();
            int rowCount = getCount(areaRows);

            for (int i = 1; i <= rowCount; i++) {
                Dispatch row = Dispatch.call(areaRows, "Item", i).toDispatch();
                Dispatch cell = Dispatch.call(row, "Cells", 1, idColIndex).toDispatch();
                Variant v = Dispatch.get(cell, "Value");
                String value = v2str(v);
                freq.put(value, freq.getOrDefault(value, 0) + 1);
            }
        }
        return freq;
    }
}
