package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.*;
import java.util.stream.Collectors;

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.splitRangeByIgnoredColumns;

@BotCommand
@CommandPkg(
        label = "Copy Table Content (deprecated)",
        name = "copyTableContent",
        description = "Copies a table (with headers range) from one sheet to another with options for headers, filtering and columns ignoring",
        icon = "excel.svg",
        return_type = DataType.BOOLEAN,
        return_required = false
)
public class CopyTable {

    @Execute
    public Value<Boolean> action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession sourceExcelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            @NotEmpty
            String originSheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            @NotEmpty
            Double originSheetIndex,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Headers Range (e.g., C9:BM9)") @NotEmpty String headersRange,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "All rows", value = "all")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "Only visible rows", value = "visible"))
            })
            @Pkg(label = "Rows to copy", default_value = "all", default_value_type = DataType.STRING)
            @SelectModes
            String rowsMode,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Include headers?", default_value = "no", default_value_type = DataType.STRING)
            @SelectModes
            String includeHeaders,

            @Idx(index = "6", type = AttributeType.SESSION)
            @Pkg(label = "Destination Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession destExcelSession,

            @Idx(index = "7", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "7.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "7.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectDestSheetBy,

            @Idx(index = "7.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Destination Sheet Name")
            @NotEmpty
            String destSheetName,

            @Idx(index = "7.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Destination Sheet Index (1-based)")
            @NotEmpty
            Double destSheetIndex,

            @Idx(index = "8", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "8.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "8.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Create sheet if not exists?", default_value = "yes", default_value_type = DataType.STRING)
            @SelectModes
            String createSheet,

            @Idx(index = "9", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "9.1", pkg = @Pkg(label = "Overwrite", value = "overwrite")),
                    @Idx.Option(index = "9.2", pkg = @Pkg(label = "Append", value = "append")),
                    @Idx.Option(index = "9.3", pkg = @Pkg(label = "Manual", value = "manual"))
            })
            @Pkg(label = "Copy mode", default_value = "overwrite", default_value_type = DataType.STRING)
            @SelectModes
            String copyMode,

            @Idx(index = "9.3.1", type = AttributeType.TEXT)
            @Pkg(label = "Manual destination cell (e.g., A2)") String manualCell,

            @Idx(index = "10", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "10.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "10.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Ignore columns?", default_value = "no", default_value_type = DataType.STRING)
            @SelectModes
            String applyIgnoreCol,

            @Idx(index = "10.1.1", type = AttributeType.LIST)
            @Pkg(label = "Columns to ignore (letters; e.g., A,B,C)")
            List<Object> ignoreColumns


    ) {
        {
            // 1) Sesión + workbook correctos
            Session sourceSession = ExcelObjects.requireSession(sourceExcelSession);
            Dispatch wb1 = ExcelObjects.requireWorkbook(sourceSession, sourceExcelSession);

            Session destSession = ExcelObjects.requireSession(destExcelSession);
            Dispatch wb2 = ExcelObjects.requireWorkbook(destSession, destExcelSession);

            Dispatch sourceSheets = Dispatch.get(wb1, "Sheets").toDispatch();
            Dispatch sourceSheet;

            if ("index".equalsIgnoreCase(selectSheetBy)) {
                sourceSheet = Dispatch.call(sourceSheets, "Item", originSheetIndex.intValue()).toDispatch();
            } else {
                sourceSheet = Dispatch.call(sourceSheets, "Item", originSheetName).toDispatch();
            }

            Dispatch destSheets = Dispatch.get(wb2, "Sheets").toDispatch();
            Dispatch destSheet = null;

            // buscar hoja por nombre o índice
            for (int i = 1; i <= Dispatch.get(destSheets, "Count").getInt(); i++) {
                Dispatch s = Dispatch.call(destSheets, "Item", i).toDispatch();
                // FIX menor: usar getString() para el nombre
                String sheetName = Dispatch.get(s, "Name").getString();
                if ("index".equalsIgnoreCase(selectDestSheetBy) && i == destSheetIndex.intValue()) {
                    destSheet = s;
                    break;
                } else if ("name".equalsIgnoreCase(selectDestSheetBy) && sheetName.equalsIgnoreCase(destSheetName)) {
                    destSheet = s;
                    break;
                }
            }

            if (destSheet == null) {
                if ("yes".equalsIgnoreCase(createSheet)) {
                    destSheet = Dispatch.call(destSheets, "Add").toDispatch();
                    Dispatch.put(destSheet, "Name", destSheetName);
                } else {
                    throw new BotCommandException("Destination sheet does not exist.");
                }
            }

            try {
                // Determinar rango completo desde headers (tu lógica actual)
                String startCell = headersRange.split(":")[0];
                String endCell = headersRange.split(":")[1];
                int headerRow = Integer.parseInt(endCell.replaceAll("\\D", ""));

                Dispatch usedRange = Dispatch.get(sourceSheet, "UsedRange").toDispatch();
                Dispatch rows = Dispatch.get(usedRange, "Rows").toDispatch();
                int lastRow = Dispatch.get(rows, "Count").getInt();


                int startRow = "yes".equalsIgnoreCase(includeHeaders) ? headerRow : headerRow + 1;
                if (lastRow < startRow) {
                    return new BooleanValue(false);
                }
                String copyRange = startCell.replaceAll("\\d", "") + startRow + ":" + endCell.replaceAll("\\d", "") + lastRow;

                Dispatch rangeToCopy = Dispatch.call(sourceSheet, "Range", copyRange).toDispatch();

                boolean isFilteredCopy = false;
                if ("visible".equalsIgnoreCase(rowsMode)) {
                    Dispatch visible = specialCellsVisibleOrNull(rangeToCopy);
                    if (visible == null) {
                        // No hay filas visibles → no copiar y devolver false
                        return new BooleanValue(false);
                    }
                    rangeToCopy = visible;
                    isFilteredCopy = true;
                }

                // Guard extra: si el rango a copiar no tiene filas útiles
                Dispatch rowsToCopy = Dispatch.get(rangeToCopy, "Rows").toDispatch();
                int rowsCountToCopy = Dispatch.get(rowsToCopy, "Count").getInt();
                if (rowsCountToCopy <= 0) {
                    return new BooleanValue(false);
                }

                // Determinar celda de inicio en destino (tu lógica actual)
                Dispatch destStart;
                if ("overwrite".equalsIgnoreCase(copyMode)) {
                    Dispatch usedRangeDest = Dispatch.get(destSheet, "UsedRange").toDispatch();
                    Dispatch.call(usedRangeDest, "Clear");
                    destStart = Dispatch.call(destSheet, "Range", "A1").toDispatch();
                } else if ("append".equalsIgnoreCase(copyMode)) {
                    Dispatch destUsedRange = Dispatch.get(destSheet, "UsedRange").toDispatch();
                    int lastDestRow = Dispatch.get(Dispatch.get(destUsedRange, "Rows").toDispatch(), "Count").getInt();
                    destStart = Dispatch.call(destSheet, "Cells", lastDestRow + 1, 1).toDispatch();
                } else {
                    destStart = Dispatch.call(destSheet, "Range", manualCell).toDispatch();
                }

                // NUEVO: detectar si es la misma instancia de Excel
                boolean sameExcelApp = isSameExcelApp(sourceSession, destSession);

                boolean copiedSomething = false;

                // Copiar ignorando columnas (compactado)
                if ("yes".equalsIgnoreCase(applyIgnoreCol) && ignoreColumns != null && !ignoreColumns.isEmpty()) {
                    List<String> ranges = splitRangeByIgnoredColumns(
                            copyRange,
                            ignoreColumns.stream().map(Object::toString).collect(Collectors.toList())
                    );

                    if (ranges.isEmpty()) {
                        return new BooleanValue(false);
                    }

                    int destColOffset = 0;
                    for (String subRange : ranges) {
                        Dispatch subRangeObj = Dispatch.call(sourceSheet, "Range", subRange).toDispatch();
                        int blockCols = Dispatch.get(Dispatch.get(subRangeObj, "Columns").toDispatch(), "Count").getInt();
                        int blockRows = Dispatch.get(Dispatch.get(subRangeObj, "Rows").toDispatch(), "Count").getInt();

                        int destRow = Dispatch.get(destStart, "Row").getInt();
                        int destCol = Dispatch.get(destStart, "Column").getInt() + destColOffset;
                        Dispatch adjustedDest = Dispatch.call(destSheet, "Cells", destRow, destCol).toDispatch();
                        Dispatch adjustedDestResized = Dispatch.call(adjustedDest, "Resize", blockRows, blockCols).toDispatch();

                        try { Dispatch.call(subRangeObj, "UnMerge"); } catch (Exception ignore) {}

                        if (isFilteredCopy) {
                            Dispatch visibleCells = specialCellsVisibleOrNull(subRangeObj);
                            if (visibleCells == null) {
                                // No hay visibles en este subrango → continuar con el siguiente
                                destColOffset += blockCols;
                                continue;
                            }

                            try {
                                if (sameExcelApp) {
                                    Dispatch.call(visibleCells, "Copy", adjustedDestResized);
                                    copiedSomething = true;
                                } else {
                                    Dispatch.call(visibleCells, "Copy");
                                    sleepQuiet(80);
                                    boolean pasted = pasteWithRetry(destSheet, adjustedDest);
                                    if (pasted) {
                                        copiedSomething = true;
                                    } else {
                                        clearCutCopyMode(sourceSession);
                                        clearCutCopyMode(destSession);
                                        throw new BotCommandException("No se pudo pegar desde portapapeles entre instancias (visibleCells).");
                                    }
                                }
                            } finally {
                                clearCutCopyMode(sourceSession);
                                clearCutCopyMode(destSession);
                            }
                        } else {
                            if (sameExcelApp) {
                                Dispatch.call(subRangeObj, "Copy", adjustedDestResized);
                                copiedSomething = true;
                                clearCutCopyMode(sourceSession);
                                clearCutCopyMode(destSession);
                            } else {
                                Dispatch.call(subRangeObj, "Copy");
                                sleepQuiet(80);
                                boolean pasted = pasteWithRetry(destSheet, adjustedDest);
                                clearCutCopyMode(sourceSession);
                                clearCutCopyMode(destSession);
                                if (pasted) {
                                    copiedSomething = true;
                                } else {
                                    throw new BotCommandException("No se pudo pegar desde portapapeles entre instancias (subRange).");
                                }
                            }
                        }

                        destColOffset += blockCols;
                    }

                    // Si no se copió nada en ninguna porción, devolver False
                    if (!copiedSomething) {
                        return new BooleanValue(false);
                    }

                } else {
                    // Copiar todo sin ignorar columnas
                    Dispatch rangeObj = rangeToCopy;
                    try {
                        Dispatch.call(rangeObj, "UnMerge");
                    } catch (Exception ignore) {
                    }

                    if (isFilteredCopy) {
                        try {
                            Dispatch visibleCells = Dispatch.call(rangeObj, "SpecialCells", new Variant(12)).toDispatch();
                            if (sameExcelApp) {
                                Dispatch.call(visibleCells, "Copy", destStart);
                            } else {
                                Dispatch.call(visibleCells, "Copy");
                                sleepQuiet(80);
                                boolean pasted = pasteWithRetry(destSheet, destStart);
                                if (!pasted) {
                                    clearCutCopyMode(sourceSession);
                                    clearCutCopyMode(destSession);
                                    return new BooleanValue(false); // No se pudo pegar visibles
                                }
                            }
                        } catch (Exception e) {
                            clearCutCopyMode(sourceSession);
                            clearCutCopyMode(destSession);
                            return new BooleanValue(false); // No hay filas visibles
                        } finally {
                            clearCutCopyMode(sourceSession);
                            clearCutCopyMode(destSession);
                        }
                    } else {
                        if (sameExcelApp) {
                            Dispatch.call(rangeObj, "Copy", destStart);
                            clearCutCopyMode(sourceSession);
                            clearCutCopyMode(destSession);
                        } else {
                            Dispatch.call(rangeObj, "Copy");
                            sleepQuiet(80);
                            boolean pasted = pasteWithRetry(destSheet, destStart);
                            clearCutCopyMode(sourceSession);
                            clearCutCopyMode(destSession);
                            if (!pasted) {
                                throw new BotCommandException("No se pudo pegar desde portapapeles entre instancias.");
                            }
                        }
                    }
                }

                // Se copiaron valores
                return new BooleanValue(true);

            } catch (Exception e) {
                // Limpieza de CutCopyMode en caso de error
                clearCutCopyMode(sourceSession);
                clearCutCopyMode(destSession);
                throw new BotCommandException("Error in CopyTable: " + e.getMessage(), e);
            }
        }
    }

        // --- Helpers para manejar copy/paste entre instancias ---
    private static boolean isSameExcelApp(Session s1, Session s2) {
        try {
            int h1 = Dispatch.get(s1.excelApp, "Hwnd").getInt();
            int h2 = Dispatch.get(s2.excelApp, "Hwnd").getInt();
            return h1 == h2;
        } catch (Exception e) {
            return false;
        }
    }

    private static void clearCutCopyMode(Session session) {
        try { Dispatch.put(session.excelApp, "CutCopyMode", false); } catch (Exception ignored) {}
    }

    private static void sleepQuiet(long ms) {
        try { Thread.sleep(ms); } catch (InterruptedException ignored) {}
    }

    /**
     * Intenta pegar desde el portapapeles del sistema en la hoja y celda dadas.
     * Hace algunos reintentos porque Excel puede demorar en disponibilizar el clipboard.
     */
    private static boolean pasteWithRetry(Dispatch destSheet, Dispatch destStart) {
        final int maxAttempts = 6;
        final long[] waits = new long[] { 60, 90, 120, 180, 250, 350 };
        for (int i = 0; i < maxAttempts; i++) {
            try {
                // Preferimos Worksheet.Paste(Destination)
                Dispatch.call(destSheet, "Paste", destStart);
                return true;
            } catch (Exception e1) {
                try {
                    // Fallback: Range.PasteSpecial() sobre la celda inicio
                    Dispatch.call(destStart, "PasteSpecial");
                    return true;
                } catch (Exception e2) {
                    // esperar y reintentar
                    if (i < waits.length) sleepQuiet(waits[i]);
                }
            }
        }
        return false;
    }

    // Devuelve las celdas visibles o null si no hay visibles (sin lanzar excepción).
    private static Dispatch specialCellsVisibleOrNull(Dispatch range) {
        try {
            return Dispatch.call(range, "SpecialCells", new Variant(12)).toDispatch(); // xlCellTypeVisible
        } catch (Exception e) {
            return null;
        }
    }

}

