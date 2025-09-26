package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ComScope;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionHelper;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Get Headers Range",
        name = "getHeadersRange",
        description = "Finds the row where a reference header exists and returns the contiguous headers range in that row. Returns empty string if not found.",
        icon = "excel.svg"
)
public class GetHeadersRange {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
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
            @NotEmpty
            String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index", description = "1-based index")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double sheetIndex,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Reference Header", description = "Exact header text to find")
            @NotEmpty
            String referenceHeader,

            @Idx(index = "4", type = AttributeType.VARIABLE)
            @Pkg(label = "Select variable")
            @NotEmpty
            Value<String> varOutput,

            @Idx(index = "5", type = AttributeType.CHECKBOX)
            @Pkg(label = "Allow discontinuous headers",
                    description = "If checked, the range will continue even if some header cells are empty",
                    default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean allowDiscontinuousHeaders
    ) {
            // 1) Sesión + workbook correctos
            Session session = ExcelObjects.requireSession(excelSession);
            Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

            // 2) Hoja a trabajar (usa helper que valida nombre/índice)
            Dispatch sheet = SessionHelper.getSheet(wb, selectSheetBy, sheetName, sheetIndex == null ? null : sheetIndex.intValue());
            try { Dispatch.call(sheet, "Activate"); } catch (Exception ignore) {}

            // 3) UsedRange y límites
            Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
            int usedFirstRow = Dispatch.get(usedRange, "Row").getInt();
            int usedFirstCol = Dispatch.get(usedRange, "Column").getInt();
            int usedCols = Dispatch.get(Dispatch.get(usedRange, "Columns").toDispatch(), "Count").getInt();
            int usedLastCol = usedFirstCol + usedCols - 1;

            // 4) Buscar el header: usamos Find y, para asegurar coincidencia exacta,
            //    validamos el valor y, si es necesario, iteramos con FindNext hasta volver al primero.
            Dispatch first = null;
            Dispatch current = null;
            try {
                current = Dispatch.call(usedRange, "Find", referenceHeader).toDispatch();
                if (current != null && current.m_pDispatch != 0) {
                    first = current;
                }
            } catch (Exception ignore) {
                current = null;
            }

            if (current == null || current.m_pDispatch == 0) {
                // no encontrado
                if (varOutput != null) varOutput.set("");
                return new StringValue("");
            }

            final String target = referenceHeader == null ? "" : referenceHeader.trim();
            String firstAddress = Dispatch.get(first, "Address").toString();

            boolean foundExact = false;
            Dispatch exactCell = current;

            while (true) {
                Variant val = Dispatch.get(current, "Value");
                String cellText = (val == null || val.isNull()) ? "" : val.toString().trim();
                if (cellText.equals(target)) {
                    foundExact = true;
                    exactCell = current;
                    break;
                }
                // avanzar al siguiente match
                Dispatch next = null;
                try { next = Dispatch.call(usedRange, "FindNext", current).toDispatch(); }
                catch (Exception ignore) { next = null; }

                if (next == null || next.m_pDispatch == 0) break;

                String nextAddr = Dispatch.get(next, "Address").toString();
                if (nextAddr.equals(firstAddress)) {
                    // volvimos al primero: cortar
                    break;
                }
                current = next;
            }

            if (!foundExact) {
                if (varOutput != null) varOutput.set("");
                return new StringValue("");
            }

            int headerRow = Dispatch.get(exactCell, "Row").getInt();
            int foundCol = Dispatch.get(exactCell, "Column").getInt();

            // 5) Expandir hacia izquierda/derecha segun allowDiscontinuousHeaders
            int startCol = foundCol;
            while (startCol - 1 >= usedFirstCol) {
                Dispatch leftCell = Dispatch.call(sheet, "Cells", headerRow, startCol - 1).toDispatch();
                Variant v = Dispatch.get(leftCell, "Value");
                boolean empty = (v == null || v.isNull() || v.toString().trim().isEmpty());
                if (!allowDiscontinuousHeaders && empty) break;
                // Si discontinuo está permitido, seguimos incluso si vacío
                startCol--;
            }

            int endCol = foundCol;
            while (endCol + 1 <= usedLastCol) {
                Dispatch rightCell = Dispatch.call(sheet, "Cells", headerRow, endCol + 1).toDispatch();
                Variant v = Dispatch.get(rightCell, "Value");
                boolean empty = (v == null || v.isNull() || v.toString().trim().isEmpty());
                if (!allowDiscontinuousHeaders && empty) break;
                endCol++;
            }

            // Ajustes de borde: recortar vacíos en extremos si quedaron
            while (startCol <= endCol) {
                Dispatch firstCell = Dispatch.call(sheet, "Cells", headerRow, startCol).toDispatch();
                Variant v = Dispatch.get(firstCell, "Value");
                if (v != null && !v.isNull() && !v.toString().trim().isEmpty()) break;
                startCol++;
            }

            while (endCol >= startCol) {
                Dispatch lastCell = Dispatch.call(sheet, "Cells", headerRow, endCol).toDispatch();
                Variant v = Dispatch.get(lastCell, "Value");
                if (v != null && !v.isNull() && !v.toString().trim().isEmpty()) break;
                endCol--;
            }

            if (endCol < startCol) {
                if (varOutput != null) varOutput.set("");
                return new StringValue("");
            }

            // 6) Construir rango final tipo "A10:F10"
            String startColLetter = getColumnLetter(startCol);
            String endColLetter = getColumnLetter(endCol);
            String range = startColLetter + headerRow + ":" + endColLetter + headerRow;

            if (varOutput != null) varOutput.set(range);
            return new StringValue(range);
    }

    private String getColumnLetter(int columnNumber) {
        StringBuilder sb = new StringBuilder();
        while (columnNumber > 0) {
            int rem = (columnNumber - 1) % 26;
            sb.insert(0, (char) ('A' + rem));
            columnNumber = (columnNumber - 1) / 26;
        }
        return sb.toString();
    }
}
