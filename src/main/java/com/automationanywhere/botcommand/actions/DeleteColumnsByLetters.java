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
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.*;
import java.util.stream.Collectors;

@BotCommand
@CommandPkg(
        label = "Delete Columns by Letters",
        name = "deleteColumnsByLetters",
        description = "Deletes whole columns by their letters (e.g., A, C, F).",
        icon = "excel.svg",
        return_type = DataType.BOOLEAN,
        return_required = true
)
public class DeleteColumnsByLetters {

    @Execute
    public Value<Boolean> action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session") @SessionObject @NotEmpty ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name") String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)") Double sheetIndex,

            @Idx(index = "3", type = AttributeType.LIST)
            @Pkg(label = "Columns to delete (letters; e.g., A,B,C)") @NotEmpty List<Object> letters
    ) {
        try {
            return run(excelSession, selectSheetBy, sheetName, sheetIndex, letters);
        } catch (Exception first) {
            try {
                ComThread.InitSTA();
                return run(excelSession, selectSheetBy, sheetName, sheetIndex, letters);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("Failed to delete columns: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private Value<Boolean> run(ExcelSession excelSession, String selectBy, String sheetName, Double sheetIndex, List<Object> letters) {
        if (letters == null || letters.isEmpty()) return new BooleanValue(false);

        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet = ExcelObjects.requireSheet(wb, selectBy, sheetName, sheetIndex);
        Dispatch app = Dispatch.get(wb, "Application").toDispatch();

        // Guardar estado y optimizar
        boolean prevAlerts = true, prevEvents = true, prevUpd = true;
        int prevCalc = -4105; // auto
        try {
            prevAlerts = Dispatch.get(app, "DisplayAlerts").getBoolean();
            prevEvents = Dispatch.get(app, "EnableEvents").getBoolean();
            prevUpd    = Dispatch.get(app, "ScreenUpdating").getBoolean();
            prevCalc   = Dispatch.get(app, "Calculation").getInt();

            Dispatch.put(app, "DisplayAlerts", false);
            Dispatch.put(app, "EnableEvents", false);
            Dispatch.put(app, "ScreenUpdating", false);
            Dispatch.put(app, "Calculation", new Variant(-4135)); // manual

            // Ordenar desc por ancho (Z..A) para evitar corrimientos
            List<String> cols = letters.stream()
                    .map(o -> o == null ? "" : o.toString().trim().toUpperCase())
                    .filter(s -> !s.isEmpty())
                    .sorted(Comparator.<String>naturalOrder().reversed())
                    .collect(Collectors.toList());

            for (String col : cols) {
                Dispatch wholeCol = Dispatch.call(sheet, "Range", col + ":" + col).toDispatch();
                try { Dispatch.call(wholeCol, "Delete"); } catch (Exception ignore) {}
            }

            return new BooleanValue(true);
        } finally {
            try { Dispatch.put(app, "Calculation", new Variant(prevCalc)); } catch (Exception ignore) {}
            try { Dispatch.put(app, "ScreenUpdating", prevUpd); } catch (Exception ignore) {}
            try { Dispatch.put(app, "EnableEvents", prevEvents); } catch (Exception ignore) {}
            try { Dispatch.put(app, "DisplayAlerts", prevAlerts); } catch (Exception ignore) {}
        }
    }
}