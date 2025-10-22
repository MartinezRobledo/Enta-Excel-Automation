package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.*;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.util.Map;


@BotCommand
@CommandPkg(
        label = "Open Workbook",
        name = "openWorkbook",
        description = "Opens an Excel workbook and creates/attaches a session",
        icon = "excel.svg",
        return_type = DataType.SESSION,
        return_label = "Workbook Session",
        return_description = "Session variable associated with the opened workbook"
)
public class OpenWorkbook {

    @Execute
    public SessionValue action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "Workbook Path", description = "Full path to the Excel workbook")
            @NotEmpty String workbookPath,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Create as Global Session", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean createGlobal,

            @Idx(index = "3", type = AttributeType.CHECKBOX)
            @Pkg(label = "Attach if session already exists", default_value_type = DataType.BOOLEAN, default_value = "true")
            Boolean attachIfExists,

            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Make Excel Visible", default_value_type = DataType.BOOLEAN, default_value = "true")
            Boolean visible
    ) {
        // 0) Carga idempotente del DLL de JACOB (no fallar si ya está cargado en otro classloader)
        safeEnsureJacobLoaded();

        File file = new File(workbookPath);
        if (!file.exists()) {
            throw new BotCommandException("Workbook file does not exist: " + workbookPath);
        }

        String sessionId   = "WB_" + Integer.toHexString(workbookPath.toLowerCase().hashCode());
        String workbookKey = SessionHelper.toWorkbookKey(workbookPath);

        // 1) Reusar sesión por ID si corresponde
        Session existingById = SessionManager.getSession(sessionId);
        if (existingById != null) {
            if (Boolean.TRUE.equals(attachIfExists)) {
                // Si el libro no está en el mapa, intentá adjuntarte SIN reabrir
                if (!existingById.openWorkbooks.containsKey(workbookKey)) {
                    Dispatch wb = findOpenWorkbook(existingById.excelApp, workbookPath);
                    if (wb == null) {
                        // No está abierto: abrir ahora
                        Dispatch workbooks = existingById.excelApp.getProperty("Workbooks").toDispatch();
                        wb = Dispatch.call(workbooks, "Open", workbookPath).toDispatch();
                    }
                    existingById.openWorkbooks.put(workbookKey, wb);
                }
                return SessionValue.builder()
                        .withSessionObject(new ExcelSession(sessionId, existingById, workbookKey))
                        .build();
            } else {
                throw new BotCommandException("Session already exists for workbook: " + workbookPath);
            }
        }

        try {
            // 2) Buscar instancia Excel compartida ya registrada en SessionManager
            Session shared = null;
            for (Map.Entry<String, Session> e : SessionManager.getSessions().entrySet()) {
                if (e.getValue() != null && e.getValue().excelApp != null) {
                    shared = e.getValue();
                    break;
                }
            }

            // 3) Si no hay sesión en el manager, intentar ATTACH a Excel ya en ejecución
            if (shared == null) {
                // Nota: en ciertos escenarios Excel tarda en registrarse en el ROT hasta perder foco. Se podría reintentar.
                // (Comportamiento documentado por Microsoft) [4](https://support.microsoft.com/en-us/topic/getobject-or-getactiveobject-cannot-find-a-running-office-application-6cdf21a3-ac90-512b-6bff-badc5f4cc215)
                try {
                    ActiveXComponent attached = ActiveXComponent.connectToActiveInstance("Excel.Application"); // [2](https://javadoc.io/static/com.hynnet/jacob/1.18/com/jacob/activeX/ActiveXComponent.html)
                    if (attached != null && attached.getObject() != null) {
                        shared = new Session(attached);
                        // No cambiamos "global" aquí: usamos el flag de creación para nuevas instancias, no para las adjuntas
                        try { attached.setProperty("Visible", Boolean.TRUE.equals(visible)); } catch (Exception ignore) {}
                    }
                } catch (Exception ignore) {
                    // Si falla el attach (p.ej. no hay Excel aún), seguimos con creación
                }
            }

            // 4) Si tampoco pudimos attach, CREAR instancia nueva (Init STA en este hilo)
            if (shared == null) {
                ComThread.InitSTA();
                ActiveXComponent excel = new ActiveXComponent("Excel.Application"); // crea nueva instancia [5](https://github.com/freemansoft/jacob-project/blob/main/samples/com/jacob/samples/office/ExcelDispatchTest.java)
                excel.setProperty("Visible", Boolean.TRUE.equals(visible));
                shared = new Session(excel);
                shared.global = createGlobal;
            } else {
                // Hay Excel en uso: ajustar visibilidad si corresponde
                try { shared.excelApp.setProperty("Visible", Boolean.TRUE.equals(visible)); } catch (Exception ignore) {}
            }

            // A) activar “quiet mode” ANTES de abrir
            QuietState quiet = enableQuietMode(shared.excelApp);

            // 5) Abrir o adjuntar el workbook
            Dispatch wb = findOpenWorkbook(shared.excelApp, workbookPath);
            if (wb == null) {
                Dispatch workbooks = shared.excelApp.getProperty("Workbooks").toDispatch();
                // B) NO actualizar vínculos y SIN diálogos al abrir (UpdateLinks := 0)
                wb = Dispatch.call(workbooks, "Open",
                        new Variant(workbookPath),
                        new Variant(0) /* UpdateLinks := 0 */,
                        new Variant(false) /* ReadOnly := false */
                ).toDispatch(); // [4](https://learn.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open)
            }
            shared.openWorkbooks.put(workbookKey, wb);

            // 6) Registrar el sessionId (per-workbook) apuntando a la misma instancia compartida
            SessionManager.addSession(sessionId, shared);

            return SessionValue.builder()
                    .withSessionObject(new ExcelSession(sessionId, shared, workbookKey))
                    .build();

        } catch (BotCommandException e) {
            throw e;
        } catch (UnsatisfiedLinkError ule) {
            // Mensaje típico: "already loaded in another classloader" -> dar hint claro
            throw new BotCommandException("Failed to open workbook (JACOB DLL conflict). " +
                    "La DLL de JACOB ya fue cargada por otro classloader. " +
                    "Asegurate de no relanzar el loader de JACOB desde otro paquete/loader. Detalle: " + ule.getMessage(), ule);
        } catch (Exception e) {
            throw new BotCommandException("Failed to open workbook: " + e.getMessage(), e);
        }
    }

    /**
     * Carga segura/idempotente de la DLL de JACOB:
     * - si la DLL ya fue cargada por OTRO classloader, ignoramos el error y continuamos,
     *   tal como sugiere la propia doc ("solo se carga una vez por classloader"). [1](https://learn.microsoft.com/en-us/office/vba/api/Excel.Range.TextToColumns)
     */
    private static void safeEnsureJacobLoaded() {
        try {
            JacobBootstrap.ensureLoaded();
        } catch (UnsatisfiedLinkError ule) {
            String msg = String.valueOf(ule.getMessage());
            if (msg != null && msg.toLowerCase().contains("already loaded in another classloader")) {
                // OK: la DLL ya está en el proceso; continuamos sin fallar.
            } else {
                throw ule;
            }
        }
    }

    /**
     * Devuelve el Dispatch del workbook ya abierto cuyo FullName coincide con workbookPath,
     * o null si no está abierto en la instancia.
     */
    private static Dispatch findOpenWorkbook(ActiveXComponent excelApp, String workbookPath) {
        try {
            Dispatch workbooks = excelApp.getProperty("Workbooks").toDispatch();
            int count = Dispatch.get(workbooks, "Count").getInt();
            for (int i = 1; i <= count; i++) {
                Dispatch wb = Dispatch.call(workbooks, "Item", i).toDispatch();
                String fullName = Dispatch.get(wb, "FullName").getString();
                if (fullName != null && fullName.equalsIgnoreCase(workbookPath)) {
                    return wb;
                }
            }
        } catch (Exception ignore) {}
        return null;
    }

    // --- QUIET MODE helpers ---
    private static class QuietState {
        Integer automationSecurityPrev;
    }
    private static QuietState enableQuietMode(ActiveXComponent excelApp) {
        QuietState st = new QuietState();
        try {
            // 1) Suprimir pop-ups genéricos (borrar hoja, sobrescribir, etc.)
            excelApp.setProperty("DisplayAlerts", false); // [1](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts)
            // 2) No preguntar por actualización de vínculos (la acción de abrir decidirá si actualiza)
            excelApp.setProperty("AskToUpdateLinks", false); // [7](https://learn.microsoft.com/en-us/office/vba/api/excel.application.asktoupdatelinks)
            // 3) Deshabilitar macros en archivos abiertos programáticamente (sin avisos)
            st.automationSecurityPrev = excelApp.getProperty("AutomationSecurity").getInt();
            excelApp.setProperty("AutomationSecurity", 3); // msoAutomationSecurityForceDisable = 3 [9](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)
            // 4) (Opcional) Evitar disparo de eventos mientras abrimos
            try { excelApp.setProperty("EnableEvents", false); } catch (Exception ignore) {}
        } catch (Exception ignore) {}
        return st;
    }
    private static void restoreQuietMode(ActiveXComponent excelApp, QuietState st) {
        try {
            if (st != null && st.automationSecurityPrev != null) {
                excelApp.setProperty("AutomationSecurity", st.automationSecurityPrev); // [9](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)
            }
            // Recuperar eventos; DisplayAlerts podés dejarlo en false durante la sesión si tus bots lo requieren
            try { excelApp.setProperty("EnableEvents", true); } catch (Exception ignore) {}
        } catch (Exception ignore) {}
    }
}