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
        label = "Create Workbook",
        name = "createWorkbook",
        description = "Crea un nuevo libro de Excel en el path indicado, asigna el nombre de la hoja por defecto y lo abre en una sesión.",
        icon = "excel.svg",
        return_type = DataType.SESSION,
        return_label = "Workbook Session",
        return_description = "Session variable associated with the newly created workbook"
)
public class CreateWorkbook {

    @Execute
    public SessionValue action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "Workbook Path (nuevo archivo)", description = "Ruta completa del archivo a crear (ej.: C:\\\\bots\\\\output.xlsx)")
            @NotEmpty String workbookPath,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Nombre de la hoja por defecto", description = "Ej.: Datos")
            @NotEmpty String defaultSheetName,


            @Idx(index = "3", type = AttributeType.CHECKBOX)
            @Pkg(label = "Overwrite if exists", default_value_type = DataType.BOOLEAN, default_value = "true")
            Boolean overwriteIfExists
    ) {
        // 0) Carga idempotente del DLL de JACOB (mismo patrón que OpenWorkbook)
        safeEnsureJacobLoaded(); // reutiliza la misma idea de OpenWorkbook [1](https://entaconsulting-my.sharepoint.com/personal/amartinez_entaconsulting_com/Documents/Archivos%20de%20chat%20de%20Microsoft%C2%A0Copilot/OpenWorkbook.java)

        // 1) Validaciones de ruta y carpeta
        if (workbookPath == null || workbookPath.trim().isEmpty()) {
            throw new BotCommandException("Debe indicar el path del archivo a crear.");
        }
        workbookPath = workbookPath.trim();

        File f = new File(workbookPath);

        if (f.exists()) {
            if (Boolean.TRUE.equals(overwriteIfExists)) {
                if (!tryDeleteOrBackup(f)) {
                    throw new BotCommandException("El archivo ya existe y está en uso o no pudo sobrescribirse: "
                            + workbookPath + ". Cerrá el archivo o quitá protección y reintentá.");
                }
            } else {
                throw new BotCommandException("El archivo ya existe: " + workbookPath
                        + ". Activá 'Overwrite if exists' para sobrescribirlo.");
            }
        }

        File parent = f.getAbsoluteFile().getParentFile();
        if (parent != null && !parent.exists()) {
            if (!parent.mkdirs()) {
                throw new BotCommandException("No se pudo crear la carpeta destino: " + parent);
            }
        }

        // 2) Preparar/obtener instancia de Excel (siguiendo estrategia de OpenWorkbook)
        Session shared = findOrCreateExcelSession(true /*visible por defecto*/); // patrón semejante a OpenWorkbook [1](https://entaconsulting-my.sharepoint.com/personal/amartinez_entaconsulting_com/Documents/Archivos%20de%20chat%20de%20Microsoft%C2%A0Copilot/OpenWorkbook.java)

        // 3) Quiet mode básico
        QuietState quiet = enableQuietMode(shared.excelApp);

        try {
            // 4) Crear libro nuevo
            Dispatch workbooks = shared.excelApp.getProperty("Workbooks").toDispatch();
            Dispatch wb = Dispatch.call(workbooks, "Add").toDispatch(); // Excel hace el trabajo de base

            // 5) Renombrar primera hoja y (opcional) dejar solo 1 hoja
            Dispatch sheets = Dispatch.get(wb, "Worksheets").toDispatch();
            int count = Dispatch.get(sheets, "Count").getInt();

            // Primera hoja
            Dispatch first = Dispatch.call(sheets, "Item", 1).toDispatch();
            try {
                Dispatch.put(first, "Name", defaultSheetName);
            } catch (Exception e) {
                throw new BotCommandException("Nombre de hoja inválido o en uso: '" + defaultSheetName + "'. Detalle: " + safeMsg(e), e);
            }

            // Eliminar hojas restantes (si las hubiera), para garantizar un único sheet con el nombre solicitado
            if (count > 1) {
                for (int i = count; i >= 2; i--) {
                    try {
                        Dispatch toDelete = Dispatch.call(sheets, "Item", i).toDispatch();
                        Dispatch.call(toDelete, "Delete");
                    } catch (Exception delEx) {
                        // Si no puede borrar por configuración, dejamos las demás hojas; no es crítico para la creación
                        // pero avisamos con un mensaje claro
                        throw new BotCommandException("No se pudo eliminar hoja extra #" + i + ". Detalle: " + safeMsg(delEx), delEx);
                    }
                }
            }

            // 6) Guardar como el path indicado (deja que Excel detecte formato por extensión)
            try {

                String normalizedPath = toWinCanonicalPath(workbookPath);

                // chequeo de permisos de carpeta
                parent = new File(normalizedPath).getAbsoluteFile().getParentFile();
                if (parent == null || !parent.exists() || !parent.canWrite()) {
                    throw new BotCommandException("La carpeta destino no es escribible: " + parent);
                }

                // Pre-chequeo: ¿ya hay un libro abierto con el mismo NOMBRE (sin ruta)?
                workbooks = shared.excelApp.getProperty("Workbooks").toDispatch();
                int openCnt = Dispatch.get(workbooks, "Count").getInt();
                String targetNameOnly = new File(normalizedPath).getName(); // p.ej. "Diferencias contabilidad estadistica.xlsx"
                for (int i = 1; i <= openCnt; i++) {
                    Dispatch wbOpen = Dispatch.call(workbooks, "Item", i).toDispatch();
                    String openName = Dispatch.get(wbOpen, "Name").getString();
                    if (openName != null && openName.equalsIgnoreCase(targetNameOnly)) {
                        throw new BotCommandException(
                                "Ya hay un libro abierto llamado '" + targetNameOnly + "'. " +
                                        "Excel no permite dos libros con el mismo nombre aunque estén en rutas distintas. " +
                                        "Cerralo o usá otro nombre e intentá nuevamente."
                        );
                    }
                }

                // Guardar como XLSX, especificando FileFormat = 51 (xlOpenXMLWorkbook)  ← clave
                // Docs SaveAs + lista de formatos: MS Learn
                Dispatch.call(wb, "SaveAs",
                        new Variant(normalizedPath),
                        new Variant(51)); // xlOpenXMLWorkbook = 51

            } catch (Exception saveEx) {
                throw new BotCommandException("No se pudo guardar el archivo en: " + workbookPath + ". Detalle: " + safeMsg(saveEx), saveEx);
            }

            // 7) Registrar en SessionManager y devolver la sesión (idéntico a OpenWorkbook)
            String sessionId = "WB_" + Integer.toHexString(workbookPath.toLowerCase().hashCode());
            String workbookKey = SessionHelper.toWorkbookKey(workbookPath);

            // Asegurar que el workbook quede mapeado
            shared.openWorkbooks.put(workbookKey, wb);

            // Registrar el sessionId per-workbook apuntando a la misma instancia compartida
            SessionManager.addSession(sessionId, shared);

            return SessionValue.builder()
                    .withSessionObject(new ExcelSession(sessionId, shared, workbookKey))
                    .build();

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException("Error creando el workbook: " + e.getMessage(), e);
        } finally {
            restoreQuietMode(shared.excelApp, quiet);
        }
    }

    // ===================== Helpers alineados a OpenWorkbook =====================

    /** Intenta reusar instancia Excel registrada, luego attach a instancia activa y finalmente crear una nueva. */
    private static Session findOrCreateExcelSession(boolean visible) {
        // 1) Buscar instancia Excel compartida ya registrada en SessionManager
        Session shared = null;
        for (Map.Entry<String, Session> e : SessionManager.getSessions().entrySet()) {
            if (e.getValue() != null && e.getValue().excelApp != null) {
                shared = e.getValue();
                break;
            }
        }

        // 2) Si no hay, intentar attach a Excel en ejecución
        if (shared == null) {
            try {
                ActiveXComponent attached = ActiveXComponent.connectToActiveInstance("Excel.Application");
                if (attached != null && attached.getObject() != null) {
                    shared = new Session(attached);
                    try { attached.setProperty("Visible", visible); } catch (Exception ignore) {}
                }
            } catch (Exception ignore) {
                // No hay Excel aún: seguimos con creación
            }
        }

        // 3) Crear instancia nueva si aún no hay
        if (shared == null) {
            ComThread.InitSTA();
            ActiveXComponent excel = new ActiveXComponent("Excel.Application");
            excel.setProperty("Visible", visible);
            shared = new Session(excel);
            shared.global = false; // por defecto
        } else {
            try { shared.excelApp.setProperty("Visible", visible); } catch (Exception ignore) {}
        }

        return shared;
    }

    /** Quiet mode similar al de OpenWorkbook: menos prompts al crear/guardar/eliminar. */
    private static class QuietState {
        Integer automationSecurityPrev;
    }
    private static QuietState enableQuietMode(ActiveXComponent excelApp) {
        QuietState st = new QuietState();
        try {
            excelApp.setProperty("DisplayAlerts", false);
            excelApp.setProperty("AskToUpdateLinks", false);
            st.automationSecurityPrev = excelApp.getProperty("AutomationSecurity").getInt();
            excelApp.setProperty("AutomationSecurity", 3); // msoAutomationSecurityForceDisable
            try { excelApp.setProperty("EnableEvents", false); } catch (Exception ignore) {}
        } catch (Exception ignore) {}
        return st;
    }
    private static void restoreQuietMode(ActiveXComponent excelApp, QuietState st) {
        try {
            if (st != null && st.automationSecurityPrev != null) {
                excelApp.setProperty("AutomationSecurity", st.automationSecurityPrev);
            }
            try { excelApp.setProperty("EnableEvents", true); } catch (Exception ignore) {}
        } catch (Exception ignore) {}
    }

    /** Carga segura/idempotente de la DLL de JACOB, mismo enfoque que OpenWorkbook. */
    private static void safeEnsureJacobLoaded() {
        try {
            JacobBootstrap.ensureLoaded(); // mismo patrón que tu OpenWorkbook
        } catch (UnsatisfiedLinkError ule) {
            String msg = String.valueOf(ule.getMessage());
            if (msg != null && msg.toLowerCase().contains("already loaded in another classloader")) {
                // OK: la DLL ya está en el proceso; continuamos sin fallar.
            } else {
                throw ule;
            }
        }
    }

    private static String safeMsg(Exception e) {
        String m = (e == null ? "" : e.getMessage());
        return (m == null ? "" : m);
    }

    private static String toWinCanonicalPath(String p) {
        try {
            // Reemplaza '/' por '\' y colapsa dobles
            String s = p.replace('/', '\\').replaceAll("\\\\{2,}", "\\\\");
            return new File(s).getCanonicalPath(); // normaliza .., .
        } catch (Exception e) {
            return p.replace('/', '\\');
        }
    }


    /** Intenta eliminar el archivo; si falla, intenta renombrarlo a .bak_timestamp. */
    private static boolean tryDeleteOrBackup(File f) {
        try { f.setWritable(true); } catch (Exception ignore) {}
        if (f.delete()) return true;
        File bak = new File(f.getAbsolutePath() + ".bak_" + System.currentTimeMillis());
        return f.renameTo(bak);
    }

}
