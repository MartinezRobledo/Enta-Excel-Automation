package com.automationanywhere.botcommand.utilities;

import java.util.concurrent.ConcurrentHashMap;
import java.util.*;
import java.util.stream.Collectors;

public class SessionManager {
    // Múltiples sessionId pueden apuntar a la MISMA instancia de Session (una sola Excel.Application)
    private static final Map<String, Session> sessions = new ConcurrentHashMap<>();

    public static void addSession(String sessionId, Session session) {
        sessions.put(sessionId, session);
    }

    public static Session getSession(String sessionId) {
        return sessions.get(sessionId);
    }

    public static Map<String, Session> getSessions() {
        return sessions;
    }

    /** Devuelve las instancias de Session (deduplicadas por identidad) */
    public static Set<Session> getDistinctSessions() {
        return new HashSet<>(sessions.values());
    }

    /** Remueve solo el id del mapa (no hace Quit). */
    public static void removeSessionIdOnly(String sessionId) {
        sessions.remove(sessionId);
    }

    /** Remueve TODAS las entradas (sessionId -> session) que apuntan a la MISMA instancia. */
    public static void removeAllByInstance(Session s) {
        if (s == null) return;
        List<String> toRemove = sessions.entrySet().stream()
                .filter(e -> e.getValue() == s)
                .map(Map.Entry::getKey)
                .collect(Collectors.toList());
        for (String key : toRemove) {
            sessions.remove(key);
        }
    }

    /** Devuelve true si aún hay sessionIds apuntando a la misma instancia. */
    public static boolean hasRefs(Session s) {
        return sessions.values().stream().anyMatch(v -> v == s);
    }
}